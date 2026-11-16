#Requires -Version 5.1
#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Installiert Intel-Ethernet-Treiber von der Intel-Driver-CD mit pnputil.
    Welche Treiberordner verwendet werden, steht in der SetupBD.ini (je nach Betriebssystem).

.DESCRIPTION
    Ablauf:
      0. CD-Pfad = <Ordner des Scripts>\<CdFolderName> (Laufwerk und Pfad des Scripts beliebig).
      1. SetupBD.ini einlesen (mehrfache Schlüssel wie InfDir/InfExclude werden als Liste gelesen).
      2. Betriebssystem erkennen (Build, Client/Server, Azure Stack HCI/Azure Local) und
         den passenden INI-Abschnitt wählen (z. B. [Win2025Server]).
      3. CD-Root über [Config] LevelsToRoot bestimmen (relativ zum Ordner von SetupBD.exe).
      4. Alle *.inf aus den InfDir-Ordnern sammeln, InfExclude-Dateien auslassen.
      5. Optional (-OnlyMatchingHardware): nur INFs, die eine im System vorhandene Intel-PCI-Hardware-ID
         in einem der InfSectionName-Abschnitte enthalten.
      6. Vorher den Treiberspeicher sichern (CSV), dann pnputil /add-driver <inf> /install pro INF.
      7. Neue oemXX.inf-Pakete ermitteln und ein Rollback-Script schreiben.

.PARAMETER CdFolderName
    Name des CD-Ordners, der im Ordner des Scripts liegt. Standardwert wird oben im param-Block festgelegt.
    Erwartete Ablage:
      <Script-Ordner>\<CdFolderName>\<Version>\APPS\SETUP\SETUPBD\Winx64\Resource\SetupBD.ini
    oder (ohne Versionsordner):
      <Script-Ordner>\<CdFolderName>\APPS\SETUP\SETUPBD\Winx64\Resource\SetupBD.ini

.PARAMETER IniPath
    Optional. Direkter Pfad zur SetupBD.ini; übersteuert die Suche über CdFolderName.

.PARAMETER Section
    INI-Abschnitt erzwingen statt automatischer OS-Erkennung (z. B. Win2022Server).

.PARAMETER OnlyMatchingHardware
    Nur INF-Dateien installieren, die zu vorhandener Intel-Hardware passen.

.PARAMETER NoInstall
    Treiber nur in den Treiberspeicher aufnehmen (pnputil /add-driver ohne /install).

.PARAMETER WhatIf
    Testlauf: zeigt erkannten Abschnitt und die INF-Dateien, ruft pnputil aber nicht auf.

.PARAMETER Verbose
    Zusätzliche Detailausgaben (z. B. welche INF wegen InfExclude oder fehlender Hardware entfällt).

.PARAMETER LogDir
    Ordner für Transcript, Ergebnis-CSV, Treiber-Sicherung und Rollback-Script.

.EXAMPLE
    .\Install-IntelNicDriver.ps1 -WhatIf
    Testlauf: zeigt erkannten Abschnitt und die INF-Dateien, ändert nichts.

.EXAMPLE
    .\Install-IntelNicDriver.ps1 -OnlyMatchingHardware

.EXAMPLE
    .\Install-IntelNicDriver.ps1 -IniPath 'D:\31.1\APPS\SETUP\SETUPBD\Winx64\Resource\SetupBD.ini' -Section Win2025Server

.NOTES
    Exit-Codes: 0 = OK, 3010 = OK, Neustart erforderlich, 1 = mindestens ein Fehler.
    ACHTUNG: Beim Installieren/Aktualisieren eines NIC-Treibers geht der Link kurz weg.
    Bei Remote-Sitzungen (RDP/PS-Remoting) über den betroffenen Adapter einplanen!
#>

# Bewusst OHNE [CmdletBinding()]: -WhatIf und -Verbose sind unten als normale Schalter definiert.
# So läuft das Script in jeder Umgebung (auch ISE/F8 oder in die Konsole eingefügt).
param(
    # ==========================================================================
    #  NAME DES CD-ORDNERS  (liegt im selben Ordner wie dieses Script)
    #  Bei einer neuen CD-Version nur diesen Wert anpassen.
    #  Hinweis: Wegen des '®' muss das Script als UTF-8 MIT BOM gespeichert bleiben.
    # ==========================================================================
    [string]$CdFolderName = 'Intel®EthernetAdaptersConnectionsCD_311_1316650',

    [string]$IniPath,                 # optional: anderer Pfad zur SetupBD.ini (Standard: CD im Script-Ordner)
    [string]$Section,                 # optional: INI-Abschnitt erzwingen (überschreibt OS-Erkennung)
    [switch]$OnlyMatchingHardware,    # nur Treiber für tatsächlich vorhandene Intel-Hardware
    [switch]$NoInstall,               # nur in den Treiberspeicher aufnehmen, nicht an Geräte binden
    [string]$LogDir = (Join-Path $env:SystemRoot 'Temp\IntelNicDriver'),  # Ablage für Logs/CSV/Rollback
    [switch]$WhatIf,                  # Testlauf: nichts installieren, nur anzeigen
    [switch]$Verbose                  # zusätzliche Detailausgaben
)

# -Verbose: Write-Verbose-Meldungen sichtbar machen
if ($Verbose) { $VerbosePreference = 'Continue' }

# Strenger Modus: Zugriff auf nicht gesetzte Variablen/Eigenschaften erzeugt einen Fehler
Set-StrictMode -Version 2.0
# Jeder Fehler eines Cmdlets wird zum Abbruch -> landet im catch-Block unten
$ErrorActionPreference = 'Stop'

# Ordner, in dem das Script liegt (= CD-Pfad). Muss auf Script-Ebene ermittelt werden.
#  - $PSScriptRoot: Normalfall (Start als .ps1-Datei, auch mit -File)
#  - $MyInvocation.MyCommand.Path: Fallback für ältere Hosts
#  - aktueller Ordner: letzter Fallback (z. B. Code markiert und mit F8 in der ISE ausgeführt)
# Pfad der Script-Datei; StrictMode-sicher gelesen (beim Einfügen in die Konsole gibt es keine Path-Eigenschaft)
$cmdPathProp = $MyInvocation.MyCommand.PSObject.Properties['Path']
$ScriptFile  = if ($cmdPathProp -and $cmdPathProp.Value) { [string]$cmdPathProp.Value } else { '' }
$RunAsFile   = [bool]$ScriptFile                     # True = als .ps1-Datei gestartet
$ScriptDir = $PSScriptRoot
if (-not $ScriptDir -and $ScriptFile) { $ScriptDir = Split-Path -Parent $ScriptFile }
if (-not $ScriptDir) { $ScriptDir = (Get-Location).ProviderPath }

#region Funktionen ---------------------------------------------------------------

function Read-MultiValueIni {
    <#
      Liest eine INI-Datei, in der ein Schlüssel MEHRFACH vorkommen darf (z. B. mehrere InfDir=).
      Normale INI-Parser würden nur den letzten Wert behalten, deshalb ein eigener Parser.
      Rückgabe: [ordered]@{ 'Abschnitt' = @{ 'Schlüssel' = List[string] } }
      Beispiel: $ini['Win2025Server']['InfDir'] -> PRO1000\Winx64\WS2025, PROXGB\Winx64\WS2025, ...
    #>
    param([Parameter(Mandatory)][string]$Path)

    $data    = [ordered]@{}   # Ergebnis; [ordered] behält die Reihenfolge der Abschnitte (Groß/Klein egal)
    $current = $null          # Name des Abschnitts, in dem wir gerade lesen

    # Datei Zeile für Zeile durchgehen
    foreach ($raw in Get-Content -LiteralPath $Path) {
        $line = $raw.Trim()                                        # Leerzeichen vorne/hinten entfernen

        # Leerzeilen und Kommentarzeilen (beginnen mit ';') überspringen
        if (-not $line -or $line.StartsWith(';')) { continue }

        # Abschnittskopf erkennen: Zeile hat die Form [Name]
        #   ^\[  = Zeile beginnt mit '['
        #   (.+) = beliebiger Inhalt -> landet in $Matches[1]
        #   \]$  = Zeile endet mit ']'
        if ($line -match '^\[(.+)\]$') {
            $current = $Matches[1].Trim()                          # neuer aktueller Abschnitt
            $data[$current] = @{}                                  # leere Schlüssel-Tabelle anlegen
            continue                                               # nächste Zeile
        }

        # Schlüssel=Wert-Zeile erkennen (nur wenn wir uns in einem Abschnitt befinden)
        #   ^([^=]+?) = alles bis zum ersten '=' (nicht gierig) -> Schlüssel in $Matches[1]
        #   \s*=\s*   = Gleichheitszeichen mit optionalen Leerzeichen
        #   (.*)$     = Rest der Zeile -> Wert in $Matches[2]
        if ($current -and $line -match '^([^=]+?)\s*=\s*(.*)$') {
            $key = $Matches[1].Trim()
            $val = $Matches[2].Trim()

            # Beim ersten Auftreten des Schlüssels eine leere Liste anlegen
            if (-not $data[$current].ContainsKey($key)) {
                $data[$current][$key] = New-Object System.Collections.Generic.List[string]
            }
            # Wert an die Liste anhängen (so bleiben mehrfache InfDir=/InfExclude= alle erhalten)
            $data[$current][$key].Add($val)
        }
    }
    return $data
}

function Find-SetupBDIni {
    <#
      $BaseDir ist der CD-Ordner (<Script-Ordner>\<CdFolderName>).
      Geprüft wird in dieser Reihenfolge:
        1. <BaseDir>\APPS\SETUP\SETUPBD\Winx64\Resource\SetupBD.ini
           (CD-Ordner enthält direkt APPS\, PRO1000\ ...)
        2. <BaseDir>\<Version>\APPS\SETUP\SETUPBD\Winx64\Resource\SetupBD.ini
           (CD-Ordner enthält einen Versionsordner, z. B. 31.1\ - so liefert Intel die CD)
    #>
    param([Parameter(Mandatory)][string]$BaseDir)

    # Relativer Pfad der INI innerhalb eines Versionsordners (fest durch Intel vorgegeben)
    $relIni = 'APPS\SETUP\SETUPBD\Winx64\Resource\SetupBD.ini'

    # Fall 1: INI direkt unterhalb des Script-Ordners
    $candidate = Join-Path $BaseDir $relIni
    if (Test-Path -LiteralPath $candidate) { return $candidate }

    # Fall 2: jeden direkten Unterordner (Versionsordner) prüfen
    $hits = @()
    foreach ($sub in Get-ChildItem -LiteralPath $BaseDir -Directory -ErrorAction SilentlyContinue) {
        $candidate = Join-Path $sub.FullName $relIni
        if (Test-Path -LiteralPath $candidate) {
            # Treffer mit Ordnername merken, um bei mehreren Versionen die höchste zu wählen
            $hits += [pscustomobject]@{ Name = $sub.Name; Path = $candidate }
        }
    }

    # Keine CD gefunden
    if ($hits.Count -eq 0) { return $null }

    # Bei mehreren Versionsordnern (z. B. 30.6 und 31.1) den höchsten nehmen.
    # Ordnernamen werden als Versionsnummer verglichen; nicht lesbare Namen landen hinten.
    $best = $hits | Sort-Object -Descending -Property @{ Expression = {
        $v = $null
        if ([version]::TryParse($_.Name, [ref]$v)) { $v } else { [version]'0.0' }
    } } | Select-Object -First 1

    if ($hits.Count -gt 1) {
        $names = ($hits | Select-Object -ExpandProperty Name) -join ', '      # z. B. "30.6, 31.1"
        Write-Warning "Mehrere Versionsordner gefunden ($names), verwende: $($best.Name)"
    }
    return $best.Path
}

function Get-OsInfo {
    <#
      Sammelt die Infos, die für die Wahl des INI-Abschnitts nötig sind.
    #>
    $os = Get-CimInstance -ClassName Win32_OperatingSystem                          # Caption, Build, ProductType
    $cv = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'  # ProductName usw.

    # Werte nur lesen, wenn sie in der Registry existieren (StrictMode würde sonst Fehler werfen)
    $productName      = if ($cv.PSObject.Properties['ProductName'])      { $cv.ProductName }      else { '' }
    $installationType = if ($cv.PSObject.Properties['InstallationType']) { $cv.InstallationType } else { '' }

    # Ergebnis als Objekt zurückgeben
    [pscustomobject]@{
        Caption          = $os.Caption
        ProductName      = $productName
        InstallationType = $installationType
        Build            = [int]$os.BuildNumber                    # z. B. 26100 = Server 2025
        IsServer         = ($os.ProductType -ne 1)                 # 1 = Workstation, 2 = DC, 3 = Server
        IsHci            = ("$($os.Caption) $productName" -match 'Azure Stack HCI|Azure Local')
    }
}

function Select-IniSection {
    <#
      Ordnet das erkannte Betriebssystem einem Abschnitt der SetupBD.ini zu.
      Die Prüfung läuft von der neuesten zur ältesten Build-Nummer; das erste Passende gewinnt.
    #>
    param([Parameter(Mandatory)]$OsInfo)
    $b = $OsInfo.Build

    # Azure Stack HCI / Azure Local hat eigene Abschnitte
    if ($OsInfo.IsHci) {
        if ($b -ge 26100) { return 'AzureStackHCI24H2AndLater' } else { return 'AzureStackHCI23H2AndEarlier' }
    }

    # Server-Betriebssysteme
    if ($OsInfo.IsServer) {
        if ($b -ge 26100) { return 'Win2025Server' }    # Server 2025
        if ($b -ge 20348) { return 'Win2022Server' }    # Server 2022 (auch 23H2 / 25398)
        if ($b -ge 17763) { return 'Win10RS5Server' }   # Server 2019 (+ SAC 1903-20H2)
        if ($b -ge 14393) { return 'Win10Server' }      # Server 2016
        throw "Nicht unterstütztes Server-Betriebssystem (Build $b)."
    }

    # Client-Betriebssysteme
    if ($b -ge 22000) { return 'Win11' }                # Windows 11
    if ($b -ge 17763) { return 'Win10RS532E' }          # Windows 10 1809+
    if ($b -ge 10240) { return 'Win1032E' }             # Windows 10 vor 1809
    throw "Nicht unterstütztes Client-Betriebssystem (Build $b)."
}

function Resolve-CdRoot {
    <#
      Ermittelt den Stammordner der CD (hier: ...\31.1), auf den sich die InfDir-Pfade beziehen.
      LevelsToRoot=4 bedeutet: vom Ordner der SetupBD.exe (Winx64) aus 4 Ebenen nach oben:
        Winx64 -> SETUPBD (1) -> SETUP (2) -> APPS (3) -> 31.1 (4)
    #>
    param([string]$IniFile, [int]$Levels, [string]$ProbeRelPath)

    # Ordner der SetupBD.exe = Elternordner von ...\Resource (in dem die INI liegt)
    $exeDir = Split-Path -Parent (Split-Path -Parent $IniFile)

    # $Levels-mal eine Ebene nach oben gehen
    $root = $exeDir
    for ($i = 0; $i -lt $Levels; $i++) { $root = Split-Path -Parent $root }

    # Plausibilitätsprüfung: existiert der erste InfDir unter diesem Root? Dann passt es.
    if ($ProbeRelPath -and (Test-Path -LiteralPath (Join-Path $root $ProbeRelPath))) { return $root }

    # Fallback: Von der INI aus Ebene für Ebene nach oben laufen, bis der erste InfDir gefunden wird
    $dir = Split-Path -Parent $IniFile
    while ($dir) {
        if ($ProbeRelPath -and (Test-Path -LiteralPath (Join-Path $dir $ProbeRelPath))) {
            Write-Warning "LevelsToRoot=$Levels passt nicht, verwende gefundenen Root: $dir"
            return $dir
        }
        $parent = Split-Path -Parent $dir
        if ($parent -eq $dir) { break }      # Laufwerkswurzel erreicht -> Schleife beenden
        $dir = $parent                       # eine Ebene höher weitermachen
    }

    # Nichts gefunden: trotzdem den berechneten Root verwenden (fehlende Ordner werden später gemeldet)
    Write-Warning "Erster InfDir '$ProbeRelPath' nicht gefunden; verwende Root laut LevelsToRoot: $root"
    return $root
}

function Get-PresentIntelHardwareIds {
    <#
      Liefert alle Hardware-IDs der im System vorhandenen Intel-PCI-Geräte (VEN_8086).
      Nur für -OnlyMatchingHardware nötig.
      Ein Gerät hat mehrere Hardware-IDs, von spezifisch bis allgemein, z. B.:
        PCI\VEN_8086&DEV_1533&SUBSYS_00008086&REV_03
        PCI\VEN_8086&DEV_1533&SUBSYS_00008086
        PCI\VEN_8086&DEV_1533
    #>

    # HashSet = Liste ohne Duplikate mit schneller Suche; Vergleich ohne Groß/Klein-Unterscheidung
    $set = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)

    Get-PnpDevice -PresentOnly -ErrorAction SilentlyContinue |      # alle aktuell vorhandenen Geräte
        Where-Object { $_.InstanceId -like 'PCI\VEN_8086*' } |      # nur Intel-PCI-Geräte (auch ohne Treiber)
        ForEach-Object {
            # Hardware-IDs des Geräts abfragen
            $prop = Get-PnpDeviceProperty -InstanceId $_.InstanceId -KeyName 'DEVPKEY_Device_HardwareIds' -ErrorAction SilentlyContinue
            # Nur verarbeiten, wenn die Eigenschaft existiert und Daten hat (StrictMode-sicher)
            if ($prop -and $prop.PSObject.Properties['Data'] -and $prop.Data) {
                # Jede ID ins HashSet aufnehmen ([void] unterdrückt die True/False-Rückgabe von Add)
                foreach ($h in @($prop.Data)) { if ($h) { [void]$set.Add([string]$h) } }
            }
        }

    # Das Komma verhindert, dass PowerShell das HashSet beim Zurückgeben in Einzelwerte "ausrollt"
    return ,$set
}

function Get-InfHardwareIds {
    <#
      Liest aus einer INF-Datei die PCI-Hardware-IDs, die in den Modell-Abschnitten stehen,
      die laut INI (InfSectionName) für dieses Betriebssystem relevant sind.
      Beispiel INF-Ausschnitt:
        [Intel.NTamd64.10.0...26100]
        %E1533NC.DeviceDesc% = E1533.10.0.1, PCI\VEN_8086&DEV_1533&SUBSYS_00008086
    #>
    param([string]$InfPath, [string[]]$SectionNames)

    $ids = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $inSection = $false        # merkt sich, ob wir gerade in einem relevanten Abschnitt sind

    # Regex für eine PCI-Hardware-ID:
    #   PCI\VEN_xxxx&DEV_xxxx            Pflichtteil (4 Hex-Ziffern)
    #   (&SUBSYS_xxxxxxxx)?              optional, 8 Hex-Ziffern
    #   (&REV_xx)?                       optional, 2 Hex-Ziffern
    #   (?i)                             Groß/Klein egal
    $rx = [regex]'(?i)PCI\\VEN_[0-9A-F]{4}&DEV_[0-9A-F]{4}(&SUBSYS_[0-9A-F]{8})?(&REV_[0-9A-F]{2})?'

    # INF Zeile für Zeile lesen (ReadLines erkennt UTF-16/UTF-8 anhand des BOM automatisch)
    foreach ($raw in [System.IO.File]::ReadLines($InfPath)) {
        # Kommentar ab ';' abschneiden und Leerzeichen entfernen
        $line = ($raw -replace ';.*$', '').Trim()
        if (-not $line) { continue }                       # Leerzeile -> überspringen

        # Abschnittskopf [Name]: prüfen, ob dieser Abschnitt in der InfSectionName-Liste steht
        if ($line -match '^\[(.+)\]$') {
            $inSection = ($SectionNames -contains $Matches[1].Trim())
            continue
        }

        # In einem relevanten Abschnitt: alle Hardware-IDs der Zeile einsammeln
        if ($inSection) {
            foreach ($m in $rx.Matches($line)) { [void]$ids.Add($m.Value) }
        }
    }
    return ,$ids               # Komma: HashSet als Ganzes zurückgeben
}

function Write-RollbackScript {
    <#
      Ermittelt die durch diesen Lauf NEU in den Treiberspeicher gekommenen Pakete (oemXX.inf)
      und schreibt ein Rollback-Script, das genau diese wieder entfernt.
      - $Before         : Treiberspeicher vor der Installation (darf auch LEER sein)
      - $InstalledInfs  : INF-Dateinamen, die dieser Lauf an pnputil übergeben hat
      Nur Pakete, deren Original-INF-Name in $InstalledInfs steht, kommen ins Rollback.
      So werden fremde Pakete (z. B. parallel durch Windows Update) nie mit gelöscht.
    #>
    param(
        [object[]]$Before,
        [System.Collections.Generic.HashSet[string]]$InstalledInfs,
        [string]$Folder,
        [string]$Stamp
    )

    # Treiberspeicher nach der Installation einlesen
    $after = @(Get-WindowsDriver -Online | Select-Object Driver, OriginalFileName, ProviderName, ClassName, Version, Date)

    # Alle vorher vorhandenen oemXX.inf-Namen in ein HashSet (schneller Vergleich, Groß/Klein egal)
    $known = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($b in @($Before)) { if ($b) { [void]$known.Add([string]$b.Driver) } }

    # Neu = jetzt vorhanden, vorher nicht
    $allNew = @($after | Where-Object { -not $known.Contains([string]$_.Driver) })

    # Aufteilen: Pakete aus diesem Lauf (Original-INF-Name passt zu einer installierten INF) und fremde Pakete
    $newPkgs = @()
    $foreign = @()
    foreach ($pkg in $allNew) {
        $leaf = Split-Path -Leaf ([string]$pkg.OriginalFileName)
        if ($InstalledInfs.Contains($leaf)) { $newPkgs += $pkg } else { $foreign += $pkg }
    }

    # Fremde neue Pakete nur melden, nicht ins Rollback aufnehmen
    foreach ($f in $foreign) {
        Write-Warning "Neues Paket stammt nicht aus diesem Lauf, wird NICHT ins Rollback aufgenommen: $($f.Driver) ($(Split-Path -Leaf ([string]$f.OriginalFileName)))"
    }

    if ($newPkgs.Count -eq 0) {
        Write-Host 'Keine neuen Treiberpakete aus diesem Lauf (alles bereits im Treiberspeicher vorhanden) - kein Rollback nötig.'
        return
    }

    $rollback = Join-Path $Folder "Rollback-IntelNicDriver_$Stamp.ps1"

    # --- Paketliste für das Rollback-Script (einfache Anführungszeichen werden verdoppelt) ---
    $pkgLines = foreach ($p in $newPkgs) {
        $inf = Split-Path -Leaf ([string]$p.OriginalFileName)
        "    @{{ Driver = '{0}'; Inf = '{1}'; Version = '{2}' }}" -f `
            ([string]$p.Driver -replace "'", "''"), ($inf -replace "'", "''"), ([string]$p.Version -replace "'", "''")
    }

    # --- Inhalt des Rollback-Scripts (Here-String: $ wird NICHT ersetzt, außer in den Platzhaltern) ---
    $template = @'
#Requires -RunAsAdministrator
# Rollback für Install-IntelNicDriver vom __STAMP__
# Entfernt die in diesem Lauf neu hinzugefügten Treiberpakete wieder.
# Windows bindet danach den nächstbesten vorhandenen Treiber (älteres Paket oder Inbox-Treiber).
# ACHTUNG: Die Netzwerkverbindung über betroffene Adapter wird kurz unterbrochen.
#
# Aufruf (als Administrator):
#   powershell.exe -NoProfile -ExecutionPolicy Bypass -File "__PATH__"
#   Mit -WhatIf wird nur angezeigt, was entfernt würde.

param([switch]$WhatIf)

# Zu entfernende Pakete (Driver = Name im Treiberspeicher, Inf = Original-INF, Version)
$packages = @(
__PACKAGES__
)

$pnputil = Join-Path $env:SystemRoot 'System32\pnputil.exe'
$errors  = 0
$reboot  = $false

# Jedes Paket einzeln entfernen und Exit-Code prüfen
foreach ($p in $packages) {
    $text = "pnputil /delete-driver $($p.Driver) /uninstall /force   ($($p.Inf) $($p.Version))"
    if ($WhatIf) { Write-Host "WhatIf: $text"; continue }
    Write-Host $text -ForegroundColor Yellow
    & $pnputil /delete-driver $p.Driver /uninstall /force
    switch ($LASTEXITCODE) {
        0       { }                                   # OK
        3010    { $reboot = $true }                   # OK, Neustart nötig
        default { Write-Warning "Fehler $LASTEXITCODE bei $($p.Driver)"; $errors++ }
    }
}

# Geräte neu erkennen, damit Windows sofort den nächstbesten Treiber bindet
# (pnputil /scan-devices gibt es erst ab Windows 10 1903 / Build 18362)
if (-not $WhatIf -and [Environment]::OSVersion.Version.Build -ge 18362) { & $pnputil /scan-devices | Out-Null }

if ($errors -gt 0) { Write-Warning "$errors Paket(e) konnten nicht entfernt werden."; exit 1 }
if ($WhatIf)       { Write-Host 'Testlauf beendet - nichts entfernt.'; exit 0 }
if ($reboot)       { Write-Warning 'Neustart erforderlich.'; exit 3010 }
Write-Host 'Rollback abgeschlossen.' -ForegroundColor Green
exit 0
'@

    # Platzhalter füllen
    $content = $template.Replace('__STAMP__', $Stamp).Replace('__PATH__', $rollback).Replace('__PACKAGES__', ($pkgLines -join "`r`n"))

    # UTF8 schreibt in PS 5.1 mit BOM -> Umlaute in den Kommentaren bleiben lesbar
    Set-Content -LiteralPath $rollback -Value $content -Encoding UTF8 -WhatIf:$false

    Write-Host ''
    Write-Host "Neue Treiberpakete: $($newPkgs.Count)  -> Rollback-Script: $rollback" -ForegroundColor Green
    # Tabelle der neuen Pakete (Format-Table -> Text -> Write-Host, damit sie sicher im Transcript landet)
    $newPkgs | Format-Table Driver, @{n='INF';e={Split-Path -Leaf ([string]$_.OriginalFileName)}}, Version -AutoSize | Out-String -Width 250 | Write-Host
}

#endregion

#region Vorbereitung -------------------------------------------------------------

# Die CD enthält nur 64-Bit-Treiber; eine 32-Bit-PowerShell würde außerdem auf SysWOW64 umgeleitet
if (-not [Environment]::Is64BitProcess) {
    throw 'Bitte in einer 64-Bit-PowerShell ausführen (die CD enthält nur Winx64-Treiber).'
}

# Zeitstempel für eindeutige Dateinamen (Log, CSV, Rollback)
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'

# Log-Ordner anlegen (auch bei -WhatIf, damit das Protokoll geschrieben werden kann)
New-Item -ItemType Directory -Path $LogDir -Force -WhatIf:$false | Out-Null
$transcript = Join-Path $LogDir "Install-IntelNicDriver_$timestamp.log"

# Transcript (komplette Konsolenausgabe in Datei) starten - auch im Testlauf
Start-Transcript -LiteralPath $transcript | Out-Null
if ($WhatIf) { Write-Host '*** TESTLAUF (-WhatIf): es wird nichts installiert ***' -ForegroundColor Magenta }

$exitCode = 0      # Rückgabewert des Scripts; wird bei Fehlern / Neustartbedarf angepasst

# Für das Rollback im finally-Block schon hier anlegen (StrictMode: Variablen müssen existieren)
$doChanges      = $false     # True = echter Lauf (kein -WhatIf)
$before         = $null      # Treiberspeicher vor der Installation; $null = noch nicht gesichert
$pnputilStarted = $false     # True, sobald mindestens ein pnputil-Aufruf erfolgt ist
# INF-Dateinamen, die an pnputil übergeben wurden (Groß/Klein egal)
$installedInfs  = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)

try {
    # --- SetupBD.ini finden ---------------------------------------------------------
    Write-Host "Script-Ordner: $ScriptDir"
    if (-not $IniPath) {
        # CD-Ordner = Script-Ordner + oben deklarierter Name
        $cdFolder = Join-Path $ScriptDir $CdFolderName
        Write-Host "CD-Ordner   : $cdFolder"
        if (-not (Test-Path -LiteralPath $cdFolder -PathType Container)) {
            throw "CD-Ordner nicht gefunden: $cdFolder  (Name oben im Script unter `$CdFolderName prüfen)"
        }
        $IniPath = Find-SetupBDIni -BaseDir $cdFolder                     # SetupBD.ini im CD-Ordner suchen
        if (-not $IniPath) {
            throw "SetupBD.ini nicht gefunden. Erwartet: $cdFolder\<Version>\APPS\SETUP\SETUPBD\Winx64\Resource\SetupBD.ini"
        }
    }
    if (-not (Test-Path -LiteralPath $IniPath)) {
        throw "SetupBD.ini nicht gefunden: $IniPath"
    }
    $IniPath = (Resolve-Path -LiteralPath $IniPath).ProviderPath      # relativen Pfad in vollen Pfad umwandeln
    Write-Host "SetupBD.ini : $IniPath"

    # --- INI einlesen ---------------------------------------------------------------
    $ini = Read-MultiValueIni -Path $IniPath

    # --- Betriebssystem erkennen ----------------------------------------------------
    $osInfo = Get-OsInfo
    Write-Host ("Betriebssystem: {0} | Build {1} | Server={2} | HCI={3}" -f $osInfo.Caption, $osInfo.Build, $osInfo.IsServer, $osInfo.IsHci)

    # --- INI-Abschnitt wählen (oder den per -Section vorgegebenen verwenden) -------
    if (-not $Section) { $Section = Select-IniSection -OsInfo $osInfo }
    if (-not $ini.Contains($Section)) {
        # Alle Abschnitte außer [Config] als Hilfe auflisten
        $available = ($ini.Keys | Where-Object { $_ -ne 'Config' }) -join ', '
        throw "Abschnitt [$Section] nicht in der INI vorhanden. Verfügbar: $available"
    }
    Write-Host "INI-Abschnitt : [$Section]" -ForegroundColor Cyan

    # --- Werte des Abschnitts auslesen ---------------------------------------------
    $sec = $ini[$Section]     # Hashtable: Schlüssel -> Liste von Werten

    # Direkt zuweisen (nicht über if/else), sonst wird aus @() $null bzw. aus einem Einzelwert ein String
    [string[]]$infDirs    = @()      # Treiberordner relativ zum CD-Root
    [string[]]$infExclude = @()      # INF-Dateinamen, die NICHT installiert werden sollen
    if ($sec.ContainsKey('InfDir'))     { $infDirs    = @($sec['InfDir']) }
    if ($sec.ContainsKey('InfExclude')) { $infExclude = @($sec['InfExclude']) }

    # InfSectionName ist eine komma-getrennte Liste von INF-Modellabschnitten
    # (z. B. Intel.NTamd64.10.0...26100). Wird nur für -OnlyMatchingHardware gebraucht.
    $sectionNames = @()
    if ($sec.ContainsKey('InfSectionName')) {
        # Alle InfSectionName-Zeilen zusammenfügen und am Komma aufteilen
        foreach ($n in ($sec['InfSectionName'] -join ',').Split(',')) {
            $n = $n.Trim()
            if (-not $n) { continue }                  # leere Einträge (z. B. ",,") überspringen
            $sectionNames += $n
            # Die INI hat teils Einträge ohne 'Intel.'-Präfix (z. B. NTamd64.10.0...26100)
            # -> zusätzlich die Variante mit Präfix aufnehmen, damit der Abschnitt in der INF gefunden wird
            if ($n -notlike 'Intel.*' -and $n -ne 'Intel') { $sectionNames += "Intel.$n" }
        }
    }
    if ($infDirs.Count -eq 0) { throw "Abschnitt [$Section] enthält keine InfDir-Einträge." }

    # --- CD-Root bestimmen ----------------------------------------------------------
    $levels = 4    # Standardwert, falls [Config] LevelsToRoot fehlt
    if ($ini.Contains('Config') -and $ini['Config'].ContainsKey('LevelsToRoot')) {
        $levels = [int]$ini['Config']['LevelsToRoot'][0]
    }
    # Der erste InfDir dient als Probe, ob der berechnete Root stimmt
    $cdRoot = Resolve-CdRoot -IniFile $IniPath -Levels $levels -ProbeRelPath $infDirs[0]
    Write-Host "CD-Root     : $cdRoot"

    #endregion

    #region INF-Dateien sammeln ---------------------------------------------------

    # Für jeden InfDir alle *.inf einsammeln; die Ausgabe der Schleife landet direkt in $infFiles
    $infFiles = foreach ($rel in $infDirs) {
        $dir = Join-Path $cdRoot $rel                                   # absoluter Treiberordner
        if (-not (Test-Path -LiteralPath $dir)) {
            Write-Warning "InfDir fehlt, übersprungen: $dir"            # Ordner nicht auf der CD -> weiter
            continue
        }
        Get-ChildItem -LiteralPath $dir -Filter '*.inf' -File |
            Where-Object {
                # Dateien aus InfExclude herausfiltern (-contains vergleicht ohne Groß/Klein)
                if ($infExclude -contains $_.Name) {
                    Write-Verbose "Ausgeschlossen (InfExclude): $($_.FullName)"
                    $false                                              # Datei verwerfen
                } else {
                    $true                                               # Datei behalten
                }
            }
    }
    $infFiles = @($infFiles)                                            # sicherstellen, dass es ein Array ist
    if ($infFiles.Count -eq 0) { throw 'Keine INF-Dateien gefunden.' }

    # --- Optional: nur INFs für vorhandene Hardware --------------------------------
    if ($OnlyMatchingHardware) {
        $present = Get-PresentIntelHardwareIds                         # IDs der eingebauten Intel-Geräte
        Write-Host "Vorhandene Intel-PCI-Hardware-IDs: $($present.Count)"

        # Jede INF prüfen: enthält sie mindestens eine ID, die auch im System vorkommt?
        $infFiles = @($infFiles | Where-Object {
            $ids = Get-InfHardwareIds -InfPath $_.FullName -SectionNames $sectionNames
            $hit = $false
            foreach ($id in $ids) {
                if ($present.Contains($id)) { $hit = $true; break }    # erster Treffer reicht
            }
            if (-not $hit) { Write-Verbose "Keine passende Hardware: $($_.Name)" }
            $hit                                                       # True = INF behalten
        })

        if ($infFiles.Count -eq 0) {
            Write-Warning 'Keine INF passt zu vorhandener Intel-Hardware. Nichts zu tun.'
            return                                                     # Script beenden (finally läuft trotzdem)
        }
    }

    # Liste der zu installierenden INFs ausgeben (relativ zum CD-Root, besser lesbar)
    Write-Host ''
    Write-Host "Zu installierende INF-Dateien ($($infFiles.Count)):" -ForegroundColor Cyan
    $infFiles | ForEach-Object { Write-Host ('  ' + $_.FullName.Substring($cdRoot.Length).TrimStart('\')) }
    Write-Host ''

    #endregion

    #region Sicherung vorher ------------------------------------------------------

    $doChanges = -not $WhatIf                # True = echter Lauf, False = Testlauf mit -WhatIf
    if ($doChanges) {
        # Aktuellen Stand des Treiberspeichers (nur Drittanbieter-Pakete oemXX.inf) sichern.
        # Daraus wird nach der Installation ermittelt, welche Pakete neu sind -> Rollback.
        Write-Host 'Sichere aktuellen Treiberspeicher (Get-WindowsDriver)...'
        $before = @(Get-WindowsDriver -Online | Select-Object Driver, OriginalFileName, ProviderName, ClassName, Version, Date)
        $backupCsv = Join-Path $LogDir "DriverStore_vorher_$timestamp.csv"
        $before | Export-Csv -LiteralPath $backupCsv -NoTypeInformation -Encoding UTF8 -WhatIf:$false
        Write-Host "  -> $backupCsv"
    }

    #endregion

    #region Installation ----------------------------------------------------------

    $pnputil = Join-Path $env:SystemRoot 'System32\pnputil.exe'        # vollständiger Pfad, unabhängig von PATH
    $results = New-Object System.Collections.Generic.List[object]       # Ergebnis je INF für die Zusammenfassung
    $rebootNeeded = $false

    # Jede INF einzeln installieren, damit Fehler pro Datei sichtbar sind
    foreach ($inf in $infFiles) {
        # pnputil-Argumente zusammenbauen: /add-driver <inf> [/install]
        $pnpArgs = @('/add-driver', $inf.FullName)
        if (-not $NoInstall) { $pnpArgs += '/install' }                 # /install = sofort an passende Geräte binden

        # Testlauf: Befehl nur anzeigen und als "WhatIf" in die Ergebnisliste aufnehmen
        if ($WhatIf) {
            Write-Host "WhatIf: pnputil $($pnpArgs -join ' ')" -ForegroundColor DarkCyan
            $results.Add([pscustomobject]@{ Inf = $inf.Name; Ordner = $inf.Directory.FullName.Substring($cdRoot.Length).TrimStart('\'); ExitCode = $null; Status = 'WhatIf' })
            continue
        }

        Write-Host "pnputil -> $($inf.Name)" -ForegroundColor Yellow
        [void]$installedInfs.Add($inf.Name)      # für das Rollback merken
        $pnputilStarted = $true

        # In PS 5.1 wird stderr eines externen Programms bei ErrorActionPreference=Stop zum Abbruch
        # -> für den pnputil-Aufruf vorübergehend auf 'Continue' setzen
        $prevEap = $ErrorActionPreference
        $ErrorActionPreference = 'Continue'
        try {
            # 2>&1 = stderr in stdout umleiten; ForEach-Object macht aus Fehlerobjekten einfachen Text
            $output = & $pnputil @pnpArgs 2>&1 | ForEach-Object { "$_" } | Out-String
            $code = $LASTEXITCODE                                       # Exit-Code von pnputil
        } finally {
            $ErrorActionPreference = $prevEap                           # ursprüngliche Einstellung wiederherstellen
        }
        Write-Host $output.TrimEnd()

        # Exit-Code von pnputil in einen Status übersetzen
        $status = switch ($code) {
            0       { 'OK' }
            3010    { $rebootNeeded = $true; 'OK - Neustart erforderlich' }       # ERROR_SUCCESS_REBOOT_REQUIRED
            259     { 'Hinzugefügt - kein passendes Gerät / bereits aktuell' }   # ERROR_NO_MORE_ITEMS
            default { 'FEHLER' }
        }
        if ($status -eq 'FEHLER') { $exitCode = 1 }                     # mindestens ein Fehler -> Exit-Code 1

        # Ergebnis dieser INF merken
        $results.Add([pscustomobject]@{ Inf = $inf.Name; Ordner = $inf.Directory.FullName.Substring($cdRoot.Length).TrimStart('\'); ExitCode = $code; Status = $status })
    }

    #endregion

    #region Rollback-Script & Zusammenfassung ------------------------------------

    # Das Rollback-Script wird im finally-Block erzeugt -> auch nach einem Abbruch

    # Ergebnisliste als CSV speichern
    $resultCsv = Join-Path $LogDir "Ergebnis_$timestamp.csv"
    $results | Export-Csv -LiteralPath $resultCsv -NoTypeInformation -Encoding UTF8 -WhatIf:$false

    # Zusammenfassung auf der Konsole
    Write-Host ''
    Write-Host 'Zusammenfassung:' -ForegroundColor Cyan
    $results | Format-Table -AutoSize | Out-String -Width 250 | Write-Host

    # Neustartbedarf nur melden, wenn es keine Fehler gab (Fehler hat Vorrang beim Exit-Code)
    if ($exitCode -eq 0 -and $rebootNeeded) {
        $exitCode = 3010
        Write-Warning 'Neustart erforderlich.'
    }
    Write-Host "Log: $transcript"
}
catch {
    # Jeder Abbruch im try-Block landet hier: Meldung + Stelle im Script ausgeben
    Write-Host "FEHLER: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor DarkGray
    $exitCode = 1
}
finally {
    # Rollback-Script schreiben - auch wenn der Lauf mit Fehler abgebrochen ist.
    # Bedingung: echter Lauf, Sicherung vorher vorhanden ($null-Prüfung: eine LEERE Liste ist gültig!)
    # und mindestens ein pnputil-Aufruf erfolgt.
    if ($doChanges -and $null -ne $before -and $pnputilStarted) {
        try {
            Write-RollbackScript -Before $before -InstalledInfs $installedInfs -Folder $LogDir -Stamp $timestamp
        } catch {
            Write-Host "FEHLER beim Erzeugen des Rollback-Scripts: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "Sicherung des Treiberspeichers vor der Installation: DriverStore_vorher_$timestamp.csv" -ForegroundColor Red
            $exitCode = 1
        }
    }

    # Wird IMMER ausgeführt (auch nach Fehler oder 'return'): Transcript sauber beenden
    try { Stop-Transcript | Out-Null } catch { }
}

# Exit-Code an den Aufrufer zurückgeben (z. B. für SCCM/MDT/Task-Sequenzen)
if ($RunAsFile) {
    exit $exitCode                        # als Datei gestartet: Exit-Code zurückgeben
} else {
    $global:LASTEXITCODE = $exitCode      # eingefügt/F8: Konsole nicht schließen, Code nur setzen
}

#endregion
