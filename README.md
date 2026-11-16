# Install-IntelNicDriver

PowerShell-Script, das Intel-Ethernet-Treiber von der **Intel Ethernet Adapter Complete Driver Pack**-CD mit `pnputil` installiert.
Welche Treiberordner für das laufende Betriebssystem verwendet werden, liest das Script aus der `SetupBD.ini` der CD – genau wie Intels eigenes Setup (`SetupBD.exe`).

## Ablage

Das Script und der CD-Ordner liegen im selben Ordner. Laufwerk und Pfad sind beliebig.

```
X:\beliebiger\Ordner\
├─ Install-IntelNicDriver.ps1
└─ Intel®EthernetAdaptersConnectionsCD_311_1316650\     <- Name oben im Script festgelegt
   └─ 31.1\
      ├─ APPS\SETUP\SETUPBD\Winx64\Resource\SetupBD.ini
      ├─ PRO1000\  PRO2500\  PROXGB\  PRO40GB\  PROCGB\  PROAVF\ ...
```

Den Namen des CD-Ordners legst du ganz oben im `param`-Block fest:

```powershell
[string]$CdFolderName = 'Intel®EthernetAdaptersConnectionsCD_311_1316650',
```

> Wegen des `®` muss das Script als **UTF-8 mit BOM** gespeichert bleiben, sonst liest Windows PowerShell 5.1 den Namen falsch.

## Aufruf

Als Administrator, in einer 64-Bit-PowerShell:

```powershell
# Testlauf – zeigt erkannten INI-Abschnitt und die INF-Dateien, ändert nichts
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-IntelNicDriver.ps1 -WhatIf

# Nur Treiber für tatsächlich eingebaute Intel-Hardware installieren
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-IntelNicDriver.ps1 -OnlyMatchingHardware

# Alle Treiber des passenden INI-Abschnitts installieren
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-IntelNicDriver.ps1
```

| Parameter | Bedeutung |
|---|---|
| `-CdFolderName` | Name des CD-Ordners im Script-Ordner (Standard: oben im Script) |
| `-IniPath` | Direkter Pfad zur `SetupBD.ini`, übersteuert die Suche |
| `-Section` | INI-Abschnitt erzwingen statt automatischer OS-Erkennung, z. B. `Win2022Server` |
| `-OnlyMatchingHardware` | Nur INFs installieren, deren Hardware-IDs im System vorhanden sind |
| `-NoInstall` | Nur in den Treiberspeicher aufnehmen (`pnputil /add-driver` ohne `/install`) |
| `-LogDir` | Ablage für Log, CSVs und Rollback-Script (Standard: `C:\Windows\Temp\IntelNicDriver`) |
| `-WhatIf` | Testlauf: zeigt die `pnputil`-Befehle nur an, installiert nichts |
| `-Verbose` | Zusätzliche Details (z. B. welche INF wegen `InfExclude` oder fehlender Hardware entfällt) |

## Betriebssystem → INI-Abschnitt

| Betriebssystem | Abschnitt |
|---|---|
| Windows Server 2016 (14393) | `Win10Server` |
| Windows Server 2019 (17763) | `Win10RS5Server` |
| Windows Server 2022 (20348) | `Win2022Server` |
| Windows Server 2025 (26100+) | `Win2025Server` |
| Azure Stack HCI / Azure Local | `AzureStackHCI23H2AndEarlier` / `AzureStackHCI24H2AndLater` |
| Windows 10 vor 1809 / ab 1809 | `Win1032E` / `Win10RS532E` |
| Windows 11 | `Win11` |

Aus dem Abschnitt werden `InfDir` (Treiberordner), `InfExclude` (ausgelassene INFs) und `InfSectionName` (für `-OnlyMatchingHardware`) gelesen. Der CD-Stammordner ergibt sich aus `[Config] LevelsToRoot`.

## Sicherung und Rollback

- Vor der Installation wird der Treiberspeicher als CSV gesichert (`DriverStore_vorher_<Zeit>.csv`).
- Nach der Installation wird ein Rollback-Script erzeugt (`Rollback-IntelNicDriver_<Zeit>.ps1`) – **auch wenn der Lauf mit einem Fehler abbricht** und auch auf Rechnern ohne vorherige Drittanbieter-Treiber.
- Ins Rollback kommen nur Pakete, die **dieser Lauf** neu hinzugefügt hat (Abgleich über den Original-INF-Namen). Pakete, die zeitgleich z. B. durch Windows Update dazukommen, werden nur gemeldet und nicht entfernt.
- Das Rollback-Script entfernt die Pakete mit `pnputil /delete-driver <oemXX.inf> /uninstall /force`, prüft jeden Exit-Code und lässt Windows anschließend die Geräte neu erkennen (`pnputil /scan-devices`, ab Build 18362). Windows bindet danach den nächstbesten vorhandenen Treiber.

```powershell
# Erst ansehen, dann ausführen
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\Windows\Temp\IntelNicDriver\Rollback-IntelNicDriver_<Zeit>.ps1 -WhatIf
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\Windows\Temp\IntelNicDriver\Rollback-IntelNicDriver_<Zeit>.ps1
```

Rollback-Exit-Codes: 0 = OK, 3010 = OK + Neustart, 1 = mindestens ein Paket nicht entfernt.

- Ergebnis je INF steht in `Ergebnis_<Zeit>.csv`, die komplette Ausgabe im Transcript-Log.

## Exit-Codes

| Code | Bedeutung |
|---|---|
| 0 | OK |
| 3010 | OK, Neustart erforderlich |
| 1 | Mindestens ein Fehler |

## Hinweise

- Das Script verwendet bewusst kein `[CmdletBinding()]`; `-WhatIf` und `-Verbose` sind normale Schalter. Dadurch läuft es auch in der ISE (F8) oder in die Konsole eingefügt.

- Beim Installieren oder Aktualisieren eines Netzwerktreibers ist die Verbindung über den Adapter kurz unterbrochen – bei RDP/PS-Remoting einplanen.
- Die Intel-CD selbst ist **nicht** Teil dieses Repos; sie wird bei Intel heruntergeladen.
- Ist die Ausführung von Scripts gesperrt, hilft `-ExecutionPolicy Bypass` beim Aufruf (siehe oben) bzw. `Unblock-File .\Install-IntelNicDriver.ps1` nach dem Download.
