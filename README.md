# Network Ping Tool

Ein eigenständiges Excel-VBA-Tool zum schnellen Pingen von Geräten in industriellen Netzwerken – z. B. Siemens-CPUs, HMIs und PLCs über Profinet.

Entstanden ist das Tool, weil das manuelle Anpingen einzelner Geräte über das CMD-Fenster auf Dauer zu lange dauert und umständlich ist – gerade wenn mehrere Geräte im Netzwerk geprüft werden müssen.

## Funktionen

- Schnelles ICMP-Ping direkt über die Windows-API (`iphlpapi.dll`) – keine externen Programme nötig
- Drei Scan-Modi
- Sauberer Abbruch laufender Scans
- Eigenständiger UserForm-Modus
- Korrekte numerische IP-Sortierung (statt alphabetischer Sortierung)
- Bewusst kein paralleles Pingen, um Konflikte mit Unternehmens-Firewalls/Endpoint-Security (z. B. Sophos) zu vermeiden

## Voraussetzungen

- Microsoft Excel (Windows)
- Makros müssen aktiviert sein

## Installation

1. Die Datei `NetworkPingTool.xlsm` herunterladen.
2. Beim Öffnen Makros aktivieren (Sicherheitswarnung von Excel bestätigen).
3. Los geht's – keine weitere Installation oder Zusatzsoftware nötig.

## Verwendung

1. Datei öffnen.
2. IP-Adressen bzw. Hostnamen der zu prüfenden Geräte eintragen.
3. Scan starten und Ergebnisse in der Liste verfolgen.
4. Scan kann jederzeit sauber abgebrochen werden.

## Versionen

- **v1.0.0** – Erste veröffentlichte Version

## Support

Falls dir das Tool nützt und du die Weiterentwicklung unterstützen möchtest, freue ich mich über einen Kaffee: [ko-fi.com/Elospeed](https://ko-fi.com/Elospeed)

## Lizenz

Dieses Tool ist völlig frei nutzbar. Es darf beliebig verwendet, verändert und weiterentwickelt werden.

Die Nutzung erfolgt auf eigene Verantwortung – jegliche Haftung für Schäden oder Folgen, die durch die Verwendung dieses Tools entstehen, wird ausgeschlossen. Es besteht keinerlei Gewährleistung, weder ausdrücklich noch stillschweigend.
