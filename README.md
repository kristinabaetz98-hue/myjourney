# MyJourney – Progressive Web App MVP

Eine deutschsprachige, ohne Build-Tools lauffähige Reiseplaner-PWA.

## Auf Windows starten

1. ZIP entpacken.
2. Python installieren, falls noch nicht vorhanden.
3. Im entpackten Ordner ein Terminal öffnen und `python -m http.server 8000` ausführen (unter Windows ggf. `py -m http.server 8000`).
4. Im Browser `http://localhost:8000` öffnen.

Alternativ den Ordner auf einen statischen HTTPS-Webhost (z. B. GitHub Pages oder Netlify) hochladen. Für die Installation auf dem iPhone ist eine **HTTPS-Adresse** erforderlich. Auf dem iPhone in Safari öffnen → Teilen → Zum Home-Bildschirm. Ein Windows-`localhost` ist nicht vom iPhone erreichbar.

## Funktionen

- Mehrere Reisen anlegen, ändern und löschen
- Flüge, Hotels, Transfers, Aktivitäten und sonstige Einträge erstellen, bearbeiten und löschen
- Automatisch chronologisch sortierter Tagesplan
- OpenStreetMap-Karte mit manuell hinterlegten Koordinaten (Leaflet, benötigt Internet)
- JSON-Buchungsimport, JSON-Datensicherung und Wiederherstellung
- Lokale Browser-Speicherung und Offline-App-Shell via Service Worker
- Responsives Design und PWA-Manifest

## JSON-Import

In „Buchungen“ `example-bookings.json` auswählen. Unterstützt ein Array oder `{ "bookings": [...] }`. Pflichtfelder: `title`, `start` (ISO-Datum oder lokale ISO-Datumzeit). Optional: `type` (`flight`, `hotel`, `transfer`, `activity`, `other`), `end`, `location`, `address`, `lat`, `lng`, `confirmation`, `cost`, `notes`.

## Grenzen der Version 0.1

Keine echten Booking.com-, Airbnb-, Flug- oder E-Mail-Verbindungen; keine PDF-Extraktion, keine KI-Planung und keine Cloud-Synchronisierung. JSON-Import ist kein direkter Anbieterimport. Keine automatisch berechneten Routen. Browserdaten können gelöscht werden; Backups exportieren. Externe Leaflet-Dateien und Kartenkacheln benötigen Internet. Ein SVG-Icon ist enthalten; für maximale iOS-Kompatibilität sollten später zusätzlich PNG-Icons erstellt werden. Keine persönlichen Buchungsdaten ohne geeignete Zugriffskontrollen auf öffentliche Server hochladen.

## Geplante Integrationsarchitektur

Buchungen besitzen `source`, `confirmation`, `tripId`, Zeitangaben und optionale Ortskoordinaten. Später können autorisierte Adapter aus APIs/OAuth, E-Mails oder Dokumenten in dieses einheitliche Datenmodell importieren. Vor Produktivbetrieb: serverseitige Tokenverwaltung, Verschlüsselung, Datenmigration, Dublettenerkennung und Benutzerkonten ergänzen.
