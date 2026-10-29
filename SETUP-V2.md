# MyJourney V2 – Account und Cloud

## Was funktioniert
- Registrierung per E-Mail/Passwort, Anmeldung, Abmeldung, Passwort-Reset-Link (Supabase Auth).
- Cloud-Speicherung der Reisen und Buchungen pro Benutzer, Synchronisation beim Laden auf anderen Geräten.
- PostgreSQL Row Level Security: jeder Nutzer sieht nur seine eigenen Daten.
- Datenmodelle für spätere Provider-Verbindungen und deduplizierte Mail-Imports.

## Einmalige Einrichtung (ohne Installation)
1. Öffne https://supabase.com/dashboard und erstelle ein kostenloses Projekt. Bewahre das Datenbankpasswort sicher auf, **niemals in GitHub einfügen**.
2. Öffne im Supabase-Projekt **SQL Editor**, füge `database.sql` ein und führe es aus.
3. Unter **Project Settings > API** (Menübezeichnungen können abweichen) kopiere **Project URL** und den **publishable / anon key**.
4. Öffne `config.js` im GitHub-Repository, klicke auf den Stift (Edit), ersetze `YOUR_SUPABASE_PROJECT_URL` und `YOUR_SUPABASE_PUBLISHABLE_KEY` und speichere mit Commit changes. Der publishable Key ist für öffentliche Webclients vorgesehen. **Niemals** den `service_role` oder Secret Key verwenden.
5. In Supabase **Authentication > URL Configuration**: Site URL auf `https://DEIN-NAME.github.io/myjourney/` setzen und dieselbe URL bei Redirect URLs ergänzen. Optional E-Mail-Bestätigung aktiv lassen (empfohlen).
6. Alle Dateien aus diesem ZIP **im Repository-Stammverzeichnis** hochladen und vorhandene Dateien ersetzen. Nicht den übergeordneten Ordner hochladen. GitHub Pages unter Settings > Pages: `main`, `/(root)`.
7. Warte auf GitHub Pages und öffne die Website in einem neuen privaten Fenster. Registriere dich, bestätige ggf. deine E-Mail, melde dich an und lege eine Testreise an. Melde dich am zweiten Gerät mit demselben Konto an.

## Achtung: bestehende lokale Daten
Die bisherige v1-App speichert Daten nur im Browser. **Vor dem Upgrade** unter Mehr > Alle Daten exportieren eine Sicherung machen. V2 beginnt für jeden neuen Cloud-Account mit einer leeren Reiseübersicht. Für eine Übernahme der v1-Daten muss ein separater sicherer Import in das Cloud-Konto ergänzt werden. Lösche das Backup nicht.

## Outlook/Hotmail: vorbereitet, noch NICHT verbunden
Geplante Architektur:
- Microsoft Entra App-Registrierung für persönliche Microsoft-Konten (`consumers`), OAuth 2.0 Authorization Code + PKCE, explizite Zustimmung zu `Mail.Read` und ggf. `offline_access`.
- Ein sicherer serverseitiger Callback (z.B. Supabase Edge Function) tauscht den Code gegen Tokens und speichert Refresh-Tokens **verschlüsselt serverseitig**, nie im Browser oder GitHub.
- Hintergrundjob liest per Microsoft Graph nur relevante Nachrichten, dedupliziert anhand der Provider-Nachrichten-ID, extrahiert Buchungsdaten, markiert unsichere Ergebnisse zur Prüfung und speichert strukturierte Einträge im jeweiligen Nutzerkonto.
- Verbindung trennen, Token widerrufen, importierte Daten löschen und minimale Aufbewahrung einplanen.
- `provider_connections` und `imported_messages` sind Datenmodelle für diese spätere Integration. Es gibt noch **keinen** funktionierenden Outlook-Import oder eine Microsoft-Anmeldeschaltfläche.

## Sicherheits- und MVP-Grenzen
- V2 synchronisiert den gesamten Reisestand per JSON-Datensatz; bei gleichzeitigen Änderungen auf mehreren Geräten gewinnt derzeit der letzte Schreibvorgang. Für echte Mehrbenutzer-/Offline-Synchronisation später einzelne Datensätze und Konfliktauflösung einführen.
- Keine sensiblen Ausweisnummern in Buchungsnotizen speichern, bis das Datenschutzkonzept geprüft ist.
- Authentifizierungs- und Cloud-Code wurde in dieser Umgebung nicht mit einem echten Supabase-Projekt getestet.
- Browser-Storage bleibt ein lokaler Cache. Auf gemeinsam genutzten Rechnern immer abmelden.
- `sw.js` der alten PWA kann Dateien zwischenspeichern: nach Deployment ggf. neu laden oder Website-Daten löschen (vorher v1-Backup!).
