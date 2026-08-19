# Sicherheitshinweise

## Keine echten Daten in dieses Repository

Dieses Repository enthaelt ausschliesslich Programmcode und synthetische
Beispieldaten. Bitte **niemals** hier ablegen oder in Meldungen einfuegen:

- echte oder auszugsweise Kunden-, Mitarbeiter- oder Vertragsdaten
- Bildschirmfotos mit Datenbank- oder Tabelleninhalten
- Verbindungszeichenfolgen, Kennwoerter, Zugangsdaten
- interne Server-, Freigabe- oder Laufwerkspfade
- produktive Dateinamen, sofern sie Rueckschluesse zulassen

Fuer eine Fehlermeldung genuegen der Ablauf, die Fehlermeldung und – falls
noetig – eine **nachgebaute** Datei mit erfundenen Inhalten.

## Sicherheitsprobleme melden

Sicherheitsrelevante Funde bitte **nicht** als oeffentliches Issue eroeffnen,
sondern direkt an die verantwortliche Person des Projekts melden. Hilfreich
sind: betroffene Fassung (`maskieren.cmd --version`), Ablauf zum Nachstellen
und die erwartete gegenueber der beobachteten Wirkung.

## Unterstuetzte Fassung

Gepflegt wird jeweils die neueste Fassung auf dem Branch `Nodejs`. Fuer den
Betrieb sollte eine feste Fassung intern uebernommen und ihre Pruefsumme
festgehalten werden (`maskieren.cmd --version`).

## Eigenschaften, die fuer eine Pruefung relevant sind

- kein Skript-Interpreter, kein Kindprozess, kein PowerShell, kein VBScript
- kein Netzwerkzugriff zur Laufzeit
- keine Installation, keine Administratorrechte, keine Registry-Aenderung
- Access-Datenbanken werden ausschliesslich **gelesen**
- Excel-Dateien werden standardmaessig **nicht** ueberschrieben
- die Ausgabe enthaelt Blatt- und Spaltennamen sowie Anzahlen, **keine
  Zellinhalte**

Nachpruefbar: siehe Abschnitt „Fuer die IT-Freigabe" in der README.
