# RC Snackbar eSG – Click & Collect

Online-Bestellplattform für die Schülergenossenschaft **RC Snackbar eSG** am Raiffeisen-Campus. 

Schülerinnen und Schüler können hier Snacks und Getränke vorbestellen und in der gewählten Pause direkt am Stand bezahlen und abholen.

---

## 🚀 Funktionen
* **Produktübersicht**: Aktuelle Snacks, Getränke und Sonderaktionen im Überblick.
* **Vorbestellung**: Auswahl der Abholpause (1. oder 2. Pause) sowie Eingabe von Name und Klasse.
* **Echtzeit-Synchronisation**: Bestellungen werden automatisch in ein Google Sheet übertragen, damit das Team die Tüten packen kann.

---

## 🛠️ Technologien
* **Frontend**: HTML5, CSS3, JavaScript (hosted via GitHub Pages)
* **Backend**: Google Apps Script & Google Sheets

---

## 📌 Einrichtung & Verwaltung für das Team

### 1. Produkte anpassen
* Die Produkte und Preise werden direkt im verbundenen Google Sheet im Tabellenblatt `Produkte` gepflegt.
* Neue Artikel oder Preisänderungen werden nach dem Speichern im Sheet automatisch auf der Website angezeigt.

### 2. Bestellungen abrufen
* Eingegangene Bestellungen erscheinen in Echtzeit im Blatt `Bestellungen`.
