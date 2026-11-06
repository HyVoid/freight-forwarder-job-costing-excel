[ 🌐 عربي ](README.ar.md) | [ 🇩🇪 Deutsch ](README.de.md) | [ 🇫🇷 Français ](README.fr.md) | [ 🇳🇱 Nederlands ](README.nl.md) | [ 🇪🇸 Español ](README.sp.md) | [ 🇬🇧 English ](README.md)

# Excel-Vorlage zur Kostenzuordnung für Speditionen: Sendungs-GuV-Rechner (Profit & Loss), Containerkosten-Verteilung & Margen-Tracker
<img width="2880" height="1228" alt="image" src="https://github.com/user-attachments/assets/f89c1a0f-f576-43af-b26a-02bae4a47348" />


[![License: Apache 2.0](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](LICENSE) [![Platform](https://img.shields.io/badge/Platform-Web%20Browser%20%2B%20Microsoft%20Excel%20(.xlsx)-217346.svg)](#) [![Tool Type](https://img.shields.io/badge/Tool%20Type-Plug--and--Play%20Freight%20Accounting%20Spreadsheet-2251FF.svg)](#)

**Eine einsatzbereite Excel-Vorlage zur Kostenzuordnung für Speditionen (`.xlsx`) und ein kostenloser Online-Rechner zur Sendungsprofitabilität — gebaut für Speditionen, NVOCCs, Zollagenten und 3PL-Betreiber, die den Nettogewinn pro Sendung berechnen müssen, ohne Formeln zu schreiben.** Verfolgen Sie **Bruttomargen von Konnossementen (B/L) und Luftfrachtbriefen (AWB)**, automatisieren Sie **FX-Umrechnungen zwischen mehreren Währungen (USD, EUR, CNY, GBP)**, splitten Sie **FCL/LCL-Containerkosten** und gleichen Sie **Kunden-AR-Rechnungen gegen AP-Rechnungen von Seefrachtführern, Drayage, Zollabfertigung und Banküberweisungen** in einer vorgefertigten Logistik-GuV-Tabelle ab.

**Testen Sie es jetzt kostenlos in Ihrem Browser oder laden Sie die entsperrte Excel-Datei herunter, um Ihre eigenen Sendungsdaten einzufügen.**

Jede Nachschlagregel, jede Währungsumrechnung und jede Formel zum Splitten von Containerkosten ist bereits eingebaut. Fügen Sie Ihre Abrechnungszeilen aus QuickBooks, Xero oder Ihrem betrieblichen Tracker ein und sehen Sie in Sekunden Ihren exakten Gewinn pro Buchung, Container und Frachtgebühren-Code — abgesichert durch eine 30-tägige Geld-zurück-Garantie.

> 🌐 **Zuerst den Online-Rechner testen** → [Starten Sie den kostenlosen interaktiven Web-Rechner für Kostenzuordnung bei Speditionen & Sendungs-GuV (Browser-Demo, keine Einrichtung erforderlich)](https://hyvoid.github.io/freight-forwarder-job-costing-excel/)  
> 📥 **Die einsatzbereite Excel-Vorlage herunterladen** → [Holen Sie sich die entsperrte Excel-Vorlage zur Kostenzuordnung & Sendungsprofitabilität für Speditionen (.xlsx mit automatisch aktualisierenden Formeln)](https://www.theseusworkshop.com/l/uyukd?utm_source=github&utm_medium=GitHub%20README&utm_campaign=readme%20new%20launch&utm_content=freight-booking-job-costing)  
>
> *Sowohl der kostenlose Web-Rechner als auch das herunterladbare Microsoft-Excel-Arbeitsbuch basieren auf exakt denselben vorkonfigurierten Frachtbuchhaltungs-Tabellen und automatisierten Margenformeln.*

---

## Was Sie damit verfolgen: Von Schmerzpunkten der Spedition zur Lösung

Zu warten auf Monatsabschlussberichte der Buchhaltung verschleiert, welche einzelnen Sendungen tatsächlich Gewinn oder Verlust gebracht haben. Diese **fertige Speditions-Kostentabelle** behebt die fünf häufigsten operativen Blindstellen der See-, Luft- und multimodalen Fracht:

| Häufiges Suchthema & Verfolgungsziel | Der tägliche Schmerzpunkt der Frachtbuchhaltung (warum Margen leaken) | Wie diese automatisierte Excel-Vorlage das behebt und verfolgt |
| :--- | :--- | :--- |
| **Gewinn pro Sendung & GuV auf Buchungsebene (`B/L`- & `AWB`-Marge)** | QuickBooks und Xero gruppieren Rechnungen nach Kalendermonat und verbergen dadurch Verlustsendungen, zu niedrige Angebote und unrentable Handelsrouten. | Ordnet **Kunden-AR-Umsatzerlöse, Lieferanten-AP-Rechnungen und Banküberweisungsgebühren** jeder einzelnen `Booking_ID` zu, damit Sie **Nettogewinn ($) und Bruttomarge (%) pro Sendung** in Echtzeit sehen. |
| **Profitabilität von Frachtgebühren-Codes & Nebenkosten** | Sie kennen den Rechnungsbetrag, können aber nicht sagen, ob Sie bei **Seefracht, lokalem Drayage, THC, Zollabfertigung, Demurrage oder Frachtversicherung** Geld verloren haben. | Schlüsselt Erlöse vs. Lieferantenkosten über granulare **Fracht-Service-Codes** auf (`Ocean Freight`, `THC`, `Origin Inspection`, `Drayage`, `Customs`, `Documentation`), um verdeckte Kostenüberschreitungen aufzudecken. |
| **Verteilung gemeinsamer Kosten bei mehreren FCL-Containern (`TEU` / `FEU`-Split)** | Deckt ein Master-Konnossement (MBL) mehrere Container ab, werden pauschale Buchungsgebühren (etwa Zollanmeldung oder B/L-Dokumentation) doppelt gezählt oder aus der Preisgestaltung pro Container weggelassen. | Trennt **direkte Containerkosten** von **gemeinsamen Gebühren auf Buchungsebene** und teilt gemeinsame Kosten automatisch gleichmäßig auf alle aktiven Container dieser Buchung auf. |
| **Umrechnung mehrwähriger Frachtrechnungen (`USD`, `EUR`, `CNY`, `GBP`)** | Seefrachtführer in USD zu bezahlen, Ursprungsagenten in CNY oder EUR und Kunden in lokaler Währung mit Taschenrechner-Kursen abzurechnen verursacht Wechselkursverluste zum Monatsende. | Rechnet jede gemischt-währige AR- und AP-Position über eine zentrale FX-Nachschlagetabelle in Ihre **einzige Basis-Berichtswährung** um und markiert jeden nicht konfigurierten Währungscode sofort. |
| **Fehlende Lieferantenrechnungen & Erkennung von Tippfehlern in Tabellen** | Ein einziger Tippfehler in einer Buchungs-ID in einer manuellen `SUMIFS`- oder `VLOOKUP`-Tabelle lässt stillschweigend eine Frachtführer-Rechnung über $2,000 auf `$0` fallen, bläht den Gewinn auf und zahlt zu hohe Verkaufsprovisionen. | Führt einen automatischen **Datenzustands-Check** aus (`ERR: UNKNOWN_CURRENCY`, `ERR: INVALID_BOOKING`), der defekte Zeilen sichtbar markiert und Ihre Dashboard-Summen intakt hält. |

---

## So berechnen Sie Sendungs-GuV und führen ein Kostenzuordnungs-Audit durch (4-Schritte-Einsteiger-Tutorial)

Sie benötigen keine Buchhaltungsausbildung oder Excel-Formelkenntnisse, um Ihre Sendungsmargen zu prüfen. Folgen Sie diesen vier Schritten, um rohe Rechnungslisten und Frachtführerrechnungen in eine präzise **GuV Buchung für Buchung** zu verwandeln:

### Schritt 1: Basiswährung, Wechselkurse & Container-Split-Regeln festlegen
Öffnen Sie den Tab `Setup_Params`, um die Grundregeln Ihres Unternehmens in Klartext festzulegen:
- Wählen Sie Ihre **Basis-Berichtswährung** (z. B. `USD`, `EUR` oder `GBP`) und tragen Sie Ihre standardmäßigen **Wechselkurse (FX)** ein.
- Prüfen Sie die vorab zugeordneten **Transaktionskategorien** (`Customer Revenue`, `Direct Vendor COGS`, `Bank Wire Fees` und `Customer Deposits`).
- Wählen Sie Ihre **Regel zur Verteilung gemeinsamer Kosten** für FCL-Sendungen mit mehreren Containern.

*Live ausprobieren:* Sehen Sie sich sofort in der [kostenlosen Browser-App zur Kostenzuordnung in der Spedition](https://hyvoid.github.io/freight-forwarder-job-costing-excel/) an, wie die Einrichtungstabelle funktioniert.

---

### Schritt 2: Buchungen, Container & rohe AR/AP-Rechnungszeilen einfügen
Tragen Sie Ihre Sendungs-IDs in die Stammlisten (`Bookings` und `Containers`) ein, dann kopieren und fügen Sie Ihre Abrechnungszeilen direkt in die Tabelle `Raw_Transactions` ein:
- **Funktioniert mit jeder Datenquelle:** QuickBooks Online-/Desktop-Exporte, Xero-CSV-Dumps, CargoWise- oder Magaya-Abrechnungsprotokolle, Frachtrechnungen von Seefrachtführern, Auszüge von Zollagenten und manuelle Dispatch-Tracker.
- **Kein manuelles Formelziehen:** Die Excel-Tabellen erweitern sich beim Einfügen neuer Zeilen automatisch.

---

### Schritt 3: Automatische Fehlerwarnungen prüfen & Ihr Sendungs-GuV-Dashboard ansehen
Sobald Sie Ihre Daten einfügen, berechnet die Tabelle alles im Hintergrund:
- Rechnet Fremdwährungs-Rechnungen von Lieferanten und Kunden in Ihre Basiswährung um.
- Markiert falsch geschriebene Buchungs-IDs oder fehlende Währungskurse in der Spalte `Data Health Status`, damit keine unverbundene Lieferantenrechnung als Nullkosten-Posten versteckt bleibt.
- Aktualisiert Ihre **GuV-Zusammenfassung auf Buchungsebene**, die **Margentabelle nach Fracht-Service-Codes** und die **Stückkosten-Aufschlüsselung pro Container**.

---

### Schritt 4: Die entsperrte Excel-Vorlage für die monatliche Wiederverwendung mit einem Klick herunterladen
Der kostenlose Browser-Rechner eignet sich hervorragend für eine schnelle Einzelprüfung. Wenn Sie Ihre Daten lokal speichern, Hunderte monatliche Buchhaltungszeilen importieren und laufende Schiffsabfahrten über Abrechnungszyklen hinweg verfolgen möchten, wechseln Sie zum herunterladbaren `.xlsx`-Arbeitsbuch.

> 💡 **Warum beschäftigte Spediteure die Excel-Datei herunterladen, statt ihr eigenes Tabellenblatt zu bauen:**  
> Mehrwährungs-Frachtabstimmung, das Splitten von Kosten mehrerer Container und das Erkennen verwaister Rechnungen bergen komplexe Buchhaltungslogik unter der Haube. **In dieser Excel-Vorlage ist jede Zuordnungsregel, jeder Nachschlagelink und jede sich automatisch erweiternde Formel bereits gebaut und getestet.**  
> 
> Sie müssen nicht einmal Zellecken ziehen, um Formeln nach unten zu kopieren. Laden Sie die Datei einfach auf Ihren Computer, fügen Sie Ihre Rohabrechnungsdaten ein und erhalten Sie sofort prüfungsfertige Gewinnzahlen. In der Hochsaison der Schifffahrt verbringen Sie Ihre Zeit damit, Befrachter-Angebote zu gewinnen, Schiffskapazitäten zu sichern und Fracht zu bewegen — während eine simple Copy-and-Paste-Vorlage Ihre Finanzanalyse in Sekunden erledigt.  
> 
> 👉 **[Laden Sie die entsperrte Excel-Vorlage zur Kostenzuordnung für Speditionen (.xlsx) herunter und automatisieren Sie Ihre Sendungs-GuV noch heute](https://www.theseusworkshop.com/l/uyukd?utm_source=github&utm_medium=GitHub%20README&utm_campaign=readme%20new%20launch&utm_content=freight-booking-job-costing)**

---

## Geschäftslogik & Methodik: 4 versteckte Gewinn-Blindstellen, die diese Excel-Vorlage für Sie löst

Die meisten Speditionen verlieren an Sendungen Geld nicht, weil sie den falschen Seefrachtsatz angeboten haben, sondern weil Standard-Buchhaltungssoftware und selbstgebaute Excel-Tabellen verbergen, wohin das Geld tatsächlich fließt. So schützt diese Vorlage Ihr Ergebnis in klaren Worten — ganz ohne formale Buchhaltungsausbildung:

### 1. Monatliche Buchhaltungssoftware verbirgt Verlustsendungen (Abgleich auf Buchungsebene vs. Monatssummen)
**Die objektive Tatsache:** Standard-Buchhaltungssoftware wie QuickBooks oder Xero organisiert Geld nach **Kalendermonaten**, während Spedition nach **Buchung, Konnossement (B/L) und Container** abläuft. Fakturieren Sie einem Kunden $8,000 im Oktober, während Seefrachtführer und Drayage-Spediteur Ihnen $7,500 im November in Rechnung stellen, erscheint Oktober künstlich profitabel und November als Verlust.  
**Ihr direkter Nutzen:** Diese Vorlage gruppiert jeden Dollar Erlöse, Lieferantenkosten und Bankgebühren nach `Booking_ID` — unabhängig davon, in welchem Monat die Rechnung eingeht. Sie sehen den echten Nettogewinn jeder einzelnen Sendung, bevor Sie die Preiskarte eines Kunden erneuern oder Verkaufsprovisionen auszahlen.

### 2. Pauschale Buchungsgebühren vs. Containerkosten verzerren Ihre Angebote (automatisches Splitten von FCL-Containerkosten)
**Die objektive Tatsache:** In der See- und multimodalen Fracht fallen manche Kosten einmal pro **Buchung** an (etwa eine Zollabfertigungsgebühr von $150 oder eine B/L-Dokumentationsgebühr), andere pro **Container** (etwa $450 Terminal-Handling-Gebühren oder lokaler Drayage pro 40'HC-Box). In einer einfachen Tabelle multiplizieren Operatoren die $150-Zollgebühr versehentlich über 5 Container (blähen Kostenschätzungen um $600 auf und verlieren das Angebot) oder vergessen, gemeinsame Buchungsgebühren bei der Gewinnberechnung pro Container einzubeziehen.  
**Ihr direkter Nutzen:** Die Vorlage trennt containerspezifische Kosten von buchungsweiten Gemeinkosten und teilt die gemeinsamen Gebühren gleichmäßig auf die exakte Anzahl der Container dieser Buchung auf:

```text
Single Container Total Cost = Direct Container Cost + (Shared Booking Costs ÷ Total Containers in Booking)

```

Sie erhalten exakte Stückkosten pro Container ohne manuelle Division.

### 3. Ein Tippfehler in einer selbstgebauten Tabelle fälscht stillschweigend hohen Gewinn (eingebauter Tippfehler- & Fehlkosten-Detektor)

**Die objektive Tatsache:** In einer Standard-Excel-Tabelle mit `SUMIFS` bewirkt ein einziger Tippfehler in einer Buchungsreferenz auf einer Drayage-Lieferantenrechnung über $1,500, dass Excel diese Rechnung komplett ignoriert und `$0` zu Ihren Kosten hinzufügt. Umgekehrt verwandelt ein defektes `VLOOKUP` Ihre gesamte Zusammenfassungstabelle in `#N/A`-Fehler. Sehen Sie, wie eine einzelne unverbundene Lieferantenrechnung eine echte Buchung verfälscht:

```text
Reported Booking Revenue:   $8,500
Reported Vendor Costs:      $5,200  <-- ($1,500 drayage bill skipped due to a Booking ID typo)
Apparent Net Profit:        $3,300  <-- (Overstated by $1,500)
Apparent Profit Margin:      38.8%  <-- (Fake margin; actual real margin is only 21.1%)

```

Wenn Sie dieser Zahl von 38.8% vertrauen, zahlen Sie Verkaufsprovisionen aus eigener Tasche und kalkulieren eine Route weiter, die kaum die Kosten deckt.

**Ihr direkter Nutzen:** Jede eingefügte Zeile durchläuft einen automatischen **Datenzustands-Check**. Enthält eine Zeile eine falsch geschriebene Buchungs-ID, einen ungültigen Gebühren-Code oder einen fehlenden Währungskurs, markiert die Vorlage sie klar (`ERR: INVALID_BOOKING` oder `ERR: UNKNOWN_CURRENCY`) und hält Ihre Zusammenfassungsmathematik sauber laufend. Sie entdecken fehlende Lieferantenrechnungen, bevor sie Geld kosten.

### 4. Fremdwährungen und Kundenvorschüsse verzerren die echte Bruttomarge (FX- & Deposit-Isolierung auf Zeilenebene)

**Die objektive Tatsache:** Grenzüberschreitende Sendungen kombinieren routinemäßig USD-Seefracht, CNY- oder EUR-Rechnungen von Ursprungsagenten, Banküberweisungsgebühren von Zwischenbanken und Kundenvorschüsse. Manuelle Währungsumrechnungen am Monatsende führen zu Wechselkursfehlern, und Kundenvorschüsse als verdiente Erlöse zu zählen lässt eine Sendung profitabel erscheinen, bevor die Lieferantenrechnungen überhaupt bezahlt sind.

**Ihr direkter Nutzen:** Die Vorlage rechnet jede Fremdwährungsrechnung auf individueller Zeilenebene in Ihre Basiswährung um und trennt **Kundenvorschüsse (gehaltenes Bargeld)** von **verdienten Frachterlösen und COGS**. Sie sehen echte Gewinnmarge und tatsächliche Cash-Position nebeneinander.

---

## Manuelle Fracht-Tabellen & Standard-Buchhaltungssoftware vs. diese automatisierte Excel-Vorlage zur Sendungs-Kostenzuordnung

| Operationaler Vergleichsfaktor | Manuelle Excel-Tabellen & generische Buchhaltung (nur QuickBooks / Xero) | Standardisierte Excel-Vorlage zur Kostenzuordnung für Speditionen (dieses Tool) |
| --- | --- | --- |
| **Gewinntransparenz pro Sendung (`MBL` / `HBL` / `AWB`)** | Vermischt alle Monatsrechnungen in Sachkonto-Buckets und verbirgt, welche Buchungen, Container oder Befrachter konkret Geld verloren. | Erzeugt automatisch eine **GuV-Tabelle Buchung für Buchung** mit Bruttoerlösen, Lieferanten-COGS, Bankgebühren, Nettogewinn ($) und Marge (%). |
| **Mehrwährige Rechnungsumrechnung (`USD`, `EUR`, `CNY`, `GBP`)** | Mitarbeiter rechnen Fremdwährungsrechnungen mit Tischrechnern und inkonsistenten Kursen um — unerklärliche Margenlücken beim Monatsabschluss sind die Folge. | Nutzt eine **zentrale FX-Kurs-Nachschlagetabelle** zur automatischen Umrechnung jeder Rechnungsposition und warnt sofort, wenn ein Währungskurs fehlt. |
| **Splitten von Kosten mehrerer FCL-Container (`20'GP` / `40'HC` / `TEU`)** | B/L-Dokumentengebühren, Zollabfertigung und Buchungszuschläge werden pro Container doppelt gezählt oder aus der Stückkostenberechnung weggelassen. | Trennt automatisch **direkte Containerrechnungen** von **gemeinsamem Buchungs-Overhead** und teilt gemeinsame Gebühren auf die exakte Containeranzahl pro Buchung auf. |
| **Nebenkosten- & Gebühren-Code-Marge (`THC`, `Drayage`, `Customs`)** | Sie sehen die Gesamtabrechnung der Sendung, aber nicht, ob **Ursprungsinspektion, Inlands-Drayage, Demurrage oder Frachtversicherung** Profit leaken. | Rollt Erlöse vs. Lieferantenkosten nach **Fracht-Service-Code** auf, damit Sie erkennen, welche Nebenkosten Geld bringen und welche verlieren. |
| **Formelpflege & Tippfehlerschutz (`VLOOKUP` / `SUMIFS`)** | Mitarbeiter müssen Formeln für jede neue Zeile manuell nach unten ziehen, und ein einziger Tippfehler in einer Buchungs-ID lässt Lieferantenkosten stillschweigend auf Null fallen. | **Vorgefertigte, sich automatisch erweiternde Tabellen & Fehlerwarnungen** aktualisieren sich beim Einfügen und quarantänen Tippfehler (`ERR: INVALID_BOOKING`), bevor sie Ihre GuV verunreinigen. |

---

## Für wen das ist: Logistik-Rollen, Tabellen-Anwendungsfälle & Vorlagen-Workflows

Dieses Toolkit ist gebaut für **kleine und mittlere Speditionen, OTI-/NVOCC-Betreiber, Zollagenten, 3PL-Teams und Frachtagenten**, die sofortige Profittransparenz auf Sendungsebene wollen, ohne $20,000+ und sechs Monate für ein Enterprise-ERP oder TMS zu investieren:

* **Speditionsinhaber & NVOCC-Geschäftsführer (`Excel-Vorlage für Speditions-GuV`):** Den echten Nettogewinn pro Sendung prüfen (nach Bankgebühren und buchungsweitem Gemeinkosten), bevor jährliche Befrachter-Preiskarten freigegeben oder monatliche Verkaufsprovisionen ausgezahlt werden.
* **Jüngere Fracht-Operations-Koordinatoren & Dispatcher (`fertige Sendungs-Kostentabelle`):** Mehrcontainer-FCL- und LCL-Jobkosten in der Hochsaison verfolgen, indem rohe Lieferantenrechnungen und Kundenpositionen in eine vorgefertigte Vorlage eingefügt werden — ganz ohne Formelschreiben.
* **Logistik-Buchhalter & Frachtbuchhaltungsangestellte (`Vorlage zur Fracht-AR/AP-Rechnungsabstimmung`):** Die Lücke zwischen QuickBooks-, Xero- oder NetSuite-Exporten und operativen Sendungsdateien schließen — unberechnete Kunden-Nebenkosten oder unverbundene Frachtführerrechnungen abfangen, bevor die Monatsbücher geschlossen werden.
* **Fracht-Preisanalysten & Handelsrouten-Manager (`Container-Stückkosten- & Nebenkosten-Margenrechner`):** Analysieren, ob Seefracht, Luftfracht, Zollabfertigung, Frachtversicherung oder lokaler Drayage die Margen auf bestimmten Ursprung-Ziel-Routen antreibt oder schrumpft.
* **Fractional CFOs & Logistikberater (`3PL- & Fracht-Margenleak-Audit-Modell`):** Kundenabrechnungsexporte in eine standardisierte Tabelle einwerfen, um zu niedrig kalkulierte Konten, Wechselkursverluste und nicht wiederverrechnete Nebengebühren an einem einzigen Nachmittag aufzudecken.

*Hinweis: Dieses Excel-Modell ergänzt Ihr betriebliches TMS (etwa CargoWise, Magaya oder GoFreight) und Ihre Steuerbuchhaltung. Es ist Ihre schnelle, fertige **Schicht für Sendungsprofitabilität und Job-Kostenzuordnung**, wenn Sie sofort klare Antworten aus Ihren Abrechnungsdaten brauchen.*

---

## Häufige Fracht-Kalkulationsszenarien: Wann Sie diese Sendungs-GuV-Tabelle nutzen

Ob Sie ein fünfköpfiges Speditionsbüro führen oder Hunderte FCL-Container pro Monat managen — diese Vorlage löst genau die alltäglichen Kalkulationsprobleme, die Logistikteams ausbremsen:

* **So berechnen Sie den Gewinn pro Sendung in Excel ohne Formeln:** Fügen Sie Kundenrechnungs- und Lieferantenrechnungszeilen in eine Tabelle ein, taggen Sie jede Zeile mit ihrer `Booking_ID` und lassen Sie das vorgefertigte Zusammenfassungsblatt Nettogewinn und Bruttomarge % automatisch berechnen.
* **So splitten Sie gemeinsame Konnossement-Kosten (B/L) auf mehrere FCL-Container:** Weisen Sie containerspezifische Gebühren (etwa Drayage oder THC) einzelnen `Container_ID`-Nummern zu und lassen Sie buchungsweite Gebühren (etwa Zollabfertigung oder Dokumentation) auf der `Booking_ID`-Ebene — die Vorlage teilt die gemeinsamen Kosten automatisch auf Ihre 20'GP- und 40'HC-Container auf.
* **So gleichen Sie verzögerte Seefrachtführer- & Drayage-Rechnungen gegen Kundenrechnungen ab:** Halten Sie offene Buchungen über Kalendermonate hinweg in Ihrem Log, damit späte Ankündigungsfristen, Detention-/Demurrage-Rechnungen oder Chassis-Split-Gebühren bei Eingang automatisch vom Gewinn der richtigen Sendung abgezogen werden.
* **So verfolgen Sie mehrwährige Frachtmargen über USD, EUR und CNY:** Erfassen Sie Ihre Referenzwechselkurse einmal in `Setup_Params`, um Rechnungen ausländischer Ursprungsagenten, Seefracht und lokale Zustellungsabrechnung in eine einzige Berichtswährung zu normieren.
* **So prüfen Sie Bruttogewinn-Provisionen von Frachtmaklern & Vertriebsmitarbeitern:** Verifizieren Sie, dass jede Lieferantenrechnung, jede Ursprungsgebühr und jede Bankgebühr verlinkt und erfasst ist, bevor Sie Verkaufsprovisionen auf den Sendungs-Bruttogewinn berechnen.

---

## Über Theseus Workshop

**Theseus Workshop** baut fertige Microsoft-Excel-Finanzmodelle und browserbasierte Rechner für Operatoren, die heute klare, verlässliche Zahlen brauchen — ohne Enterprise-Software-Ballast und ohne komplexe Tabellenprogrammierung.

Jede Vorlage ist um ein praktisches Ziel herum entworfen:

> **„Legen Sie die exakten operativen und finanziellen Zahlen in eine einzige vorgefertigte Tabelle, damit Ihre nächste Geschäftsentscheidung Sekunden statt Stunden dauert.“**

Wenn Sie aufhören wollen, mit kaputten `VLOOKUP`-Tabellen und manueller Währungsmathematik zu kämpfen, holen Sie sich die vollständige Excel-Vorlage unten oder erkunden Sie unsere komplette Bibliothek betrieblicher Tabellen und Business-Rechner:

* 📥 **Diese Fracht-Kalkulationsvorlage holen:** [Entsperrte Excel-Vorlage zur Kostenzuordnung & Sendungs-GuV für Speditionen herunterladen (.xlsx)](https://www.theseusworkshop.com/l/uyukd?utm_source=github&utm_medium=GitHub%20README&utm_campaign=readme%20new%20launch&utm_content=freight-booking-job-costing)
* 🌐 **Den kostenlosen Web-Rechner testen:** [Die interaktive Browser-App zur Kostenzuordnung in der Spedition öffnen](https://hyvoid.github.io/freight-forwarder-job-costing-excel/)
* 🧰 **Alle fertigen Excel-Modelle & Business-Toolkits erkunden:** [Den offiziellen Theseus Workshop Store auf TheseusWorkshop.com besuchen](https://www.google.com/search?q=https://www.theseusworkshop.com%3Futm_source%3Dgithub%26utm_medium%3DGitHub%2520README%26utm_campaign%3Dreadme%2520new%2520launch%26utm_content%3Dabout-section-store-link)

---

## Technische Details

<details>
<summary>Für technische Prüfer, Excel-Praktiker und Mitarbeiter</summary>

### Arbeitsmappen-Architektur

Das Arbeitsbuch ist um vier funktionale Schichten organisiert: **Regeln und Parameter → Stammdaten → Transaktionseingabe → Berechnung und analytische Ausgabe**.

Das Design enthält 11 Blätter:

| Schicht              | Blatt                   | Rolle                                                                             |
| ------------------ | ----------------------- | -------------------------------------------------------------------------------- |
| Regeln & Parameter | `Instructions`          | Betriebsanweisungen, Dateneingaberegeln und Geschäftsdefinitionen               |
| Regeln & Parameter | `Setup_Params`          | Basiswährung, FX-Kurse, Transaktionszuordnungen und Richtlinie zur Verteilung gemeinsamer Kosten |
| Stammdaten        | `Booking_Master`        | Buchungsindex, Kunde, Status und Containeranzahl                             |
| Stammdaten        | `Container_Master`      | 1:N-Beziehung zwischen Buchung und Container                                            |
| Stammdaten        | `Service_Master`        | Standard-Service-Code, Dienstleistungsname und Kategorie                                |
| Transaktionseingabe  | `Raw_Transactions`      | Einheitliches Finanztransaktionsregister plus Datenzustandsdiagnostik                |
| Berechnungs-Engine | `Job_Costing_Engine`    | Aggregation von Erlösen, COGS, Bankgebühren, Gewinn, Marge und Einlagen auf Buchungsebene  |
| Analyse           | `Booking_PL`            | GuV einer einzelnen Buchung und Drill-down auf Serviceebene                                  |
| Analyse           | `Container_Analysis`    | Direkte und zugeteilte gemeinsame Kosten auf Containerebene                                |
| Analyse           | `Service_Profitability` | Profitabilität nach Dienstleistung über Buchungen hinweg                                           |
| Managementausgabe   | `Executive_Dashboard`   | Globale KPIs, Profitabilitäts-Rankings und Datenzustandswarnungen                    |

Der Kernablauf ist:

```text
Setup_Params
      │
      ├──────────────┐
      ▼              ▼
Booking_Master   Service_Master
      │              │
      └──────┬───────┘
             ▼
     Raw_Transactions
             │
     Linkage + FX + Financial
        Treatment Diagnosis
             │
             ▼
    Job_Costing_Engine
             │
    ┌────────┼───────────┐
    ▼        ▼           ▼
Booking_PL  Container_  Service_
            Analysis    Profitability
             │
             └──────┬──────┘
                    ▼
          Executive_Dashboard
```

Die Quellarchitektur identifiziert drei wichtige mehrstufige Nachschlageketten: eine vierstufige Währungs-/Finanzbehandlungs-Kette, eine vierstufige Gemeinkosten-/Container-Drill-down-Kette und eine dreistufige Service-Klassifizierungskette. 

Die kritische Designregel lautet, dass **Berechnungswerte und Diagnosezustände getrennt bleiben**. `USD Equivalent` bleibt numerisch, während `Data Health Status` Fehler wie `ERR: UNKNOWN_CURRENCY` trägt. Dies bewahrt die `SUMIFS`-Kompatibilität und verhindert, dass defekte Transaktionen in die GuV-Engine gelangen. 

### Drei Fallen, die selbst erfahrene Speditionsoperatoren erwischen

#### Falle 1 — Einen fehlenden FX-Kurs als gültigen Betrag behandeln

**1. Entscheidung:** Eine Buchung scheint bestimmte Kosten und Marge zu haben.

**2. Verdeckter Fehler:** Eine Lieferantentransaktion ist in einer Währung denominiert, die in der Referenz-FX-Tabelle nicht konfiguriert ist.

**3. Auswirkung:** Wird der fehlende Kurs fälschlich als `1.0` behandelt, kann ein Fremdbetrag wie `50,000` zu `$50,000 USD` werden und das Buchungsergebnis radikal verzerren.

**4. Warum das falsch ist:** Ein fehlender Wechselkurs ist ein Datenintegritätsversagen, kein wirtschaftlicher Wechselkurs.

**5. Richtiger Ansatz:** Im Berechnungsfeld einen neutralen numerischen Wert von `0` zurückgeben und `ERR: UNKNOWN_CURRENCY` separat melden.

**6. Korrigierte Entscheidung:** Die Buchung wird zur Datenergänzung markiert, statt einen falschen Gewinn oder Verlust zugewiesen zu bekommen.

```text
Wrong:
Unknown Currency → FX = 1.0 → USD Equivalent → P&L

Correct:
Unknown Currency → FX = 0
                 → USD Equivalent = 0
                 → Data Health = ERR: UNKNOWN_CURRENCY
                 → P&L excludes invalid transaction
                 → Management fixes source data
```

<details>
<summary>Formellogik</summary>

```excel
=MAP(K2:K10000,L2:L10000,
 LAMBDA(curr,m_rate,
   IF(curr="",
      0,
      IF(m_rate>0,
         m_rate,
         XLOOKUP(
           curr,
           Setup_Params!$A$10:$A$20,
           Setup_Params!$B$10:$B$20,
           0
         )
      )
   )
 ))
```

Die wichtige Regel lautet: **Eine unbekannte Währung niemals standardmäßig auf 1.0 setzen**. 

</details>

#### Falle 2 — Die GuV einer Buchung lesen, bevor der Datenzustand geprüft ist

**1. Entscheidung:** Eine Buchung mit `$8,500` Erlösen und `$5,200` COGS gilt als `$3,300` Gewinn.

**2. Verdeckter Fehler:** Eine oder mehrere Quelltransaktionen haben ungültige Buchungs- oder Service-Verknüpfungen oder eine unaufgelöste Währung.

**3. Auswirkung:** Die scheinbare Marge kann überzeichnet sein, weil eine Kostenzeile stillschweigend nicht in die Aggregation gelangt ist.

**4. Warum das falsch ist:** Ein mathematisch korrektes `SUMIFS`-Ergebnis kann wirtschaftlich falsch sein, wenn die Quellzeilen unvollständig sind.

**5. Richtiger Ansatz:** Die Engine aggregiert nur Transaktionen mit `Data Health Status = "OK"` und meldet die Anzahl ungültiger Zeilen separat.

**6. Korrigierte Entscheidung:** Die Buchung gilt als **vorläufig**, bis die Datenausnahmen geklärt sind.

```text
Wrong:
Transactions → SUMIFS → Profit → Decision

Correct:
Transactions → Linkage / FX validation
             → Healthy rows only
             → Profit + Health Warning
             → Decision
```

<details>
<summary>Formellogik</summary>

```excel
rev =
SUMIFS(
  Raw_Transactions!O2:O10000,
  Raw_Transactions!B2:B10000,bk,
  Raw_Transactions!N2:N10000,"Revenue",
  Raw_Transactions!P2:P10000,"OK"
)
```

Derselbe Zustandsfilter wird auf COGS, Bankgebühren und Einlagen angewendet. 

</details>

#### Falle 3 — Gemeinsame Kosten auf Buchungsebene direkt jedem Container zuweisen

**1. Entscheidung:** Eine Buchung mit zwei Containern benötigt eine Profitabilitätsansicht auf Containerebene.

**2. Verdeckter Fehler:** Eine COGS-Zeile auf Buchungsebene gehört zur Sendung als Ganzes, nicht zu einem bestimmten Container.

**3. Auswirkung:** Die vollen gemeinsamen Kosten beiden Containern zuzuweisen verdoppelt die scheinbaren Kosten.

**4. Warum das falsch ist:** Die wirtschaftliche Zugehörigkeit der Kosten liegt auf Buchungsebene, während die Analyse eine kontrollierte Zuteilungsregel erfordert.

**5. Richtiger Ansatz:** Die konfigurierte Richtlinie `EQUAL_SHARE` verwenden und die gemeinsamen COGS durch die Containeranzahl der Buchung teilen.

**6. Korrigierte Entscheidung:** Jeder Container erhält seine direkten Kosten plus seinen zugeteilten Anteil an den Gemeinkosten, wodurch das Gesamtergebnis auf Buchungsebene erhalten bleibt.

```text
Booking shared COGS = $1,000
Container count     = 2

Allocated shared cost per container
= $1,000 / 2
= $500
```

Das Arbeitsbuch trennt direkte Containerkosten ausdrücklich von gemeinsamen Kosten auf Buchungsebene und verwendet einen sicheren Nenner aus der Containeranzahl. 

<details>
<summary>Formellogik</summary>

```excel
alloc_cogs =
IF(
  alloc_policy="EQUAL_SHARE",
  direct_cogs + (shared_cogs / cntr_count),
  direct_cogs
)
```

`Setup_Params!B6` steuert, ob die gemeinsamen Kosten per `EQUAL_SHARE` verteilt oder per `NO_ALLOCATE` unverteilt belassen werden. 

</details>

### Beispielszenario

Betrachten Sie eine Speditionsbuchung mit den folgenden gültigen Transaktionen:

| Position                      | Typ     | Originalbetrag | Währung | USD-Äquivalent |
| ------------------------- | -------- | --------------: | -------- | -------------: |
| Kundenrechnung          | Revenue  |           8,500 | USD      |         $8,500 |
| Seefracht-Lieferantenkosten | COGS     |           3,000 | USD      |         $3,000 |
| Ursprungsinspektion         | COGS     |             900 | USD      |           $900 |
| Frachtversicherung           | Revenue  |             400 | USD      |           $400 |
| Bankgebühr               | Bank Fee |              75 | USD      |            $75 |
| Rückerstattbare Einlage       | Deposit  |             500 | USD      |           $500 |

Die gültige GuV lautet:

```text
Revenue
= 8,500 + 400
= $8,900

COGS
= 3,000 + 900
= $3,900

Bank Fees
= $75

Net Profit
= $8,900 - $3,900 - $75
= $4,925

Profit Margin
= $4,925 / $8,900
= 55.34%
```

Die rückerstattbare Einlage von `$500` wird separat verfolgt statt als Betriebserlös oder COGS behandelt zu werden.

Nehmen Sie nun an, eine zusätzliche Lieferantentransaktion wird in einer nicht konfigurierten Währung erfasst. Das System fügt die Transaktion **nicht** stillschweigend zu COGS hinzu. Ihr angewandter FX-Kurs wird `0`, ihre USD-Berechnung bleibt numerisch, und das Diagnosefeld meldet `ERR: UNKNOWN_CURRENCY`. Die Buchung erhält eine Zustandswarnung, statt als sauber präsentiert zu werden.

Die betriebliche Empfehlung ändert sich entsprechend:

> **Verwenden Sie die Marge von 55.34% nicht als endgültige Profitabilitätszahl, bis die unaufgelöste Transaktion korrigiert ist.**

Dies ist der zentrale Unterschied zwischen einem Profitabilitätsrechner und einem kontrollierten Job-Costing-Workflow: Das Ergebnis ist von Belegen darüber begleitet, ob die zugrunde liegenden Daten vollständig genug sind, um ihnen zu vertrauen.

### Formelreferenz

<details>
<summary>Booking Master — Containeranzahl</summary>

**Zweck:** Bestimmen, wie viele Container jeder Buchung zugeordnet sind und Division-durch-Null-Fehler verhindern.

```excel
=MAP(
  A2:A10000,
  LAMBDA(
    bk,
    IF(
      bk="",
      "",
      LET(
        cnt,
        COUNTIF(Container_Master!A2:A10000,bk),
        IF(cnt=0,1,cnt)
      )
    )
  )
)
```

Ist noch kein Container erfasst, fällt der Nenner auf `1` zurück. Das ist eine Berechnungssicherung; betrieblich sollte der fehlende Container weiterhin als Datenqualitätsproblem behandelt werden. 

</details>

<details>
<summary>Raw Transactions — Angewandter FX-Kurs</summary>

**Zweck:** Einen manuellen FX-Kurs bevorzugen, andernfalls den konfigurierten Währungskurs abrufen.

```excel
=MAP(
  K2:K10000,
  L2:L10000,
  LAMBDA(
    curr,m_rate,
    IF(
      curr="",
      0,
      IF(
        m_rate>0,
        m_rate,
        XLOOKUP(
          curr,
          Setup_Params!$A$10:$A$20,
          Setup_Params!$B$10:$B$20,
          0
        )
      )
    )
  )
)
```

Eine unbekannte Währung liefert `0`, nicht `1.0`. 

</details>

<details>
<summary>Raw Transactions — Finanzielle Behandlung</summary>

**Zweck:** Betriebliche Transaktionstypen auf Finanzkategorien abbilden.

```excel
=MAP(
  H2:H10000,
  LAMBDA(
    type,
    IF(
      type="",
      "",
      XLOOKUP(
        type,
        Setup_Params!$E$10:$E$20,
        Setup_Params!$F$10:$F$20,
        "COGS"
      )
    )
  )
)
```

Das Quelldesign verwendet `COGS` als konservativen Fallback für einen unbekannten Transaktionstyp und verringert so das Risiko einer Gewinnüberzeichnung. 

</details>

<details>
<summary>Raw Transactions — USD-Äquivalent</summary>

**Zweck:** Jeden ursprünglichen Transaktionsbetrag in die USD-Basiswährung des Arbeitsbuchs umrechnen und dabei ein numerisches Ergebnis behalten.

```excel
=MAP(
  J2:J10000,
  M2:M10000,
  LAMBDA(
    amt,rate,
    IF(
      amt="",
      0,
      amt*rate
    )
  )
)
```

Diese Spalte numerisch zu halten ist wichtig, weil nachgelagerte `SUMIFS`-Berechnungen niemals Diagnosetext wie `"INVALID"` anstelle einer Zahl erhalten dürfen. 

</details>

<details>
<summary>Raw Transactions — Datenzustandsstatus</summary>

**Zweck:** Buchungs-, Service- und FX-Verknüpfung unabhängig von den finanziellen Berechnungsspalten validieren.

```excel
=MAP(
  B2:B10000,
  F2:F10000,
  K2:K10000,
  M2:M10000,
  LAMBDA(
    bk,srv,curr,rate,
    IF(
      bk="",
      "",
      IF(
        ISNA(XLOOKUP(
          bk,
          Booking_Master!A2:A10000,
          Booking_Master!A2:A10000
        )),
        "ERR: INVALID_BOOKING",
        IF(
          ISNA(XLOOKUP(
            srv,
            Service_Master!A2:A500,
            Service_Master!A2:A500
          )),
          "ERR: INVALID_SERVICE",
          IF(rate=0,
             "ERR: UNKNOWN_CURRENCY",
             "OK"
          )
        )
      )
    )
  )
)
```

Die Diagnoseebene ist absichtlich von den numerischen Berechnungsfeldern getrennt. 

</details>

<details>
<summary>Job Costing Engine — GuV auf Buchungsebene</summary>

**Zweck:** Eine dynamische Buchungsliste erzeugen und nur gesunde Transaktionen aggregieren.

```excel
=LET(
  bk_list,
  UNIQUE(
    FILTER(
      Booking_Master!A2:A10000,
      Booking_Master!A2:A10000<>""
    )
  ),
  BYROW(
    bk_list,
    LAMBDA(
      bk,
      LET(
        cust,
        XLOOKUP(
          bk,
          Booking_Master!A2:A10000,
          Booking_Master!C2:C10000,
          "Unknown"
        ),
        rev,
        SUMIFS(
          Raw_Transactions!O2:O10000,
          Raw_Transactions!B2:B10000,bk,
          Raw_Transactions!N2:N10000,"Revenue",
          Raw_Transactions!P2:P10000,"OK"
        ),
        cogs,
        SUMIFS(
          Raw_Transactions!O2:O10000,
          Raw_Transactions!B2:B10000,bk,
          Raw_Transactions!N2:N10000,"COGS",
          Raw_Transactions!P2:P10000,"OK"
        ),
        fees,
        SUMIFS(
          Raw_Transactions!O2:O10000,
          Raw_Transactions!B2:B10000,bk,
          Raw_Transactions!N2:N10000,"Bank Fee",
          Raw_Transactions!P2:P10000,"OK"
        ),
        profit,
        rev-cogs-fees,
        margin,
        IF(rev=0,0,profit/rev),
        bad_rows,
        COUNTIFS(
          Raw_Transactions!B2:B10000,bk,
          Raw_Transactions!P2:P10000,"<>OK"
        ),
        warn_msg,
        IF(
          bad_rows>0,
          "⚠️ "&bad_rows&" TXN DATA ERRORS",
          "CLEAN"
        ),
        HSTACK(
          bk,cust,rev,cogs,fees,
          profit,margin,
          dep,warn_msg
        )
      )
    )
  )
)
```

Die wesentliche Kontrolle ist das wiederholte Kriterium `Data Health Status = "OK"`. 

</details>

### Validierungsregeln

| Feld / Bedingung      | Regel                                                            | Fehlerverhalten                                                                     |
| ---------------------- | --------------------------------------------------------------- | ---------------------------------------------------------------------------------- |
| `Booking ID`           | Muss in `Booking_Master` existieren                                  | `ERR: INVALID_BOOKING`; Transaktion wird von der GuV ausgeschlossen                              |
| `Service Code`         | Muss in `Service_Master` existieren                                  | `ERR: INVALID_SERVICE`; Transaktion wird von der GuV ausgeschlossen                              |
| `Currency`             | Muss einen gültigen konfigurierten FX-Kurs oder einen positiven manuellen FX-Kurs haben | Angewandter FX wird `0`; Diagnose meldet `ERR: UNKNOWN_CURRENCY`                 |
| `Manual FX Rate`       | Positiver Wert überschreibt den Referenz-FX-Kurs                  | Ungültiger/nicht positiver Wert fällt auf den konfigurierten FX zurück                             |
| `Transaction Type`     | Sollte der `Setup_Params`-Zuordnung entsprechen                             | Unbekannter Typ fällt konservativ auf `COGS` zurück                                   |
| `USD Equivalent`       | Muss numerisch bleiben                                             | Diagnosetext darf niemals in das numerische Aggregationsfeld geschrieben werden           |
| `Data Health Status`   | Muss für die finanzielle Aggregation `OK` sein                          | Nicht-OK-Zeilen werden von der Buchungs-GuV ausgeschlossen                                          |
| `Container Count`      | Darf als Zuteilungsnenner nicht Null werden        | Fällt auf `1` zurück, um `#DIV/0!` zu verhindern                                             |
| Verteilung gemeinsamer Kosten | Gesteuert durch `Setup_Params!B6`                                 | `EQUAL_SHARE` verteilt nach Containeranzahl; `NO_ALLOCATE` lässt nur direkte Kosten |
| Buchungsprofitabilität  | Erlöse von Null dürfen keinen Divisionsfehler erzeugen               | Marge liefert `0`                                                                 |
| Management-Datenzustand  | Jeder nicht leere Status ungleich `OK` wird gezählt                 | Dashboard zeigt eine Warnung statt eines pauschal sauberen Zustands      |

Die Quell-Kreuzvalidierung bestätigt die Abdeckung der Kern-Eingaben für Buchung, Container, Service und Transaktion sowie der Ausgaben USD-Äquivalent, Nettogewinn, Marge und zugeteilte Kosten. Sie prüft zudem speziell die Trennung zwischen Diagnosetext und numerischen `SUMIFS`-Feldern. 

<!-- Technical Details parent <details> intentionally remains open for Part 2. -->
### Kreuzvalidierung & Einsatzbereitschaft

Die Arbeitsmappenarchitektur ist darauf ausgelegt, die analytische Kette zu validieren, **bevor** das Management sich auf die resultierende GuV verlässt.

Die Quellimplementierung prüft ausdrücklich die Transaktionstyp-Zuordnungskette, die FX-Nachschlagekette, die Service-Master-Verknüpfung, die Gemeinkosten-Zuteilungskette und die Kompatibilität zwischen Diagnosetext und numerischen Aggregationsfeldern. 

| Validierungsdimension         | Was geprüft wird                                                                  | Erwartetes Ergebnis                                          |
| ---------------------------- | -------------------------------------------------------------------------------- | -------------------------------------------------------- |
| **Buchungsverknüpfung**          | Jede `Booking ID` einer Transaktion existiert in `Booking_Master`                        | Ungültige Zeilen erhalten einen Fehlerzustand                      |
| **Service-Verknüpfung**          | Jeder `Service Code` einer Transaktion existiert in `Service_Master`                      | Ungültige Zeilen werden von der GuV ausgeschlossen                       |
| **FX-Abdeckung**              | Jede Transaktion hat entweder einen gültigen manuellen FX-Kurs oder einen konfigurierten Referenzkurs | Unbekannte Währungen lösen eine Zustandswarnung aus              |
| **Finanzielle Klassifizierung** | Transaktionstyp wird auf Revenue / COGS / Bank Fee / Deposit abgebildet                     | Unbekannte Typen nutzen den konservativen COGS-Fallback         |
| **Numerische Aggregation**      | `USD Equivalent` bleibt unter allen Bedingungen numerisch                            | `SUMIFS` trifft nie auf Diagnosetext              |
| **Zustandsgesteuerte Aggregation** | Die GuV enthält nur `Data Health Status = "OK"`                                    | Defekte Quellzeilen können nicht stillschweigend in den ausgewiesenen Gewinn gelangen |
| **Container-Zuteilung**     | Gemeinsame Kosten auf Buchungsebene nutzen die konfigurierte Zuteilungsrichtlinie                  | Keine wiederholte Vollkosten-Zuweisung über Container hinweg       |
| **Nullerlös-Schutz**  | Die Margenberechnung behandelt Revenue = 0                                    | Kein `#DIV/0!`                                             |
| **Management-Überwachung**     | Globale Transaktionsfehler werden separat gezählt                                 | Dashboard signalisiert ungeklärte Quelldatenprobleme          |

Die Kreuzvalidierungsmatrix der Implementierung besagt, dass die Kerneingabefelder — Buchung, Container, Service und Transaktionsdaten — und die wesentlichen Ausgaben — USD-Äquivalent, Nettogewinn, Marge und zugeteilte Kosten — von der Architektur abgedeckt sind. 

### Betrieblicher SOP

Das beabsichtigte Betriebsmodell ist bewusst einfach:

```text
1. Establish Master Data
        ↓
2. Record Financial Transactions
        ↓
3. Check Data Health Status
        ↓
4. Review Booking / Service / Container Profitability
        ↓
5. Resolve Exceptions
        ↓
6. Use Clean Results for Management Decisions
```

#### 1. Stammdaten anlegen

Legen Sie die Buchung in `Booking_Master` an.

Wenn Container relevant sind, legen Sie ihre Buchung-zu-Container-Beziehungen in `Container_Master` an.

Pflegen Sie den standardisierten Dienstleistungskatalog in `Service_Master`.

Beginnen Sie die Finanzanalyse nicht damit, freitextliche Dienstleistungsnamen oder Buchungskennungen direkt in Berichtsblätter zu tippen. Die Stammtabellen sind die Referenzebene für nachgelagerte Nachschlageketten. 

#### 2. Finanztransaktionen erfassen

Erfassen Sie Erlöse, Lieferantenkosten, Bankgebühren, Einlagen und andere relevante Finanzaktivitäten in `Raw_Transactions`.

Jede Transaktion sollte die jeweiligen Angaben enthalten:

* `Booking ID`
* `Container ID`, wenn zutreffend
* `Service Code`
* `Transaction Type`
* `Allocation Level`
* `Original Amount`
* `Currency`
* `Manual FX Rate`, wenn ein transaktionsspezifischer Kurs erforderlich ist

Die Transaktionstabelle ist die einzige Quelle, die detaillierte Finanzereignis-Datensätze erhalten soll. Die Berechnungs-Engine sollte nicht manuell bearbeitet werden. 

#### 3. Datenzustand prüfen

Prüfen Sie `Raw_Transactions[Data Health Status]`.

Eine saubere Zeile liefert:

```text
OK
```

Ausnahmen benennen die gebrochene Beziehung, etwa:

```text
ERR: INVALID_BOOKING
ERR: INVALID_SERVICE
ERR: UNKNOWN_CURRENCY
```

Die Implementierung behandelt dies bewusst als separate Kontrollebene. Eine Transaktion kann daher numerisch sicher bleiben und dennoch sichtbar vom Finanzergebnis ausgeschlossen sein. 

#### 4. Profitabilität überprüfen

Verwenden Sie:

* `Job_Costing_Engine` für die Profitabilität auf Buchungsebene.
* `Booking_PL` für die Aufschlüsselung einer Buchung Dienstleistung für Dienstleistung.
* `Container_Analysis`, wenn die Buchung mehrere Container enthält.
* `Service_Profitability` für die Service-Ökonomie über Buchungen hinweg.
* `Executive_Dashboard` für die Überwachung auf Managementebene.

Dies schafft einen Drill-down-Pfad von **globaler Profitabilität → Buchung → Dienstleistung → Container → Transaktion**, statt jede Frage aus einem einzigen Übersichts-Dashboard beantworten zu müssen.

### Regeln zur Datenpflege

Die Architektur stützt sich auf kontrollierte Stammdaten, weil Excel keine ebenso harte Fremdschlüssel-Durchsetzung bietet wie eine Datenbank- oder ERP-Plattform.

Die wichtigsten Pflegeregeln sind:

1. **Überschreiben Sie keine formelgenerierten Felder.**
2. **Tippen Sie keinen Dienstleistungsnamen von Hand ein, wo ein Service-Code stehen sollte.**
3. **Führen Sie keine neue Währung ein, ohne vorher ihren FX-Kurs zu `Setup_Params` hinzuzufügen.**
4. **Verwenden Sie `1.0` nicht als Fallback für fehlende FX-Kurse.**
5. **Schreiben Sie keinen Fehlertext in numerische Aggregationsspalten.**
6. **Bearbeiten Sie die Ausgabe von `Job_Costing_Engine` nicht manuell.**
7. **Klären Sie Nicht-OK-Transaktionszeilen, bevor Sie eine Buchungsmarge als endgültig behandeln.**
8. **Halten Sie die Richtlinie zur Verteilung gemeinsamer Kosten zentral in `Setup_Params`.**

Das Quelldesign untersagt speziell das Hartkodieren globaler Parameter in Formeln und zentralisiert Basiswährung, FX-Kurse, Transaktionszuordnungen und Zuteilungsrichtlinie in `Setup_Params`. 

### Fehlerbehebung

#### `⚠️ X TXN DATA ERRORS` erscheint auf dem Dashboard

**Wahrscheinliche Ursachen:**

* Die Währung fehlt in der FX-Tabelle.
* Die Buchungs-ID existiert nicht in `Booking_Master`.
* Der Service-Code existiert nicht in `Service_Master`.
* Eine Quelltransaktion enthält eine ungültige Verknüpfung.

**Lösung:**

Filtern Sie die Spalte `Data Health Status` in `Raw_Transactions`, um Werte ungleich `OK` anzuzeigen.

Lösen Sie dann jedes zugrunde liegende Quelldatenproblem, statt die GuV-Ausgabe manuell zu ändern.

Das Executive Dashboard ist darauf ausgelegt, diesen Zustand über eine globale Zustandskarte offenzulegen:

```excel
=LET(
    total_errs,
    COUNTIFS(
        Raw_Transactions!P2:P10000,
        "<>OK",
        Raw_Transactions!P2:P10000,
        "<>"
    ),
    IF(
        total_errs=0,
        "✅ ALL DATA HEALTHY",
        "⚠️ ATTENTION: " &
        total_errs &
        " ERRORS IN RAW TRANSACTIONS"
    )
)
```



#### Dynamische Array-Formeln geben `#SPILL!` zurück

Eine Überlauf-Formel benötigt einen freien Zielbereich.

Prüfen Sie:

* Zellen direkt unterhalb der Formel;
* Zellen rechts neben der Formel;
* manuell eingegebene Werte, die den beabsichtigten Überlaufbereich belegen;
* Restformatierungen oder kopierte Inhalte, die ein Hindernis gebildet haben.

Leeren Sie die hindernden Zellen und lassen Sie das dynamische Array erneut expandieren.

#### Profitabilität erscheint unerwartet hoch

Ändern Sie die Gewinnformel nicht sofort.

Prüfen Sie zuerst:

```text
Raw Transactions
       ↓
Data Health Status
       ↓
Unknown Currency?
Invalid Booking?
Invalid Service?
       ↓
Only OK rows enter P&L
```

Eine verdächtig hohe Marge kann durch eine Kostentransaktion verursacht werden, die die Verknüpfungskette nicht bestanden und daher aus der Berechnung ausgeschlossen wurde.

Genau deshalb unterscheidet das Arbeitsbuch zwischen **„Null als numerischem Berechnungszustand“** und **„Null als wirtschaftlicher Schlussfolgerung“**. Ein fehlender FX-Kurs kann `$0.00` in einem numerischen Feld erzeugen und gleichzeitig `ERR: UNKNOWN_CURRENCY` im Diagnosefeld. 

#### Container-Profitabilität erscheint zu hoch oder zu niedrig

Prüfen Sie drei Dinge der Reihe nach:

1. Ist die richtige `Booking ID` ausgewählt?
2. Ist die richtige `Container ID` ausgewählt?
3. Verwendet `Setup_Params!B6` die beabsichtigte Zuteilungsrichtlinie?

Unter `EQUAL_SHARE` werden die gemeinsamen COGS auf Buchungsebene durch die Containeranzahl der Buchung geteilt, bevor sie zu den direkten Containerkosten addiert werden.

Unter `NO_ALLOCATE` werden gemeinsame Kosten auf Buchungsebene keinen einzelnen Containern zugewiesen.

Der Unterschied ist wesentlich, weil die Zuweisung der vollen buchungsweiten Gemeinkosten an jeden Container dieselbe wirtschaftliche Kostenposition duplizieren würde. 

### Grenzen der Implementierung

Dieses Arbeitsbuch ist bewusst als **leichte analytische Kontrollebene** konzipiert, nicht als Ersatz für eine Buchhaltungsplattform oder ein Transportmanagementsystem.

Es ist am stärksten, wenn:

* betriebliche und finanzielle Daten bereits in Tabellen oder Exporten vorliegen;
* die zentrale Analyseeinheit eine Buchung oder Sendung ist;
* das Management Job-Level-Profitabilität benötigt;
* Dienstleistungen und Gebühren über Buchungen hinweg verglichen werden müssen;
* mehrere Währungen eine konsistente Berichtsbasis benötigen;
* gemeinsame Kosten eine explizite Zuteilungsregel erfordern;
* Nutzer transparente Formeln statt opaker Automatisierung brauchen.

Es ist nicht dafür ausgelegt, Folgendes zu bieten:

* mehrbenutzer-fähige transaktionelle Parallelität;
* ERP-taugliche Berechtigungen;
* automatisierten Bankabgleich;
* automatisierte Frachtführer-Rechnungserfassung;
* Fremdschlüssel-Durchsetzung auf Datenbankebene;
* Enterprise-Workflow-Orchestrierung.

Der Wert des Arbeitsbuchs liegt daher nicht in der Anzahl seiner Blätter, sondern in der kontrollierten analytischen Kette, die **Buchung → Transaktion → Klassifizierung → Währungsumrechnung → validierte Aggregation → Profitabilität → Drill-down** verbindet.

### Analytisches Kontrollprinzip

Die wichtigste Implementierungsregel lässt sich so zusammenfassen:

```text
Never let a broken data relationship
look like a valid financial result.
```

Eine konventionelle Tabelle könnte Folgendes zulassen:

```text
Missing FX
    ↓
0 USD
    ↓
SUMIFS
    ↓
Higher Profit
```

Dieses Arbeitsbuch macht den Zustand stattdessen explizit:

```text
Missing FX
    ↓
0 USD Equivalent
    +
ERR: UNKNOWN_CURRENCY
    ↓
Excluded from P&L
    +
Booking Health Warning
    ↓
Correct the source data
    ↓
Recalculate
```

Dieser Unterschied ist die Grundlage des Zuverlässigkeitsmodells der Job Costing Engine. Die Architektur erkennt ausdrücklich, dass `SUMIFS` numerische Werte sicher aggregieren kann, während Text stillschweigend ignoriert wird — Diagnosezustände dürfen daher niemals direkt in numerische Berechnungsspalten eingebettet werden. 

### Empfohlene Einführungsreihenfolge

Für eine neue Arbeitsmappen-Implementierung ist die sicherste Reihenfolge:

```text
Phase 1
Setup_Params
    ↓
Define Base Currency
Define FX Rates
Define Transaction Types
Define Financial Treatments
Define Allocation Policy

Phase 2
Booking_Master
Container_Master
Service_Master
    ↓
Establish Reference Data

Phase 3
Raw_Transactions
    ↓
Load Historical / Current Financial Data

Phase 4
Data Health Review
    ↓
Resolve Invalid Bookings
Resolve Invalid Services
Resolve Unknown Currencies

Phase 5
Job_Costing_Engine
    ↓
Validate Revenue / COGS / Fees / Profit

Phase 6
Booking_PL
Container_Analysis
Service_Profitability
Executive_Dashboard
    ↓
Operational Review
```

Das Ergebnis ist ein wiederholbarer Ablauf statt einer einmaligen Tabellenberechnung: Legen Sie die Referenzdaten einmal an, hängen Sie neue betriebliche Transaktionen an, klären Sie Ausnahmen und lassen Sie die Berechnungs- und Analyseebenen aus demselben zugrunde liegenden Datenmodell aktualisieren. 

<!-- Part 2 ends with the Technical Details parent <details> intentionally OPEN. Do not add </details> here. -->
### Abschließende Architektur-Hinweise

Das Arbeitsbuch ist bewusst um **eine einzige Profitabilitäts-Engine auf Buchungsebene** herum gebaut, nicht um eine Sammlung unverbundener Rechner.

Die analytische Hierarchie ist:

```text
Booking
 ├── Customer
 ├── Containers
 │    └── Container-level direct + allocated costs
 └── Transactions
      ├── Revenue
      ├── COGS
      ├── Bank Fees
      └── Deposits
           ↓
      Service-level profitability
           ↓
      Booking-level P&L
           ↓
      Executive-level profitability & data health
```

Diese Struktur hält die Buchung als zentrale wirtschaftliche Einheit und erlaubt Nutzern zugleich, nach unten in Dienstleistungen und Container oder nach oben in die buchungsübergreifende Profitabilität zu bohren. 

Die Implementierung erkennt zudem ausdrücklich drei mehrstufige Nachschlageketten:

1. **Transaktion → Finanzielle Behandlung → FX → USD-Äquivalent → Buchungs-GuV**
2. **Buchung → Containeranzahl → Gemeinkosten-Richtlinie → Verteilung gemeinsamer Kosten → Containeranalyse**
3. **Service-Code → Service Master → Service-Kategorie → Service-Profitabilität**

Diese Ketten werden validiert, bevor die analytischen Ausgaben als produktionsreif gelten. 

### Wovor die Engine schützt

Das zentrale Designprinzip ist nicht bloß Automatisierung. Es geht darum, **zu verhindern, dass eine gebrochene Eingabebeziehung sich als gültige finanzielle Schlussfolgerung ausgibt**.

| Fehlerfall                                              | Schutzmechanismus                                                           |
| ---------------------------------------------------- | ------------------------------------------------------------------------------ |
| Unbekannte Währung                                     | FX fällt auf `0` zurück, während ein separater Diagnosezustand den fehlenden Kurs meldet |
| Unbekannter Transaktionstyp                             | Konservativer `COGS`-Fallback verhindert versehentliche Gewinnaufblähung              |
| Ungültige Buchung                                      | `ERR: INVALID_BOOKING`                                                         |
| Ungültiger Service-Code                                 | `ERR: INVALID_SERVICE`                                                         |
| Gebrochene Transaktionskette                             | Nicht-`OK`-Zeilen werden von der GuV-Aggregation ausgeschlossen                                |
| Gemeinkosten über Container dupliziert             | Explizite Zuteilungsrichtlinie und Nenner aus der Containeranzahl                     |
| Nullerlös                                         | Margenberechnung liefert `0` statt `#DIV/0!`                            |
| Diagnosetext verunreinigt Finanzberechnungen | Numerische Berechnungsspalten und Diagnosespalten bleiben getrennt             |

Die Fallback-Matrix der Implementierung weist einen `1.0`-FX-Fallback speziell zurück, weil er einen unbekannten Fremdwährungsbetrag in einen unplausibel großen USD-Betrag verwandeln könnte. 

### Empfohlenes Nutzungsmuster

Für den laufenden Betrieb soll sich das Arbeitsbuch wie ein **aktualisierbares Betriebsmodell** verhalten, nicht wie ein einmaliger Bericht.

```text
Maintain Setup_Params
        ↓
Maintain Master Data
        ↓
Append New Transactions
        ↓
Resolve Data Health Exceptions
        ↓
Review Booking Profitability
        ↓
Drill Into Service / Container
        ↓
Use Clean Results for Decisions
```

Das macht das Arbeitsbuch nützlich für wiederkehrende Fragen wie:

* Welche Buchungen sind tatsächlich profitabel?
* Welche Dienstleistungen erzeugen die stärksten Margen?
* Welche Sendungskosten treiben einen unerwarteten Verlust?
* Welche Kosten auf Buchungsebene müssen auf Container verteilt werden?
* Sind die ausgewiesenen Margen vertrauenswürdig, oder verbergen ungeklärte Quelldatenprobleme Kosten?
* Wo sollten sich Preisgestaltung, Lieferantenverhandlung oder betriebliche Überprüfung als Nächstes konzentrieren?

Der Quell-SOP folgt derselben Reihenfolge: Stammdaten anlegen, Finanztransaktionen erfassen, `Data Health Status` prüfen und dann die Profitabilitäts-Dashboards zur Überprüfung nutzen. 

### Abschließendes Designprinzip

Das Arbeitsbuch versucht nicht, die Finanzanalyse einfacher erscheinen zu lassen, indem es ihre Unsicherheit verbirgt.

Stattdessen macht es die Unsicherheit sichtbar.

Eine Buchung mit berechneter Marge wird nicht automatisch als vertrauenswürdige Buchung behandelt.

Eine **saubere Buchung** bedeutet:

```text
Valid Booking
+ Valid Service
+ Valid Currency / FX
+ Valid Transaction Classification
+ Healthy Transaction Chain
= Profitability Result That Can Be Reviewed
```

Dieser Unterschied ist der Kern der Implementierung.

</details>

## Weitere Tools dieser Serie

Eine Sammlung leichter Excel-Entscheidungsunterstützungs-Tools für Betriebssteuerung, Profitabilitätsanalyse, Bestandsabstimmung, Personalplanung und andere Geschäftsabläufe.

* **[rolling-cashflow-forecast](https://github.com/HyVoid/rolling-cashflow-forecast)** — Warum ist Cash immer knapp? — Zukünftige Liquidität unter wechselnden operativen und finanziellen Annahmen prognostizieren.
* **[Multi-Currency-Trade-Financial-Analyzer](https://github.com/HyVoid/Multi-Currency-Trade-Financial-Analyzer)** — Welche Handelsströme sind nach Währungseffekten tatsächlich profitabel? — Mehrwährige Handelsprofitabilität über grenzüberschreitende Operationen messen.
* **[RTL_Loan_Sizer_Pricer](https://github.com/HyVoid/RTL_Loan_Sizer_Pricer)** — Welche Finanzierungsstruktur ist unter verschiedenen Kreditnehmer- und Sicherheitsszenarien tragfähig? — Kreditszenarien mit risikoadjustierten Finanzierungsannahmen vergleichen.
* **[Personal-Business-Dual-Track-Budget-Manager](https://github.com/HyVoid/Personal-Business-Dual-Track-Budget-Manager)** — Wo vermischt oder fehlallokiert sich privates und geschäftliches Cash? — Private und geschäftliche Cashflows in unabhängig verfolgte Budget-Tracks trennen.
* **[Multifamily-Value-Add-Acquisition-Model](https://github.com/HyVoid/Multifamily-Value-Add-Acquisition-Model)** — Schafft diese Immobilienakquisition nach Sanierungs- und Betriebsannahmen noch Eigenkapitalwert? — Mehrfamilien-Akquisitionsrenditen unter Sanierungsfinanzierung und betrieblichen Szenarien bewerten.
* **[Real-Estate-Valuation-all-in-one-calculator](https://github.com/HyVoid/Real-Estate-Valuation-all-in-one-calculator)** — Was ist eine Immobilie unter verschiedenen Bewertungsmethoden wert? — Immobilienwert über mehrere Bewertungsmethoden mit Steuer- und Capex-Modellierung berechnen.
* **[multi-entity-expense-allocation](https://github.com/HyVoid/multi-entity-expense-allocation)** — Werden gemeinsame Kosten fair über Gesellschaften verteilt, oder werden manche Assets überlastet, während andere ihrer wahren Kostenexposure entkommen? — Lieferantenrechnungen über mehrere Rechtsträger und Betriebsassets mit konfigurierbaren Zuteilungsmethoden verteilen.
* **[investment-pool-nav-excel-template](https://github.com/HyVoid/investment-pool-nav-excel-template)** — Was schuldet man jedem Teilnehmer tatsächlich, basierend auf seinem Einstiegszeitpunkt und dem aktuellen Poolwert? — Investmentpool-Anteile kontinuierlich über NAV-basierte Stückrechnung verfolgen.
* **[shopify-cash-flow-forecast-excel-toolkit](https://github.com/HyVoid/shopify-cash-flow-forecast-excel-toolkit)** — Welche operativen Verpflichtungen erzeugen in den nächsten zwölf Wochen ein Cash-Problem? — Shopify-Verkäufe, Banktransaktionen, Werbeausgaben und Inventareinkäufe in einen einheitlichen Cashflow-Entscheidungsrahmen verbinden.
* **[Revenue-Management-Decision-Engine](https://github.com/HyVoid/Revenue-Management-Decision-Engine)** — Hat sich genug substanzielle Information geändert, um eine Preisanpassung zu rechtfertigen? — Preisempfehlungen über konfigurierbare Leitplanken steuern, die unnötige Umsatzvolatilität verhindern.
* **[amazon-wbr-sqp-analytics-template](https://github.com/HyVoid/amazon-wbr-sqp-analytics-template)** — Was hat sich im Geschäft in dieser Periode tatsächlich geändert und welche Produkte oder Keywords trieben die Änderung? — Amazon Seller Central SQP- und Sellerboard-Exporte in standardisierte Executive Business Reviews umwandeln.
* **[food-manufacturing-cost-calculator-excel](https://github.com/HyVoid/food-manufacturing-cost-calculator-excel)** — Welche Produkte sind nach vollständiger Verteilung aller Herstellungskosten tatsächlich profitabel und welche zerstören die Marge? — Fabrikarbeit, Nebenkosten und Materialkosten auf jede SKU verteilen für echte Stückkostenökonomie.
* **[ecommerce-profit-calculator-excel](https://github.com/HyVoid/ecommerce-profit-calculator-excel)** — Welche Bestellungen sind nach allen versteckten Kosten tatsächlich profitabel und wo leakt Geld über Kanäle und Währungen? — Echte Bestell-Level-Profitabilität rekonstruieren durch Abgleich von Auszahlungen, Versandkosten, Zahlungsgebühren und Marktplatzprovisionen über mehrkanalige E-Commerce-Operationen.
* **[construction-job-costing-excel-template](https://github.com/HyVoid/construction-job-costing-excel-template)** — Welche Bauunternehmer-Projekte und Kunden sind nach allen Jobkosten tatsächlich profitabel? — Lead-Kalkulation, Jobausführung und Ist-Kostenverfolgung in einen einheitlichen Rahmen für Bauunternehmer-Profitabilität verbinden.
* **[Flora-Pastry-Coffee-Shop-Operating-Control-System](https://github.com/HyVoid/Flora-Pastry-Coffee-Shop-Operating-Control-System)** — Ist der Laden nach allen Zutaten-, Arbeits- und Betriebskosten wirklich profitabel? — Rezeptkalkulation, Verkauf, Inventarverbrauch und Betriebsausgaben in einen einheitlichen Rahmen für Profitabilität und Cashflow verbinden.
* **[cross-entity-logistics-console](https://github.com/HyVoid/cross-entity-logistics-console)** — Welche Container und Kunden sind über Zoll- und Flottenoperationen hinweg tatsächlich profitabel? — Zollabfertigungs- und Flottenausführungsdaten zu einer einzigen Container-Level-Profitabilitätssicht konsolidieren.
* **[family-loan-tracker-excel](https://github.com/HyVoid/family-loan-tracker-excel)** — Was ist nach Jahren von Transfers und aufgelaufenen Zinsen heute tatsächlich geschuldet? — Langfristige Geldtransfers und Zinsansammlungen zu einem aktuellen Offensaldo abstimmen.
* **[3-statement-financial-model-excel](https://github.com/HyVoid/3-statement-financial-model-excel)** — Wie werden unsere Finanzberichte unter verschiedenen Betriebs- und Finanzierungsszenarien aussehen? — Integrierte Finanzberichte prognostizieren, indem Erlöse, Aufwendungen, Working Capital und Schulden in ein einheitliches Modell verlinkt werden.
* **[assisted-living-cost-allocation-excel](https://github.com/HyVoid/assisted-living-cost-allocation-excel)** — Wie viel der gemeldeten Marge jedes Gebäudes erzeugt tatsächlich dieses Gebäude, und wie viel die Art, wie Zentralkosten verteilt wurden? — Gemeinkosten über Betriebseinheiten verteilen, Quelltransaktionen erhalten und Einheiten-Ergebnisse mit Gesamtsummen abstimmen.
* **[real-estate-development-feasibility-excel](https://github.com/HyVoid/real-estate-development-feasibility-excel)** — Bei welchem maximalen Grunderwerbspreis wird dieses Entwicklungsprojekt zu einer bankfähigen Investition? — Wohnungsbau-Entwicklungschancen bewerten, indem Landvergleiche gescreent, Baukosten geschätzt und Profit-on-Cost-Schwellen berechnet werden.
* **[uk-residential-costing-software](https://github.com/HyVoid/uk-residential-costing-software)** — Erreicht dieses Bauprojekt seine Zielmarge wirklich, wenn Standortlogistik, Ausführungsfaktoren und Gewerkerate-Änderungen eingerechnet werden? — Physische Gebäudemengen von volatilen Gewerkeraten und Baustellenvorhaltung trennen, um kaufmännische Aufschläge über Ausschreibungsbudgets und Kundenangebote zu kaskadieren.
* **[b2b-pricing-landed-cost-excel-model](https://github.com/HyVoid/b2b-pricing-landed-cost-excel-model)** — Welche Großhandelspreise und Kundensegmente bewahren die Marge wirklich, wenn Fracht, Zölle, Zahlungsbedingungen und FX vollständig kalkuliert sind? — SKU-Level-Landed-Costs, Cost-to-Serve und Pocket-Margin-Preiswasserfälle berechnen, um grenzüberschreitende Großhandelsprofitabilität zu simulieren.
* **[custom-label-pricing-calculator](https://github.com/HyVoid/custom-label-pricing-calculator)** — Was sollten wir für einen Custom-Label-Auftrag je Mengenstaffel verlangen, um wettbewerbsfähig und profitabel zu bleiben? — Parametrische Label-Preise kalkulieren, indem Substrat-Quadratzoll-Kosten, Rüstgebühren und Mengenstaffel-Rabatte gegen Wettbewerber-Benchmarks modelliert werden.
* **[landed-cost-calculator-retail-deduction-excel](https://github.com/HyVoid/landed-cost-calculator-retail-deduction-excel)** — Wie hoch ist die echte Nettomarge pro SKU, wenn tatsächliche Landed Costs und Handelsabzüge eingerechnet werden? — PO-geschätzte Landed Costs gegen tatsächliche Fracht-, Zoll-, FX- und Drayage-Gebühren abstimmen und Handelsabzüge prüfen, um SKU-Level-Nettomarge offenzulegen.
* **[saas-finance-mrr-burn-rate-excel](https://github.com/HyVoid/saas-finance-mrr-burn-rate-excel)** — Können wir die Einstellungs-, Marketing- und Lieferantenverpflichtungen dieser Woche eingehen, ohne vor dem Bücherabschluss das Cash auszugeben? — Wochen-Cash-Burn-Takt, MRR/ARR-Wasserfall-Abweichung, Retention-Kennzahlen und Lohnkosten gegen Budget verfolgen.
* **[startup-financial-model-excel](https://github.com/HyVoid/startup-financial-model-excel)** — Wird unser Wachstumsplan das Cash aufbrauchen, bevor die Kohortenmonetarisierung aufholt? — Mehrjährigen Cash Burn, Runway, Unit Economics und Fundraising-Dilution aus Bottom-up-Kohorten- und Headcount-Annahmen prognostizieren.
* **[shopify-pnl-reconciliation-excel](https://github.com/HyVoid/shopify-pnl-reconciliation-excel)** — Ist jeder Shopify-Store tatsächlich profitabel, wenn echte SKU-Level-Kosten, Gemeinkosten und Bankgebühren abgestimmt sind? — Multi-Store-Shopify-Verkaufsexporte gegen SKU-Level-COGS, zugerechnete mehrwährige OPEX und Bankauszüge abstimmen für eine vertrauenswürdige Monats-GuV.
* **[travel-agency-accounting-excel-template](https://github.com/HyVoid/travel-agency-accounting-excel-template)** — Welche Tourbuchungen erzeugen tatsächlich Nettogewinn und positiven Cashflow nach Lieferantenkosten und B2B-Einzügen? — Tour-Jobkosten pro Buchung, Bruttomargen, AR/AP-Abwicklung und operativen Cashflow über mehrere Währungen berechnen.

## Lizenz

Dieses Projekt wird unter der **Apache License 2.0** veröffentlicht.

Sie dürfen das Werk gemäß den Bedingungen der Apache License 2.0 nutzen, verändern, verbreiten und darauf aufbauen.
