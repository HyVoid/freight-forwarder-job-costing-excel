[ 🌐 عربي ](README.ar.md) | [ 🇩🇪 Deutsch ](README.de.md) | [ 🇫🇷 Français ](README.fr.md) | [ 🇳🇱 Nederlands ](README.nl.md) | [ 🇪🇸 Español ](README.sp.md) | [ 🇬🇧 English ](README.md)

# Excel-sjabloon voor opdrachtkostenberekening bij expéditeurs: winst-en-verliescalculator per zending, kostenverdeling over containers & margemonitor
<img width="2880" height="1228" alt="image" src="https://github.com/user-attachments/assets/f89c1a0f-f576-43af-b26a-02bae4a47348" />


[![License: Apache 2.0](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](LICENSE) [![Platform](https://img.shields.io/badge/Platform-Web%20Browser%20%2B%20Microsoft%20Excel%20(.xlsx)-217346.svg)](#) [![Tool Type](https://img.shields.io/badge/Tool%20Type-Plug--and--Play%20Freight%20Accounting%20Spreadsheet-2251FF.svg)](#)

**Een kant-en-klaar Excel-sjabloon voor opdrachtkostenberekening bij expéditeurs (`.xlsx`) en een gratis online calculator voor de winstgevendheid van zendingen, gebouwd voor expéditeurs, NVOCC's, douanemakelaars en 3PL-exploitanten die de nettowinst per zending moeten berekenen zonder formules te schrijven.** Volg de **brutomarges van cognossementen (B/L) en luchtvrachtbrieven (AWB)**, automatiseer **FX-omrekeningen tussen meerdere valuta (USD, EUR, CNY, GBP)**, splits de **containerkosten van FCL/LCL** en reken **klantfacturen (AR) af tegen leveranciersfacturen (AP) van oceaanvervoerders, drayage, douaneafhandeling en bankoverschrijvingen** in één kant-en-klaar logistiek winst-en-verlosspreadsheet.

**Test het nu gratis in uw browser, of download het ontgrendelde Excel-bestand om uw eigen zendingengegevens te plakken.**

Elke opzoekregel, valutaconversie en formule voor het splitsen van containerkosten zit al ingebouwd. Plak uw factureringsregels uit QuickBooks, Xero of uw operationele tracking-sheet en zie binnen enkele seconden uw exacte winst per boeking, container en vrachtcode, met een geld-terug-garantie van 30 dagen.

> 🌐 **Probeer eerst de online calculator** → [Start de gratis interactieve webcalculator voor opdrachtkostenberekening en winst-en-verliesrekening van zendingen voor expéditeurs (browserdemo, nul installatie vereist)](https://hyvoid.github.io/freight-forwarder-job-costing-excel/)  
> 📥 **Download het kant-en-klare Excel-sjabloon** → [Ontvang het ontgrendelde Excel-sjabloon voor opdrachtkostenberekening en winstgevendheid van zendingen voor expéditeurs (.xlsx met automatisch bijwerkende formules)](https://www.theseusworkshop.com/l/uyukd?utm_source=github&utm_medium=GitHub%20README&utm_campaign=readme%20new%20launch&utm_content=freight-booking-job-costing)  
>
> *Zowel de gratis webcalculator als het downloadbare Microsoft Excel-werkboek draaien op exact dezelfde vooraf geconfigureerde vrachtboekhoudtabellen en geautomatiseerde margeformules.*

---

## Waar dit u bij helpt: afbeelding van pijnpunten in het goederenvervoer op oplossingen

Wachten op maandafsluitingsrapporten van de boekhouding verbergt welke individuele zendingen echt geld hebben verdiend of verloren. Dit **plug-and-play spreadsheet voor opdrachtkostenberekening in het vervoer** lost de vijf meest voorkomende operationele boekhoudkundige blinde vlekken op in oceaanvracht, luchtvracht en multimodaal vervoer:

| Veelgezocht onderwerp & volgdoel | Het dagelijkse pijnpunt in de vrachtboekhouding (waarom marge weg lekt) | Hoe dit geautomatiseerde Excel-sjabloon het oplost en volgt |
| :--- | :--- | :--- |
| **Winst per zending & winst-en-verliesrekening op boekingniveau (`B/L` & `AWB`-marge)** | QuickBooks en Xero groeperen facturen per kalendermaand, waardoor verlieslatende zendingen, te laag geprijsde offertes en onrendabele handelsroutes verborgen blijven. | Koppelt **klantomzet (AR), leveranciersfacturen (AP) en bankoverschrijvingskosten** aan elke individuele `Booking_ID` zodat u in realtime de **nettowinst ($) en brutomarge (%) per zending** ziet. |
| **Winstgevendheid per vrachtkosten- en nevenkostensoort** | U kent het totale factuurbedrag, maar kunt niet zeggen of u geld verloor aan **oceaanvracht, lokale drayage, THC, douaneafhandeling, demurrage of cargoverzekering**. | Splitst omzet vs. leverancierskosten uit over gedetailleerde **vrachtdienstcodes** (`Ocean Freight`, `THC`, `Origin Inspection`, `Drayage`, `Customs`, `Documentation`) om verborgen kostenoverschrijdingen bloot te leggen. |
| **Verdeling van gedeelde kosten over meerdere FCL-containers (`TEU` / `FEU`-splitsing)** | Wanneer één master-cognossement (MBL) meerdere containers dekt, worden vaste boekingskosten (zoals douaneaangifte of B/L-documentatie) dubbel geteld of weggelaten uit de prijzen per container. | Scheidt **directe kosten per container** van **gedeelde kosten op boekingniveau** en verdeelt de gedeelde kosten automatisch gelijkmatig over alle actieve containers in die boeking. |
| **Omrekening van vrachtfacturen in meerdere valuta (`USD`, `EUR`, `CNY`, `GBP`)** | Oceaanvervoerders betalen in USD, herkomstagenten in CNY of EUR, en klanten factureren in lokale valuta met handmatig opgezochte wisselkoersen veroorzaakt maandelijkse wisselkoersverliezen. | Converteert elke gemengde AR- en AP-regel naar uw **enkele basisrapportagevaluta** via een centrale FX-opzoektabel en signaleert elke niet-geconfigureerde valutacode onmiddellijk. |
| **Detectie van ontbrekende leveranciersfacturen & typefouten in spreadsheets** | Eén typefout in een boeking-ID in een handmatige `SUMIFS`- of `VLOOKUP`-sheet laat een vervoersfactuur van $2,000 stilzwijgend naar `$0` vallen, waardoor de winst opbloeit en er te veel verkoopcommissies worden uitbetaald. | Voert een automatische **datagezondheidscontrole** uit (`ERR: UNKNOWN_CURRENCY`, `ERR: INVALID_BOOKING`) die defecte regels zichtbaar markeert terwijl uw dashboards totalen intact blijven. |

---

## Hoe berekent u de winst van een zending en voert u een opdrachtkostenaudit uit (beginnershandleiding in 4 stappen)

U hebt geen boekhoudopleiding of Excel-formulevaardigheden nodig om uw zendingsmarges te auditen. Volg deze vier stappen om ruwe factuurlijsten en vervoerdersfacturen om te zetten in een nauwkeurige **winst-en-verliesrekening per boeking**:

### Stap 1: stel uw basisvaluta, wisselkoersen & containersplitsingsbeleid in
Open het tabblad `Setup_Params` om de basisregels van uw bedrijf in gewone taal vast te leggen:
- Kies uw **basisrapportagevaluta** (zoals `USD`, `EUR` of `GBP`) en vul uw standaard **wisselkoersen (FX)** in.
- Controleer de vooraf toegewezen **transactiecategorieën** (`Customer Revenue`, `Direct Vendor COGS`, `Bank Wire Fees` en `Customer Deposits`).
- Selecteer uw **regel voor verdeling van gedeelde kosten** voor FCL-zendingen met meerdere containers.

*Probeer het live:* bekijk nu hoe de instellingentabel werkt in de [gratis online browsercalculator voor opdrachtkostenberekening in het goederenvervoer](https://hyvoid.github.io/freight-forwarder-job-costing-excel/).

---

### Stap 2: plak uw boekingen, containers & ruwe AR/AP-factuurregels
Zet uw zendings-ID's in de stamlijsten (`Bookings` en `Containers`), en kopieer en plak daarna uw factureringsregels rechtstreeks in de tabel `Raw_Transactions`:
- **Werkt met elke gegevensbron:** QuickBooks Online / Desktop-exports, Xero CSV-dumps, CargoWise- of Magaya-factureringslogboeken, vrachtfacturen van oceaanvervoerders, afschriften van douanemakelaars en handmatige expeditieoverzichten.
- **Nul handmatig formules slepen:** de Excel-tabellen worden automatisch uitgebreid wanneer u nieuwe regels plakt.

---

### Stap 3: controleer automatische foutwaarschuwingen & bekijk uw winst-en-verliesdashboard voor zendingen
Zodra u uw gegevens plakt, berekent het spreadsheet alles op de achtergrond:
- Converteert leveranciersfacturen en klantfacturen in vreemde valuta naar uw basisvaluta.
- Markeert elke verkeerd gespelde boeking-ID of ontbrekende valutakoers in de kolom `Data Health Status`, zodat een niet-gekoppelde leveranciersfactuur nooit als nulkostenpost kan wegmuggen.
- Werkt uw **winst-en-verliesoverzicht op boekingniveau**, uw **margetabel per vrachtdienstcode** en uw **uitsplitsing van eenheidskosten per container** bij.

---

### Stap 4: download het ontgrendelde Excel-sjabloon voor hergebruik per maand met één klik
De gratis browsercalculator is ideaal voor een snelle controle in één sessie. Wanneer u uw gegevens lokaal wilt opslaan, honderden maandelijkse boekhoudregels wilt importeren en lopende schijfvertrekken over factureringscycli wilt volgen, schakelt u over op het downloadbare `.xlsx`-werkboek.

> 💡 **Waarom drukke expéditeurs het Excel-bestand downloaden in plaats van hun eigen sheet te bouwen:**  
> Afstemming van vracht in meerdere valuta, het splitsen van kosten over meerdere containers en het opsporen van weesfacturen omvatten complexe boekhoudlogica onder de motorkap. **In dit Excel-sjabloon is elke koppelingsregel, opzoeklink en automatisch uitbreidbare formule al voor u gebouwd en getest.**  
> 
> U hoeft zelfs niet celhoeken te klikken en slepen om formules naar beneden te kopiëren. Download gewoon het bestand op uw computer, plak uw ruwe factuurgegevens en ontvang direct auditklare winstcijfers. Tijdens de piek van het vrachtseizoen besteedt u uw tijd aan het winnen van verzenderoffertes, het veiligstellen van scheepsruimte en het verplaatsen van cargo—terwijl een eenvoudig kopieer-en-plak-sjabloon uw financiële analyse in enkele seconden regelt.  
> 
> 👉 **[Download het ontgrendelde Excel-sjabloon voor opdrachtkostenberekening bij expéditeurs (.xlsx) en automatiseer vandaag nog uw winst-en-verliesrekening per zending](https://www.theseusworkshop.com/l/uyukd?utm_source=github&utm_medium=GitHub%20README&utm_campaign=readme%20new%20launch&utm_content=freight-booking-job-costing)**

---

## De bedrijfslogica & methodiek: 4 verborgen winstblinde vlekken die dit Excel-sjabloon voor u oplost

De meeste expéditeurs verliezen geld op zendingen niet omdat ze het verkeerde oceaantarief hebben geofferteerd, maar omdat standaard boekhoudsoftware en zelfgemaakte Excel-sheets verbergen waar het geld werkelijk naartoe gaat. Hier is hoe dit sjabloon uw nettoresultaat beschermt, in gewone taal—zelfs zonder enige formele boekhoudachtergrond:

### 1. Maandelijkse boekhoudsoftware verbergt verlieslatende zendingen (koppeling op boekingniveau vs. maandtotalen)
**Het objectieve feit:** standaard boekhoudsoftware zoals QuickBooks of Xero organiseert geld per **kalendermaand**, terwijl goederenvervoer draait om **boekingen, cognossementen (B/L) en containers**. Als u een klant in oktober $8,000 factureert, maar de oceaanvervoerder en de drayage-trucker u in november $7,500 factureren, dan lijkt oktober kunstmatig winstgevend en november als verlies.  
**Uw directe voordeel:** dit sjabloon groepeert elke dollar aan omzet, leverancierskosten en bankoverschrijvingskosten per `Booking_ID`, ongeacht in welke maand de factuur binnenkomt. U ziet de echte nettowinst van elke afzonderlijke zending voordat u het tariefkaart van een klant verlengt of verkoopcommissies uitbetaalt.

### 2. Vaste boekingskosten vs. kosten per container verstoren uw offertes (automatische splitsing van FCL-containerkosten)
**Het objectieve feit:** in oceaan- en multimodaal vracht ontstaan sommige kosten één keer per **boeking** (zoals een douaneafhandelingskosten van $150 of een B/L-documentatiekosten), terwijl andere kosten per **container** ontstaan (zoals $450 aan terminalafhandelingskosten of lokale drayage per 40'HC-container). In een eenvoudig spreadsheet vermenigvuldigen operators de $150 douanekosten per ongeluk over 5 containers (waardoor kostenramingen met $600 opbloeien en de aanbesteding verloren gaat) of vergeten ze gedeelde boekingskosten mee te tellen bij het berekenen van de winst per container.  
**Uw directe voordeel:** het sjabloon scheidt containerspecifieke kosten van gedeelde kosten op boekingniveau en verdeelt de gedeelde kosten gelijkmatig over het exacte aantal containers in die boeking:

```text
Single Container Total Cost = Direct Container Cost + (Shared Booking Costs ÷ Total Containers in Booking)

```

U krijgt exacte eenheidskosten per container zonder handmatig te delen.

### 3. Eén typefout in een zelfgemaakt spreadsheet smeedt stilzwijgend een hoge vals winst (ingebouwde detectie van typefouten en ontbrekende kosten)

**Het objectieve feit:** in een standaard Excel-sheet gebouwd met `SUMIFS` veroorzaakt één typefout in een boekingsreferentienummer op een drayage-leveranciersfactuur van $1,500 dat Excel die factuur volledig negeert en `$0` aan uw kosten toevoegt. Omgekeerd verandert een kapotte `VLOOKUP` uw hele overzichtstabel in `#N/A`-fouten. Kijk hoe één niet-gekoppelde leveranciersfactuur een echte boeking verstoort:

```text
Reported Booking Revenue:   $8,500
Reported Vendor Costs:      $5,200  <-- ($1,500 drayage bill skipped due to a Booking ID typo)
Apparent Net Profit:        $3,300  <-- (Overstated by $1,500)
Apparent Profit Margin:      38.8%  <-- (Fake margin; actual real margin is only 21.1%)

```

Als u dat cijfer van 38.8% vertrouwt, betaalt u verkoopcommissies uit eigen zak en blijft u een route quoteren die nauwelijks quitatie oplevert.

**Uw directe voordeel:** elke regel die u plakt, gaat door een automatische **datagezondheidscontrole**. Als een regel een verkeerd gespelde boeking-ID, een ongeldige kostensoort of een ontbrekende valutakoers bevat, markeert het sjabloon die duidelijk (`ERR: INVALID_BOOKING` of `ERR: UNKNOWN_CURRENCY`) terwijl uw overzichtsberekeningen netjes doorlopen. U vangt ontbrekende leveranciersfacturen voordat ze u geld kosten.

### 4. Vreemde valuta en voorafbetaalde klantaanbetalingen verstoren de echte brutomarge (FX op regelniveau & isolatie van aanbetalingen)

**Het objectieve feit:** grensoverschrijdende zendingen combineren doorgaans oceaanvracht in USD, facturen van herkomstagenten in CNY of EUR, bankoverschrijvingskosten van tussenpersonen en voorafbetaalde klantaanbetalingen. Valutatotalen handmatig omrekenen aan het einde van de maand introduceert wisselkoersfouten, en voorafbetaalde klantaanbetalingen als verdiende omzet tellen laat een zending winstgevend lijken voordat de leveranciersfacturen zelfs maar zijn betaald.

**Uw directe voordeel:** het sjabloon converteert elke vreemdvalutafactuur naar uw basisvaluta op regelniveau en scheidt **klantaanbetalingen (vastgehouden contant geld)** van **verdiende vrachtomzet en COGS**. U ziet uw echte winstmarge en uw werkelijke kaspositie naast elkaar.

---

## Handmatige vrachtspreadsheets & generieke boekhoudsoftware vs. dit geautomatiseerde Excel-sjabloon voor opdrachtkostenberekening van zendingen

| Operationele vergelijkingsfactor | Handmatige Excel-sheets & generieke boekhouding (alleen QuickBooks / Xero) | Gestandaardiseerd Excel-sjabloon voor opdrachtkostenberekening bij expéditeurs (deze tool) |
| --- | --- | --- |
| **Zichtbaarheid van winst per zending (`MBL` / `HBL` / `AWB`)** | Mengt alle maandfacturen in grootboekcategorieën, waardoor verborgen blijft welke specifieke boekingen, containers of verzenders geld verliezen. | Genereert automatisch een **winst-en-verliestabel per boeking** met brutoomzet, leveranciers-COGS, bankoverschrijvingskosten, nettowinst ($) en marge (%). |
| **Omrekening van facturen in meerdere valuta (`USD`, `EUR`, `CNY`, `GBP`)** | Medewerkers rekenen facturen van buitenlandse agenten om met bureaucalculators en inconsistente koersen, wat onverklaarbare margekloven veroorzaakt bij de maandafsluiting. | Gebruikt één **gecentraliseerde FX-koersopzoektabel** om elke factuurregel automatisch te converteren en toont direct een waarschuwing als een valutakoers ontbreekt. |
| **Kostensplitsing over meerdere FCL-containers (`20'GP` / `40'HC` / `TEU`)** | B/L-documentatiekosten, douanebedeling en boekingstoeslagen worden per container gedupliceerd of weggelaten uit de eenheidskostenberekeningen van containers. | Scheidt automatisch **directe containerfacturen** van **gedeelde overhead op boekingniveau** en verdeelt de gedeelde kosten over het exacte aantal containers per boeking. |
| **Nevenkosten & marge per kostensoort (`THC`, `Drayage`, `Customs`)** | U ziet de totale facturering van de zending, maar kunt niet zeggen of **herkomstinspectie, binnenlandse drayage, demurrage of cargoverzekering** winst laat lekken. | Rollt omzet vs. leverancierskosten op per **vrachtdienstcode** zodat u ziet welke nevenkosten geld opleveren en welke geld kosten. |
| **Formuleonderhoud & typefoutenbescherming (`VLOOKUP` / `SUMIFS`)** | Medewerkers moeten formules handmatig naar beneden slepen voor elke nieuwe regel, en één verkeerd gespelde boeking-ID laat leverancierskosten stilzwijgend naar nul vallen. | **Vooraf gebouwde automatisch uitbreidende tabellen & foutmeldingen** werken automatisch bij bij het plakken en zetten typefouten in quarantaine (`ERR: INVALID_BOOKING`) voordat ze uw P&L corrumperen. |

---

## Voor wie dit is: logistiekrollen, spreadsheet-gebruikscenario's & sjabloonworkflows

Deze toolkit is gebouwd voor **kleine en middelgrote expéditeurs, OTI/NVOCC-exploitanten, douanemakelaars, 3PL-teams en vrachtagenten** die onmiddellijke winstzichtbaarheid op zendingsniveau willen zonder $20,000+ en zes maanden te besteden aan het opzetten van een enterprise ERP of TMS:

* **Eigenaren van expediteurbedrijven & managing directors van NVOCC's (`Freight Forwarder Profit & Loss Excel Template`):** audit de echte nettowinst per zending (na bankoverschrijvingskosten en gedeelde overhead op boekingniveau) voordat u jaarlijkse verzendertariefkaarten goedkeurt of maandelijkse verkoopcommissies uitbetaalt.
* **Junior coördinatoren vrachtoperaties & dispatchers (`Plug-and-Play Shipment Costing Spreadsheet`):** volg de opdrachtkosten van FCL- en LCL-transporten met meerdere containers tijdens drukke vrachtseizoenen door ruwe leveranciersfacturen en klantkosten te plakken in een kant-en-klaar sjabloon, zonder één formule te schrijven.
* **Boekhouders logistiek & vrachtadministrateurs (`Freight AR/AP Invoice Reconciliation Template`):** overbrug de kloof tussen QuickBooks-, Xero- of NetSuite-exports en operationele zendingsdossiers—en vang ongefactureerde klantnevenkosten of niet-gekoppelde vervoerdersfacturen op voordat de maandelijkse boeken worden gesloten.
* **Tariefanalisten vracht & managers handelsroutes (`Container Unit Cost & Accessorial Margin Calculator`):** analyseer of Oceaanvracht, Luchtvracht, Douanebedeling, Cargoverzekering of Lokale drayage de marges op specifieke herkomst-bestemmingsroutes aandrijft of krimpt.
* **Deeltijd-CFO's & logistiekconsultants (`3PL & Freight Margin Leakage Audit Model`):** zet factureringsexports van klanten in een gestandaardiseerd spreadsheet om in één middag ondergeprijsde accounts, wisselkoersverliezen en onherstelde nevenkosten op te sporen.

*Let op: dit Excel-model vult uw operationele TMS (zoals CargoWise, Magaya of GoFreight) en uw fiscale boekhoudsoftware aan. Het fungeert als uw snelle plug-and-play **laag voor zendingswinstgevendheid en opdrachtkostenberekening** wanneer u meteen heldere antwoorden uit uw factureringsgegevens nodig hebt.*

---

## Veelvoorkomende scenario's voor vrachtkostenberekening: wanneer dit winst-en-verlosspreadsheet voor zendingen gebruiken

Of u nu een expeditiekantoor van vijf personen runt of honderden maandelijkse FCL-containers beheert, dit sjabloon lost de exacte dagelijkse kostenproblemen op die logistiketeams vertragen:

* **Hoe bereken ik de winst per zending in Excel zonder formules te schrijven:** plak uw klantfactuurregels en leveranciersfactuurregels in één tabel, label elke regel met zijn `Booking_ID` en laat het vooraf gebouwde overzichtsblad automatisch de nettowinst en brutomarge % berekenen.
* **Hoe verdeel ik de gedeelde kosten van het cognossement (B/L) over meerdere FCL-containers:** wijs containerspecifieke kosten (zoals drayage of THC) toe aan individuele `Container_ID`-nummers en laat kosten op boekingniveau (zoals douaneafhandeling of documentatie) op het niveau van het `Booking_ID`—het sjabloon verdeelt de gedeelde kosten automatisch over uw 20'GP- en 40'HC-containers.
* **Hoe stem ik vertraagde facturen van oceaanvervoerders en drayage af tegen klantfacturen:** houd openstaande boekingen in uw logboek over kalendermaanden heen zodat late aankomstmeldingen, demurrage/detentiefacturen of chassis-splitskosten automatisch worden afgetrokken van de winstmarge van de juiste zending zodra ze binnenkomen.
* **Hoe volg ik vrachtmarges in meerdere valuta over USD, EUR en CNY:** voer uw referentiewisselkoersen één keer in bij `Setup_Params` om facturen van overzeese herkomstagenten, oceaanvervoerdersvracht en lokale leveringsfacturering te normaliseren naar één rapportagevaluta.
* **Hoe audit ik de brutowinstcommissies van vrachtbemiddelaars en verkoopvertegenwoordigers:** verifieer dat elke leveranciersfactuur, herkomstkostenpost en bankoverschrijving gekoppeld en verwerkt is voordat u verkoopcommissies uitbetaalt op de brutowinst van zendingen.

---

## Over Theseus Workshop

**Theseus Workshop** bouwt kant-en-klare financiële modellen in Microsoft Excel en browsergebaseerde calculators voor operators die vandaag duidelijke, betrouwbare cijfers nodig hebben—zonder opgeblazen enterprisesoftware of complexe spreadsheetprogrammering.

Elk sjabloon is ontworpen rond één praktisch doel:

> **"Zet de exacte operationele en financiële cijfers in één kant-en-klare tabel zodat uw volgende bedrijfsbeslissing seconden in plaats van uren kost."**

Als u wilt stoppen met worstelen met kapotte `VLOOKUP`-sheets en handmatige valutaberekeningen, download dan hieronder het volledige Excel-sjabloon of verken onze complete bibliotheek met operationele spreadsheets en zakelijke calculators:

* 📥 **Dit vrachtkostensjabloon ophalen:** [Download het ontgrendelde Excel-sjabloon voor opdrachtkostenberekening en winst-en-verliesrekening van zendingen voor expéditeurs (.xlsx)](https://www.theseusworkshop.com/l/uyukd?utm_source=github&utm_medium=GitHub%20README&utm_campaign=readme%20new%20launch&utm_content=freight-booking-job-costing)
* 🌐 **De gratis webcalculator testen:** [Open de interactieve browserapp voor opdrachtkostenberekening in het goederenvervoer](https://hyvoid.github.io/freight-forwarder-job-costing-excel/)
* 🧰 **Alle kant-en-klare Excel-modellen & zakelijke toolkits verkennen:** [Bezoek de officiële Theseus Workshop-winkel op TheseusWorkshop.com](https://www.google.com/search?q=https://www.theseusworkshop.com%3Futm_source%3Dgithub%26utm_medium%3DGitHub%2520README%26utm_campaign%3Dreadme%2520new%2520launch%26utm_content%3Dabout-section-store-link)

---

## Technische details

<details>
<summary>Voor technische reviewers, Excel-practitioners en medewerkers</summary>

### Architectuur van het werkboek

Het werkboek is georganiseerd rond vier functionele lagen: **regels en parameters → stamdata → transactie-invoer → berekening en analytische uitvoer**.

Het ontwerp bevat 11 bladen:

| Laag              | Blad                   | Rol                                                                             |
| ------------------ | ----------------------- | -------------------------------------------------------------------------------- |
| Regels & parameters | `Instructions`          | Bedieningsinstructies, regels voor gegevensinvoer en bedrijfsdefinities               |
| Regels & parameters | `Setup_Params`          | Basisvaluta, wisselkoersen, transactietoewijzingen en beleid voor verdeling van gedeelde kosten |
| Stamdata        | `Booking_Master`        | Boekingsindex, klant, status en aantal containers                             |
| Stamdata        | `Container_Master`      | 1:N-relatie tussen boeking en container                                            |
| Stamdata        | `Service_Master`        | Standaard servicecode, servicenaam en categorie                                |
| Transactie-invoer  | `Raw_Transactions`      | Uniform grootboek van financiële transacties plus diagnostiek van datagezondheid                |
| Berekeningsengine | `Job_Costing_Engine`    | Aggregatie op boekingniveau van omzet, COGS, bankkosten, winst, marge en aanbetalingen  |
| Analyse            | `Booking_PL`            | Winst-en-verliesrekening van één boeking en uitsplitsing per service                                  |
| Analyse            | `Container_Analysis`    | Directe kosten op containerniveau en toegekende gedeelde kosten                                |
| Analyse            | `Service_Profitability` | Winstgevendheid per service over boekingen heen                                           |
| Directie-uitvoer   | `Executive_Dashboard`   | Mondiale KPI's, winstgevendheidsrangschikkingen en waarschuwingen over datagezondheid                    |

De kernstroom is:

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

De bronarchitectuur benoemt drie belangrijke opzoekketens met meerdere hops: een valuta/financiële-behandelingsketen met vier hops, een opzoekketen van gedeelde kosten/container met vier hops en een serviceclassificatieketen met drie hops. 

De kritieke ontwerpregel is dat **berekeningswaarden en diagnostische toestanden gescheiden blijven**. `USD Equivalent` blijft numeriek, terwijl `Data Health Status` fouten draagt zoals `ERR: UNKNOWN_CURRENCY`. Dit behoudt `SUMIFS`-compatibiliteit en voorkomt tegelijkertijd dat defecte transacties de engine van de winst-en-verliesrekening binnenkomen. 

### Drie valstrikken die zelfs ervaren vrachtoperatoren te pakken nemen

#### Valstrik 1 — Een ontbrekende wisselkoers als geldig bedrag behandelen

**1. Beslissing:** Een boeking lijkt een bepaalde kostenpost en marge te hebben.

**2. Verborgen gebrek:** Een leverancierstransactie is uitgedrukt in een valuta die niet is geconfigureerd in de referentietabel van wisselkoersen.

**3. Impact van de beslissing:** Als de ontbrekende koers onjuist als `1.0` wordt behandeld, kan een vreemd bedrag zoals `50,000` veranderen in `$50,000 USD`, waardoor het boekingsresultaat radicaal wordt vervormd.

**4. Waarom dit fout is:** Een ontbrekende wisselkoers is een falen van data-integriteit, geen economische wisselkoers.

**5. Juiste aanpak:** Geef een neutrale numerieke waarde van `0` terug voor het berekeningsveld en rapporteer `ERR: UNKNOWN_CURRENCY` afzonderlijk.

**6. Gecorrigeerde beslissing:** De boeking wordt gemarkeerd voor gegevensaanvulling, in plaats van een valse winst of verlies toegewezen te krijgen.

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
<summary>Formulelogica</summary>

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

De belangrijke regel is: **wijs een onbekende valuta nooit standaard 1.0 toe**. 

</details>

#### Valstrik 2 — De winst-en-verliesrekening van een boeking lezen voordat de datagezondheid is gecontroleerd

**1. Beslissing:** Een boeking met `$8,500` omzet en `$5,200` COGS wordt beschouwd als een winst van `$3,300`.

**2. Verborgen gebrek:** Eén of meer brontransacties hebben ongeldige boekings- of servicekoppelingen, of een onopgeloste valuta.

**3. Impact van de beslissing:** De schijnbare marge kan worden overschat omdat een kostenregel stilzwijgend niet in de aggregatie is opgenomen.

**4. Waarom dit fout is:** Een wiskundig correct `SUMIFS`-resultaat kan economisch fout zijn als de bronregels onvolledig zijn.

**5. Juiste aanpak:** De engine aggregeert alleen transacties met `Data Health Status = "OK"` en rapporteert het aantal ongeldige regels afzonderlijk.

**6. Gecorrigeerde beslissing:** De boeking wordt behandeld als **voorlopig** totdat de gegevensuitzonderingen zijn opgelost.

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
<summary>Formulelogica</summary>

```excel
rev =
SUMIFS(
  Raw_Transactions!O2:O10000,
  Raw_Transactions!B2:B10000,bk,
  Raw_Transactions!N2:N10000,"Revenue",
  Raw_Transactions!P2:P10000,"OK"
)
```

Hetzelfde gezondheidsfilter wordt toegepast op COGS, bankkosten en aanbetalingen. 

</details>

#### Valstrik 3 — Gedeelde kosten op boekingniveau direct aan elke container toewijzen

**1. Beslissing:** Een boeking met twee containers heeft een winstgevendheidsoverzicht op containerniveau nodig.

**2. Verborgen gebrek:** Een COGS-regel op boekingniveau behoort tot de zending als geheel en niet tot een specifieke container.

**3. Impact van de beslissing:** De volledige gedeelde kostenpost aan beide containers toewijzen verdubbelt de schijnbare kostenpost.

**4. Waarom dit fout is:** Het economische eigendom van de kostenpost ligt op boekingniveau, terwijl de analyse een gecontroleerde toewijzingsregel vereist.

**5. Juiste aanpak:** Gebruik het geconfigureerde beleid `EQUAL_SHARE` en deel de gedeelde COGS door het aantal containers van de boeking.

**6. Gecorrigeerde beslissing:** Elke container ontvangt zijn directe kosten plus zijn toegekende aandeel in de gemeenschappelijke kostenpost, waarbij het totaal op boekingniveau behouden blijft.

```text
Booking shared COGS = $1,000
Container count     = 2

Allocated shared cost per container
= $1,000 / 2
= $500
```

Het werkboek scheidt directe containerkosten expliciet van gedeelde kosten op boekingniveau en gebruikt een veilige noemer op basis van het aantal containers. 

<details>
<summary>Formulelogica</summary>

```excel
alloc_cogs =
IF(
  alloc_policy="EQUAL_SHARE",
  direct_cogs + (shared_cogs / cntr_count),
  direct_cogs
)
```

`Setup_Params!B6` bepaalt of de gedeelde kostenpost wordt toegekend met `EQUAL_SHARE` of niet-toegewezen blijft met `NO_ALLOCATE`. 

</details>

### Voorbeeldscenario

Beschouw een boeking in het goederenvervoer met de volgende geldige transacties:

| Post                      | Type     | Oorspronkelijk bedrag | Valuta | USD-equivalent |
| ------------------------- | -------- | --------------: | -------- | -------------: |
| Klantfactuur          | Revenue  |           8,500 | USD      |         $8,500 |
| Leverancierskosten oceaanvracht | COGS     |           3,000 | USD      |         $3,000 |
| Herkomstinspectie         | COGS     |             900 | USD      |           $900 |
| Cargoverzekering           | Revenue  |             400 | USD      |           $400 |
| Bankkosten               | Bank Fee |              75 | USD      |            $75 |
| Terugvorderbare aanbetaling       | Deposit  |             500 | USD      |           $500 |

De geldige winst-en-verliesrekening is:

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

De terugvorderbare aanbetaling van `$500` wordt apart gevolgd, in plaats van te worden behandeld als bedrijfsomzet of COGS.

Stel nu dat één aanvullende leverancierstransactie wordt ingevoerd in een niet-geconfigureerde valuta. Het systeem voegt de transactie **niet** stilzwijgend toe aan de COGS. De toegepaste wisselkoers wordt `0`, de USD-berekening blijft numeriek, en het diagnostiekveld rapporteert `ERR: UNKNOWN_CURRENCY`. De boeking ontvangt een gezondheidswaarschuwing, in plaats van als schoon te worden gepresenteerd.

De operationele aanbeveling verandert dienovereenkomstig:

> **Gebruik de marge van 55.34% niet als eindcijfer voor winstgevendheid tot de onopgeloste transactie is gecorrigeerd.**

Dit is het centrale onderscheid tussen een winstgevendheidsrekenmachine en een gecontroleerde workflow voor opdrachtkostenberekening: het resultaat gaat vergezeld van bewijs over de vraag of de onderliggende gegevens volledig genoeg zijn om te vertrouwen.

### Formulereferentie

<details>
<summary>Booking Master — Aantal containers</summary>

**Doel:** Bepaal het aantal containers dat aan elke boeking is gekoppeld en voorkom fouten door deling door nul.

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

Als er nog geen container is ingevoerd, valt de noemer terug op `1`. Dit is een rekenbeveiliging, terwijl de ontbrekende-containerconditie operationeel toch als een gegevenskwaliteitsprobleem moet worden behandeld. 

</details>

<details>
<summary>Raw Transactions — Toegepaste wisselkoers</summary>

**Doel:** Geef voorrang aan een handmatige wisselkoers en haal anders de geconfigureerde valutakoers op.

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

Een onbekende valuta levert `0`, niet `1.0`. 

</details>

<details>
<summary>Raw Transactions — Financiële behandeling</summary>

**Doel:** Koppel operationele transactietypes aan financiële categorieën.

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

Het bronontwerp gebruikt `COGS` als conservatieve terugval voor een niet-herkend transactietype, wat het risico verkleint de winst te overschatten. 

</details>

<details>
<summary>Raw Transactions — USD-equivalent</summary>

**Doel:** Zet elk oorspronkelijk transactiebedrag om in de USD-basisvaluta van het werkboek met behoud van een numeriek resultaat.

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

Deze kolom numeriek houden is belangrijk omdat stroomafwaartse `SUMIFS`-berekeningen nooit diagnosetekst zoals `"INVALID"` in plaats van een getal mogen ontvangen. 

</details>

<details>
<summary>Raw Transactions — Data Health Status</summary>

**Doel:** Valideer de koppelingen van boeking, service en wisselkoers onafhankelijk van de financiële berekeningskolommen.

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

De diagnostieklaag is opzettelijk gescheiden van de numerieke berekeningsvelden. 

</details>

<details>
<summary>Job Costing Engine — Winst-en-verliesrekening op boekingniveau</summary>

**Doel:** Genereer een dynamische lijst van boekingen en aggregeer alleen gezonde transacties.

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

De essentiële controle is het herhaalde criterium `Data Health Status = "OK"`. 

</details>

### Validatieregels

| Veld / voorwaarde      | Regel                                                            | Foutgedrag                                                                     |
| ---------------------- | --------------------------------------------------------------- | ---------------------------------------------------------------------------------- |
| `Booking ID`           | Moet bestaan in `Booking_Master`                                  | `ERR: INVALID_BOOKING`; transactie uitgesloten van de P&L                              |
| `Service Code`         | Moet bestaan in `Service_Master`                                  | `ERR: INVALID_SERVICE`; transactie uitgesloten van de P&L                              |
| `Currency`             | Moet een geldige geconfigureerde wisselkoers of positieve handmatige wisselkoers hebben | Toegepaste koers wordt `0`; diagnostiek rapporteert `ERR: UNKNOWN_CURRENCY`                 |
| `Manual FX Rate`       | Positieve waarde overschrijft de referentiewisselkoers                  | Ongeldige/niet-positieve waarde valt terug op de geconfigureerde koers                             |
| `Transaction Type`     | Hoort te matchen met de `Setup_Params`-toewijzing                             | Onbekend type valt conservatief terug op `COGS`                                   |
| `USD Equivalent`       | Moet numeriek blijven                                             | Diagnosetekst mag nooit in het numerieke aggregatieveld worden geschreven           |
| `Data Health Status`   | Moet `OK` zijn voor financiële aggregatie                          | Niet-OK-regels zijn uitgesloten van de P&L van de boeking                                      |
| `Container Count`      | Mag niet nul worden wanneer het als toewijzingsnoemer dient        | Valt terug op `1` om `#DIV/0!` te voorkomen                                             |
| Verdeling van gedeelde kosten | Gecontroleerd door `Setup_Params!B6`                                 | `EQUAL_SHARE` kent toe op basis van het aantal containers; `NO_ALLOCATE` laat alleen directe kosten |
| Winstgevendheid van de boeking  | Omzet van nul mag geen delingsfout produceren               | De marge levert `0` op                                                                 |
| Gezondheid directiegegevens  | Elke niet-lege status anders dan `OK` wordt geteld                 | Het dashboard toont een waarschuwing in plaats van een onvoorwaardelijk schone toestand                      |

De bronkruisvalidatie bevestigt de dekking van de kerninvoer van boeking, container, service en transactie, plus de uitvoer van USD-equivalent, nettowinst, marge en toegekende kosten. Ze verifieert ook specifiek de scheiding tussen diagnosetekst en numerieke `SUMIFS`-velden. 

<!-- Technical Details parent <details> intentionally remains open for Part 2. -->
### Kruisvalidatie en paraatheid voor implementatie

De werkboekarchitectuur is ontworpen om de analytische keten te valideren **voordat** het management op de resulterende P&L vertrouwt.

De bronimplementatie controleert expliciet de toewijzingsketen van transactietypes, de opzoekketen van wisselkoersen, de koppeling met Service Master, de toewijzingsketen van gedeelde kosten en de compatibiliteit tussen diagnosetekst en numerieke aggregatievelden. 

| Validatiedimensie         | Wat wordt gecontroleerd                                                                  | Verwachte uitkomst                                          |
| ---------------------------- | -------------------------------------------------------------------------------- | -------------------------------------------------------- |
| **Boekingskoppeling**          | Elke `Booking ID` van een transactie bestaat in `Booking_Master`                        | Ongeldige regels krijgen een fouttoestand                      |
| **Servicekoppeling**          | Elke `Service Code` van een transactie bestaat in `Service_Master`                      | Ongeldige regels zijn uitgesloten van de P&L                       |
| **Wisselkoersdekking**              | Elke transactie heeft een geldige handmatige wisselkoers of een geconfigureerde referentiekoers | Onbekende valuta triggeren een gezondheidswaarschuwing              |
| **Financiële classificatie** | Transactietype wordt toegewezen aan Revenue / COGS / Bank Fee / Deposit                     | Onbekende typen gebruiken de conservatieve COGS-terugval         |
| **Numerieke aggregatie**      | `USD Equivalent` blijft onder alle omstandigheden numeriek                                    | `SUMIFS` stuit niet op diagnosetekst              |
| **Door gezondheid gefilterde aggregatie** | De P&L omvat alleen `Data Health Status = "OK"`                                    | Kapotte bronregels kunnen niet stilzwijgend in de gerapporteerde winst terechtkomen |
| **Containertoewijzing**     | Gedeelde kosten op boekingniveau gebruiken het geconfigureerde toewijzingsbeleid                  | Geen herhaalde toewijzing van de volledige kostenpost over containers heen       |
| **Nul-omzetbescherming**  | De margeberekening gaat om met Revenue = 0                                    | Geen `#DIV/0!`                                             |
| **Directiemonitoring**     | Mondiale transactiefouten worden apart geteld                                 | Het dashboard signaleert onopgeloste problemen in de brondata          |

De kruisvalidatiematrix van de implementatie stelt dat de kerninvoervelden — boeking, container, service en transactiegegevens — en de belangrijkste uitvoer — USD-equivalent, nettowinst, marge en toegekende kosten — door de architectuur zijn gedekt. 

### Operationele SOP

Het beoogde bedrijfsmodel is opzettelijk simpel:

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

#### 1. Stamdata vastleggen

Leg de boeking aan in `Booking_Master`.

Als containers relevant zijn, leg dan hun boeking-naar-container-relaties aan in `Container_Master`.

Onderhoud de gestandaardiseerde servicecatalogus in `Service_Master`. 

Begin de financiële analyse niet door vrije servicenamen of boeking-identificatoren direct in rapportagebladen te typen. De stamtabellen zijn de referentielaag voor stroomafwaartse opzoekketens. 

#### 2. Financiële transacties vastleggen

Voer omzet, leverancierskosten, bankkosten, aanbetalingen en andere relevante financiële activiteit in `Raw_Transactions` in.

Elke transactie moet het passende van het volgende bevatten:

* `Booking ID`
* `Container ID`, indien van toepassing
* `Service Code`
* `Transaction Type`
* `Allocation Level`
* `Original Amount`
* `Currency`
* `Manual FX Rate`, wanneer een transactiespecifieke koers vereist is

De transactietabel is de enige bron die de gedetailleerde records van financiële gebeurtenissen zou moeten ontvangen. De berekeningsengine mag niet handmatig worden bewerkt. 

#### 3. Gezondheid van gegevens controleren

Controleer `Raw_Transactions[Data Health Status]`.

Een schone regel levert:

```text
OK
```

Uitzonderingen benoemen de gebroken relatie, zoals:

```text
ERR: INVALID_BOOKING
ERR: INVALID_SERVICE
ERR: UNKNOWN_CURRENCY
```

De implementatie behandelt dit opzettelijk als een aparte controlelaag. Een transactie kan daardoor numeriek veilig blijven en tegelijkertijd zichtbaar uit het financiële resultaat worden uitgesloten. 

#### 4. Winstgevendheid beoordelen

Gebruik:

* `Job_Costing_Engine` voor winstgevendheid op boekingniveau.
* `Booking_PL` voor de uitsplitsing per service van één boeking.
* `Container_Analysis` wanneer de boeking meerdere containers bevat.
* `Service_Profitability` voor service-economie over boekingen heen.
* `Executive_Dashboard` voor monitoring op directieniveau.

Dit creëert een inzoompad van **mondiale winstgevendheid → boeking → service → container → transactie**, in plaats van elke vraag uit één samenvattend dashboard te laten beantwoorden.

### Regels voor gegevensonderhoud

De architectuur steunt op gecontroleerde stamdata omdat Excel niet dezelfde harde handhaving van vreemde sleutels biedt als een database- of ERP-platform.

De belangrijkste onderhoudsregels zijn:

1. **Overschrijf geen door formules gegenereerde velden.**
2. **Typ servicenamen niet handmatig in waar een servicecode moet worden gebruikt.**
3. **Voer geen nieuwe valuta in zonder eerst haar wisselkoers aan `Setup_Params` toe te voegen.**
4. **Gebruik `1.0` niet als terugval voor een ontbrekende wisselkoers.**
5. **Schrijf geen fouttekst in numerieke aggregatiekolommen.**
6. **Bewerk de uitvoer van `Job_Costing_Engine` niet handmatig.**
7. **Los niet-OK-transactieregels op voordat u de marge van een boeking als eindgeldig behandelt.**
8. **Houd het beleid voor verdeling van gedeelde kosten gecentraliseerd in `Setup_Params`.**

Het bronontwerp verbiedt specifiek het hard coderen van globale parameters in formules en centraliseert basisvaluta, wisselkoersen, transactietoewijzingen en toewijzingsbeleid in `Setup_Params`. 

### Problemen oplossen

#### `⚠️ X TXN DATA ERRORS` verschijnt op het dashboard

**Mogelijke oorzaken:**

* De valuta ontbreekt in de wisselkoerstabel.
* De `Booking ID` bestaat niet in `Booking_Master`.
* De `Service Code` bestaat niet in `Service_Master`.
* Een brontransactie bevat een ongeldige koppeling.

**Oplossing:**

Filter de kolom `Data Health Status` in `Raw_Transactions` om waarden anders dan `OK` te tonen.

Los daarna elk onderliggend probleem in de brongegevens op, in plaats van de uitvoer van de P&L handmatig te wijzigen.

Het Executive Dashboard is ontworpen om deze toestand via een mondiale gezondheidskaart bloot te leggen:

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



#### Dynamische matrixformules geven `#SPILL!` terug

Een overloopformule vereist dat het doelgebied vrij blijft.

Controleer:

* cellen direct onder de formule;
* cellen rechts van de formule;
* handmatig ingevoerde waarden die het beoogde overloopbereik bezetten;
* resterende opmaak of gekopieerde inhoud die een obstructie heeft veroorzaakt.

Maak de hinderende cellen leeg en laat de dynamische matrix opnieuw expanderen.

#### Winstgevendheid verschijnt onverwacht hoog

Wijzig de winstformule niet onmiddellijk.

Controleer eerst:

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

Een verdacht hoge marge kan worden veroorzaakt door een kostentransactie die in de koppelingsketen is mislukt en daarom uit de berekening is uitgesloten.

Dit is precies waarom het werkboek onderscheid maakt tussen **"nul als numerieke berekeningstoestand"** en **"nul als economische conclusie"**. Een ontbrekende wisselkoers kan `$0.00` in een numeriek veld produceren en tegelijkertijd `ERR: UNKNOWN_CURRENCY` in het diagnostiekveld produceren. 

#### Winstgevendheid van de container verschijnt over- of onderschat

Controleer drie dingen in volgorde:

1. Is de juiste `Booking ID` geselecteerd?
2. Is de juiste `Container ID` geselecteerd?
3. Gebruikt `Setup_Params!B6` het beoogde toewijzingsbeleid?

Onder `EQUAL_SHARE` worden gedeelde COGS op boekingniveau gedeeld door het aantal containers van de boeking voordat ze aan directe containerkosten worden toegevoegd.

Onder `NO_ALLOCATE` worden gedeelde kosten op boekingniveau niet aan individuele containers toegekend. 

Het onderscheid is van belang omdat het toewijzen van de volledige gedeelde kostenpost op boekingniveau aan elke container dezelfde economische kostenpost zou dupliceren. 

### Grenzen van de implementatie

Dit werkboek is opzettelijk ontworpen als een **lichte analytische controlelaag**, niet als vervanging van een boekhoudplatform of transportmanagementsysteem.

Het is het sterkst wanneer:

* operationele en financiële gegevens al in spreadsheets of exports bestaan;
* de belangrijkste analyse-eenheid een boeking of zending is;
* het management winstgevendheid op opdrachtniveau nodig heeft;
* services en kostenposten over boekingen heen vergeleken moeten worden;
* meerdere valuta een consistente rapportagebasis nodig hebben;
* gedeelde kosten een expliciete toewijzingsregel vereisen;
* gebruikers transparante formules nodig hebben in plaats van ondoorzichtige automatisering.

Het is niet ontworpen om te bieden:

* transactionele concurrency voor meerdere gebruikers;
* machtigingen op ERP-niveau;
* geautomatiseerde bankafstemming;
* geautomatiseerde verwerking van vervoerdersfacturen;
* handhaving van vreemde sleutels op databaseniveau;
* orchestratie van bedrijfsworkflows.

De waarde van het werkboek ligt daarom niet in het aantal bladen dat het bevat. Het is de gecontroleerde analytische keten die **boeking → transactie → classificatie → valutaomrekening → gevalideerde aggregatie → winstgevendheid → inzoomen** verbindt.

### Analytisch controleprincipe

De belangrijkste implementatieregel kan als volgt worden samengevat:

```text
Never let a broken data relationship
look like a valid financial result.
```

Een conventioneel spreadsheet zou kunnen toestaan:

```text
Missing FX
    ↓
0 USD
    ↓
SUMIFS
    ↓
Higher Profit
```

Dit werkboek maakt de toestand daarentegen expliciet:

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

Dat onderscheid is de fundering van het betrouwbaarheidsmodel van de Job Costing Engine. De architectuur erkent expliciet dat `SUMIFS` numerieke waarden veilig kan aggregeren terwijl tekst stilzwijgend wordt genegeerd — diagnostische toestanden mogen daarom nooit direct in numerieke berekeningskolommen worden ingebed. 

### Aanbevolen implementatievolgorde

Voor een nieuwe werkboekimplementatie is de veiligste volgorde:

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

Het resultaat is een herhaalbare workflow en geen eenmalige spreadsheetberekening: leg de referentiegegevens eenmaal vast, voeg nieuwe operationele transacties toe, los uitzonderingen op en laat de berekenings- en analyselagen bijwerken vanuit hetzelfde onderliggende gegevensmodel. 

<!-- Part 2 ends with the Technical Details parent <details> intentionally OPEN. Do not add </details> here. -->
### Slotopmerkingen over de architectuur

Het werkboek is opzettelijk gebouwd rond een **enkele engine voor winstgevendheid op boekingniveau**, in plaats van een verzameling losse rekenmachines.

De analytische hiërarchie is:

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

Deze structuur houdt de boeking als centrale economische eenheid en staat gebruikers tegelijkertijd toe naar beneden te zoomen in services en containers of naar boven te zoomen naar winstgevendheid over boekingen heen. 

De implementatie erkent ook expliciet drie opzoekketens met meerdere hops:

1. **Transactie → Financiële behandeling → FX → USD-equivalent → P&L van de boeking**
2. **Boeking → Aantal containers → Beleid gedeelde kosten → Toewijzing gedeelde kosten → Containeranalyse**
3. **Servicecode → Service Master → Servicecategorie → Winstgevendheid per service**

Deze ketens worden gevalideerd voordat de analytische uitvoer als productierijp wordt beschouwd. 

### Waartegen de engine beschermt

Het centrale ontwerpprincipe is niet louter automatisering. Het is **voorkomen dat een gebroken invoerrelatie zich voordoet als een geldige financiële conclusie**.

| Faalwijze                                              | Beschermingsmechanisme                                                           |
| ---------------------------------------------------- | ------------------------------------------------------------------------------ |
| Onbekende valuta                                     | Koers valt terug op `0`, terwijl een aparte diagnostische toestand de ontbrekende koers meldt |
| Onbekend transactietype                                     | Conservatieve `COGS`-terugval voorkomt accidentele winstinflatie              |
| Ongeldige boeking                                      | `ERR: INVALID_BOOKING`                                                         |
| Ongeldige servicecode                                      | `ERR: INVALID_SERVICE`                                                         |
| Gebroken transactieketen                             | Niet-`OK`-regels zijn uitgesloten van de P&L-aggregatie                                |
| Gedeelde kostenpost gedupliceerd over containers             | Expliciet toewijzingsbeleid en noemer op basis van het aantal containers                     |
| Nul-omzet                                         | Margeberekening levert `0` op in plaats van `#DIV/0!`                            |
| Diagnosetekst die financiële berekeningen verontreinigt | Numerieke berekeningskolommen en diagnostiekkolommen blijven gescheiden             |

De terugvalmatrix van de implementatie wijst specifiek een `1.0`-wisselkoersterugval af omdat die een niet-herkend bedrag in vreemde valuta in een ongeloofwaardig groot USD-bedrag zou kunnen veranderen. 

### Aanbevolen gebruikspatroon

Voor doorlopende operaties is het werkboek bedoeld om zich te gedragen als een **verversbaar bedrijfsmodel**, niet als een eenmalig rapport.

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

Dit maakt het werkboek nuttig voor terugkerende vragen zoals:

* Welke boekingen zijn werkelijk winstgevend?
* Welke services produceren de sterkste marges?
* Welke zendingskosten veroorzaken een onverwacht verlies?
* Welke kosten op boekingniveau moeten aan containers worden toegekend?
* Zijn de gerapporteerde marges betrouwbaar, of verbergen onopgeloste problemen in de brongegevens kosten?
* Waar moeten prijsstelling, leveranciersonderhandeling of operationele review zich vervolgens op richten?

De bron-SOP volgt dezelfde volgorde: leg stamdata vast, registreer financiële transacties, inspecteer `Data Health Status` en gebruik daarna de winstgevendheidsdashboards voor beoordeling. 

### Slotontwerpprincipe

Het werkboek probeert de financiële analyse niet eenvoudiger te laten lijken door zijn onzekerheid te verbergen.

Het maakt de onzekerheid daarentegen zichtbaar.

Een boeking met een berekende marge wordt niet automatisch als een betrouwbare boeking behandeld.

Een **schone boeking** betekent:

```text
Valid Booking
+ Valid Service
+ Valid Currency / FX
+ Valid Transaction Classification
+ Healthy Transaction Chain
= Profitability Result That Can Be Reviewed
```

Dat onderscheid is de kern van de implementatie.

</details>

## Andere tools in deze serie

Een verzameling lichtgewicht Excel-tools voor beslissingsondersteuning die operationele controle, winstgevendheidsanalyse, voorraadafstemming, personeelsplanning en andere bedrijfsworkflows dekken.

* **[rolling-cashflow-forecast](https://github.com/HyVoid/rolling-cashflow-forecast)** — Waarom is contant geld altijd krap? — Prognoseer toekomstige liquiditeit onder wisselende operationele en financiële aannames.
* **[Multi-Currency-Trade-Financial-Analyzer](https://github.com/HyVoid/Multi-Currency-Trade-Financial-Analyzer)** — Welke handelsstromen zijn echt winstgevend na valutaeffecten? — Meet de winstgevendheid van handel in meerdere valuta over grensoverschrijdende operaties heen.
* **[RTL_Loan_Sizer_Pricer](https://github.com/HyVoid/RTL_Loan_Sizer_Pricer)** — Welke financieringsstructuur is houdbaar onder verschillende leners- en onderpandscenario's? — Vergelijk leenscenario's met risicogecorrigeerde financieringsaannames.
* **[Personal-Business-Dual-Track-Budget-Manager](https://github.com/HyVoid/Personal-Business-Dual-Track-Budget-Manager)** — Waar wordt persoonlijk en zakelijk geld gemengd of verkeerd toegewezen? — Scheid persoonlijke en zakelijke geldstromen in onafhankelijk gevolgde budgetsporen.
* **[Multifamily-Value-Add-Acquisition-Model](https://github.com/HyVoid/Multifamily-Value-Add-Acquisition-Model)** — Creëert deze vastgoedaankoop nog aandelenwaarde na renovatie- en exploitatieaannames? — Evalueer de rendementen van multifamily-aankopen onder renovatiefinanciering en exploitscenario's.
* **[Real-Estate-Valuation-all-in-one-calculator](https://github.com/HyVoid/Real-Estate-Valuation-all-in-one-calculator)** — Wat is een pand waard onder verschillende waarderingsmethoden? — Bereken de waarde van een pand volgens meerdere waarderingsmethoden met belasting- en kapitaaluitgavenmodellering.
* **[multi-entity-expense-allocation](https://github.com/HyVoid/multi-entity-expense-allocation)** — Worden gedeelde kosten eerlijk verdeeld over entiteiten, of worden sommige activa overbelast terwijl andere aan hun echte kostenblootstelling ontsnappen? — Verdeel leveranciersfacturen over meerdere rechtspersonen en operationele activa met configureerbare toewijzingsmethodologieën.
* **[investment-pool-nav-excel-template](https://github.com/HyVoid/investment-pool-nav-excel-template)** — Wat is elke deelnemer werkelijk verschuldigd op basis van zijn instapmoment en de huidige poolwaarde? — Volg het eigendom van een investeringspool continu met op NAV gebaseerde unit-boekhouding.
* **[shopify-cash-flow-forecast-excel-toolkit](https://github.com/HyVoid/shopify-cash-flow-forecast-excel-toolkit)** — Welke operationele verplichtingen veroorzaken de komende twaalf weken een kasprobleem? — Verbind Shopify-verkoopbanktransacties, advertentie-uitgaven en voorraadaankopen tot één beslissingskader voor kasstromen.
* **[Revenue-Management-Decision-Engine](https://github.com/HyVoid/Revenue-Management-Decision-Engine)** — Heeft er voldoende betekenisvolle informatie veranderd om een prijswijziging te rechtvaardigen? — Stuur prijssamenstellingen aan met configureerbare vangrails die onnodige omzetvolatiliteit voorkomen.
* **[amazon-wbr-sqp-analytics-template](https://github.com/HyVoid/amazon-wbr-sqp-analytics-template)** — Wat is er deze periode werkelijk veranderd in het bedrijf en welke producten of zoekwoorden dreven die verandering? — Transformeer Amazon Seller Central SQP- en Sellerboard-exports naar gestandaardiseerde executive business reviews.
* **[food-manufacturing-cost-calculator-excel](https://github.com/HyVoid/food-manufacturing-cost-calculator-excel)** — Welke producten zijn echt winstgevend nadat alle fabricagekosten zijn toegewezen en welke vernietigen de marge? — Wijs fabrieksarbeid, nutsvoorzieningen, overhead en materiaalkosten toe aan elke SKU voor echte eenheidsfabricage-economie.
* **[ecommerce-profit-calculator-excel](https://github.com/HyVoid/ecommerce-profit-calculator-excel)** — Welke orders zijn echt winstgevend na alle verborgen kosten en waar lekt geld weg over kanalen en valuta? — Reconstructeer de echte winstgevendheid op orderniveau door verkoopuitbetalingen, verzendkosten, betaalkosten en marktplaatscommissies af te stemmen over multichannel e-commerceoperaties.
* **[construction-job-costing-excel-template](https://github.com/HyVoid/construction-job-costing-excel-template)** — Welke aannemersprojecten en klanten zijn echt winstgevend na alle opdrachtkosten? — Verbind lead-estimatie, opdrachtuitvoering en werkelijke kostenopvolging tot één uniform kader voor aannemerswinstgevendheid.
* **[Flora-Pastry-Coffee-Shop-Operating-Control-System](https://github.com/HyVoid/Flora-Pastry-Coffee-Shop-Operating-Control-System)** — Is de winkel echt winstgevend na alle ingrediënt-, arbeids- en exploitatiekosten? — Verbind receptkosten, verkopen, voorraadverbruik en exploitatiekosten tot één uniform kader voor winstgevendheid en kasstromen.
* **[cross-entity-logistics-console](https://github.com/HyVoid/cross-entity-logistics-console)** — Welke containers en klanten zijn echt winstgevend over douane- en vlootoperaties heen? — Consolideer douaneafhandeling- en vlootuitvoeringsgegevens tot één overzicht van winstgevendheid op containerniveau.
* **[family-loan-tracker-excel](https://github.com/HyVoid/family-loan-tracker-excel)** — Wat is er vandaag werkelijk verschuldigd na jaren overboekingen en opgelopen rente? — Stem langetermijnfondsoverboekingen en rentetoerekening af tot een actueel openstaand saldo.
* **[3-statement-financial-model-excel](https://github.com/HyVoid/3-statement-financial-model-excel)** — Hoe zien onze financiële overzichten eruit onder verschillende exploitatie- en financieringsscenario's? — Prognoseer geïntegreerde financiële overzichten door omzet, kosten, werkkapitaal en schuld te koppelen tot één uniform model.
* **[assisted-living-cost-allocation-excel](https://github.com/HyVoid/assisted-living-cost-allocation-excel)** — Welk deel van de gerapporteerde marge van elk gebouw wordt werkelijk door dat gebouw gegenereerd en welk door de manier waarop hoofdkantoorkosten werden toegewezen? — Wijs gedeelde kosten toe over operationele eenheden, behoud brontransacties en stem resultaten op eenheidsniveau af op bedrijfstotalen.
* **[real-estate-development-feasibility-excel](https://github.com/HyVoid/real-estate-development-feasibility-excel)** — Bij welke maximale grondaanschafprijs wordt dit project een bankable investering? — Evalueer kansrijke woningbouwprojecten door landcomparabelen te screenen, bouwkosten te ramen en winst-op-kostendrempels te berekenen.
* **[uk-residential-costing-software](https://github.com/HyVoid/uk-residential-costing-software)** — Haalt dit bouwproject zijn doelmarge werkelijk zodra site-logistiek, afvalfactoren en gewijzigde vakwerkkoersen zijn verwerkt? — Scheid fysieke bouwvolumes van vluchtige vakwerktarieven en algemene bouwplaatskosten om commerciële opslagen te cascaderen over aanbestedingsbudgetten en klantenoffertes.
* **[b2b-pricing-landed-cost-excel-model](https://github.com/HyVoid/b2b-pricing-landed-cost-excel-model)** — Welke groothandelsprijzen en klantsegmenten behouden werkelijk de marge zodra vracht, rechten, betaalvoorwaarden en FX volledig zijn gekost? — Bereken SKU-niveau landed costs, cost-to-serve en pocket-marge-prijskaskades om grensoverschrijdende groothandelswinstgevendheid te simuleren.
* **[custom-label-pricing-calculator](https://github.com/HyVoid/custom-label-pricing-calculator)** — Wat moeten we in rekening brengen voor een op maat gedrukte etiketbestelling per hoeveelheidstrap om concurrentiekrachtig en winstgevend te blijven? — Bereken parametrische etiketprijzen door substraatkosten per vierkante inch, instelkosten en kortingsdrempels per hoeveelheidstrap te modelleren tegen concurrentbenchmarks.
* **[landed-cost-calculator-retail-deduction-excel](https://github.com/HyVoid/landed-cost-calculator-retail-deduction-excel)** — Wat is de echte nettomarge per SKU zodra werkelijke landed costs en retailerinhoudingen zijn verwerkt? — Stem PO-geschatte landed costs af op werkelijke vracht, rechten, FX en drayage terwijl u retailerafdrachten auditeert om de nettomarge op SKU-niveau bloot te leggen.
* **[saas-finance-mrr-burn-rate-excel](https://github.com/HyVoid/saas-finance-mrr-burn-rate-excel)** — Kunnen we deze week instroom-, marketing- en leveranciersverplichtingen nakomen zonder cash te kort te komen voordat de boeken sluiten? — Volg het wekelijkse cashburn-tempo, MRR/ARR-waterfallvariantie, retentiemetrics en salariskosten tegen het budget vóór de maandafsluiting.
* **[startup-financial-model-excel](https://github.com/HyVoid/startup-financial-model-excel)** — Raakt ons groeiplan zonder cash voordat cohortmonetisatie inhaalt? — Prognoseer meerjarige cashburn, runway, unit-economics en financieringsverdunning op basis van bottom-up cohort- en personeelsaannames.
* **[shopify-pnl-reconciliation-excel](https://github.com/HyVoid/shopify-pnl-reconciliation-excel)** — Is elke Shopify-winkel echt winstgevend zodra echte kosten op SKU-niveau, gedeelde uitgaven en bankkosten zijn afgestemd? — Stem multi-store Shopify-verkoopexports af op SKU-niveau COGS, toegewezen multivaluta-OPEX en bankafschriften om een betrouwbare maandelijkse P&L te produceren.
* **[travel-agency-accounting-excel-template](https://github.com/HyVoid/travel-agency-accounting-excel-template)** — Welke rondreisboekingen genereren werkelijk nettowinst en positieve kasstroom na leverancierskosten en B2B-incasso's? — Bereken opdrachtkosten per boeking van rondreizen, brutomarges, AR/AP-afrekening en operationele kasstromen over meerdere valuta.

## Licentie

Dit project wordt uitgebracht onder de **Apache License 2.0**.

U mag het werk gebruiken, wijzigen, distribueren en erop voortbouwen in overeenstemming met de voorwaarden van de Apache License 2.0.
