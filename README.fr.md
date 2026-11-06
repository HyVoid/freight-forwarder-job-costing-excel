[ 🌐 عربي ](README.ar.md) | [ 🇩🇪 Deutsch ](README.de.md) | [ 🇫🇷 Français ](README.fr.md) | [ 🇳🇱 Nederlands ](README.nl.md) | [ 🇪🇸 Español ](README.sp.md) | [ 🇬🇧 English ](README.md)

# Modèle Excel de calcul des coûts par expédition pour transitaires : calculateur de profits et pertes (P&L) des envois, allocation des coûts de conteneurs & suivi des marges
<img width="2880" height="1228" alt="image" src="https://github.com/user-attachments/assets/f89c1a0f-f576-43af-b26a-02bae4a47348" />


[![License: Apache 2.0](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](LICENSE) [![Platform](https://img.shields.io/badge/Platform-Web%20Browser%20%2B%20Microsoft%20Excel%20(.xlsx)-217346.svg)](#) [![Tool Type](https://img.shields.io/badge/Tool%20Type-Plug--and--Play%20Freight%20Accounting%20Spreadsheet-2251FF.svg)](#)

**Un modèle Excel de calcul des coûts par expédition pour transitaires (`.xlsx`) prêt à l'emploi et un calculateur en ligne gratuit de rentabilité des envois, conçus pour les transitaires, les NVOCC, les courtiers en douane et les opérateurs 3PL qui doivent calculer le bénéfice net par expédition sans écrire de formules.** Suivez les **marges brutes des connaissements (B/L) et des lettres de transport aérien (AWB)**, automatisez les **conversions de change multidevises (USD, EUR, CNY, GBP)**, répartissez les **coûts des conteneurs FCL/LCL** et réconciliez les **factures clients AR avec les factures fournisseurs AP des transporteurs maritimes, du camionnage, du courtage en douane et des virements bancaires** dans un seul tableur logistique de profits et pertes préconfiguré.

**Testez-le gratuitement dans votre navigateur dès maintenant, ou téléchargez le fichier Excel déverrouillé pour coller vos propres données d'expédition.**

Chaque règle de correspondance, conversion de devise et formule de répartition des coûts de conteneurs est déjà intégrée. Collez vos lignes de facturation depuis QuickBooks, Xero ou votre tableur de suivi opérationnel pour voir votre bénéfice exact par réservation, par conteneur et par code de frais de fret en quelques secondes, avec une garantie de remboursement de 30 jours.

> 🌐 **Essayez d'abord le calculateur en ligne** → [Lancez le calculateur web interactif gratuit de calcul des coûts par expédition et de compte de résultat des envois pour transitaires (démo navigateur, aucune configuration requise)](https://hyvoid.github.io/freight-forwarder-job-costing-excel/)  
> 📥 **Téléchargez le modèle Excel prêt à l'emploi** → [Obtenez le modèle Excel déverrouillé de calcul des coûts par expédition et de rentabilité des envois pour transitaires (.xlsx avec formules de mise à jour automatique)](https://www.theseusworkshop.com/l/uyukd?utm_source=github&utm_medium=GitHub%20README&utm_campaign=readme%20new%20launch&utm_content=freight-booking-job-costing)  
>
> *Le calculateur web gratuit et le classeur Microsoft Excel téléchargeable fonctionnent sur exactement les mêmes tables de comptabilité fret préconfigurées et les mêmes formules de marge automatisées.*

---

## Ce que cet outil vous aide à suivre : correspondance entre les points de douleur du transit et les solutions

Attendre les rapports comptables de fin de mois masque quelles expéditions individuelles ont réellement gagné ou perdu de l'argent. Ce **tableur de calcul des coûts par expédition prêt à l'emploi** corrige les cinq angles morts comptables opérationnels les plus fréquents dans le fret maritime, le fret aérien et le transport multimodal :

| Sujet de recherche fréquent & objectif de suivi | Le point de douleur quotidien de la comptabilité fret (pourquoi la marge fuit) | Comment ce modèle Excel automatisé le corrige et le suit |
| :--- | :--- | :--- |
| **Bénéfice par expédition & P&L au niveau de la réservation (`B/L` & `AWB`)** | QuickBooks et Xero regroupent les factures par mois civil, masquant les expéditions déficitaires, les devis sous-évaluées et les lignes de trafic non rentables. | Rapproche **les revenus AR clients, les factures AP fournisseurs et les frais de virement bancaire** de chaque `Booking_ID` individuel pour que vous voyiez en temps réel le **bénéfice net ($) et la marge brute (%) par expédition**. |
| **Rentabilité par code de frais de fret et services accessoires** | Vous connaissez le montant total de la facture, mais vous ne pouvez pas dire si vous avez perdu de l'argent sur le **fret maritime, le camionnage local, la THC, le dédouanement, la surestarie ou l'assurance cargo**. | Décompose les revenus vs. les coûts fournisseurs selon des **codes de services de fret** granulaires (`Ocean Freight`, `THC`, `Origin Inspection`, `Drayage`, `Customs`, `Documentation`) pour révéler les dépassements de coûts cachés. |
| **Répartition des coûts partagés multi-conteneurs FCL (`TEU` / `FEU`)** | Lorsqu'un connaissement principal (MBL) couvre plusieurs conteneurs, les frais forfaitaires de réservation (comme l'entrée en douane ou la documentation B/L) sont doublement comptés ou omis du prix par conteneur. | Sépare les **coûts directs par conteneur** des **frais partagés au niveau de la réservation**, en divisant automatiquement les coûts partagés de manière égale entre tous les conteneurs actifs de cette réservation. |
| **Conversion des factures de fret multidevises (`USD`, `EUR`, `CNY`, `GBP`)** | Payer les transporteurs maritimes en USD, les agents d'origine en CNY ou EUR, et facturer les clients en devise locale avec des taux de change calculés à la main provoque des pertes de change en fin de mois. | Convertit chaque ligne AR et AP multidevise en votre **devise de reporting de base unique** via une table de recherche de change centralisée, en signalant immédiatement tout code de devise non configuré. |
| **Détection des factures fournisseurs manquantes et des fautes de frappe dans le tableur** | Une seule faute de frappe dans un identifiant de réservation dans une feuille manuelle `SUMIFS` ou `VLOOKUP` fait silencieusement tomber une facture transporteur de 2 000 $ à `$0`, gonflant le bénéfice et entraînant un trop-perçu de commissions commerciales. | Exécute une **vérification automatique de l'intégrité des données** (`ERR: UNKNOWN_CURRENCY`, `ERR: INVALID_BOOKING`) qui met en évidence les lignes défectueuses bien en vue tout en préservant l'intégrité des totaux de votre tableau de bord. |

---

## Comment calculer le profit d'une expédition et réaliser un audit de calcul des coûts de fret (tutoriel débutant en 4 étapes)

Vous n'avez besoin d'aucune formation comptable ni de compétences en formules Excel pour auditer vos marges d'expédition. Suivez ces quatre étapes pour transformer des listes de factures brutes et des factures transporteurs en un **état de profits et pertes réservation par réservation** précis :

### Étape 1 : définissez votre devise de base, vos taux de change et votre politique de répartition des conteneurs
Ouvrez l'onglet `Setup_Params` pour établir les règles de référence de votre entreprise en langage clair :
- Choisissez votre **devise de reporting de base** (par ex. `USD`, `EUR` ou `GBP`) et saisissez vos **taux de change (FX) standard**.
- Vérifiez les **catégories de transactions** pré-mappées (`Customer Revenue`, `Direct Vendor COGS`, `Bank Wire Fees` et `Customer Deposits`).
- Sélectionnez votre **règle de répartition des coûts partagés** pour les expéditions FCL multi-conteneurs.

*Essayez en direct :* prévisualisez dès maintenant le fonctionnement de la table de configuration dans le [calculateur web gratuit de calcul des coûts par expédition pour transitaires](https://hyvoid.github.io/freight-forwarder-job-costing-excel/).

---

### Étape 2 : collez vos réservations, conteneurs et lignes de factures AR/AP brutes
Déposez vos identifiants d'expédition dans les listes maîtres (`Bookings` et `Containers`), puis copiez-collez vos lignes de facturation directement dans la table `Raw_Transactions` :
- **Compatible avec toute source de données :** exports QuickBooks Online / Desktop, dumps CSV Xero, journaux de facturation CargoWise ou Magaya, factures de fret des transporteurs maritimes, relevés de courtiers en douane et suivis d'expédition manuels.
- **Zéro glissement manuel de formules :** les tables Excel s'étendent automatiquement lorsque vous collez de nouvelles lignes.

---

### Étape 3 : vérifiez les alertes d'erreur automatiques et consultez votre tableau de bord P&L des expéditions
Dès que vous collez vos données, le tableur calcule tout en arrière-plan :
- Convertit les factures fournisseurs et clients en devises étrangères dans votre devise de base.
- Signale tout identifiant de réservation mal orthographié ou taux de devise manquant dans la colonne `Data Health Status`, afin qu'une facture fournisseur non liée ne se cache jamais comme un élément à coût nul.
- Met à jour votre **synthèse de profits et pertes au niveau des réservations**, votre **table de marges par code de service de fret** et votre **ventilation des coûts unitaires par conteneur**.

---

### Étape 4 : téléchargez le modèle Excel déverrouillé pour une réutilisation mensuelle en un clic
Le calculateur navigateur gratuit est idéal pour une vérification rapide en une seule session. Lorsque vous devez enregistrer vos données localement, importer des centaines de lignes comptables mensuelles et suivre les départs de navires en cours à travers les cycles de facturation, passez au classeur `.xlsx` téléchargeable.

> 💡 **Pourquoi les transitaires débordés téléchargent le fichier Excel plutôt que de construire leur propre feuille :**  
> La réconciliation de fret multidevises, la répartition des coûts multi-conteneurs et la détection des factures orphelines impliquent une logique comptable complexe sous le capot. **Dans ce modèle Excel, chaque règle de correspondance, chaque lien de recherche et chaque formule auto-extensible est déjà construit et testé pour vous.**  
> 
> Vous n'avez même pas besoin de cliquer-glisser les coins de cellules pour copier les formules vers le bas. Téléchargez simplement le fichier sur votre ordinateur, collez vos données de facturation brutes et obtenez instantanément des chiffres de bénéfice prêts pour l'audit. Pendant les pics de la saison d'expédition, consacrez votre temps à remporter des devis de chargeurs, à sécuriser l'espace navire et à déplacer la marchandise—tout en laissant un simple modèle à copier-coller gérer votre analyse financière en quelques secondes.  
> 
> 👉 **[Téléchargez le modèle Excel de calcul des coûts par expédition pour transitaires déverrouillé (.xlsx) et automatisez dès aujourd'hui le P&L de vos expéditions](https://www.theseusworkshop.com/l/uyukd?utm_source=github&utm_medium=GitHub%20README&utm_campaign=readme%20new%20launch&utm_content=freight-booking-job-costing)**

---

## La logique métier et la méthodologie : 4 angles morts de profit cachés que ce modèle Excel résout pour vous

La plupart des transitaires perdent de l'argent sur des expéditions non pas parce qu'ils ont mal coté le taux maritime, mais parce que les outils comptables standard et les feuilles Excel maison masquent où va réellement l'argent. Voici comment ce modèle protège votre résultat net, en termes simples—même si vous n'avez aucune formation comptable formelle :

### 1. Les logiciels comptables mensuels masquent les expéditions déficitaires (correspondance au niveau de la réservation vs. totaux mensuels)
**Le fait objectif :** les logiciels comptables standard comme QuickBooks ou Xero organisent l'argent par **mois civil**, tandis que le transit s'organise par **réservation, connaissement (B/L) et conteneur**. Si vous facturez 8 000 $ à un client en octobre, mais que le transporteur maritime et le transporteur routier vous facturent 7 500 $ en novembre, octobre paraît artificiellement rentable et novembre affiche une perte.  
**Votre bénéfice direct :** ce modèle regroupe chaque dollar de revenu, de coût fournisseur et de frais de virement bancaire par `Booking_ID`, quel que soit le mois d'arrivée de la facture. Vous voyez le véritable bénéfice net de chaque expédition avant de renouveler la grille tarifaire d'un client ou de verser des commissions commerciales.

### 2. Les frais forfaitaires de réservation vs. les coûts par conteneur faussent vos devis (répartition automatique des coûts de conteneurs FCL)
**Le fait objectif :** dans le fret maritime et multimodal, certains coûts surviennent une fois par **réservation** (comme des frais de dédouanement de 150 $ ou des frais de documentation B/L), tandis que d'autres surviennent par **conteneur** (comme des frais de manutention terminale de 450 $ ou le camionnage local par conteneur 40'HC). Dans un tableur basique, les opérateurs multiplient par erreur les 150 $ de frais de douane sur 5 conteneurs (gonflant les estimations de coûts de 600 $ et perdant l'appel d'offres) ou oublient d'inclure les frais partagés de réservation lors du calcul du profit par conteneur.  
**Votre bénéfice direct :** le modèle sépare les coûts spécifiques aux conteneurs des coûts partagés à l'échelle de la réservation et répartit les frais partagés de manière égale sur le nombre exact de conteneurs de cette réservation :

```text
Single Container Total Cost = Direct Container Cost + (Shared Booking Costs ÷ Total Containers in Booking)

```

Vous obtenez des coûts unitaires exacts par conteneur sans effectuer de division manuelle.

### 3. Une seule faute de frappe dans un tableur maison fabrique silencieusement un faux bénéfice élevé (détection intégrée des fautes de frappe et des coûts manquants)

**Le fait objectif :** dans une feuille Excel standard construite avec `SUMIFS`, une seule faute de frappe dans un numéro de référence de réservation sur une facture de camionnage fournisseur de 1 500 $ fait qu'Excel ignore complètement cette facture et ajoute `$0` à vos coûts. Inversement, un `VLOOKUP` cassé transforme toute votre table de synthèse en erreurs `#N/A`. Voyez comment une seule facture fournisseur non liée distord une réservation réelle :

```text
Reported Booking Revenue:   $8,500
Reported Vendor Costs:      $5,200  <-- ($1,500 drayage bill skipped due to a Booking ID typo)
Apparent Net Profit:        $3,300  <-- (Overstated by $1,500)
Apparent Profit Margin:      38.8%  <-- (Fake margin; actual real margin is only 21.1%)

```

Si vous faites confiance à ce chiffre de 38,8 %, vous payez les commissions de vos commerciaux sur votre propre poche et continuez de coter une ligne de trafic qui est tout juste à l'équilibre.

**Votre bénéfice direct :** chaque ligne que vous collez passe par une **vérification automatique de l'intégrité des données**. Si une ligne contient un identifiant de réservation mal orthographié, un code de frais invalide ou un taux de devise manquant, le modèle la signale clairement (`ERR: INVALID_BOOKING` ou `ERR: UNKNOWN_CURRENCY`) tout en gardant vos calculs de synthèse opérationnels. Vous repérez les factures fournisseurs manquantes avant qu'elles ne vous coûtent de l'argent.

### 4. Les devises étrangères et les acomptes clients à l'avance faussent la marge brute réelle (FX au niveau de la ligne et isolation des acomptes)

**Le fait objectif :** les expéditions transfrontalières combinent systématiquement du fret maritime en USD, des factures d'agents d'origine en CNY ou EUR, des frais de virement bancaires intermédiaires et des acomptes clients à l'avance. Convertir les totaux de devises à la main en fin de mois introduit des erreurs de taux de change, et compter les acomptes clients comme revenus acquis fait paraître une expédition rentable avant même que les factures fournisseurs ne soient payées.

**Votre bénéfice direct :** le modèle convertit chaque facture en devise étrangère dans votre devise de base au niveau de chaque ligne et sépare les **acomptes clients (espèces détenues)** des **revenus de fret acquis et du COGS**. Vous voyez votre marge de profit réelle et votre position de trésorerie réelle côte à côte.

---

## Tableurs de fret manuels et logiciels comptables génériques vs. ce modèle Excel automatisé de calcul des coûts par expédition

| Facteur de comparaison opérationnelle | Feuilles Excel manuelles & comptabilité générique (QuickBooks / Xero seuls) | Modèle Excel standardisé de calcul des coûts par expédition pour transitaires (cet outil) |
| --- | --- | --- |
| **Visibilité du profit par expédition (`MBL` / `HBL` / `AWB`)** | Mélange toutes les factures mensuelles dans des postes de grand livre général, masquant quelles réservations, conteneurs ou chargeurs précis perdent de l'argent. | Génère une **table de P&L automatique réservation par réservation** affichant les revenus bruts, le COGS fournisseurs, les frais de virement bancaire, le bénéfice net ($) et la marge (%). |
| **Conversion de factures multidevises (`USD`, `EUR`, `CNY`, `GBP`)** | Le personnel convertit les factures d'agents étrangers avec des calculatrices de bureau et des taux incohérents, provoquant des écarts de marge inexpliqués à la clôture de fin de mois. | Utilise une **table de recherche de taux de change centralisée** pour convertir chaque ligne de facture automatiquement et affiche un avertissement instantané si un taux de devise manque. |
| **Répartition des coûts multi-conteneurs FCL (`20'GP` / `40'HC` / `TEU`)** | Les frais de documentation B/L, le courtage en douane et les surtaxes de réservation sont dupliqués par conteneur ou omis des calculs de coût unitaire des conteneurs. | Sépare automatiquement les **factures directes par conteneur** des **frais généraux partagés de réservation**, en répartissant les frais partagés sur le nombre exact de conteneurs par réservation. |
| **Marge des services accessoires et codes de frais (`THC`, `Drayage`, `Customs`)** | Vous voyez la facturation totale de l'expédition, mais vous ne pouvez pas dire si **l'inspection à l'origine, le camionnage intérieur, la surestarie ou l'assurance cargo** fait fuir le profit. | Agrège les revenus vs. les coûts fournisseurs par **code de service de fret** pour repérer quels frais accessoires rapportent de l'argent et lesquels en font perdre. |
| **Maintenance des formules & protection contre les fautes de frappe (`VLOOKUP` / `SUMIFS`)** | Le personnel doit glisser les formules manuellement pour chaque nouvelle ligne, et un seul identifiant de réservation mal orthographié fait silencieusement tomber les coûts fournisseurs à zéro. | **Tables auto-extensibles préconstruites & alertes d'erreur** se mettent à jour automatiquement au collage et mettent en quarantaine les fautes de frappe (`ERR: INVALID_BOOKING`) avant qu'elles ne corrompent votre P&L. |

---

## À qui cela s'adresse : rôles logistiques, cas d'usage de tableurs & flux de travail du modèle

Cette boîte à outils est conçue pour les **transitaires de petite et moyenne taille, les opérateurs OTI/NVOCC, les courtiers en douane, les équipes 3PL et les agents de fret** qui veulent une visibilité immédiate du profit au niveau des expéditions sans dépenser plus de 20 000 $ et six mois à déployer un ERP ou un TMS d'entreprise :

* **Propriétaires d'entreprises de transit et directeurs généraux de NVOCC (`Freight Forwarder Profit & Loss Excel Template`) :** auditez le bénéfice net réel par expédition (après frais de virement bancaire et frais généraux partagés de réservation) avant d'approuver des grilles tarifaires annuelles de chargeurs ou de verser les commissions mensuelles des commerciaux.
* **Coordinateurs d'opérations juniors et répartiteurs (`Plug-and-Play Shipment Costing Spreadsheet`) :** suivez les coûts des missions FCL et LCL multi-conteneurs pendant les saisons de pointe en collant des factures fournisseurs brutes et des frais clients dans un modèle préconstruit, sans écrire la moindre formule.
* **Comptables logistiques et commis comptables fret (`Freight AR/AP Invoice Reconciliation Template`) :** comblez l'écart entre les exports QuickBooks, Xero ou NetSuite et les dossiers opérationnels d'expédition—en détectant les services accessoires clients non facturés ou les factures transporteurs non liées avant la clôture mensuelle des comptes.
* **Analystes de tarification fret et gestionnaires de lignes de trafic (`Container Unit Cost & Accessorial Margin Calculator`) :** analysez si le fret maritime, le fret aérien, le courtage en douane, l'assurance cargo ou le camionnage local stimule ou réduit les marges sur des lignes d'origine à destination spécifiques.
* **CFO à temps partiel et consultants logistiques (`3PL & Freight Margin Leakage Audit Model`) :** déposez les exports de facturation clients dans un tableur standardisé pour découvrir en un seul après-midi les comptes sous-évalués, les pertes de change et les frais accessoires non récupérés.

*Remarque : ce modèle Excel complète votre TMS opérationnel (comme CargoWise, Magaya ou GoFreight) et votre logiciel de comptabilité fiscale. Il agit comme votre **couche rapide et prête à l'emploi de rentabilité des expéditions et de calcul des coûts** lorsque vous avez besoin de réponses claires de vos données de facturation immédiatement.*

---

## Scénarios courants de calcul des coûts de fret : quand utiliser ce tableur de P&L des expéditions

Que vous gériez un bureau de transit de cinq personnes ou des centaines de conteneurs FCL mensuels, ce modèle résout les problèmes exacts de calcul des coûts quotidiens qui ralentissent les équipes logistiques :

* **Comment calculer le profit par expédition dans Excel sans écrire de formules :** collez vos lignes de factures clients et de factures fournisseurs dans une seule table, étiquetez chaque ligne avec son `Booking_ID`, et laissez la feuille de synthèse préconstruite calculer automatiquement le bénéfice net et la marge brute %.
* **Comment répartir les coûts partagés du connaissement (B/L) sur plusieurs conteneurs FCL :** attribuez les frais spécifiques aux conteneurs (comme le camionnage ou la THC) à des numéros `Container_ID` individuels et laissez les frais à l'échelle de la réservation (comme le dédouanement ou la documentation) au niveau du `Booking_ID`—le modèle répartit automatiquement les coûts partagés entre vos conteneurs 20'GP et 40'HC.
* **Comment réconcilier les factures retardées des transporteurs maritimes et du camionnage avec les factures clients :** conservez les réservations ouvertes dans votre journal à travers les mois civils afin que les avis d'arrivée tardifs, les factures de détention/surestarie ou les frais de fractionnement de châssis se déduisent automatiquement de la marge de la bonne expédition lorsqu'ils arrivent.
* **Comment suivre les marges de fret multidevises entre USD, EUR et CNY :** saisissez une seule fois vos taux de change de référence dans `Setup_Params` pour normaliser les factures d'agents d'origine à l'étranger, le fret des transporteurs maritimes et la facturation de livraison locale dans une seule devise de reporting.
* **Comment auditer les commissions de courtiers de fret et de commerciaux sur le bénéfice brut :** vérifiez que chaque facture fournisseur, frais d'origine et frais de virement bancaire est lié et comptabilisé avant de calculer les paiements de commissions commerciales sur le bénéfice brut des expéditions.

---

## À propos de Theseus Workshop

**Theseus Workshop** construit des modèles financiers Microsoft Excel prêts à l'emploi et des calculateurs basés sur navigateur pour les opérateurs qui ont besoin de chiffres clairs et fiables dès aujourd'hui—sans gonflement de logiciel d'entreprise ni programmation complexe de tableurs.

Chaque modèle est conçu autour d'un objectif pratique unique :

> **« Placer les chiffres opérationnels et financiers exacts dans une seule table préconstruite pour que votre prochaine décision métier prenne des secondes au lieu d'heures. »**

Si vous voulez arrêter de vous battre avec des feuilles `VLOOKUP` cassées et des calculs de devises manuels, téléchargez le modèle Excel complet ci-dessous ou explorez notre bibliothèque complète de tableurs opérationnels et de calculateurs d'affaires :

* 📥 **Obtenir ce modèle de calcul des coûts de fret :** [Téléchargez le modèle Excel déverrouillé de calcul des coûts par expédition et de P&L des envois pour transitaires (.xlsx)](https://www.theseusworkshop.com/l/uyukd?utm_source=github&utm_medium=GitHub%20README&utm_campaign=readme%20new%20launch&utm_content=freight-booking-job-costing)
* 🌐 **Tester le calculateur web gratuit :** [Ouvrez l'application navigateur interactive de calcul des coûts par expédition pour transitaires](https://hyvoid.github.io/freight-forwarder-job-costing-excel/)
* 🧰 **Explorer tous les modèles Excel prêts à l'emploi et boîtes à outils d'affaires :** [Visitez la boutique officielle Theseus Workshop sur TheseusWorkshop.com](https://www.google.com/search?q=https://www.theseusworkshop.com%3Futm_source%3Dgithub%26utm_medium%3DGitHub%2520README%26utm_campaign%3Dreadme%2520new%2520launch%26utm_content%3Dabout-section-store-link)

---

## Détails techniques

<details>
<summary>Pour les réviseurs techniques, les praticiens d'Excel et les collaborateurs</summary>

### Architecture du classeur

Le classeur est organisé autour de quatre couches fonctionnelles : **règles et paramètres → données de base → saisie des transactions → calcul et sortie analytique**.

La conception comporte 11 feuilles :

| Couche              | Feuille                   | Rôle                                                                             |
| ------------------ | ----------------------- | -------------------------------------------------------------------------------- |
| Règles et paramètres | `Instructions`          | Instructions d'utilisation, règles de saisie et définitions métier               |
| Règles et paramètres | `Setup_Params`          | Devise de base, taux de change, correspondances de transactions et politique de répartition des coûts partagés |
| Données de base        | `Booking_Master`        | Index des réservations, client, statut et nombre de conteneurs                             |
| Données de base        | `Container_Master`      | Relation réservation-conteneur 1:N                                            |
| Données de base        | `Service_Master`        | Code de service standard, nom du service et catégorie                                |
| Saisie des transactions  | `Raw_Transactions`      | Registre unifié des transactions financières plus diagnostics d'intégrité des données                |
| Moteur de calcul | `Job_Costing_Engine`    | Agrégation au niveau de la réservation des revenus, COGS, frais bancaires, profit, marge et acomptes  |
| Analyse           | `Booking_PL`            | P&L d'une seule réservation et analyse descendante au niveau du service                                  |
| Analyse           | `Container_Analysis`    | Coûts directs au niveau du conteneur et coûts partagés alloués                                |
| Analyse           | `Service_Profitability` | Rentabilité par service transverse aux réservations                                           |
| Sortie exécutive   | `Executive_Dashboard`   | KPI globaux, classements de rentabilité et avertissements d'intégrité des données                    |

Le flux central est :

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

L'architecture source identifie trois chaînes de recherche importantes à sauts multiples : une chaîne devise/traitement financier à quatre sauts, une chaîne d'analyse descendante coûts partagés/conteneur à quatre sauts et une chaîne de classification des services à trois sauts. 

La règle de conception critique est que **les valeurs de calcul et les états de diagnostic restent séparés**. `USD Equivalent` reste numérique, tandis que `Data Health Status` porte les erreurs telles que `ERR: UNKNOWN_CURRENCY`. Cela préserve la compatibilité `SUMIFS` tout en empêchant les transactions défectueuses d'être incluses dans le moteur de compte de résultat. 

### Trois pièges qui surprennent même les opérateurs de transit expérimentés

#### Piège 1 — Traiter un taux de change manquant comme un montant valide

**1. Décision :** Une réservation semble avoir un certain coût et une certaine marge.

**2. Défaut caché :** Une transaction fournisseur est libellée dans une devise qui n'est pas configurée dans la table de change de référence.

**3. Impact de la décision :** Si le taux manquant est incorrectement traité comme `1.0`, un montant étranger tel que `50,000` peut devenir `$50,000 USD`, déformant radicalement le résultat de la réservation.

**4. Pourquoi c'est faux :** Un taux de change manquant est un échec d'intégrité des données, pas un taux de change économique.

**5. Approche correcte :** Retourner une valeur neutre numérique de `0` pour le champ de calcul tout en signalant séparément `ERR: UNKNOWN_CURRENCY`.

**6. Décision corrigée :** La réservation est signalée pour complétion des données au lieu de se voir attribuer un profit ou une perte faux.

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
<summary>Logique de la formule</summary>

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

La règle importante est de **ne jamais prendre par défaut une devise inconnue à 1.0**. 

</details>

#### Piège 2 — Lire le P&L d'une réservation avant de vérifier l'intégrité des données

**1. Décision :** Une réservation avec `$8,500` de revenus et `$5,200` de COGS est considérée comme un profit de `$3,300`.

**2. Défaut caché :** Une ou plusieurs transactions sources ont des liens de réservation ou de service invalides, ou une devise non résolue.

**3. Impact de la décision :** La marge apparente peut être surestimée parce qu'une ligne de coût a silencieusement échoué à intégrer l'agrégation.

**4. Pourquoi c'est faux :** Un résultat `SUMIFS` mathématiquement correct peut encore être économiquement faux si les lignes sources sont incomplètes.

**5. Approche correcte :** Le moteur n'agrège que les transactions dont le `Data Health Status = "OK"` et signale séparément le nombre de lignes invalides.

**6. Décision corrigée :** La réservation est traitée comme **provisoire** jusqu'à ce que les exceptions de données soient résolues.

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
<summary>Logique de la formule</summary>

```excel
rev =
SUMIFS(
  Raw_Transactions!O2:O10000,
  Raw_Transactions!B2:B10000,bk,
  Raw_Transactions!N2:N10000,"Revenue",
  Raw_Transactions!P2:P10000,"OK"
)
```

Le même filtre d'intégrité est appliqué aux COGS, aux frais bancaires et aux acomptes. 

</details>

#### Piège 3 — Attribuer directement le coût partagé au niveau de la réservation à chaque conteneur

**1. Décision :** Une réservation à deux conteneurs nécessite une vue de rentabilité au niveau du conteneur.

**2. Défaut caché :** Une ligne de COGS au niveau de la réservation appartient à l'expédition dans son ensemble plutôt qu'à un conteneur spécifique.

**3. Impact de la décision :** Attribuer la totalité du coût partagé aux deux conteneurs double le coût apparent.

**4. Pourquoi c'est faux :** La propriété économique du coût est au niveau de la réservation, tandis que l'analyse requiert une règle d'allocation contrôlée.

**5. Approche correcte :** Utiliser la politique configurée `EQUAL_SHARE` et diviser les COGS partagés par le nombre de conteneurs de la réservation.

**6. Décision corrigée :** Chaque conteneur reçoit ses coûts directs plus sa part allouée du coût commun, préservant le total au niveau de la réservation.

```text
Booking shared COGS = $1,000
Container count     = 2

Allocated shared cost per container
= $1,000 / 2
= $500
```

Le classeur sépare explicitement les coûts directs des conteneurs des coûts partagés au niveau de la réservation et utilise un dénominateur sûr du nombre de conteneurs. 

<details>
<summary>Logique de la formule</summary>

```excel
alloc_cogs =
IF(
  alloc_policy="EQUAL_SHARE",
  direct_cogs + (shared_cogs / cntr_count),
  direct_cogs
)
```

`Setup_Params!B6` contrôle si le coût partagé est alloué avec `EQUAL_SHARE` ou laissé non alloué avec `NO_ALLOCATE`. 

</details>

### Exemple de scénario

Considérez une réservation de transit avec les transactions valides suivantes :

| Article                      | Type     | Montant original | Devise | Équivalent USD |
| ------------------------- | -------- | --------------: | -------- | -------------: |
| Facture client          | Revenue  |           8,500 | USD      |         $8,500 |
| Coût fournisseur de fret maritime | COGS     |           3,000 | USD      |         $3,000 |
| Inspection à l'origine         | COGS     |             900 | USD      |           $900 |
| Assurance cargo           | Revenue  |             400 | USD      |           $400 |
| Frais bancaire               | Bank Fee |              75 | USD      |            $75 |
| Acompte récupérable       | Deposit  |             500 | USD      |           $500 |

Le P&L valide est :

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

L'acompte récupérable de `$500` est suivi séparément plutôt que traité comme un revenu d'exploitation ou un COGS.

Supposez maintenant qu'une transaction fournisseur supplémentaire soit saisie dans une devise non configurée. Le système n'ajoute **pas** silencieusement la transaction aux COGS. Son taux de change appliqué devient `0`, son calcul USD reste numérique et le champ de diagnostic signale `ERR: UNKNOWN_CURRENCY`. La réservation reçoit un avertissement d'intégrité au lieu d'être présentée comme propre.

La recommandation opérationnelle change en conséquence :

> **N'utilisez pas la marge de 55.34% comme chiffre de rentabilité final tant que la transaction non résolue n'est pas corrigée.**

C'est la distinction clé entre un calculateur de rentabilité et un flux de travail de calcul des coûts contrôlé : le résultat est accompagné de preuves sur le fait que les données sous-jacentes sont suffisamment complètes pour être dignes de confiance.

### Référence des formules

<details>
<summary>Booking Master — Nombre de conteneurs</summary>

**Objectif :** Déterminer le nombre de conteneurs attachés à chaque réservation et prévenir les échecs de division par zéro.

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

Si aucun conteneur n'a encore été saisi, le dénominateur se rabat sur `1`. C'est une sauvegarde de calcul, tandis que la condition de conteneur manquant doit encore être traitée comme un problème de qualité des données opérationnellement. 

</details>

<details>
<summary>Raw Transactions — Taux de change appliqué</summary>

**Objectif :** Prioriser un taux de change manuel et sinon récupérer le taux de devise configuré.

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

Une devise inconnue retourne `0`, pas `1.0`. 

</details>

<details>
<summary>Raw Transactions — Traitement financier</summary>

**Objectif :** Cartographier les types de transaction opérationnels vers des catégories financières.

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

La conception source utilise `COGS` comme repli conservateur pour un type de transaction non reconnu, réduisant le risque de surestimer le profit. 

</details>

<details>
<summary>Raw Transactions — Équivalent USD</summary>

**Objectif :** Convertir chaque montant de transaction original en devise de base USD du classeur tout en conservant un résultat numérique.

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

Garder cette colonne numérique est important parce que les calculs `SUMIFS` en aval ne devraient jamais recevoir de texte de diagnostic tel que `"INVALID"` à la place d'un nombre. 

</details>

<details>
<summary>Raw Transactions — Data Health Status</summary>

**Objectif :** Valider les liens de réservation, de service et de change indépendamment des colonnes de calcul financier.

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

La couche de diagnostic est intentionnellement séparée des champs de calcul numériques. 

</details>

<details>
<summary>Job Costing Engine — P&L au niveau de la réservation</summary>

**Objectif :** Générer une liste de réservations dynamique et n'agréger que les transactions saines.

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

Le contrôle essentiel est le critère répété `Data Health Status = "OK"`. 

</details>

### Règles de validation

| Champ / condition      | Règle                                                            | Comportement en cas d'erreur                                                                     |
| ---------------------- | --------------------------------------------------------------- | ---------------------------------------------------------------------------------- |
| `Booking ID`           | Doit exister dans `Booking_Master`                                  | `ERR: INVALID_BOOKING` ; transaction exclue du P&L                              |
| `Service Code`         | Doit exister dans `Service_Master`                                  | `ERR: INVALID_SERVICE` ; transaction exclue du P&L                              |
| `Currency`             | Doit avoir un taux de change configuré valide ou un taux de change manuel positif | Le change appliqué devient `0` ; le diagnostic signale `ERR: UNKNOWN_CURRENCY`                 |
| `Manual FX Rate`       | Une valeur positive remplace le taux de change de référence                  | Une valeur invalide/non positive se rabat sur le change configuré                             |
| `Transaction Type`     | Devrait correspondre au mappage `Setup_Params`                             | Un type inconnu se rabat conservativement sur `COGS`                                   |
| `USD Equivalent`       | Doit rester numérique                                             | Le texte de diagnostic ne doit jamais être écrit dans le champ d'agrégation numérique           |
| `Data Health Status`   | Doit être `OK` pour l'agrégation financière                          | Les lignes non OK sont exclues du P&L de la réservation                                          |
| `Container Count`      | Ne doit pas devenir zéro lorsqu'il sert de dénominateur d'allocation        | Se rabat sur `1` pour prévenir `#DIV/0!`                                             |
| Répartition des coûts partagés | Contrôlée par `Setup_Params!B6`                                 | `EQUAL_SHARE` alloue selon le nombre de conteneurs ; `NO_ALLOCATE` laisse uniquement les coûts directs |
| Rentabilité de la réservation  | Un revenu de zéro ne doit pas produire d'erreur de division               | La marge retourne `0`                                                                 |
| Intégrité des données exécutive  | Tout état non vide autre que `OK` est compté                 | Le tableau de bord affiche un avertissement au lieu de présenter un état propre non qualifié      |

La validation croisée source confirme la couverture des entrées principales de réservation, de conteneur, de service et de transaction, ainsi que des sorties d'équivalent USD, de bénéfice net, de marge et de coût alloué. Elle valide aussi spécifiquement la séparation entre le texte de diagnostic et les champs numériques `SUMIFS`. 

<!-- Technical Details parent <details> intentionally remains open for Part 2. -->
### Validation croisée et disponibilité au déploiement

L'architecture du classeur est conçue pour valider la chaîne analytique **avant** que la direction ne s'appuie sur le P&L qui en résulte.

L'implémentation source vérifie explicitement la chaîne de correspondance des types de transaction, la chaîne de recherche de change, la liaison Service Master, la chaîne de répartition des coûts partagés et la compatibilité entre le texte de diagnostic et les champs d'agrégation numériques. 

| Dimension de validation         | Ce qui est vérifié                                                                  | Résultat attendu                                          |
| ---------------------------- | -------------------------------------------------------------------------------- | -------------------------------------------------------- |
| **Liaison de réservation**          | Chaque `Booking ID` de transaction existe dans `Booking_Master`                        | Les lignes invalides reçoivent un état d'erreur                      |
| **Liaison de service**          | Chaque `Service Code` de transaction existe dans `Service_Master`                      | Les lignes invalides sont exclues du P&L                       |
| **Couverture de change**              | Chaque transaction a soit un taux de change manuel valide, soit un taux de référence configuré | Les devises inconnues déclenchent un avertissement d'intégrité              |
| **Classification financière** | Le type de transaction se mappe en Revenue / COGS / Bank Fee / Deposit                     | Les types inconnus utilisent le repli conservateur COGS         |
| **Agrégation numérique**      | `USD Equivalent` reste numérique en toutes conditions                            | `SUMIFS` ne rencontre pas de texte de diagnostic              |
| **Agrégation conditionnée par l'intégrité** | Le P&L n'inclut que `Data Health Status = "OK"`                                    | Les lignes sources défectueuses ne peuvent pas entrer silencieusement dans le profit déclaré |
| **Allocation des conteneurs**     | Les coûts partagés au niveau de la réservation utilisent la politique d'allocation configurée                  | Aucune attribution répétée du coût total entre les conteneurs       |
| **Protection revenu zéro**  | Le calcul de marge gère Revenu = 0                                    | Pas de `#DIV/0!`                                             |
| **Surveillance exécutive**     | Les erreurs de transaction globales sont comptées séparément                                 | Le tableau de bord signale les problèmes de données source non résolus          |

La matrice de validation croisée de l'implémentation indique que les champs d'entrée principaux — réservation, conteneur, service et données de transaction — et les sorties majeures — équivalent USD, bénéfice net, marge et coût alloué — sont couverts par l'architecture. 

### SOP opérationnelle

Le modèle opérationnel envisagé est délibérément simple :

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

#### 1. Établir les données de base

Créer la réservation dans `Booking_Master`.

Si les conteneurs sont pertinents, créer leurs relations réservation-conteneur dans `Container_Master`.

Maintenir le catalogue de services standardisé dans `Service_Master`.

Ne pas commencer l'analyse financière en tapant des noms de service libres ou des identifiants de réservation directement dans les feuilles de reporting. Les tables maîtresses sont la couche de référence pour les chaînes de recherche en aval. 

#### 2. Enregistrer les transactions financières

Saisir les revenus, coûts fournisseurs, frais bancaires, acomptes et autres activités financières pertinentes dans `Raw_Transactions`.

Chaque transaction doit contenir les éléments appropriés :

* `Booking ID`
* `Container ID`, le cas échéant
* `Service Code`
* `Transaction Type`
* `Allocation Level`
* `Original Amount`
* `Currency`
* `Manual FX Rate`, lorsqu'un taux spécifique à la transaction est requis

La table des transactions est la seule source qui devrait recevoir les enregistrements détaillés des événements financiers. Le moteur de calcul ne devrait pas être édité manuellement. 

#### 3. Vérifier l'intégrité des données

Examiner `Raw_Transactions[Data Health Status]`.

Une ligne propre retourne :

```text
OK
```

Les exceptions identifient la relation défectueuse, telles que :

```text
ERR: INVALID_BOOKING
ERR: INVALID_SERVICE
ERR: UNKNOWN_CURRENCY
```

L'implémentation traite délibérément ceci comme une couche de contrôle séparée. Une transaction peut donc rester numériquement sûre tout en étant visiblement exclue du résultat financier. 

#### 4. Examiner la rentabilité

Utiliser :

* `Job_Costing_Engine` pour la rentabilité au niveau de la réservation.
* `Booking_PL` pour la ventilation service par service d'une réservation.
* `Container_Analysis` lorsque la réservation contient plusieurs conteneurs.
* `Service_Profitability` pour l'économie des services transverse aux réservations.
* `Executive_Dashboard` pour la surveillance au niveau de la direction.

Cela crée un chemin d'exploration descendante de la **rentabilité globale → réservation → service → conteneur → transaction** plutôt que de forcer chaque question à être répondue depuis un seul tableau de bord de synthèse.

### Règles de maintenance des données

L'architecture s'appuie sur des données de base contrôlées parce qu'Excel ne fournit pas la même contrainte stricte de clé étrangère qu'une base de données ou une plateforme ERP.

Les règles de maintenance les plus importantes sont :

1. **Ne pas écraser les champs générés par formule.**
2. **Ne pas taper manuellement des noms de service là où un code de service devrait être utilisé.**
3. **Ne pas introduire une nouvelle devise sans d'abord ajouter son taux de change à `Setup_Params`.**
4. **Ne pas utiliser `1.0` comme repli de change manquant.**
5. **Ne pas écrire de texte d'erreur dans les colonnes d'agrégation numériques.**
6. **Ne pas éditer manuellement la sortie `Job_Costing_Engine`.**
7. **Résoudre les lignes de transaction non OK avant de traiter une marge de réservation comme finale.**
8. **Garder la politique de répartition des coûts partagés centralisée dans `Setup_Params`.**

La conception source interdit spécifiquement de coder en dur les paramètres globaux dans les formules et centralise la devise de base, les taux de change, les correspondances de transactions et la politique d'allocation dans `Setup_Params`. 

### Dépannage

#### `⚠️ X TXN DATA ERRORS` apparaît sur le tableau de bord

**Causes probables :**

* La devise est manquante dans la table de change.
* Le `Booking ID` n'existe pas dans `Booking_Master`.
* Le `Service Code` n'existe pas dans `Service_Master`.
* Une transaction source contient une liaison invalide.

**Résolution :**

Filtrer la colonne `Data Health Status` dans `Raw_Transactions` pour afficher les valeurs autres que `OK`.

Puis résoudre chaque problème de données source sous-jacent plutôt que de modifier manuellement la sortie du P&L.

Le tableau de bord exécutif est conçu pour exposer cette condition via une carte de santé globale :

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



#### Les formules de tableau dynamique retournent `#SPILL!`

Une formule de déversement nécessite que sa zone de destination reste dégagée.

Vérifier :

* les cellules directement sous la formule ;
* les cellules à droite de la formule ;
* les valeurs saisies manuellement occupant la plage de déversement prévue ;
* le formatage résiduel ou le contenu copié qui a créé une obstruction.

Effacer les cellules obstruantes et permettre au tableau dynamique de s'étendre à nouveau.

#### La rentabilité apparaît anormalement élevée

Ne pas modifier immédiatement la formule de profit.

Vérifier d'abord :

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

Une marge anormalement élevée peut être causée par une transaction de coût qui a échoué à la chaîne de liaison et a donc été exclue du calcul.

C'est exactement pourquoi le classeur distingue **« zéro comme état de calcul numérique »** de **« zéro comme conclusion économique »**. Un taux de change manquant peut produire `$0.00` dans un champ numérique tout en produisant simultanément `ERR: UNKNOWN_CURRENCY` dans le champ de diagnostic. 

#### La rentabilité des conteneurs apparaît surestimée ou sous-estimée

Vérifier trois choses dans l'ordre :

1. Le bon `Booking ID` est-il sélectionné ?
2. Le bon `Container ID` est-il sélectionné ?
3. `Setup_Params!B6` utilise-t-il la politique d'allocation prévue ?

Sous `EQUAL_SHARE`, les COGS partagés au niveau de la réservation sont divisés par le nombre de conteneurs de la réservation avant d'être ajoutés aux coûts directs des conteneurs.

Sous `NO_ALLOCATE`, les coûts partagés au niveau de la réservation ne sont pas attribués aux conteneurs individuels.

La distinction importe parce qu'attribuer la totalité du coût partagé au niveau de la réservation à chaque conteneur dupliquerait le même coût économique. 

### Limites de l'implémentation

Ce classeur est délibérément conçu comme une **couche de contrôle analytique légère**, et non comme un remplacement d'une plateforme comptable ou d'un système de gestion du transport.

Il est le plus fort lorsque :

* les données opérationnelles et financières existent déjà dans des tableurs ou des exports ;
* l'unité d'analyse clé est une réservation ou une expédition ;
* la direction a besoin de la rentabilité au niveau des missions ;
* les services et frais doivent être comparés entre les réservations ;
* plusieurs devises nécessitent une base de reporting cohérente ;
* les coûts partagés requièrent une règle d'allocation explicite ;
* les utilisateurs ont besoin de formules transparentes plutôt que d'une automatisation opaque.

Il n'est pas conçu pour fournir :

* la concurrence transactionnelle multi-utilisateurs ;
* des permissions de qualité ERP ;
* la réconciliation bancaire automatisée ;
* l'ingestion automatisée des factures transporteurs ;
* la contrainte de clé étrangère au niveau de la base de données ;
* l'orchestration de flux de travail d'entreprise.

La valeur du classeur n'est donc pas le nombre de feuilles qu'il contient. C'est la chaîne analytique contrôlée reliant **réservation → transaction → classification → conversion de devise → agrégation validée → rentabilité → exploration descendante**.

### Principe de contrôle analytique

La règle d'implémentation la plus importante peut être résumée ainsi :

```text
Never let a broken data relationship
look like a valid financial result.
```

Un tableur conventionnel pourrait permettre :

```text
Missing FX
    ↓
0 USD
    ↓
SUMIFS
    ↓
Higher Profit
```

Ce classeur rend plutôt l'état explicite :

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

Cette distinction est le fondement du modèle de fiabilité du moteur de calcul des coûts. L'architecture reconnaît explicitement que `SUMIFS` peut agréger en sécurité les valeurs numériques tout en ignorant silencieusement le texte, de sorte que les états de diagnostic ne doivent jamais être intégrés directement dans les colonnes de calcul numériques. 

### Séquence de déploiement recommandée

Pour une nouvelle implémentation de classeur, la séquence la plus sûre est :

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

Le résultat est un flux de travail reproductible plutôt qu'un calcul de tableur ponctuel : établir les données de référence une fois, ajouter de nouvelles transactions opérationnelles, résoudre les exceptions et laisser les couches de calcul et d'analyse se mettre à jour à partir du même modèle de données sous-jacent. 

<!-- Part 2 ends with the Technical Details parent <details> intentionally OPEN. Do not add </details> here. -->
### Notes finales d'architecture

Le classeur est délibérément construit autour d'un **moteur unique de rentabilité au niveau de la réservation**, plutôt que d'une collection de calculateurs déconnectés.

La hiérarchie analytique est :

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

Cette structure garde la réservation comme unité économique centrale tout en permettant aux utilisateurs de descendre dans les services et les conteneurs ou de remonter vers la rentabilité transverse aux réservations. 

L'implémentation reconnaît aussi explicitement trois chaînes de recherche à sauts multiples :

1. **Transaction → Traitement financier → FX → Équivalent USD → P&L de la réservation**
2. **Réservation → Nombre de conteneurs → Politique de coûts partagés → Répartition des coûts partagés → Analyse des conteneurs**
3. **Code de service → Service Master → Catégorie de service → Rentabilité du service**

Ces chaînes sont validées avant que les sorties analytiques ne soient considérées comme prêtes pour la production. 

### Ce contre quoi le moteur protège

Le principe de conception central n'est pas seulement l'automatisation. C'est **empêcher une relation d'entrée défectueuse de se déguiser en conclusion financière valide**.

| Échec                                              | Mécanisme de protection                                                           |
| ---------------------------------------------------- | ------------------------------------------------------------------------------ |
| Devise inconnue                                     | Le change prend par défaut `0`, tandis qu'un état de diagnostic séparé signale le taux manquant |
| Type de transaction inconnu                             | Le repli conservateur `COGS` prévient l'inflation accidentelle du profit              |
| Réservation invalide                                      | `ERR: INVALID_BOOKING`                                                         |
| Code de service invalide                                 | `ERR: INVALID_SERVICE`                                                         |
| Chaîne de transactions rompue                             | Les lignes non `OK` sont exclues de l'agrégation du P&L                                |
| Coût partagé dupliqué entre les conteneurs             | Politique d'allocation explicite et dénominateur de nombre de conteneurs                     |
| Revenu zéro                                         | Le calcul de marge retourne `0` au lieu de `#DIV/0!`                            |
| Texte de diagnostic contaminant les calculs financiers | Les colonnes de calcul numériques et les colonnes de diagnostic restent séparées             |

La matrice de repli de l'implémentation rejette spécifiquement un repli de change à `1.0` parce que cela pourrait transformer un montant en devise étrangère non reconnue en un montant USD implausiblement grand. 

### Modèle d'utilisation recommandé

Pour les opérations continues, le classeur est destiné à se comporter comme un **modèle opérationnel rafraîchissable**, et non comme un rapport ponctuel.

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

Cela rend le classeur utile pour des questions récurrentes telles que :

* Quelles réservations sont réellement rentables ?
* Quels services produisent les marges les plus fortes ?
* Quels coûts d'expédition provoquent une perte inattendue ?
* Quels coûts au niveau de la réservation doivent être alloués aux conteneurs ?
* Les marges déclarées sont-elles dignes de confiance, ou des problèmes de données source non résolus cachent-ils des coûts ?
* Sur quoi la tarification, la négociation fournisseur ou la revue opérationnelle devrait-elle se concentrer ensuite ?

La SOP source suit la même séquence : établir les données de base, enregistrer les transactions financières, examiner `Data Health Status`, puis utiliser les tableaux de bord de rentabilité pour la revue. 

### Principe de conception final

Le classeur ne tente pas de faire paraître l'analyse financière plus simple en cachant son incertitude.

Il rend plutôt l'incertitude visible.

Une réservation avec une marge calculée n'est pas automatiquement traitée comme une réservation digne de confiance.

Une **réservation propre** signifie :

```text
Valid Booking
+ Valid Service
+ Valid Currency / FX
+ Valid Transaction Classification
+ Healthy Transaction Chain
= Profitability Result That Can Be Reviewed
```

Cette distinction est le cœur de l'implémentation.

</details>

## Autres outils de cette série

Une collection d'outils légers d'aide à la décision sous Excel couvrant le contrôle opérationnel, l'analyse de rentabilité, la réconciliation des stocks, la planification de la main-d'œuvre et d'autres flux de travail métier.

* **[rolling-cashflow-forecast](https://github.com/HyVoid/rolling-cashflow-forecast)** — Pourquoi la trésorerie est-elle toujours serrée ? — Prévoyez la liquidité future sous des hypothèses opérationnelles et financières changeantes.
* **[Multi-Currency-Trade-Financial-Analyzer](https://github.com/HyVoid/Multi-Currency-Trade-Financial-Analyzer)** — Quels flux commerciaux sont réellement rentables après effets de change ? — Mesurez la rentabilité commerciale multidevises à travers des opérations transfrontalières.
* **[RTL_Loan_Sizer_Pricer](https://github.com/HyVoid/RTL_Loan_Sizer_Pricer)** — Quelle structure de financement est soutenable selon différents scénarios d'emprunteur et de garantie ? — Comparez des scénarios de prêt avec des hypothèses de financement ajustées au risque.
* **[Personal-Business-Dual-Track-Budget-Manager](https://github.com/HyVoid/Personal-Business-Dual-Track-Budget-Manager)** — Où la trésorerie personnelle et professionnelle est-elle mélangée ou mal allouée ? — Séparez les flux de trésorerie personnels et professionnels en pistes budgétaires suivies indépendamment.
* **[Multifamily-Value-Add-Acquisition-Model](https://github.com/HyVoid/Multifamily-Value-Add-Acquisition-Model)** — Cette acquisition immobilière crée-t-elle encore de la valeur après rénovation et hypothèses d'exploitation ? — Évaluez les rendements d'acquisition multifamiliale sous financement de rénovation et scénarios d'exploitation.
* **[Real-Estate-Valuation-all-in-one-calculator](https://github.com/HyVoid/Real-Estate-Valuation-all-in-one-calculator)** — Quelle est la valeur d'un bien selon les différentes méthodes d'évaluation ? — Calculez la valeur d'un bien selon plusieurs méthodes d'évaluation avec modélisation fiscale et des dépenses en capital.
* **[multi-entity-expense-allocation](https://github.com/HyVoid/multi-entity-expense-allocation)** — Les coûts partagés sont-ils alloués équitablement entre les entités ou certains actifs sont-ils surchargés tandis que d'autres échappent à leur véritable exposition aux coûts ? — Répartissez les factures fournisseurs entre plusieurs entités juridiques et actifs opérationnels selon des méthodologies d'allocation configurables.
* **[investment-pool-nav-excel-template](https://github.com/HyVoid/investment-pool-nav-excel-template)** — À combien s'élève réellement la part de chaque participant selon son moment d'entrée et la valeur actuelle du pool ? — Suivez la participation à un pool d'investissement en continu grâce à une comptabilité unitaire basée sur la NAV.
* **[shopify-cash-flow-forecast-excel-toolkit](https://github.com/HyVoid/shopify-cash-flow-forecast-excel-toolkit)** — Quels engagements opérationnels créeront un problème de trésorerie dans les douze prochaines semaines ? — Connectez les transactions bancaires de ventes Shopify, les dépenses publicitaires et les achats de stock dans un cadre unifié de décision de flux de trésorerie.
* **[Revenue-Management-Decision-Engine](https://github.com/HyVoid/Revenue-Management-Decision-Engine)** — Est-ce que suffisamment d'informations significatives ont changé pour justifier un ajustement de prix ? — Gouvernez les recommandations tarifaires via des garde-fous configurables qui préviennent la volatilité inutile des revenus.
* **[amazon-wbr-sqp-analytics-template](https://github.com/HyVoid/amazon-wbr-sqp-analytics-template)** — Qu'a réellement changé dans l'activité cette période et quels produits ou mots-clés en ont été les moteurs ? — Transformez les exports Amazon Seller Central SQP et Sellerboard en revues d'affaires exécutives standardisées.
* **[food-manufacturing-cost-calculator-excel](https://github.com/HyVoid/food-manufacturing-cost-calculator-excel)** — Quels produits sont réellement rentables après allocation de tous les coûts de fabrication et lesquels détruisent la marge ? — Allouez la main-d'œuvre d'usine, les services, les frais généraux et les coûts de matériaux à chaque SKU pour une véritable économie unitaire de fabrication.
* **[ecommerce-profit-calculator-excel](https://github.com/HyVoid/ecommerce-profit-calculator-excel)** — Quelles commandes sont réellement rentables après tous les coûts cachés et où l'argent fuit-il à travers les canaux et les devises ? — Reconstituez la véritable rentabilité au niveau des commandes en réconciliant les versements de ventes, les frais d'expédition, les frais de paiement et les commissions de marketplaces à travers des opérations e-commerce multicanal.
* **[construction-job-costing-excel-template](https://github.com/HyVoid/construction-job-costing-excel-template)** — Quels projets d'entrepreneur et quels clients sont réellement rentables après tous les coûts de chantier ? — Connectez l'estimation des prospects, l'exécution des chantiers et le suivi des coûts réels dans un cadre unifié de rentabilité des entrepreneurs.
* **[Flora-Pastry-Coffee-Shop-Operating-Control-System](https://github.com/HyVoid/Flora-Pastry-Coffee-Shop-Operating-Control-System)** — La boutique est-elle réellement rentable après tous les coûts d'ingrédients, de main-d'œuvre et d'exploitation ? — Connectez le calcul des coûts des recettes, les ventes, la consommation de stock et les charges d'exploitation dans un cadre unifié de rentabilité et de trésorerie.
* **[cross-entity-logistics-console](https://github.com/HyVoid/cross-entity-logistics-console)** — Quels conteneurs et quels clients sont réellement rentables à travers les opérations de douane et de flotte ? — Consolidez les données de dédouanement et d'exécution de flotte en une vue unique de rentabilité au niveau des conteneurs.
* **[family-loan-tracker-excel](https://github.com/HyVoid/family-loan-tracker-excel)** — Que doit-on réellement aujourd'hui après des années de transferts et d'intérêts accumulés ? — Réconciliez des transferts de fonds à long terme et des intérêts courus en un solde impayé actuel.
* **[3-statement-financial-model-excel](https://github.com/HyVoid/3-statement-financial-model-excel)** — À quoi ressembleront nos états financiers selon différents scénarios d'exploitation et de financement ? — Prévoyez des états financiers intégrés en reliant revenus, charges, fonds de roulement et dette dans un modèle unifié.
* **[assisted-living-cost-allocation-excel](https://github.com/HyVoid/assisted-living-cost-allocation-excel)** — Quelle part de la marge déclarée de chaque établissement est réellement générée par cet établissement plutôt que par la façon dont les coûts du siège ont été alloués ? — Allouez les dépenses partagées entre unités opérationnelles tout en préservant les transactions sources et en réconciliant les résultats au niveau des unités avec les totaux de l'entreprise.
* **[real-estate-development-feasibility-excel](https://github.com/HyVoid/real-estate-development-feasibility-excel)** — À quel prix maximal d'acquisition du foncier ce projet devient-il un investissement bancable ? — Évaluez les opportunités de développement résidentiel en filtrant les comparables fonciers, en estimant les coûts de construction et en calculant les seuils de profit sur coûts.
* **[uk-residential-costing-software](https://github.com/HyVoid/uk-residential-costing-software)** — Ce projet de construction atteindra-t-il réellement sa marge cible une fois pris en compte la logistique de chantier, les coefficients de déchets et les variations de tarifs corps d'état ? — Séparez les quantités physiques du bâtiment des tarifs corps d'état volatils et des frais préliminaires de chantier pour cascader les majorations commerciales à travers les budgets d'appel d'offres et les devis clients.
* **[b2b-pricing-landed-cost-excel-model](https://github.com/HyVoid/b2b-pricing-landed-cost-excel-model)** — Quels prix de gros et segments clients préservent réellement la marge une fois le fret, les droits, les conditions de paiement et le change entièrement coûtés ? — Calculez les coûts de revient au niveau SKU, le coût de service et les cascades de prix de marge nette pour simuler la rentabilité du commerce de gros transfrontalier.
* **[custom-label-pricing-calculator](https://github.com/HyVoid/custom-label-pricing-calculator)** — Que devrions-nous facturer pour une commande d'étiquettes personnalisées à chaque palier de quantité pour rester compétitif et rentable ? — Calculez la tarification paramétrique des étiquettes en modélisant les coûts au pouce carré des substrats, les frais de mise en route et les remises par palier de quantité par rapport aux références concurrentielles.
* **[landed-cost-calculator-retail-deduction-excel](https://github.com/HyVoid/landed-cost-calculator-retail-deduction-excel)** — Quelle est la véritable marge nette sur chaque SKU une fois les coûts de revient réels et les déductions des détaillants pris en compte ? — Réconciliez les coûts de revient estimés sur PO avec le fret réel, les droits, le change et le camionnage tout en auditant les déductions de reversement des détaillants pour révéler la marge nette au niveau SKU.
* **[saas-finance-mrr-burn-rate-excel](https://github.com/HyVoid/saas-finance-mrr-burn-rate-excel)** — Pouvons-nous honorer cette semaine les engagements d'embauche, de marketing et fournisseurs sans manquer de trésorerie avant la clôture ? — Suivez le rythme hebdomadaire de consommation de trésorerie, l'écart du waterfall MRR/ARR, les métriques de rétention et les coûts de paie par rapport au budget avant la clôture de fin de mois.
* **[startup-financial-model-excel](https://github.com/HyVoid/startup-financial-model-excel)** — Notre plan de croissance manquera-t-il de trésorerie avant que la monétisation des cohortes ne suive ? — Prévoyez la consommation de trésorerie pluriannuelle, la runway, l'économie unitaire et la dilution de levée de fonds à partir d'hypothèses ascendantes de cohortes et d'effectifs.
* **[shopify-pnl-reconciliation-excel](https://github.com/HyVoid/shopify-pnl-reconciliation-excel)** — Chaque boutique Shopify est-elle réellement rentable une fois les véritables coûts au niveau SKU, les dépenses partagées et les frais bancaires réconciliés ? — Réconciliez les exports de ventes Shopify multi-boutiques avec les COGS au niveau SKU, les OPEX multidevises alloués et les relevés bancaires pour produire un P&L mensuel digne de confiance.
* **[travel-agency-accounting-excel-template](https://github.com/HyVoid/travel-agency-accounting-excel-template)** — Quelles réservations de circuits génèrent réellement un bénéfice net et un flux de trésorerie positif après coûts fournisseurs et encaissements B2B ? — Calculez les coûts de mission par réservation de circuit, les marges brutes, le règlement AR/AP et la trésorerie d'exploitation à travers plusieurs devises.

## Licence

Ce projet est publié sous la **licence Apache 2.0**.

Vous pouvez utiliser, modifier, distribuer et bâtir sur l'œuvre conformément aux termes de la licence Apache 2.0.
