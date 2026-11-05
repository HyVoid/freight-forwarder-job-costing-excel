[ 🌐 عربي ](README.ar.md) | [ 🇩🇪 Deutsch ](README.de.md) | [ 🇫🇷 Français ](README.fr.md) | [ 🇳🇱 Nederlands ](README.nl.md) | [ 🇪🇸 Español ](README.sp.md) | [ 🇬🇧 English ](README.md)

# Freight Forwarder Job Costing Excel Template: Shipment Profit & Loss (P&L) Calculator, Container Cost Allocation & Margin Tracker
<img width="2880" height="1228" alt="image" src="https://github.com/user-attachments/assets/f89c1a0f-f576-43af-b26a-02bae4a47348" />


[![License: Apache 2.0](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](LICENSE) [![Platform](https://img.shields.io/badge/Platform-Web%20Browser%20%2B%20Microsoft%20Excel%20(.xlsx)-217346.svg)](#) [![Tool Type](https://img.shields.io/badge/Tool%20Type-Plug--and--Play%20Freight%20Accounting%20Spreadsheet-2251FF.svg)](#)

**A ready-to-use Freight Forwarder Job Costing Excel Template (`.xlsx`) and Free Online Shipment Profitability Calculator built for freight forwarders, NVOCCs, customs brokers, and 3PL operators who need to calculate net profit per shipment without writing formulas.** Track **Bill of Lading (B/L) and Air Waybill (AWB) gross margins**, automate **multi-currency FX conversions (USD, EUR, CNY, GBP)**, split **FCL/LCL container costs**, and reconcile **customer AR invoices against ocean carrier, drayage, customs brokerage, and bank wire AP bills** in one pre-built logistics profit and loss spreadsheet.

**Test it free in your browser right now, or download the unlocked Excel file to paste your own shipment data.**

Every lookup rule, currency conversion, and container cost-splitting formula is already built in. Paste your billing rows from QuickBooks, Xero, or your operational tracking sheet to see your exact profit per booking, container, and freight charge code in seconds, backed by a 30-day money-back guarantee.

> 🌐 **Try the Online Calculator First** → [Launch the Free Interactive Freight Forwarder Job Costing & Shipment P&L Web Calculator (Browser Demo, Zero Setup Required)](https://hyvoid.github.io/freight-forwarder-job-costing-excel/)  
> 📥 **Download the Ready-to-Use Excel Template** → [Get the Unlocked Freight Forwarder Job Costing & Shipment Profitability Excel Template (.xlsx with Auto-Updating Formulas)](https://www.theseusworkshop.com/l/uyukd?utm_source=github&utm_medium=GitHub%20README&utm_campaign=readme%20new%20launch&utm_content=freight-booking-job-costing)  
>
> *Both the free web calculator and the downloadable Microsoft Excel workbook run on the exact same pre-configured freight accounting tables and automated margin formulas.*

---

## What It Helps You Track: Freight Forwarding Pain-Point to Solution Mapping

Waiting for month-end accounting reports hides which individual shipments actually made or lost money. This **plug-and-play freight job costing spreadsheet** fixes the five most common operational accounting blind spots in ocean freight, air cargo, and multimodal forwarding:

| High-Frequency Search Topic & Tracking Goal | The Daily Freight Accounting Pain Point (Why Margin Leaks) | How This Automated Excel Template Fixes & Tracks It |
| :--- | :--- | :--- |
| **Profit Per Shipment & Booking-Level P&L (`B/L` & `AWB` Margin)** | QuickBooks and Xero group invoices by calendar month, hiding money-losing shipments, underpriced quotes, and unprofitable trade lanes. | Matches **Customer AR Revenue, Vendor AP Bills, and Bank Wire Fees** to each individual `Booking_ID` so you see real-time **Net Profit ($) and Gross Margin (%) per shipment**. |
| **Freight Charge Code & Accessorial Profitability** | You know the total invoice amount, but cannot tell if you lost money on **Ocean Freight, Local Drayage, THC, Customs Clearance, Demurrage, or Cargo Insurance**. | Breaks down revenue vs. vendor cost across granular **Freight Service Codes** (`Ocean Freight`, `THC`, `Origin Inspection`, `Drayage`, `Customs`, `Documentation`) to expose hidden cost overruns. |
| **FCL Multi-Container Shared Cost Allocation (`TEU` / `FEU` Split)** | When one Master Bill of Lading (MBL) covers multiple containers, flat booking fees (like customs entry or B/L documentation) get double-counted or left out of per-container pricing. | Separates **direct per-container costs** from **shared booking-level fees**, automatically dividing shared costs evenly across all active containers in that booking. |
| **Multi-Currency Freight Invoice Conversion (`USD`, `EUR`, `CNY`, `GBP`)** | Paying ocean carriers in USD, origin agents in CNY or EUR, and billing clients in local currency using manual calculator rates causes month-end exchange rate losses. | Converts every mixed-currency AR and AP line item into your **single Base Reporting Currency** using a central FX lookup table, flagging any unconfigured currency code immediately. |
| **Missing Vendor Bills & Broken Spreadsheet Typo Detection** | A single typo in a Booking ID inside a manual `SUMIFS` or `VLOOKUP` sheet silently drops a $2,000 carrier bill to `$0`, inflating profit and overpaying sales commissions. | Runs an automatic **Data-Health Check** (`ERR: UNKNOWN_CURRENCY`, `ERR: INVALID_BOOKING`) that highlights broken rows in plain sight while keeping your dashboard totals intact. |

---

## How to Calculate Shipment Profit & Run a Freight Job Costing Audit (4-Step Beginner Tutorial)

You do not need accounting training or Excel formula skills to audit your shipment margins. Follow these four steps to turn raw invoice lists and carrier bills into an accurate **Booking-by-Booking Profit and Loss statement**:

### Step 1: Set Your Base Currency, Exchange Rates & Container Split Policy
Open the `Setup_Params` tab to set your company's baseline rules in plain English:
- Choose your **Base Reporting Currency** (such as `USD`, `EUR`, or `GBP`) and enter your standard **Foreign Exchange (FX) Rates**.
- Review the pre-mapped **Transaction Categories** (`Customer Revenue`, `Direct Vendor COGS`, `Bank Wire Fees`, and `Customer Deposits`).
- Select your **Shared-Cost Allocation Rule** for multi-container FCL shipments.

*Try it live:* Preview how the setup table works right now in the [Free Online Freight Forwarding Job Costing Browser Calculator](https://hyvoid.github.io/freight-forwarder-job-costing-excel/).

---

### Step 2: Paste Your Bookings, Containers & Raw AR/AP Invoice Lines
Drop your shipment IDs into the master lists (`Bookings` and `Containers`), then copy and paste your billing rows directly into the `Raw_Transactions` table:
- **Works with any data source:** QuickBooks Online / Desktop exports, Xero CSV dumps, CargoWise or Magaya billing logs, ocean carrier freight bills, customs broker statements, and manual dispatch trackers.
- **Zero manual formula dragging:** The Excel tables expand automatically when you paste new rows.

---

### Step 3: Check Automatic Error Alerts & View Your Shipment P&L Dashboard
The moment you paste your data, the spreadsheet calculates everything in the background:
- Converts foreign-currency vendor bills and customer invoices into your base currency.
- Flags any misspelled Booking IDs or missing currency rates in the `Data Health Status` column so an unlinked vendor bill never hides as a zero-cost item.
- Updates your **Booking-Level Profit & Loss Summary**, **Freight Service-Code Margin Table**, and **Per-Container Unit Cost Breakdown**.

---

### Step 4: Download the Unlocked Excel Template for One-Click Monthly Reuse
The free browser calculator is great for a quick single-session check. When you need to save your data locally, import hundreds of monthly accounting rows, and track ongoing vessel sailings across billing cycles, switch to the downloadable `.xlsx` workbook.

> 💡 **Why Busy Forwarders Download the Excel File Instead of Building Their Own Sheet:**  
> Cross-currency freight reconciliation, multi-container cost splitting, and orphan bill detection involve complex accounting logic under the hood. **In this Excel template, every matching rule, lookup link, and auto-expanding formula is already built and tested for you.**  
> 
> You do not even need to click and drag cell corners to copy formulas down. Simply download the file to your computer, paste your raw billing data, and get instant, audit-ready profit numbers. During peak shipping seasons, spend your time winning shipper quotes, securing vessel space, and moving cargo—while letting a simple copy-and-paste template handle your financial analysis in seconds.  
> 
> 👉 **[Download the Unlocked Freight Forwarder Job Costing Excel Template (.xlsx) and Automate Your Shipment P&L Today](https://www.theseusworkshop.com/l/uyukd?utm_source=github&utm_medium=GitHub%20README&utm_campaign=readme%20new%20launch&utm_content=freight-booking-job-costing)**

---

## The Business Logic & Methodology: 4 Hidden Profit Blind Spots This Excel Template Solves for You

Most freight forwarders lose money on shipments not because they quoted the wrong ocean rate, but because standard accounting tools and homemade Excel sheets hide where the money actually goes. Here is how this template protects your bottom line in plain terms—even if you have zero formal accounting background:

### 1. Monthly Accounting Software Hides Money-Losing Shipments (Booking-Level Matching vs. Monthly Totals)
**The Objective Fact:** Standard accounting software like QuickBooks or Xero organizes money by **calendar month**, while freight forwarding happens by **Booking, Bill of Lading (B/L), and Container**. If you invoice a customer $8,000 in October, but the ocean carrier and drayage trucker bill you $7,500 in November, October looks artificially profitable and November looks like a loss.  
**Your Direct Benefit:** This template groups every dollar of revenue, vendor cost, and bank wire fee by `Booking_ID` regardless of what month the invoice arrives. You see the true net profit of every single shipment before you renew a client's rate card or pay out sales commissions.

### 2. Flat Booking Fees vs. Per-Container Costs Distort Your Quotes (Automatic FCL Container Cost Splitting)
**The Objective Fact:** In ocean and multimodal freight, some costs happen once per **Booking** (like a $150 customs clearance fee or B/L documentation charge), while other costs happen per **Container** (like $450 terminal handling charges or local drayage per 40'HC box). In a basic spreadsheet, operators either multiply the $150 customs fee across 5 containers by mistake (inflating cost estimates by $600 and losing the bid) or forget to include shared booking fees when calculating per-container profit.  
**Your Direct Benefit:** The template separates container-specific costs from booking-wide shared costs and splits shared fees evenly across the exact number of containers in that booking:

```text
Single Container Total Cost = Direct Container Cost + (Shared Booking Costs ÷ Total Containers in Booking)

```

You get exact per-container unit costs without doing manual division.

### 3. One Typo in a Homemade Spreadsheet Silently Fakes High Profit (Built-In Typo & Missing-Cost Detection)

**The Objective Fact:** In a standard Excel sheet built with `SUMIFS`, a single typo in a Booking reference number on a $1,500 drayage vendor bill causes Excel to ignore that bill completely and add `$0` to your costs. Conversely, a broken `VLOOKUP` turns your entire summary table into `#N/A` errors. Look at how a single unlinked vendor invoice distorts a real booking:

```text
Reported Booking Revenue:   $8,500
Reported Vendor Costs:      $5,200  <-- ($1,500 drayage bill skipped due to a Booking ID typo)
Apparent Net Profit:        $3,300  <-- (Overstated by $1,500)
Apparent Profit Margin:      38.8%  <-- (Fake margin; actual real margin is only 21.1%)

```

If you trust that 38.8% number, you overpay sales rep commissions out of your own pocket and keep quoting a lane that barely breaks even.

**Your Direct Benefit:** Every row you paste goes through an automatic **Data-Health Check**. If a row has a misspelled Booking ID, an invalid charge code, or a missing currency rate, the template flags it clearly (`ERR: INVALID_BOOKING` or `ERR: UNKNOWN_CURRENCY`) while keeping your summary math running cleanly. You catch missing vendor bills before they cost you money.

### 4. Foreign Currencies and Upfront Customer Deposits Distort Real Gross Margin (Row-Level FX & Deposit Isolation)

**The Objective Fact:** Cross-border shipments routinely combine USD ocean freight, CNY or EUR origin agent bills, intermediary bank wire fees, and upfront customer deposits. Converting currency totals by hand at the end of the month introduces exchange-rate errors, and counting upfront customer deposits as earned revenue makes a shipment look profitable before vendor bills are even paid.

**Your Direct Benefit:** The template converts every foreign-currency bill into your base currency at the individual row level and separates **Customer Deposits (Cash Held)** from **Earned Freight Revenue and COGS**. You see both your real profit margin and your actual cash position side by side.

---

## Manual Freight Spreadsheets & Generic Accounting Software vs. This Automated Shipment Job Costing Excel Template

| Operational Comparison Factor | Manual Excel Sheets & Generic Accounting (QuickBooks / Xero Alone) | Standardized Freight Forwarder Job Costing Excel Template (This Tool) |
| --- | --- | --- |
| **Per-Shipment Profit Visibility (`MBL` / `HBL` / `AWB`)** | Blends all monthly invoices into general ledger buckets, hiding which specific bookings, containers, or shippers lost money. | Generates an automatic **Booking-by-Booking P&L Table** showing Gross Revenue, Vendor COGS, Bank Wire Fees, Net Profit ($), and Margin (%). |
| **Multi-Currency Invoice Conversion (`USD`, `EUR`, `CNY`, `GBP`)** | Staff convert foreign agent bills with desk calculators and inconsistent rates, causing unexplained margin gaps at month-end close. | Uses one **Centralized FX Rate Lookup Table** to convert every invoice line automatically and displays an instant warning if a currency rate is missing. |
| **FCL Multi-Container Cost Splitting (`20'GP` / `40'HC` / `TEU`)** | B/L document fees, customs brokerage, and booking surcharges get duplicated per container or omitted from container unit cost calculations. | Automatically separates **Direct Container Bills** from **Shared Booking Overheads**, splitting shared fees across the exact container count per booking. |
| **Accessorial & Charge-Code Margin (`THC`, `Drayage`, `Customs`)** | You see total shipment billing, but cannot tell if **Origin Inspection, Inland Drayage, Demurrage, or Cargo Insurance** is leaking profit. | Rolls up revenue vs. vendor cost by **Freight Service Code** so you can spot which accessorial charges make money and which ones lose money. |
| **Formula Maintenance & Typo Protection (`VLOOKUP` / `SUMIFS`)** | Staff have to drag formulas down manually for every new row, and a single misspelled Booking ID silently drops vendor costs to zero. | **Pre-built Auto-Expanding Tables & Error Alerts** update automatically on paste and quarantine typos (`ERR: INVALID_BOOKING`) before they corrupt your P&L. |

---

## Who This Is For: Logistics Roles, Spreadsheet Use Cases & Template Workflows

This toolkit is built for **small-to-midsize freight forwarders, OTI/NVOCC operators, customs brokers, 3PL teams, and freight agents** who want immediate shipment-level profit visibility without spending $20,000+ and six months setting up an enterprise ERP or TMS:

* **Freight Forwarding Owners & NVOCC Managing Directors (`Freight Forwarder Profit & Loss Excel Template`):** Audit real net profit per shipment (after bank wire fees and shared booking overheads) before approving annual shipper rate cards or paying monthly sales rep commissions.
* **Junior Freight Operations Coordinators & Dispatchers (`Plug-and-Play Shipment Costing Spreadsheet`):** Track multi-container FCL and LCL job costs during busy shipping seasons by pasting raw vendor bills and customer charges into a pre-built template with zero formula writing.
* **Logistics Bookkeepers & Freight Accounting Clerks (`Freight AR/AP Invoice Reconciliation Template`):** Bridge the gap between QuickBooks, Xero, or NetSuite exports and operational shipment files—catching unbilled customer accessorials or unlinked carrier bills before closing the monthly books.
* **Freight Pricing Analysts & Trade-Lane Managers (`Container Unit Cost & Accessorial Margin Calculator`):** Analyze whether Ocean Freight, Air Cargo, Customs Brokerage, Cargo Insurance, or Local Drayage is driving or shrinking margins across specific origin-to-destination lanes.
* **Fractional CFOs & Logistics Consultants (`3PL & Freight Margin Leakage Audit Model`):** Drop client billing exports into a standardized spreadsheet to uncover underpriced accounts, foreign-exchange losses, and unrecovered accessorial fees in a single afternoon.

*Note: This Excel model complements your operational TMS (such as CargoWise, Magaya, or GoFreight) and tax accounting software. It acts as your fast, plug-and-play **shipment profitability and job costing layer** when you need clear answers from your billing data right away.*

---

## Common Freight Costing Scenarios: When to Use This Shipment P&L Spreadsheet

Whether you are running a five-person forwarding office or managing hundreds of monthly FCL containers, this template solves the exact day-to-day costing problems that slow down logistics teams:

* **How to Calculate Profit Per Shipment in Excel Without Writing Formulas:** Paste your customer invoice lines and vendor bill lines into one table, tag each row with its `Booking_ID`, and let the pre-built summary sheet calculate Net Profit and Gross Margin % automatically.
* **How to Split Shared Bill of Lading (B/L) Costs Across Multiple FCL Containers:** Assign container-specific charges (like drayage or THC) to individual `Container_ID` numbers and leave booking-wide charges (like customs clearance or documentation) at the `Booking_ID` level—the template splits shared costs across your 20'GP and 40'HC containers automatically.
* **How to Reconcile Delayed Ocean Carrier & Drayage Bills Against Customer Invoices:** Keep open bookings in your log across calendar months so late arrival notices, detention/demurrage bills, or chassis split charges automatically deduct from the right shipment's profit margin when they arrive.
* **How to Track Multi-Currency Freight Margins Across USD, EUR, and CNY:** Enter your reference exchange rates once in `Setup_Params` to normalize overseas origin agent invoices, ocean carrier freight, and local delivery billing into a single reporting currency.
* **How to Audit Freight Broker & Sales Rep Gross Profit Commissions:** Verify that every vendor bill, origin fee, and bank wire charge is linked and accounted for before calculating sales commission payouts on shipment gross profit.

---

## About Theseus Workshop

**Theseus Workshop** builds plug-and-play Microsoft Excel financial models and browser-based calculators for operators who need clear, reliable numbers today—without enterprise software bloat or complex spreadsheet programming.

Every template is designed around one practical goal:

> **"Put the exact operational and financial numbers in one pre-built table so your next business decision takes seconds instead of hours."**

If you want to stop wrestling with broken `VLOOKUP` sheets and manual currency math, grab the full Excel template below or explore our complete library of operational spreadsheets and business calculators:

* 📥 **Get This Freight Costing Template:** [Download the Unlocked Freight Forwarder Job Costing & Shipment P&L Excel Template (.xlsx)](https://www.theseusworkshop.com/l/uyukd?utm_source=github&utm_medium=GitHub%20README&utm_campaign=readme%20new%20launch&utm_content=freight-booking-job-costing)
* 🌐 **Test the Free Web Calculator:** [Open the Interactive Freight Forwarding Job Costing Browser App](https://hyvoid.github.io/freight-forwarder-job-costing-excel/)
* 🧰 **Explore All Ready-to-Use Excel Models & Business Toolkits:** [Visit the Official Theseus Workshop Store at TheseusWorkshop.com](https://www.google.com/search?q=https://www.theseusworkshop.com%3Futm_source%3Dgithub%26utm_medium%3DGitHub%2520README%26utm_campaign%3Dreadme%2520new%2520launch%26utm_content%3Dabout-section-store-link)

---

## Technical Details

<details>
<summary>For technical reviewers, Excel practitioners, and collaborators</summary>

### Workbook Architecture

The workbook is organized around four functional layers: **rules and parameters → master data → transaction input → calculation and analytical output**.

The design contains 11 Sheets:

| Layer              | Sheet                   | Role                                                                             |
| ------------------ | ----------------------- | -------------------------------------------------------------------------------- |
| Rules & Parameters | `Instructions`          | Operating instructions, data-entry rules, and business definitions               |
| Rules & Parameters | `Setup_Params`          | Base currency, FX rates, transaction mappings, and shared-cost allocation policy |
| Master Data        | `Booking_Master`        | Booking index, customer, status, and container count                             |
| Master Data        | `Container_Master`      | Booking-to-container 1:N relationship                                            |
| Master Data        | `Service_Master`        | Standard Service Code, service name, and category                                |
| Transaction Input  | `Raw_Transactions`      | Unified financial transaction ledger plus data-health diagnostics                |
| Calculation Engine | `Job_Costing_Engine`    | Booking-level Revenue, COGS, Bank Fees, Profit, Margin, and Deposit aggregation  |
| Analysis           | `Booking_PL`            | Single-Booking P&L and service-level drill-down                                  |
| Analysis           | `Container_Analysis`    | Container-level direct and allocated shared costs                                |
| Analysis           | `Service_Profitability` | Cross-Booking profitability by service                                           |
| Executive Output   | `Executive_Dashboard`   | Global KPIs, profitability rankings, and data-health warnings                    |

The core flow is:

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

The source architecture identifies three important multi-hop lookup chains: a four-hop currency/financial-treatment chain, a four-hop shared-cost/container drill-down chain, and a three-hop service-classification chain. 

The critical design rule is that **calculation values and diagnostic states remain separate**. `USD Equivalent` stays numeric, while `Data Health Status` carries errors such as `ERR: UNKNOWN_CURRENCY`. This preserves `SUMIFS` compatibility while preventing broken transactions from being included in the P&L engine. 

### Three Traps That Catch Even Experienced Freight-Forwarding Operators

#### Trap 1 — Treating a Missing FX Rate as a Valid Amount

**1. Decision:** A Booking appears to have a certain cost and margin.

**2. Hidden flaw:** A vendor transaction is denominated in a currency that is not configured in the reference FX table.

**3. Decision impact:** If the missing rate is incorrectly treated as `1.0`, a foreign amount such as `50,000` can become `$50,000 USD`, radically distorting the Booking result.

**4. Why it is wrong:** A missing exchange rate is a data-integrity failure, not an economic exchange rate.

**5. Correct approach:** Return a numeric neutral value of `0` for the calculation field while separately reporting `ERR: UNKNOWN_CURRENCY`.

**6. Corrected decision:** The Booking is flagged for data completion instead of being assigned a false profit or loss.

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
<summary>Formula logic</summary>

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

The important rule is **never default an unknown currency to 1.0**. 

</details>

#### Trap 2 — Reading a Booking P&L Before Checking Data Health

**1. Decision:** A Booking with `$8,500` revenue and `$5,200` COGS is considered a `$3,300` profit.

**2. Hidden flaw:** One or more source transactions have invalid Booking or Service links, or an unresolved currency.

**3. Decision impact:** The apparent margin may be overstated because a cost line has silently failed to enter the aggregation.

**4. Why it is wrong:** A mathematically correct `SUMIFS` result can still be economically wrong if the source rows are incomplete.

**5. Correct approach:** The engine aggregates only transactions whose `Data Health Status = "OK"` and reports the number of invalid rows separately.

**6. Corrected decision:** The Booking is treated as **provisional** until the data exceptions are resolved.

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
<summary>Formula logic</summary>

```excel
rev =
SUMIFS(
  Raw_Transactions!O2:O10000,
  Raw_Transactions!B2:B10000,bk,
  Raw_Transactions!N2:N10000,"Revenue",
  Raw_Transactions!P2:P10000,"OK"
)
```

The same health filter is applied to COGS, Bank Fees, and Deposits. 

</details>

#### Trap 3 — Assigning Booking-Level Shared Cost Directly to Every Container

**1. Decision:** A two-container Booking needs a container-level profitability view.

**2. Hidden flaw:** A Booking-level COGS line belongs to the shipment as a whole rather than to a specific container.

**3. Decision impact:** Assigning the full shared cost to both containers doubles the apparent cost.

**4. Why it is wrong:** The cost's economic ownership is Booking-level, while the analysis requires a controlled allocation rule.

**5. Correct approach:** Use the configured `EQUAL_SHARE` policy and divide shared COGS by the Booking's container count.

**6. Corrected decision:** Each container receives its direct costs plus its allocated share of the common cost, preserving the Booking-level total.

```text
Booking shared COGS = $1,000
Container count     = 2

Allocated shared cost per container
= $1,000 / 2
= $500
```

The workbook explicitly separates direct container costs from Booking-level shared costs and uses a safe container-count denominator. 

<details>
<summary>Formula logic</summary>

```excel
alloc_cogs =
IF(
  alloc_policy="EQUAL_SHARE",
  direct_cogs + (shared_cogs / cntr_count),
  direct_cogs
)
```

`Setup_Params!B6` controls whether the shared cost is allocated using `EQUAL_SHARE` or left unallocated with `NO_ALLOCATE`. 

</details>

### Example Scenario

Consider a freight-forwarding Booking with the following valid transactions:

| Item                      | Type     | Original Amount | Currency | USD Equivalent |
| ------------------------- | -------- | --------------: | -------- | -------------: |
| Customer invoice          | Revenue  |           8,500 | USD      |         $8,500 |
| Ocean freight vendor cost | COGS     |           3,000 | USD      |         $3,000 |
| Origin inspection         | COGS     |             900 | USD      |           $900 |
| Cargo insurance           | Revenue  |             400 | USD      |           $400 |
| Bank charge               | Bank Fee |              75 | USD      |            $75 |
| Recoverable deposit       | Deposit  |             500 | USD      |           $500 |

The valid P&L is:

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

The `$500` recoverable deposit is tracked separately rather than treated as operating revenue or COGS.

Now assume one additional vendor transaction is entered in an unconfigured currency. The system does **not** silently add the transaction to COGS. Its applied FX rate becomes `0`, its USD calculation remains numeric, and the diagnostic field reports `ERR: UNKNOWN_CURRENCY`. The Booking receives a health warning instead of being presented as clean.

The operational recommendation changes accordingly:

> **Do not use the 55.34% margin as a final profitability figure until the unresolved transaction is corrected.**

This is the key distinction between a profitability calculator and a controlled job-costing workflow: the result is accompanied by evidence about whether the underlying data is sufficiently complete to trust.

### Formula Reference

<details>
<summary>Booking Master — Container Count</summary>

**Purpose:** Determine the number of containers attached to each Booking and prevent division-by-zero failures.

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

If no container has yet been entered, the denominator falls back to `1`. This is a calculation safeguard, while the missing-container condition should still be treated as a data-quality issue operationally. 

</details>

<details>
<summary>Raw Transactions — Applied FX Rate</summary>

**Purpose:** Prioritize a manual FX rate and otherwise retrieve the configured currency rate.

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

An unknown currency returns `0`, not `1.0`. 

</details>

<details>
<summary>Raw Transactions — Financial Treatment</summary>

**Purpose:** Map operational transaction types to financial categories.

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

The source design uses `COGS` as the conservative fallback for an unrecognized transaction type, reducing the risk of overstating profit. 

</details>

<details>
<summary>Raw Transactions — USD Equivalent</summary>

**Purpose:** Convert each original transaction amount into the workbook's USD base currency while retaining a numeric result.

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

Keeping this column numeric is important because downstream `SUMIFS` calculations should never receive diagnostic text such as `"INVALID"` in place of a number. 

</details>

<details>
<summary>Raw Transactions — Data Health Status</summary>

**Purpose:** Validate Booking, Service, and FX linkage independently from financial calculation columns.

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

The diagnostic layer is intentionally separate from numeric calculation fields. 

</details>

<details>
<summary>Job Costing Engine — Booking-Level P&L</summary>

**Purpose:** Generate a dynamic Booking list and aggregate only healthy transactions.

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

The essential control is the repeated `Data Health Status = "OK"` criterion. 

</details>

### Validation Rules

| Field / Condition      | Rule                                                            | Error Behavior                                                                     |
| ---------------------- | --------------------------------------------------------------- | ---------------------------------------------------------------------------------- |
| `Booking ID`           | Must exist in `Booking_Master`                                  | `ERR: INVALID_BOOKING`; transaction excluded from P&L                              |
| `Service Code`         | Must exist in `Service_Master`                                  | `ERR: INVALID_SERVICE`; transaction excluded from P&L                              |
| `Currency`             | Must have a valid configured FX rate or positive manual FX rate | Applied FX becomes `0`; diagnostic reports `ERR: UNKNOWN_CURRENCY`                 |
| `Manual FX Rate`       | Positive value overrides the reference FX rate                  | Invalid/non-positive value falls back to configured FX                             |
| `Transaction Type`     | Should match `Setup_Params` mapping                             | Unknown type conservatively falls back to `COGS`                                   |
| `USD Equivalent`       | Must remain numeric                                             | Diagnostic text must never be written into the numeric aggregation field           |
| `Data Health Status`   | Must be `OK` for financial aggregation                          | Non-OK rows are excluded from Booking P&L                                          |
| `Container Count`      | Must not become zero when used as allocation denominator        | Falls back to `1` to prevent `#DIV/0!`                                             |
| Shared-cost allocation | Controlled by `Setup_Params!B6`                                 | `EQUAL_SHARE` allocates by container count; `NO_ALLOCATE` leaves direct costs only |
| Booking profitability  | Revenue of zero must not produce a division error               | Margin returns `0`                                                                 |
| Executive data health  | Any non-empty status other than `OK` is counted                 | Dashboard displays a warning instead of presenting an unqualified clean state      |

The source cross-validation confirms coverage of the core Booking, Container, Service, and transaction inputs, plus USD Equivalent, Net Profit, Margin, and allocated cost outputs. It also specifically validates the separation between diagnostic text and numeric `SUMIFS` fields. 

<!-- Technical Details parent <details> intentionally remains open for Part 2. -->
### Cross-Validation and Deployment Readiness

The workbook architecture is designed to validate the analytical chain **before** management relies on the resulting P&L.

The source implementation explicitly checks the transaction-type mapping chain, FX lookup chain, Service Master linkage, shared-cost allocation chain, and the compatibility between diagnostic text and numeric aggregation fields. 

| Validation Dimension         | What Is Checked                                                                  | Expected Result                                          |
| ---------------------------- | -------------------------------------------------------------------------------- | -------------------------------------------------------- |
| **Booking linkage**          | Every transaction `Booking ID` exists in `Booking_Master`                        | Invalid rows receive an error state                      |
| **Service linkage**          | Every transaction `Service Code` exists in `Service_Master`                      | Invalid rows are excluded from P&L                       |
| **FX coverage**              | Every transaction has either a valid manual FX rate or configured reference rate | Unknown currencies trigger a health warning              |
| **Financial classification** | Transaction Type maps to Revenue / COGS / Bank Fee / Deposit                     | Unknown types use the conservative COGS fallback         |
| **Numeric aggregation**      | `USD Equivalent` remains numeric under all conditions                            | `SUMIFS` does not encounter diagnostic text              |
| **Health-gated aggregation** | P&L only includes `Data Health Status = "OK"`                                    | Broken source rows cannot silently enter reported profit |
| **Container allocation**     | Booking-level shared costs use the configured allocation policy                  | No repeated full-cost assignment across containers       |
| **Zero-revenue protection**  | Profit margin calculation handles Revenue = 0                                    | No `#DIV/0!`                                             |
| **Executive monitoring**     | Global transaction errors are counted separately                                 | Dashboard signals unresolved source-data issues          |

The implementation's cross-validation matrix states that the core input fields — Booking, Container, Service, and transaction data — and the major outputs — USD Equivalent, Net Profit, Margin, and Allocated Cost — are covered by the architecture. 

### Operational SOP

The intended operating model is deliberately simple:

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

#### 1. Establish Master Data

Create the Booking in `Booking_Master`.

If containers are relevant, create their Booking-to-Container relationships in `Container_Master`.

Maintain the standardized service catalog in `Service_Master`.

Do not begin financial analysis by typing free-form service names or Booking identifiers directly into reporting sheets. The master tables are the reference layer for downstream lookup chains. 

#### 2. Record Financial Transactions

Enter revenue, vendor costs, bank fees, deposits, and other relevant financial activity into `Raw_Transactions`.

Each transaction should contain the appropriate:

* `Booking ID`
* `Container ID`, when applicable
* `Service Code`
* `Transaction Type`
* `Allocation Level`
* `Original Amount`
* `Currency`
* `Manual FX Rate`, when a transaction-specific rate is required

The transaction table is the only source that should receive the detailed financial event records. The calculation engine should not be manually edited. 

#### 3. Check Data Health

Review `Raw_Transactions[Data Health Status]`.

A clean row returns:

```text
OK
```

Exceptions identify the broken relationship, such as:

```text
ERR: INVALID_BOOKING
ERR: INVALID_SERVICE
ERR: UNKNOWN_CURRENCY
```

The implementation deliberately treats this as a separate control layer. A transaction can therefore remain numerically safe while still being visibly excluded from the financial result. 

#### 4. Review Profitability

Use:

* `Job_Costing_Engine` for Booking-level profitability.
* `Booking_PL` for one Booking's service-by-service breakdown.
* `Container_Analysis` when the Booking contains multiple containers.
* `Service_Profitability` for cross-Booking service economics.
* `Executive_Dashboard` for management-level monitoring.

This creates a drill-down path from **global profitability → Booking → service → container → transaction** rather than forcing every question to be answered from one summary dashboard.

### Data Maintenance Rules

The architecture relies on controlled master data because Excel does not provide the same hard foreign-key enforcement as a database or ERP platform.

The most important maintenance rules are:

1. **Do not overwrite formula-generated fields.**
2. **Do not manually type Service Names where a Service Code should be used.**
3. **Do not introduce a new currency without first adding its FX rate to `Setup_Params`.**
4. **Do not use `1.0` as a missing-FX fallback.**
5. **Do not write error text into numeric aggregation columns.**
6. **Do not manually edit the `Job_Costing_Engine` output.**
7. **Resolve non-OK transaction rows before treating a Booking margin as final.**
8. **Keep the shared-cost allocation policy centralized in `Setup_Params`.**

The source design specifically prohibits hard-coding global parameters into formulas and centralizes base currency, FX rates, transaction mappings, and allocation policy in `Setup_Params`. 

### Troubleshooting

#### `⚠️ X TXN DATA ERRORS` appears on the dashboard

**Likely causes:**

* Currency is missing from the FX table.
* Booking ID does not exist in `Booking_Master`.
* Service Code does not exist in `Service_Master`.
* A source transaction contains an invalid linkage.

**Resolution:**

Filter the `Data Health Status` column in `Raw_Transactions` to show values other than `OK`.

Then resolve each underlying source-data issue rather than manually changing the P&L output.

The Executive Dashboard is designed to expose this condition through a global health card:

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



#### Dynamic-array formulas return `#SPILL!`

A spill formula requires its destination area to remain clear.

Check:

* cells directly below the formula;
* cells to the right of the formula;
* manually entered values occupying the intended spill range;
* residual formatting or copied content that has created an obstruction.

Clear the obstructing cells and allow the dynamic array to expand again.

#### Profitability appears unexpectedly high

Do not immediately change the profit formula.

First check:

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

A suspiciously high margin can be caused by a cost transaction that failed the linkage chain and was therefore excluded from the calculation.

This is exactly why the workbook distinguishes **"zero as a numeric calculation state"** from **"zero as an economic conclusion."** A missing FX rate can produce `$0.00` in a numeric field while simultaneously producing `ERR: UNKNOWN_CURRENCY` in the diagnostic field. 

#### Container profitability appears overstated or understated

Check three things in order:

1. Is the correct `Booking ID` selected?
2. Is the correct `Container ID` selected?
3. Is `Setup_Params!B6` using the intended allocation policy?

Under `EQUAL_SHARE`, Booking-level shared COGS is divided by the Booking's container count before being added to direct container costs.

Under `NO_ALLOCATE`, shared Booking-level costs are not assigned to individual containers.

The distinction matters because assigning the full Booking-level shared cost to every container would duplicate the same economic cost. 

### Implementation Boundaries

This workbook is intentionally designed as a **lightweight analytical control layer**, not as a replacement for an accounting platform or transportation management system.

It is strongest when:

* operational and financial data already exists in spreadsheets or exports;
* the key unit of analysis is a Booking or shipment;
* management needs job-level profitability;
* services and charges need to be compared across Bookings;
* multiple currencies need a consistent reporting basis;
* shared costs require an explicit allocation rule;
* users need transparent formulas rather than opaque automation.

It is not designed to provide:

* multi-user transactional concurrency;
* ERP-grade permissions;
* automated bank reconciliation;
* automated carrier invoice ingestion;
* database-level foreign-key enforcement;
* enterprise workflow orchestration.

The value of the workbook is therefore not the number of sheets it contains. It is the controlled analytical chain connecting **Booking → transaction → classification → currency conversion → validated aggregation → profitability → drill-down**.

### Analytical Control Principle

The most important implementation rule can be summarized as:

```text
Never let a broken data relationship
look like a valid financial result.
```

A conventional spreadsheet might allow:

```text
Missing FX
    ↓
0 USD
    ↓
SUMIFS
    ↓
Higher Profit
```

This workbook instead makes the state explicit:

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

That distinction is the foundation of the Job Costing Engine's reliability model. The architecture explicitly recognizes that `SUMIFS` can safely aggregate numeric values while silently ignoring text, so diagnostic states must never be embedded directly inside numeric calculation columns. 

### Recommended Deployment Sequence

For a new workbook implementation, the safest sequence is:

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

The result is a repeatable workflow rather than a one-time spreadsheet calculation: establish the reference data once, append new operational transactions, resolve exceptions, and let the calculation and analysis layers update from the same underlying data model. 

<!-- Part 2 ends with the Technical Details parent <details> intentionally OPEN. Do not add </details> here. -->
### Final Architecture Notes

The workbook is deliberately built around a **single Booking-level profitability engine**, rather than a collection of disconnected calculators.

The analytical hierarchy is:

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

This structure keeps the Booking as the central economic unit while allowing users to drill downward into services and containers or upward into cross-Booking profitability. 

The implementation also explicitly recognizes three multi-hop lookup chains:

1. **Transaction → Financial Treatment → FX → USD Equivalent → Booking P&L**
2. **Booking → Container Count → Shared Cost Policy → Shared Cost Allocation → Container Analysis**
3. **Service Code → Service Master → Service Category → Service Profitability**

These chains are validated before the analytical outputs are considered production-ready. 

### What the Engine Protects Against

The central design principle is not merely automation. It is **preventing a broken input relationship from masquerading as a valid financial conclusion**.

| Failure                                              | Protective Mechanism                                                           |
| ---------------------------------------------------- | ------------------------------------------------------------------------------ |
| Unknown currency                                     | FX defaults to `0`, while a separate diagnostic state reports the missing rate |
| Unknown transaction type                             | Conservative `COGS` fallback prevents accidental profit inflation              |
| Invalid Booking                                      | `ERR: INVALID_BOOKING`                                                         |
| Invalid Service Code                                 | `ERR: INVALID_SERVICE`                                                         |
| Broken transaction chain                             | Non-`OK` rows are excluded from P&L aggregation                                |
| Shared cost duplicated across containers             | Explicit allocation policy and container-count denominator                     |
| Zero revenue                                         | Margin calculation returns `0` instead of `#DIV/0!`                            |
| Diagnostic text contaminating financial calculations | Numeric calculation columns and diagnostic columns remain separate             |

The implementation's fallback matrix specifically rejects a `1.0` FX fallback because that could turn an unrecognized foreign-currency amount into an implausibly large USD amount. 

### Recommended Use Pattern

For ongoing operations, the workbook is intended to behave as a **refreshable operating model**, not a one-time report.

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

This makes the workbook useful for recurring questions such as:

* Which Bookings are actually profitable?
* Which services produce the strongest margins?
* Which shipment costs are driving an unexpected loss?
* Which Booking-level costs need to be allocated to containers?
* Are reported margins trustworthy, or are unresolved source-data issues hiding costs?
* Where should pricing, vendor negotiation, or operational review focus next?

The source SOP follows the same sequence: establish master data, register financial transactions, inspect `Data Health Status`, and then use the profitability dashboards for review. 

### Final Design Principle

The workbook does not attempt to make financial analysis look simpler by hiding its uncertainty.

Instead, it makes the uncertainty visible.

A Booking with a calculated margin is not automatically treated as a trustworthy Booking.

A **clean Booking** means:

```text
Valid Booking
+ Valid Service
+ Valid Currency / FX
+ Valid Transaction Classification
+ Healthy Transaction Chain
= Profitability Result That Can Be Reviewed
```

That distinction is the core of the implementation.

</details>

## Other Tools in This Series

A collection of lightweight Excel decision-support tools covering operational control, profitability analysis, inventory reconciliation, workforce planning, and other business workflows.

* **[rolling-cashflow-forecast](https://github.com/HyVoid/rolling-cashflow-forecast)** — Why is cash always tight? — Forecast future liquidity under changing operational and financial assumptions.
* **[Multi-Currency-Trade-Financial-Analyzer](https://github.com/HyVoid/Multi-Currency-Trade-Financial-Analyzer)** — Which trade flows are actually profitable after currency effects? — Measure multi-currency trade profitability across cross-border operations.
* **[RTL_Loan_Sizer_Pricer](https://github.com/HyVoid/RTL_Loan_Sizer_Pricer)** — What financing structure is supportable under different borrower and collateral scenarios? — Compare lending scenarios using risk-adjusted financing assumptions.
* **[Personal-Business-Dual-Track-Budget-Manager](https://github.com/HyVoid/Personal-Business-Dual-Track-Budget-Manager)** — Where is personal and business cash being mixed or misallocated? — Separate personal and business cash flows into independently tracked budget tracks.
* **[Multifamily-Value-Add-Acquisition-Model](https://github.com/HyVoid/Multifamily-Value-Add-Acquisition-Model)** — Does this real estate acquisition still create equity value after renovation and operating assumptions? — Evaluate multifamily acquisition returns under renovation financing and operating scenarios.
* **[Real-Estate-Valuation-all-in-one-calculator](https://github.com/HyVoid/Real-Estate-Valuation-all-in-one-calculator)** — What is a property worth under different valuation methods? — Calculate property value across multiple valuation methods with tax and capital expenditure modeling.
* **[multi-entity-expense-allocation](https://github.com/HyVoid/multi-entity-expense-allocation)** — Are shared costs being allocated fairly across entities or are some assets overburdened while others escape their true cost exposure? — Distribute supplier invoices across multiple legal entities and operating assets using configurable allocation methodologies.
* **[investment-pool-nav-excel-template](https://github.com/HyVoid/investment-pool-nav-excel-template)** — What is each participant actually owed based on their entry timing and current pool value? — Track investment pool ownership continuously using NAV-based unit accounting.
* **[shopify-cash-flow-forecast-excel-toolkit](https://github.com/HyVoid/shopify-cash-flow-forecast-excel-toolkit)** — Which operational commitments will create a cash problem in the next twelve weeks? — Connect Shopify sales bank transactions advertising spend and inventory purchasing into a unified cash flow decision framework.
* **[Revenue-Management-Decision-Engine](https://github.com/HyVoid/Revenue-Management-Decision-Engine)** — Has enough meaningful information changed to justify a pricing adjustment? — Govern pricing recommendations through configurable guard rails that prevent unnecessary revenue volatility.
* **[amazon-wbr-sqp-analytics-template](https://github.com/HyVoid/amazon-wbr-sqp-analytics-template)** — What actually changed in the business this period and which products or keywords drove the change? — Transform Amazon Seller Central SQP and Sellerboard exports into standardized executive business reviews.
* **[food-manufacturing-cost-calculator-excel](https://github.com/HyVoid/food-manufacturing-cost-calculator-excel)** — Which products are actually profitable after all manufacturing costs are allocated and which are destroying margin? — Allocate factory labor utilities overhead and material costs to every SKU for true unit manufacturing economics.
* **[ecommerce-profit-calculator-excel](https://github.com/HyVoid/ecommerce-profit-calculator-excel)** — Which orders are actually profitable after all hidden costs and where is money leaking across channels and currencies? — Reconstruct true order-level profitability by reconciling sales payouts shipping costs payment fees and marketplace commissions across multi-channel e-commerce operations.
* **[construction-job-costing-excel-template](https://github.com/HyVoid/construction-job-costing-excel-template)** — Which contractor projects and customers are actually profitable after all job costs? — Connect lead estimation, job execution, and actual cost tracking into a unified contractor profitability framework.
* **[Flora-Pastry-Coffee-Shop-Operating-Control-System](https://github.com/HyVoid/Flora-Pastry-Coffee-Shop-Operating-Control-System)** — Is the shop genuinely profitable after all ingredient, labor, and operating costs? — Connect recipe costing, sales, inventory consumption, and operating expenses into a unified profitability and cash flow framework.
* **[cross-entity-logistics-console](https://github.com/HyVoid/cross-entity-logistics-console)** — Which containers and customers are actually profitable across customs and fleet operations? — Consolidate customs clearance and fleet execution data into a single container-level profitability view.
* **[family-loan-tracker-excel](https://github.com/HyVoid/family-loan-tracker-excel)** — What is actually owed today after years of transfers and accrued interest? — Reconcile long-term fund transfers and interest accruals into a current outstanding balance.
* **[3-statement-financial-model-excel](https://github.com/HyVoid/3-statement-financial-model-excel)** — What will our financial statements look like under different operating and financing scenarios? — Forecast integrated financial statements by linking revenue, expenses, working capital, and debt into a unified model.
* **[assisted-living-cost-allocation-excel](https://github.com/HyVoid/assisted-living-cost-allocation-excel)** — How much of each building's reported margin is actually generated by that building versus by the way corporate costs were allocated? — Allocate shared expenses across operating units while preserving source transactions and reconciling unit-level results to company totals.
* **[real-estate-development-feasibility-excel](https://github.com/HyVoid/real-estate-development-feasibility-excel)** — At what maximum land acquisition price does this development project become a bankable investment? — Evaluate residential development opportunities by screening land comparables, estimating construction costs, and calculating profit-on-cost thresholds.
* **[uk-residential-costing-software](https://github.com/HyVoid/uk-residential-costing-software)** — Will this construction project actually hit its target margin once site logistics, waste factors, and trade rate changes are accounted for? — Separate physical building quantities from volatile trade rates and site preliminaries to cascade commercial markups across tender budgets and client quotations.
* **[b2b-pricing-landed-cost-excel-model](https://github.com/HyVoid/b2b-pricing-landed-cost-excel-model)** — Which wholesale prices and customer segments actually preserve margin once freight, duties, payment terms, and FX are fully costed? — Calculate SKU-level landed costs, cost-to-serve, and pocket-margin price waterfalls to simulate cross-border wholesale profitability.
* **[custom-label-pricing-calculator](https://github.com/HyVoid/custom-label-pricing-calculator)** — What should we charge for a custom label order at each quantity tier to stay competitive and profitable? — Calculate parametric label pricing by modeling substrate square-inch costs, setup charges, and quantity-tier discount breaks against competitor benchmarks.
* **[landed-cost-calculator-retail-deduction-excel](https://github.com/HyVoid/landed-cost-calculator-retail-deduction-excel)** — What is the true net margin on each SKU once actual landed costs and retailer deductions are accounted for? — Reconcile PO-estimated landed costs against actual freight, duty, FX, and drayage charges while auditing retailer remittance deductions to expose SKU-level net margin.
* **[saas-finance-mrr-burn-rate-excel](https://github.com/HyVoid/saas-finance-mrr-burn-rate-excel)** — Can we make this week's hiring, marketing, and vendor commitments without running out of cash before the books close? — Track weekly cash burn pacing, MRR/ARR waterfall variance, retention metrics, and payroll costs against budget before month-end close.
* **[startup-financial-model-excel](https://github.com/HyVoid/startup-financial-model-excel)** — Will our growth plan run out of cash before cohort monetization catches up? — Forecast multi-year cash burn, runway, unit economics, and fundraising dilution from bottom-up cohort and headcount assumptions.
* **[shopify-pnl-reconciliation-excel](https://github.com/HyVoid/shopify-pnl-reconciliation-excel)** — Is each Shopify storefront actually profitable once true SKU-level costs, shared expenses, and bank charges are reconciled? — Reconcile multi-store Shopify sales exports against SKU-level COGS, allocated multi-currency OPEX, and bank statements to produce a trustworthy monthly P&L.
* **[travel-agency-accounting-excel-template](https://github.com/HyVoid/travel-agency-accounting-excel-template)** — Which tour bookings actually generate net profit and positive cash flow after supplier costs and B2B collections? — Calculate per-booking tour job costs, gross margins, AR/AP settlement, and operating cash flow across multiple currencies.

## License

This project is released under the **Apache License 2.0**.

You may use, modify, distribute, and build upon the work in accordance with the terms of the Apache License 2.0.
