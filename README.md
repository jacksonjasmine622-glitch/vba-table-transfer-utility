**Project: Excel VBA – Form to Master Table Loader**

**What it does:** Reads shared fields once, scans multiple line-item sections, appends rows into a master ListObject table, writes totals only on the last row.

**Skills demonstrated:** ListObject automation, merged-cell safe reading, section-driven mapping, performance toggles, error handling.

**How to use:**
Open MasterTracker.xlsm

Open the workbook containing the InputForm sheet

Update constants at top of module if needed

Run TransferFormDataToMasterTable

**Expected table headers:**
Header-Level Fields

QuoteDate

SalesRep

ReferenceNumber

ContactName

CustomerName

MaterialType

ShippingCost

AdditionalCharge

QualitySpec1

QualitySpec2

Line Item Fields

Quantity

Shape

Surface

ItemType

Thickness

OuterDiameter

Width

Length

Finish

EdgeType

Direction

Protection

UnitPrice

Totals

TotalAmount
