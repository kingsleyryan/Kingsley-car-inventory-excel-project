# Car Inventory Data Analysis Project

A comprehensive data cleaning, inventory tracking, and formula implementation project completed using Microsoft Excel. This project showcases how to systematically parse abbreviations into meaningful business data.

---

## 🛠️ Step-by-Step Project Workflow

### 1. Manufacturer Code Resolution
The project began by resolving the two-letter manufacturer codes into their full, recognizable brand names to build a clear reference key:
* **FD** $\rightarrow$ Ford
* **HO** $\rightarrow$ Honda
* **HY** $\rightarrow$ Hyundai
* **CR** $\rightarrow$ Chrysler
* **TY** $\rightarrow$ Toyota
* **GM** $\rightarrow$ General Motors

### 2. Model Name Extraction
Using the core vehicle abbreviations from the raw inventory dataset, the 3-letter codes were systematically decoded into full vehicle model names:
* **MTG** $\rightarrow$ Mustang (Ford)
* **CAR** $\rightarrow$ Caravan (Chrysler)
* **CMR** $\rightarrow$ Camaro (Chevrolet / GM)
* **SLV** $\rightarrow$ Silverado (Chevrolet / GM)
* **PTC** $\rightarrow$ PT Cruiser (Chrysler)

**Formula Used:**
Implemented a precise `VLOOKUP` range lookup to automate this mapping across the entire dataset:
`=VLOOKUP(D2, D$55:E$66, 2, FALSE)`
*(Using `FALSE`/`0` ensures strict exact matches for data integrity).*

### 3. Chronological Calculations (Vehicle Age Logic)
Faced with two-digit shorthand manufacturing years (e.g., `06` for 2006, `98` for 1998), an advanced conditional logic formula was built to handle the threshold relative to the current tracking baseline (**2026**):
`=IF(26-F2<0, 100-F2+26, 26-F2)`

### 4. Metrics & Operational Tracking
* **Miles Driven Per Year:** Calculated usage intensity by mapping odometer readings against asset age (`=H2/G2`).
* **Warranty Coverage Audit:** Used conditional logic to automatically cross-examine current mileage against coverage limits:
  `=IF(H2<=L2, "Yes", "Not Covered")`

### 5. Custom Composite Identity Mapping
Engineered a brand-new unique identifier by extracting parts of the original Car ID, isolating the color attribute, converting it to uppercase, and placing it directly into the center text string:
`=CONCATENATE(LEFT(A2,7), UPPER(LEFT(J2,3)), RIGHT(A2,3))`
* *Transforms:* `FD06MTG001` + `Black` $\rightarrow$ **`FD06MTGBLA001`**
*
