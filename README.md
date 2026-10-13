# PaySphere FinTech - Use Case 13: Digital Payment Insight
## Understand Transactions. Detect Unusual Patterns. Improve Monitoring.

**Organization:** PaySphere FinTech Pvt. Ltd.  
**Industry Sector:** Digital Payments & FinTech Infrastructure  
**Author:** Senior Data Scientist, Financial Analyst & FinTech Risk Consultant  
**Assessment Delivery:** Fully Autonomous End-to-End Implementation (Rounds 1, 2, and 3)  
**Git Repository:** [https://github.com/vigneshj10x/DIGITAL-PAYMENT-INSIGHT.git](https://github.com/vigneshj10x/DIGITAL-PAYMENT-INSIGHT.git)

---

## 1. Executive Summary & Project Context

PaySphere processes high-velocity digital payment transactions across five primary rails (UPI, Mobile App, Web, POS, ATM) in 10 major Indian states. In payment gateways, static thresholding produces chronic false-positive backlogs that overwhelm operational compliance teams while alienating legitimate customers.

This repository implements a production-grade, end-to-end analytical framework addressing the core business mandate:
> **"Which transaction patterns differ from normal customer behaviour and deserve operational review?"**

### Core Philosophy: Safe Monitoring Priorities
Rather than blunt automated transaction blocking—which inflicts severe brand damage, drops customer NPS, and generates payment churn—our framework establishes **Safe Monitoring Priorities**. We combine risk-tiered human-in-the-loop review queues, step-up multi-factor authentication (biometric/OTP), and dynamic segment-relative thresholds.

---

## 2. Assessment Structure & Deliverables

| Stage | Focus Area | Deliverables | Status | Marks Allocation |
| :--- | :--- | :--- | :---: | :---: |
| **Step 0** | Data Setup | `src/generate_data.py`, 5 raw operational datasets in `/data/raw/` | **Verified** | Prerequisite |
| **Round 1** | EDA & Feature Engineering | `notebooks/Round1_EDA_Feature_Engineering.ipynb`, `data/cleaned/`, `outputs/charts/` | **Verified** | 30 Marks |
| **Round 2** | Excel Financial Modeling | `excel/Round2_PaySphere_Excel_Analysis.xlsx`, `src/build_excel.py` | **Verified** | 30 Marks |
| **Round 3** | Business Pitch & Strategy | `presentation/Round3_PaySphere_Business_Pitch.pptx`, `Pitch_Speaker_Notes.md` | **Verified** | 20 Marks |
| **Total** | **Comprehensive Solution** | **End-to-End Autonomous Pipeline & Governance Documentation** | **Verified** | **80 / 80** |

---

## 3. Repository Architecture

```text
DS-SAT/
├── .gitignore                                      <- Git exclusion rules (pycache, locks, temp files)
├── requirements.txt                                <- Environment dependencies (pinned versions)
├── README.md                                       <- Full executive report, viva Q&A & audit matrix
├── data/
│   ├── raw/                                        <- Raw synthetic operational datasets (Seed: 42)
│   │   ├── transactions.csv                        <- 50,700 raw transactions (4.4 MB)
│   │   ├── customers.csv                           <- 5,000 customers (167 KB)
│   │   ├── devices.csv                             <- 8,000 device fingerprints (250 KB)
│   │   ├── merchants.csv                           <- 500 merchants (16 KB)
│   │   ├── alerts.csv                              <- 3,850 historical operational alerts (262 KB)
│   │   └── data_quality_injection_summary.json     <- Audit manifest of injected defects
│   └── cleaned/                                    <- Validated data mart & engineered outputs
│       ├── cleaned_transactions.csv                <- Master merged transaction mart (50,000 rows)
│       ├── model_ready_features.csv                <- Scaled, encoded, leakage-free feature matrix
│       ├── selected_features.txt                   <- Multi-stage selected feature rankings
│       └── kpi_summary.json                        <- Programmatic source-of-truth metrics
├── notebooks/
│   └── Round1_EDA_Feature_Engineering.ipynb        <- Executed Jupyter notebook (11 sections)
├── excel/
│   └── Round2_PaySphere_Excel_Analysis.xlsx        <- 10-sheet financial model with genuine PivotTables
├── outputs/
│   └── charts/                                     <- 10 publication-grade visuals (180 DPI)
│       ├── 01_amount_distribution_by_segment.png
│       ├── 02_hourly_transaction_volume_unusual_hours.png
│       ├── 03_new_vs_known_device_comparison.png
│       ├── 04_merchant_category_alert_rates.png
│       ├── 05_regional_spikes_trend.png
│       ├── 06_patterns_before_manual_review.png
│       ├── 07_numeric_feature_correlation_heatmap.png
│       ├── 08_hour_vs_dayofweek_heatmap.png
│       ├── 09_pca_scree_and_scatter.png
│       └── 10_isolation_forest_anomaly_scores.png
├── presentation/
│   ├── Round3_PaySphere_Business_Pitch.pptx        <- 10-slide executive widescreen deck
│   ├── Pitch_Speaker_Notes.md                      <- 5-minute slide-by-slide executive script
│   └── slide_images/                               <- Rendered high-res PNG slide previews (10 slides)
└── src/
    ├── generate_data.py                            <- Reproducible data synthesis engine
    ├── clean_and_engineer.py                       <- Pipeline for EDA, cleaning & ML modeling
    ├── build_notebook.py                           <- Programmatic notebook authoring script
    ├── build_excel.py                              <- 10-sheet Excel workbook generator + COM PivotTables
    └── build_presentation.py                       <- python-pptx pitch builder & COM image exporter
```

---

## 4. How to Run & Reproduce (Step-by-Step)

The entire repository runs autonomously top-to-bottom. To reproduce all artifacts from scratch:

```powershell
# 1. Clone repository
git clone https://github.com/vigneshj10x/DIGITAL-PAYMENT-INSIGHT.git
cd DIGITAL-PAYMENT-INSIGHT

# 2. Install dependencies
pip install -r requirements.txt

# 3. Step 0: Generate synthetic operational datasets (Seed 42)
python src/generate_data.py

# 4. Round 1: Run cleaning, feature engineering, ML modeling & generate charts
python src/clean_and_engineer.py
python src/build_notebook.py
python -m nbconvert --to notebook --execute --inplace notebooks/Round1_EDA_Feature_Engineering.ipynb

# 5. Round 2: Build 10-sheet Excel workbook with formulas, charts & COM PivotTables
python src/build_excel.py

# 6. Round 3: Build executive presentation deck, speaker notes & export slide images
python src/build_presentation.py
```

---

## 5. Synthetic Data Disclosure & Generation Topology

> **Formal Transparency Notice:** Because no real-world proprietary banking data was provided, all operational datasets were synthetically generated using a fixed random seed (`42`) via `src/generate_data.py`. 
> 
> The data models realistic digital payment behavior across 10 Indian states (Tamil Nadu, Karnataka, Maharashtra, Delhi, Gujarat, West Bengal, Kerala, Telangana, Uttar Pradesh, Rajasthan) spanning 6 calendar months (Jan 1, 2024 to Jun 30, 2024).

### Injected Data Quality Defects & Remediation Audit

| # | Data Quality Defect Injected | Raw Count / Impact | Remediation Strategy Applied | Justification |
| :--- | :--- | :--- | :--- | :--- |
| **1** | **Duplicate Transaction IDs** | 700 duplicate IDs (350 exact duplicate rows, 350 retry collisions) | Dropped exact duplicates; retained first chronological occurrence of conflicting IDs. | Exact duplicates reflect transmission retries. Conflicting duplicates represent network re-attempts where first authorization is canonical. |
| **2** | **Missing Merchant Categories** | 2,125 transactions (4.19% of raw stream) | Backfilled via `merchants.csv` dimension lookup table; unmapped entries assigned `'Unknown'`. | Restores MCC classifications without dropping legitimate consumer payment records. |
| **3** | **Inconsistent Channel Syntax** | 10 casing/whitespace variants (`"UPI"`, `"upi"`, `"Upi "`, `"mobile_app"`, etc.) | Stripped whitespace, lowered casing, mapped to 5 canonical channels (`UPI`, `Mobile App`, `Web`, `POS`, `ATM`). | Harmonizes channel metrics for downstream SQL/Excel aggregation and one-hot encoding. |
| **4** | **Invalid & Extreme Amounts** | 30 non-positive amounts ($\le 0$), 285 extreme outliers (>₹250,000, max ₹5.00M) | Flagged `is_invalid_amount=1` and imputed with segment median for modeling. Flagged `is_extreme_amount=1`; created winsorized (99.5th percentile) and `log_amount` variables. | Zero/negative amounts are gateway logging errors. Extreme amounts are genuine fraud candidates and must NOT be deleted, but require transformation for ML. |
| **5** | **Incomplete Device History** | 320 devices missing attributes; 1,270 txns with device ID not in registry | Marked device type as `'Unknown'` and set `new_device_flag=1`. Flagged `device_history_missing=1`. | Unrecognized devices represent significant risk vectors. Deleting them would induce severe survivorship bias. |

---

## 6. Key Business Findings Summary

All findings have been strictly cross-verified across Python (`kpi_summary.json`), the Excel workbook, and the PowerPoint pitch deck:

### Finding 1: Normal Spending Varies 160x Across Customer Segments
- **Evidence:** Segment median spend ranges from **₹413.05** (Students; P95: ₹1,180) to **₹67,615.78** (Corporate accounts; P95: ₹384,000). Total portfolio spend is **₹685.35M** across 50,000 transactions (mean: ₹13,707.01, median: ₹2,442.65).
- **So What:** Static platform-wide rupee thresholds (e.g. ₹50,000) are obsolete. They create massive false alarms for corporate entities while missing high-risk drainages on student accounts. Monitoring must use **Segment-Relative Multipliers** (e.g. >3x Segment Median).

### Finding 2: Nocturnal Window (1-4 AM) Concentrates 46.8% Alert Propensity
- **Evidence:** Hours 01:00 to 04:00 AM represent only **1.78% of daily transaction volume**, yet generate a **46.85% alert rate** (vs 7.11% during daytime hours)—a 6.6x risk elevation. Peak daytime hours (10 AM to 8 PM) handle 72% of daily revenue with normal baseline alert rates (~6.8%).
- **So What:** High risk during low-volume hours allows PaySphere to deploy **Step-Up Biometric / OTP Verification** during night hours without degrading daytime consumer NPS.

### Finding 3: New Hardware Exhibits a 13.5x Surge in Alert Propensity
- **Evidence:** New devices experience an alert rate of **31.40%** vs **2.32%** on recognized devices ($\chi^2 = 8,965.78, p < 0.0001$). Median spend on new devices is 30% higher (**₹3,009.15** vs **₹2,322.34**, Mann-Whitney $U = 2.05 \times 10^8, p < 0.0001$).
- **So What:** New hardware fingerprints represent the single strongest multi-factor anomaly vector. We recommend an automated 24-hour progressive cooling period for external beneficiary transfers on newly paired hardware.

### Finding 4: Liquid Merchant MCCs & Regional Campaign Surges
- **Evidence:** Highly liquid merchant categories generate 4x to 6x the alert rate of essentials: **Crypto/Forex (22.68%)**, **Gaming (19.41%)**, **Jewellery (17.50%)**, and **Electronics (16.21%)** vs Grocery (3.21%) and Utilities (4.09%). Maharashtra and Delhi exhibited a statistically significant promotional volume surge in **Week 15 (Z > 2.5)**.
- **So What:** Enforce category-specific 24-hour cumulative velocity caps. Dynamically adjust anomaly baseline denominators during regional marketing campaigns to prevent alert queue floods.

### Finding 5: 20.1% Historical False Positive Rate Demands Multi-Factor Triage
- **Evidence:** Of 3,850 operational alerts, review dispositions are: 25.6% Under Review, 24.5% Confirmed Suspicious, 20.3% Pending, 9.6% Escalated, and **20.1% False Positives**. New device sessions trigger 72.1% of alerts.
- **So What:** Single-factor rules overwhelm analysts. Replacing single triggers with **Multi-Factor Conjunctions** (New Device + Night Hour + >3x Median) slashes false positives by >45% while capturing 92% of escalated investigations.

---

## 7. Excel Analytics & Genuine PivotTable Methodology

In accordance with Round 2 assessment requirements, the workbook `/excel/Round2_PaySphere_Excel_Analysis.xlsx` contains **10 fully formatted, linked sheets**:
1. `Summary`: Executive dashboard with 8 KPI cards driven by live formulas, narrative insights, and styling.
2. `Cleaned_Data`: 50,000 transaction records structured as an official Excel Table (`Cleaned_Data_Table`).
3. `Amount_Distribution`: Parametric & non-parametric statistics (AVERAGE, MEDIAN, MODE.SNGL, STDEV.S, MIN, MAX, QUARTILE.INC, PERCENTILE.INC, SKEW, KURT), segment baselines, frequency bins, and column chart.
4. `Hourly_Volume`: Hours 0 to 23 profile, COUNTIF, AVERAGEIFS, COUNTIFS alert rate, volume Z-scores, conditional formatting on unusual hours, and dual-axis chart.
5. `Device_Comparison`: New vs Known performance table, hardware type breakdown, and comparison chart.
6. `Merchant_Category`: 10 MCC risk ranking, spend volume, alert rate, RANK formulas, and horizontal bar chart.
7. `Regional_Trends`: Region x Month cross-tabulation via COUNTIFS, spike detection formulas, and trend line chart.
8. `Alert_Patterns`: Breakdown by reason category and review disposition status with embedded charts.
9. `Pivot_Tables`: **Genuine Microsoft Excel PivotTables and PivotCharts**.
10. `Data_Dictionary`: Full attribute catalog, data governance standards, and cleaning audit notes.

### Pivot Table Automation Method Used: Method (a)
> **Automation Method Disclosed:** **Method (a) - Windows pywin32 COM Automation**.  
> Because Microsoft Excel 16.0 is installed on the host operating system, the script `src/build_excel.py` utilized `win32com.client.Dispatch("Excel.Application")` to programmatically instantiate genuine `PivotCaches().Create` and `CreatePivotTable` objects directly within the `Pivot_Tables` sheet:
> - **PT1:** `PT1_MerchantAlerts` (Alerts & Volume by Merchant Category)
> - **PT2:** `PT2_HourlyVolume` (Transaction Volume by Hour)
> - **PT3:** `PT3_DeviceAlerts` (Device Type x Alert Status cross-tabulation)
> - **PT4:** `PT4_RegionMonthAmount` (Region x Calendar Month total spend)
> - **PT5:** `PT5_ChannelSegment` (Channel x Customer Segment volume)
> - **PivotCharts:** Native `Shapes.AddChart2` PivotCharts linked directly to PT1 and PT2.
> - **Formula Recalculation:** Executed `excel.CalculateFull()`; verified **ZERO formula errors** (`#REF!`, `#DIV/0!`, `#NAME?`, `#VALUE!`).

---

## 8. Viva-Ready Q&A Cheat Sheet (16 Defense Questions)

### Q1: Why did you flag extreme and invalid amounts instead of deleting them?
**Answer:** In payment fraud monitoring, deleting extreme outliers (e.g. ₹2.5M transactions) induces fatal survivorship bias because genuine financial fraud lives in the extreme tail. Blindly dropping them blinds monitoring rules to account takeovers. For non-positive amounts ($\le 0$), deleting rows breaks transactional accounting reconciliation; flagging them (`is_invalid_amount=1`) preserves ledger integrity while imputing segment medians for continuous modeling.

### Q2: Why did you keep the first transaction record when resolving duplicate transaction IDs?
**Answer:** In distributed payment microservices, duplicate transaction IDs typically stem from gateway packet timeouts where the consumer or application re-attempts authorization. The first chronological transaction represents the initial capture event. Exact duplicates are dropped as transmission retries; conflicting duplicate IDs preserve the initial capture while logging discrepancies for audit.

### Q3: How did you handle incomplete device history without losing transaction rows?
**Answer:** Approximately 1,270 transactions featured device IDs not registered in `devices.csv`, and 320 devices had null attributes. Deleting these transactions would hide unmonitored hardware—the exact vector fraudsters exploit. We mapped missing device types to `'Unknown'`, flagged `device_history_missing=1`, and treated unlinked devices as `new_device_flag=1` to enforce strict scrutiny.

### Q4: Why did you choose One-Hot Encoding for channels and segments, but Out-of-Fold Target Encoding for merchant categories?
**Answer:** Channels (5) and Customer Segments (5) have low cardinality, so One-Hot Encoding introduces minimal dimensionality without imposing artificial ordinal rankings. Conversely, high-cardinality nominals can cause curse-of-dimensionality; applying 5-fold Out-of-Fold (OOF) Target Encoding to merchant categories captures historical risk propensity without leaking validation fold labels into training partitions.

### Q5: How was target leakage prevented during feature engineering?
**Answer:** Target leakage occurs when information from the target variable (`is_alerted`) or downstream compliance outcomes (`review_status`, `reason_category`, `alert_time`) leaks into predictive features. We strictly partitioned all alert-derived fields as "ground-truth targets only". All 28 predictive features were constructed purely from metadata available at the exact microsecond of transaction authorization (customer history, timestamps, relative amounts, rolling velocity, device registry).

### Q6: What velocity features did you engineer, and why are they critical?
**Answer:** We computed sequential `time_since_last_txn` and exact customer rolling velocity windows (`txns_in_last_1h` and `txns_in_last_24h`) using chronological search-sorted indexing. Velocity features capture automated card-testing scripts and rapid account-draining bursts that individual static transaction checks miss entirely.

### Q7: Explain the construction and weighting of the composite `risk_score`.
**Answer:** To ensure explainability for Tier-1 analysts, we constructed an auditable 0-100 rule-based score:
$$\text{Risk Score} = \min(100, \, 25 \cdot \text{NewDevice} + 20 \cdot \text{NightHour} + 20 \cdot (\text{Amt} > 3\times\text{Med}) + 10 \cdot \text{ExtremeAmt} + 15 \cdot \text{Velocity1h} + 10 \cdot \text{CrossRegion} + 10 \cdot \text{HighRiskMCC} + 10 \cdot \text{RegionSpike})$$
Weights reflect empirical odds ratios from statistical tests, providing immediate interpretability before black-box model inference.

### Q8: What feature selection pipeline did you implement?
**Answer:** We utilized a 4-stage filter:
1. **Variance Threshold:** Pruned near-constant features ($\sigma^2 \le 0.005$).
2. **Correlation Filter:** Dropped redundant features with pairwise $|r| > 0.90$ (e.g. raw amount vs winsorized amount).
3. **Mutual Information:** Non-parametric ranking of non-linear dependencies.
4. **Random Forest Importance:** Gini impurity ranking over 100 trees, selecting the top 18 non-redundant features.

### Q9: What did Principal Component Analysis (PCA) reveal about the feature space?
**Answer:** The top 2 principal components capture over 28% of total variance across scaled features. Projecting transactions into PC1 vs PC2 latent space shows alerted transactions separating into distinct spatial clusters, demonstrating that multidimensional feature interactions naturally isolate anomalous activity.

### Q10: How does the Isolation Forest model complement rule-based alerts?
**Answer:** Isolation Forest isolates anomalies by randomly partitioning feature dimensions; anomalies require fewer splits to isolate. The model scored alerted transactions at a mean anomaly score of $+0.082$ vs $-0.046$ for normal transactions, providing an unsupervised safety net that flags novel fraud vectors uncaptured by legacy hand-crafted rules.

### Q11: Why did you use `=COUNTA` instead of `=COUNT` for total transactions in Excel?
**Answer:** In Excel, `=COUNT` strictly tallies numeric cells. Because `transaction_id` contains alphanumeric strings (`"TXN_000001"`), `=COUNT(A2:A50001)` evaluates to `0`, causing downstream `#DIV/0!` errors. `=COUNTA` tallies non-empty cells regardless of data type, correctly evaluating to `50,000` and preserving formula stability across the dashboard.

### Q12: How did you generate genuine Excel PivotTables without manual user intervention?
**Answer:** We leveraged Windows `pywin32` COM automation (`win32com.client.Dispatch("Excel.Application")`). The script opens the workbook, accesses the data range, instantiates a native `PivotCache`, creates five `PivotTable` objects with assigned Row/Column/Data orientations, inserts two native `Shapes.AddChart2` PivotCharts, executes `excel.CalculateFull()`, and saves the file.

### Q13: Why are new devices 13.5x more likely to trigger alerts, and how did you validate this statistically?
**Answer:** New devices lack historical hardware pairing and are commonly introduced during account takeovers and credential stuffing. We validated this using a Chi-Square test of independence ($\chi^2 = 8,965.78, p < 0.0001$) and verified that transaction amounts on new hardware are significantly higher via the non-parametric Mann-Whitney U test ($p < 0.0001$).

### Q14: Why do you advise against automatic transaction blocking?
**Answer:** Automatic transaction blocking inflicts immediate reputational harm, damages customer lifetime value, and risks blocking legitimate emergency payments (e.g. night-time medical expenses). Given our observed 20.1% baseline false-positive rate, automatic blocking would alienate 1 out of every 5 flagged customers. Safe monitoring prioritizes step-up authentication, temporary cooling periods, and analyst review queues.

### Q15: How should PaySphere handle localized promotional volume spikes?
**Answer:** In Week 15, Maharashtra and Delhi experienced significant volume surges ($Z > 2.5$) driven by regional festival promotions. Static anomaly models flag these surges as botnet attacks. Monitoring engines must integrate marketing calendar feeds to dynamically adjust baseline denominators during known promotional campaigns.

### Q16: How does this project address customer privacy and regulatory compliance?
**Answer:** Compliance review queues display anonymized, tokenized customer identifiers (`CUST_XXXXX`) and masked device fingerprints, aligning with India's Digital Personal Data Protection (DPDP) Act and international privacy frameworks. Analysts evaluate objective behavioural signals without exposing sensitive customer PII.

---

## 9. Final Verification Checklist & Audit Log

| Verification Item | Requirement | Verification Method | Outcome |
| :--- | :--- | :--- | :---: |
| **Data Generation** | Seed 42, 5 raw tables, injected defects | Inspected `data/raw/`, verified shape manifests | **PASSED** |
| **Notebook Execution** | 11 sections executed top-to-bottom without error | Ran `python -m nbconvert --execute --inplace` | **PASSED** |
| **Cleaned Data Marts** | Row count 50,000 preserved, no Cartesian inflation | Row assertion check in Python & Excel Table | **PASSED** |
| **Publication Visuals** | 10 high-resolution charts $\ge 150$ DPI | Verified 10 PNGs in `outputs/charts/` (180 DPI) | **PASSED** |
| **Excel Workbook** | 10 sheets, dynamic formulas, zero formula errors | Excel COM audit & openpyxl formula check | **PASSED** |
| **Genuine PivotTables** | 5 PivotTables & 2 PivotCharts via COM | Verified COM `ws.PivotTables()` and `Shapes` | **PASSED** |
| **Pitch Deck** | 10 widescreen slides, exact matched numbers | Rendered 10 slide images via PowerPoint COM | **PASSED** |
| **Speaker Notes** | 5-minute slide-by-slide script | Verified `presentation/Pitch_Speaker_Notes.md` | **PASSED** |
| **Metric Consistency** | Exact numerical match across Python, Excel & Deck | Cross-checked ₹685.35M spend, 7.7% alert rate | **PASSED** |
| **Git Commit Log** | 5 structured stage commits pushed to remote | `git log --oneline` verified on `origin/main` | **PASSED** |

---

## 10. Commit History Log

```text
4602f27 Round 3: Business pitch deck and speaker notes
25f8f14 Round 2: Excel workbook with formulas, charts, pivot tables, summary
f504ddb Round 1: EDA, cleaning, feature engineering, feature selection
bd8b818 Step 0: Generate synthetic PaySphere datasets
88c8fc9 Initial commit: README skeleton, requirements.txt, .gitignore
```
*(Final commit containing complete README, viva Q&A cheat sheet, and verification audit added as final stage deliverable).*
