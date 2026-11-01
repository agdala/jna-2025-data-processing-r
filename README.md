# JNA 2025 Data Processing in R

## Overview

This repository contains an R script developed to clean, recode, and label data from the **Joint Needs Assessment (JNA) 2025**, supporting humanitarian analysis and evidence-based decision making.

The script documents the full data preparation workflow applied to the JNA 2025 dataset in Brazil, ensuring transparency, reproducibility, and methodological consistency across analyses and visual products.

While the main analytical outputs for JNA 2025 were produced using Power BI dashboards, this repository preserves the underlying data transformation logic implemented in R.

---

## My role

I developed this script as the person responsible for preparing the JNA 2025 data in Brazil, within the Information Management working group of R4V Brazil. Its outputs fed the Power BI dashboards used for the analysis.

## Context

The Joint Needs Assessment is a coordinated, inter-agency household survey designed to identify priority needs, vulnerabilities, and access to services among affected populations.

In humanitarian settings, clear documentation of data processing steps is critical to:

- Ensure analytical transparency  
- Enable peer review and validation  
- Facilitate reuse of methodologies across assessments  
- Support institutional memory and continuity  

This repository focuses exclusively on **data processing and preparation**, not on data dissemination.

---

## Scope of the script

The `JNA_2025.R` script performs the following tasks:

- Renames raw survey variables using meaningful, standardized names  
- Imports reference tables and codebooks  
- Recodes numeric survey responses into labeled categorical variables  
- Applies harmonized labels across individual and household-level variables  
- Structures household composition data for up to 12 household members  
- Prepares the dataset for statistical analysis and visualization  

No analytical indicators or dashboards are generated directly by this script.

---

## Script structure

The script is organized into logical sections:

1. Package loading  
2. Dataset preparation  
3. Systematic variable renaming based on the original questionnaire  
4. Import of reference tables and codebooks  
5. Recoding and labeling of categorical variables  
6. Standardization of demographic and household-level variables  

This modular structure allows easy adaptation for future assessments or country contexts.

Data folders are intentionally excluded from version control.


## Requirements

### Software

- R version 4.0 or higher  

### R packages

The script relies on the following packages:

- `tidyverse`  
- `readxl`  
- `writexl`  
- `utf8`  
- `lookup`  

All packages are available from CRAN.

---

## How to use

1. Clone this repository  
2. Store raw survey data and reference tables locally, following the suggested folder structure  
3. Update file paths in the script to match your local environment  
4. Run `scripts/JNA_2025.R`  

The script assumes that reference tables are structured consistently with the original JNA 2025 codebook.

---

## Data protection and ethical considerations

This repository **does not contain any personal or sensitive data**.

Users must ensure that:

- Personally identifiable information is not committed to version control  
- Data sharing complies with organizational data protection policies  
- Outputs derived from this script respect ethical and confidentiality standards  

The repository is intended for methodological documentation and code reuse only.

---

## Reusability

The script can be adapted for:

- Future JNA rounds  
- Similar household surveys  
- Other humanitarian assessments requiring systematic recoding and labeling  

Users are encouraged to modify variable names, reference tables, and folder structures as needed.

---

## Limitations

- File paths are not fully parameterized and may require manual adjustment  
- The script assumes a fixed maximum household size of 12 members  
- Analytical indicators are not calculated within this repository  

These design choices reflect the scope of the original implementation.

---

## License

This project is licensed under the MIT License.

You are free to use, modify, and distribute the code, provided that proper attribution is given.

---

## Acknowledgements

This script supports the Joint Needs Assessment 2025 data processing workflow and reflects collaboration across humanitarian information management and analysis teams.


