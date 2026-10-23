# Healthcare Data Cleaning & Understanding

**Task ID:** task1  
**Tool:** Microsoft Excel workbook, prepared with automated data checks and workbook generation  
**Dataset:** Synthetic Medical Dataset v2

## Objective

Explore patient information, identify column types, check missing and duplicate records, and calculate patient counts, diagnosis frequencies, and average age.

## Dataset and source

- [Original dataset and dataset card](https://huggingface.co/datasets/Sumitkumar098/Synthetic_Medical_Dataset_v2)
- [Download source CSV](https://huggingface.co/datasets/Sumitkumar098/Synthetic_Medical_Dataset_v2/resolve/main/Synthetic_Medical_Dataset_v2.csv)
- Author: Sumitkumar098 on Hugging Face.
- Retrieved: 23 October 2026.
- Source: 15,000 synthetic records and 36 columns. The dataset page identifies the license as MIT; consult the original source for its terms.
- Source CSV SHA-256: `47176c567e40e8ea259d17018931e23ee3404e5bd82bec21ff0b3d3b6cd9b91a`.

These are fictional practice records. Medication entries are source data, not clinical recommendations.

## Project file

[Download the Excel workbook](Healthcare_Data_Cleaning.xlsx)

| Worksheet | Contents |
|---|---|
| Summary | Formula-based metrics, diagnosis counts, and a linked chart |
| Clean_Data | Six selected fields for all 15,000 patients, in a filterable Excel Table named Patients |
| Raw_Data | All original records and all 36 source fields, with source values retained |
| Cleaning_Log | Cleaning decisions, missing-value formulas, selected-column dictionary, and source details |

The Summary appears first. Download the workbook and open it in Excel to inspect the formulas and chart.

## Column types

| Column | Type | Interpretation |
|---|---|---|
| Patient_ID | Identifier, stored as text | Unique synthetic patient key |
| Age | Numeric whole number | Age in years |
| Gender | Categorical text | Source gender label |
| Predicted_Disease | Categorical text | Synthetic diagnosis label, including Healthy |
| Medicine_1 | Categorical text | First medication; None is an explicit source category |
| Dosage_1 | Text | Dose and units or instructions, retained as supplied |

## Cleaning and quality checks

1. Compared all 36 original fields to check for exact duplicate records before selecting columns. **No exact duplicates were found**, so no records were removed.
2. Confirmed that all 15,000 patient IDs are populated and unique. Each row represents one patient in this file.
3. Selected the six assignment fields and retained every patient. Other source fields remain available in Raw_Data.
4. Checked surrounding whitespace and missing tokens in selected fields. **No whitespace changes were required.**
5. Stored age as a number and patient ID as an identifier. All ages are whole numbers, with an observed range of 18–90 years. No age corrections or imputation were applied.
6. Reviewed gender and diagnosis labels for case and spacing variants. No standardisation changes were required.
7. Found **1,084 blank Dosage_1 values**, all in Healthy records where Medicine_1 is None. Retained these blanks rather than inventing a dose or replacing them with zero. There were no missing values in the other five selected fields.

The dataset required limited cleaning. Zero duplicate removals and zero category corrections are reported honestly rather than introducing artificial errors.

## Results

| Metric | Result |
|---|---:|
| Source records | 15,000 |
| Cleaned records | 15,000 |
| Unique patients | 15,000 |
| Exact duplicate records | 0 |
| Average age | 44.8 years |
| Numeric ages included in average | 15,000 |
| Minimum age | 18 years |
| Maximum age | 90 years |
| Most common disease label | Anxiety Disorder |
| Patients with Anxiety Disorder | 1,222 |
| Blank first dosages | 1,084 |

Anxiety Disorder is the most frequent disease label, followed jointly by Diabetes Type 2 and Urinary Tract Infection, with 1,195 patients each. Healthy accounts for 1,084 records and is included in the diagnosis summary for completeness; it is not described as a disease.

## How to reproduce and review in Excel

1. Download the source CSV and open it in Excel, using comma-delimited import if needed. Keep an unchanged copy as Raw_Data.
2. Copy the source sheet. Use **Data → Remove Duplicates**, selecting all 36 columns and confirming that the data has headers. Record the number removed.
3. Retain Patient_ID, Age, Gender, Predicted_Disease, Medicine_1, and Dosage_1 in that order. Convert Age to whole numbers; keep IDs and dosage strings as text.
4. Select the cleaned range and use **Insert → Table** with headers. Name the table Patients.
5. Inspect filters and use `COUNTBLANK` on each selected field. Confirm that blank dosages belong to Healthy records with medication None. Check distinct patient IDs separately before treating rows as patients.
6. The workbook uses these formulas over the bounded data range:

   - Patient count: `=COUNTA(Clean_Data!A2:A15001)`; valid here because IDs were independently verified as populated and unique.
   - Average age: `=AVERAGE(Clean_Data!B2:B15001)`.
   - Numeric age count: `=COUNT(Clean_Data!B2:B15001)`.
   - Minimum and maximum: `=MIN(Clean_Data!B2:B15001)` and `=MAX(Clean_Data!B2:B15001)`.
   - Diagnosis count: `=COUNTIFS(Clean_Data!$D$2:$D$15001,A14)`, copied down the diagnosis list.
   - Blank dosages: `=COUNTBLANK(Clean_Data!F2:F15001)`.

7. The chart references the diagnosis summary cells. A formula-based frequency table is used rather than a PivotTable, so existing category counts recalculate when data values change.
8. Verify that diagnosis counts sum to 15,000 and match the cleaned row count. Compare the source and cleaned sheets, and inspect the cleaning log.

The formulas cover the supplied 15,000 records. If adding rows or new diagnosis categories, extend the formula ranges, diagnosis list, and chart range. The supplied diagnosis order reflects the original counts and does not automatically re-sort after edits.

## Validation and limitations

- Counts and age statistics were independently calculated from the downloaded CSV and compared with workbook formula results.
- Workbook sheets were rendered for visual review. This does not establish testing in a live Microsoft Excel session; open the downloaded file in Excel before submission.
- Synthetic frequencies do not estimate real-world disease prevalence. The source does not support clinical conclusions about treatment effectiveness or appropriate dosing.
- Dosages use mixed units and wording. No average dose across different medicines was calculated.
- This workbook and documentation were prepared with AI assistance. Review the formulas and findings and follow your course's rules for assisted work.

## Upload and submission

1. Create a GitHub repository named `healthcare-data-cleaning-excel` and add a README.
2. Choose Public if your assessor needs access without an invitation.
3. Use **Add file → Upload files** to upload Healthcare_Data_Cleaning.xlsx and this README.md. Replace the starter README if one already exists, then commit the changes.
4. Confirm that both files appear on the repository's main page. Open the repository in a private browser window and test the workbook download.
5. Submit your repository URL in the assignment's submission field. The link should point to your own completed project, not the dataset provider's page.
