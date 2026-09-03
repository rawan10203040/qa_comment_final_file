# 📊 QA Comment Updater & Summary Generator

A Streamlit application that processes Excel files and automatically updates the `QAComment` column based on specific QA rules.

The application also uses a predefined Function Reference to add:

* `Area`
* `Group`
* `Team leader`

Then it removes unnecessary rows and generates a consolidated summary of the remaining QA issues.

---

## 🚀 Features

### 1. Upload Multiple Excel Files

The application supports uploading one or multiple `.xlsx` files at the same time.

Each uploaded file is processed independently.

---

### 2. Update `QAComment`

The application evaluates the following columns:

* `FunctionName`
* `IsMultiValue`
* `HasBlankValue`
* `QAComment`

The `QAComment` value is automatically updated according to the following rules.

### Packing

| IsMultiValue | HasBlankValue | QAComment  |
| ------------ | ------------- | ---------- |
| TRUE         | FALSE         | `ok`       |
| TRUE         | TRUE          | `null`     |
| FALSE        | FALSE         | `conflict` |
| FALSE        | TRUE          | `conflict` |

### Other Functions

| IsMultiValue | HasBlankValue | QAComment  |
| ------------ | ------------- | ---------- |
| TRUE         | FALSE         | `conflict` |
| TRUE         | TRUE          | `conflict` |
| FALSE        | FALSE         | `ok`       |
| FALSE        | TRUE          | `null`     |

If none of the conditions are matched, the original `QAComment` value is kept.

---

## 📋 Function Reference

The application contains a predefined reference mapping based on `FunctionName`.

The reference contains:

| FunctionName                 | Action   | Area          | Group         | Team leader          |
| ---------------------------- | -------- | ------------- | ------------- | -------------------- |
| Part Options                 | analysis | -             | -             | -                    |
| Other Qualification Features | ok       | Qualification | Qualification | Mohamed Farouk       |
| Lifecycle                    | ok       | Risk          | Lifecycle     | Hamed mohamed        |
| Die Family                   | ok       | Parts         | Parts         | Mohamed Sayed Farouk |
| Forecast PLP                 | ok       | Forecast      | Forecast      | Ahmed Abotaleb       |
| TradeCodes                   | ok       | Trade Codes   | Trade Codes   | Ahmed Abotaleb       |
| Part Marking Image           | ok       | Part Image    | Parts         | Abdelrahman Ata      |
| Chemical                     | ok       | Chemical      | Compliance    | Mohamed Nabil        |
| Country of Origin            | ok       | SC location   | SC sites      | Mostafa Ahmed        |
| MSL                          | ok       | MFG - Main    | MFG           | Ahmed ibraheem       |
| Package                      | ok       | Package       | Parts         | Abdelrahman Ata      |
| ProcessTechnology            | ok       | Technology    | Technology    | Mostafa Ahmed        |

The complete reference mapping is included directly inside the application.

---

## ➕ Add Reference Columns

After updating `QAComment`, the application adds the following columns based on `FunctionName`:

```text
Area
Group
Team leader
```

The columns are positioned next to `FunctionName`.

Example:

```text
FunctionName
Area
Group
Team leader
IsMultiValue
HasBlankValue
QAComment
...
```

---

## 🗑️ Remove Unnecessary Rows

After applying the reference mapping, the application removes rows according to the following rules.

### Remove `Action = analysis`

Any row where:

```text
Action = analysis
```

is removed.

### Remove `QAComment = ok`

Any row where:

```text
QAComment = ok
```

is removed.

Therefore, the final output contains the rows that require attention, mainly:

```text
conflict
null
```

---

## 📊 Summary

After processing all uploaded files, the application creates one consolidated Summary.

The Summary is based on rows where:

```text
QAComment = conflict
OR
QAComment = null
```

The Summary contains:

| Area          | Counts | Team Leader     |
| ------------- | -----: | --------------- |
| Part Image    |  22460 | Abdelrahman Ata |
| Package       |    185 | Abdelrahman Ata |
| PinOut        |    195 | Abdelrahman Ata |
| Forecast      |  12278 | Ahmed Abotaleb  |
| Trade Codes   |   7650 | Ahmed Abotaleb  |
| MFG - Main    |  27442 | Ahmed ibraheem  |
| MFG - Others  |   2744 | Ahmed ibraheem  |
| Risk          |  78965 | Hamed mohamed   |
| Market        |    827 | Hatem Hussein   |
| Qualification |  13065 | Mohamed Farouk  |

If the same `Area` and `Team Leader` appear in multiple uploaded files, their counts are combined.

---

## 📥 Output Files

### One Input File

If only one Excel file is uploaded, the application generates:

```text
OriginalFile_Updated.xlsx
```

---

### Multiple Input Files

If multiple Excel files are uploaded, each file gets its own updated output:

```text
File1_Updated.xlsx
File2_Updated.xlsx
File3_Updated.xlsx
```

The application also creates:

```text
Updated_Output_Files.zip
```

containing all updated Excel files.

---

## 📋 Summary Output

The Summary can be downloaded separately as:

```text
QA_Summary.xlsx
```

The Summary contains:

```text
Area
Counts
Team Leader
```

---

## 📈 Processing Statistics

For each processed file, the application displays:

* Original Rows
* Analysis Removed
* OK Removed
* Final Rows

Example:

```text
Original Rows:       100,000
Analysis Removed:     20,000
OK Removed:           30,000
Final Rows:           50,000
```

---

## 👀 Data Preview

After processing each file, the application displays a preview of the first 100 rows.

This allows the user to verify:

* `FunctionName`
* `Area`
* `Group`
* `Team leader`
* `QAComment`
* Other original columns

before downloading the result.

---

# 🛠️ Installation

Clone the repository:

```bash
git clone <YOUR_GITHUB_REPOSITORY_URL>
cd <YOUR_PROJECT_FOLDER>
```

Install the required libraries:

```bash
pip install -r requirements.txt
```

---

# 📦 Requirements

The project uses:

```text
streamlit
polars
fastexcel
openpyxl
xlsxwriter
pandas
```

These packages are listed in:

```text
requirements.txt
```

---

# ▶️ Run Locally

Run the Streamlit application using:

```bash
streamlit run app.py
```

The application will open in your browser.

---

# ☁️ Deploy on Streamlit Cloud

### 1. Push the project to GitHub

Your repository should contain:

```text
QA-Comment-Updater/
│
├── app.py
│
└── requirements.txt
```

### 2. Open Streamlit Community Cloud

Go to Streamlit Community Cloud and connect your GitHub repository.

### 3. Select the repository

Choose the repository containing:

```text
app.py
```

### 4. Set the Main File

Use:

```text
app.py
```

### 5. Deploy

Streamlit will automatically install the packages from:

```text
requirements.txt
```

and start the application.

---

# 📁 Project Structure

```text
QA-Comment-Updater/
│
├── app.py
│
├── requirements.txt
│
└── README.md
```

---

# 🔄 Processing Flow

The complete processing flow is:

```text
Upload Excel Files
        ↓
Read Excel
        ↓
Validate Required Columns
        ↓
Clean Data
        ↓
Update QAComment
        ↓
Match FunctionName with Reference
        ↓
Add Area
        ↓
Add Group
        ↓
Add Team leader
        ↓
Remove Action = analysis
        ↓
Remove QAComment = ok
        ↓
Keep conflict / null
        ↓
Generate Updated Excel Files
        ↓
Generate Consolidated Summary
        ↓
Download Excel / ZIP / Summary
```

---

# ⚠️ Required Input Columns

Every input Excel file must contain these columns:

```text
FunctionName
IsMultiValue
HasBlankValue
QAComment
```

If any required column is missing, the file will be skipped and an error message will be displayed.

---

# 🧪 Example

### Input

| FunctionName | IsMultiValue | HasBlankValue | QAComment |
| ------------ | ------------ | ------------- | --------- |
| Lifecycle    | TRUE         | FALSE         |           |
| Die Family   | FALSE        | FALSE         |           |
| Packing      | TRUE         | TRUE          |           |
| Package      | FALSE        | TRUE          |           |

### After QAComment Rules

| FunctionName | IsMultiValue | HasBlankValue | QAComment |
| ------------ | ------------ | ------------- | --------- |
| Lifecycle    | TRUE         | FALSE         | conflict  |
| Die Family   | FALSE        | FALSE         | ok        |
| Packing      | TRUE         | TRUE          | null      |
| Package      | FALSE        | TRUE          | null      |

### After Reference Mapping

| FunctionName | Area    | Group     | Team leader          | QAComment |
| ------------ | ------- | --------- | -------------------- | --------- |
| Lifecycle    | Risk    | Lifecycle | Hamed mohamed        | conflict  |
| Die Family   | Parts   | Parts     | Mohamed Sayed Farouk | ok        |
| Packing      | -       | -         | -                    | null      |
| Package      | Package | Parts     | Abdelrahman Ata      | null      |

### Final Output

Rows with:

```text
Action = analysis
```

and:

```text
QAComment = ok
```

are removed.

The final file contains only the rows requiring attention.

---

# 💡 Purpose

This tool is designed to automate repetitive QA data-cleaning tasks and make it easier to:

* Process large Excel files
* Apply consistent QA rules
* Map Functions to Areas and Teams
* Identify `conflict` and `null` cases
* Assign issues to the correct Team Leader
* Generate consolidated QA statistics
* Reduce manual Excel work

---

## 👩‍💻 Author

**Rawan Osama**

QA / Data Analysis Automation

```

لو هتحطيه على GitHub، احفظيه باسم **`README.md`** في نفس الـ repository الموجود فيه `app.py` و`requirements.txt`.
```
