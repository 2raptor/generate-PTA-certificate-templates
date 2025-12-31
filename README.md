# Generate Templates in Word Documents

This project generates student certificates from a Microsoft Word template.

It uses the **Reflections Certificate Pack** from PTA, available here:\
[PTA Reflections – Certificate Packs 2025-26](https://capta.shoppta.com/category/Reflections)

The script `populateTemplates.js` contains a `sheetURL`, which points to a Google Sheet published to the web as a CSV:

**Google Sheets → File → Share → Publish to web**

---

## How to Run

### 1. Update `.env` file

#### Create or edit a `.env` file in the project root and set the following variables:

```
SHEET_URL_WINNERS={CSV_EXPORTED_URL}
SHEET_URL_PARTICIPANTS={CSV_EXPORTED_URL}
GROUP_CERTIFICATE_OF_PARTICIPATION_INTO_ONE={true or false} # e.g.,: false
CERTIFICATE_YEAR={SOME_YEAR_VALUE} # e.g.,: 2026-27
SCHOOL_NAME={SOME SCHOOL NAME}
```

### 2. Prepare the directory

```bash
# Set the CERTIFICATE_YEAR variable. Use the same value used in `.env` file from Step 1
CERTIFICATE_YEAR="2026-27"

# Create directories using the variable
mkdir -p "template/${CERTIFICATE_YEAR}"
mkdir -p "certificates/${CERTIFICATE_YEAR}"

# Add .gitkeep inside the leaf folders
touch "template/${CERTIFICATE_YEAR}/.gitkeep"
touch "certificates/${CERTIFICATE_YEAR}/.gitkeep"
```

### 3. Adjusting the template placeholder

Template provided on the PTA website usually has only the following two placeholders in form of **Text Box**.
<img width="600" alt="template-from-website" src="https://github.com/user-attachments/assets/c51af8e6-5a56-4b5a-988a-b88a7f7309f1" />


**How to locate Text Box in Microsoft Word Menu:**  
`Menu → Insert → Draw Text Box` but it is easier to copy paste the existing text box and reposition them as per your need.

<img width="350" height="147" alt="microsoft-word-text-box-menu" src="https://github.com/user-attachments/assets/ad524ede-a34c-4fd7-8dd6-2b7b899126f9" />



Update those text boxes with placeholders values wrapped inside curly braces `{}` like `{name}`, `({grade})`, `{artcategoryandaward}` and `{schoolname}`. Add additional text box if needed. The current script support 3 fields but it is easy to update the script if you are adding additional text box or placeholders within the existing text box.
<img width="600" alt="template-placehoder-customized" src="https://github.com/user-attachments/assets/ca8fb27e-6fe4-4ba8-b8d9-ee42bde2f92e" />


Align additional text boxes carefully, do test prints, and once ready, save as `template.docx` inside `template/${CERTIFICATE_YEAR}/`.

### 4. Run these commands

```bash
# Be in the root directory
cd ~/generate-PTA-certificate-templates
npm install
node ./scripts/populateTemplates.js
```

### 5. Open generated certificates templates for print

Inspect for long student names or multiple participation entries and adjust font size if needed. Test print before using actual certificates to avoid paper/ink wastage.

Inspect the generated certificate templates. Some students' names may be long, and since the name and grade are combined in the same text box, text might overflow to the next line. Reduce font size as needed to fit all text in one line. The same may occur when `GROUP_CERTIFICATE_OF_PARTICIPATION_INTO_ONE` is `true`; the second text box may overflow if a student has multiple participation entries. Although uncommon, inspect each certificate and adjust accordingly.

Before printing, understand how your certificates need to be positioned in your printer. Take test prints to verify alignment. If your printer struggles to pull multiple sheets from the feed, load one certificate at a time to avoid wastage. Check test prints to ensure ink or toner is sufficient. Being vigilant helps prevent wasting certificates due to printer issues.

---

## Important Notes

- The script **overwrites files** if they already exist\
  (This makes re-running the script convenient.)
- If you change the value of `GROUP_CERTIFICATE_OF_PARTICIPATION_INTO_ONE`, it is recommended to:
  - Delete previously generated certificates in:
    ```
    #(Use the appropriate year directory)
    certificates/{CERTIFICATE_YEAR}
    ```
  - Then re-run the script
  
---
# Understanding the script
## `fetchCsvData` Script

This script fetches CSV data from the `sheetURL` using Axios API.

The script is currently written to match the specific columns and rows in the Google Sheet.\
If your sheet structure changes, update the `selectedData` configuration accordingly.

### How the Script Works (Conceptually)

Think of the CSV as an array of rows, where each row is converted into an object using column names as keys.

| Column 1  | Column 2  |
| --------- | --------- |
| Row1 Val1 | Row1 Val2 |
| Row2 Val1 | Row2 Val2 |

### First Loop (Row 1)

```js
cleanedRow['Column 1'] => Row1 Val1
cleanedRow['Column 2'] => Row1 Val2
```

### Second Loop (Row 2)

```js
cleanedRow['Column 1'] => Row2 Val1
cleanedRow['Column 2'] => Row2 Val2
```

Each row is processed in the same way.

---

## `populateTemplates` Script

This script:

1. Calls `fetchCsvData` for two different sheet URLs:
   - `SHEET_URL_WINNERS`
   - `SHEET_URL_PARTICIPANTS`
2. Merges the data
3. Formats and sanitizes the values
4. Generates certificate files based on rules and flags

### Certificate Grouping Behavior

The script attempts to create a unique filename per student per certificate type.\
This behavior is controlled by the flag:

```js
GROUP_CERTIFICATE_OF_PARTICIPATION_INTO_ONE
```

---

## When `GROUP_CERTIFICATE_OF_PARTICIPATION_INTO_ONE` is `true`

Each student may receive the following **certificate types**:

1. **Certificate of Participation**

   - `Certificate of Participation in Music Composition`

2. **Award of Excellence**

   - `Award of Excellence in Music Composition`

3. **Honorable Mention**

   - `Honorable Mention in Visual Arts`

### Participation Certificate Combinations

Only **Certificate of Participation** can span multiple categories.

#### One Category

```
Certificate of Participation in Music Composition
```

#### Two Categories

```
Certificate of Participation in Music Composition and Photography
```

#### Multiple Categories

```
Certificate of Participation in Literature, Visual Arts, Music Composition, and Photography
```

In this mode, a student who submitted artwork in multiple categories will receive **one participation certificate** listing all categories instead of multiple participation certificates.

However, they may still receive **separate certificates** for awards or honorable mentions.

---

### Internal Handling

- The script uses a `userMap` per certificate type
- Participation categories are accumulated in `participationArray`
- This prevents duplicate participation certificates for the same student

Example generated files when `GROUP_CERTIFICATE_OF_PARTICIPATION_INTO_ONE` is `true`:

```
# for participation certificate no category name will be appended as all will be merged into one certificate
John-Doe-participating.docx
# when they win a award of excellence or honorable mentions only category name will be appended
John-Doe-Visual-Arts.docx
```

---

#### When `GROUP_CERTIFICATE_OF_PARTICIPATION_INTO_ONE` is `false`

Each participant receives **one certificate per category**, regardless of awards.

If a student participates in five categories, they will receive **five certificates**.

Example generated files when `GROUP_CERTIFICATE_OF_PARTICIPATION_INTO_ONE` is `false`:

```
# for participation certificate the category name will be appended
John-Doe-participating-Literature.docx
John-Doe-participating-Choreography.docx
# when they win a award of excellence or honorable mentions only category name will be appended
John-Doe-Visual-Arts.docx
```


---


## Template Processing Explained

The script opens template downloaded from the [PTA website](https://capta.shoppta.com/category/Reflections):

```
24-25_Reflections_Certificate_Template.doc
```

The script replaces:

- `placeholder1`
- `placeholder2`

The generated certificates are stored in:

```
./certificates
```

Templates are placed under this directory with school year as sub-directory. Use `CERTIFICATE_YEAR` from `.env` to point to the correct template:

```
./template/
```

---

## Examples that explain the script output

### Student Participated in 5 Categories, No Awards

#### `GROUP_CERTIFICATE_OF_PARTICIPATION_INTO_ONE = false`

They receive **5 certificates**:

1. Certificate of Participation in Music Composition
2. Certificate of Participation in Photography
3. Certificate of Participation in Visual Arts
4. Certificate of Participation in Dance Choreography
5. Certificate of Participation in Literature

#### `GROUP_CERTIFICATE_OF_PARTICIPATION_INTO_ONE = true`

They receive **1 certificate**:

1. Certificate of Participation in Music Composition, Photography, Visual Arts, Dance Choreography, and Literature

### Student Participated in 5 Categories and Won 2 Awards

#### `GROUP_CERTIFICATE_OF_PARTICIPATION_INTO_ONE = false`

They receive **5 certificates**:

1. Award of Excellence in Music Composition
2. Honorable Mention in Photography
3. Certificate of Participation in Visual Arts
4. Certificate of Participation in Dance Choreography
5. Certificate of Participation in Literature

#### `GROUP_CERTIFICATE_OF_PARTICIPATION_INTO_ONE = true`

They receive **3 certificates**:

1. Award of Excellence in Music Composition
2. Honorable Mention in Photography
3. Certificate of Participation in Visual Arts, Dance Choreography, and Literature

### Student Participated in 5 Categories and Won 5 Awards

#### `GROUP_CERTIFICATE_OF_PARTICIPATION_INTO_ONE = false`

They receive **5 certificates**:

1. Award of Excellence in Music Composition
2. Honorable Mention in Photography
3. Award of Excellence in Visual Arts
4. Honorable Mention in Dance Choreography
5. Honorable Mention in Literature

#### `GROUP_CERTIFICATE_OF_PARTICIPATION_INTO_ONE = true`

They receive **5 certificates**:

1. Award of Excellence in Music Composition
2. Honorable Mention in Photography
3. Award of Excellence in Visual Arts
4. Honorable Mention in Dance Choreography
5. Honorable Mention in Literature

