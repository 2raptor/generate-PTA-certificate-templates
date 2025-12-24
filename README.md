# generate templates in word doc
This project was create to generate certificates from a Microsoft word template for students. 
This project used the reflection certification pack from here [PTA Reflections- Certificate Packs] (https://stores.shoppta.com/product_detail.lasso?id=1022&-session=KI_Clients:18067DA402b7525DDEtn4902A3EA)

`populateTemplates.js` has sheetURL. This is Google Sheet which is set to "Published to Web" as csv. (File->Share->Publish to Web). 

## `fetchCsvData` Script
This script fetches data from the sheetURL using axios APIs. The script is coded for the specific column/rows currently in the sheetURL. To suit your needs remember please udpate the `selectedData` accordingly.

Simple way to understand the script like reading array of index with key as column name for each row.

| Column 1  | Column 2  |
| Row1 Val1 | Row1 Val2 |
| Row2 Val1 | Row2 Val2 |

In the 1st loop row 1 will be read as using the column name as key. Likewise each row is read.
cleanedRow['Column 1'] => Row1 Val1
cleanedRow['Column 2'] => Row1 Val2

In the 2nd loop
cleanedRow['Column 1'] => Row2 Val1
cleanedRow['Column 2'] => Row2 Val2

## `populateTemplates` Script
This script gets the content of the file and then does formatting and sanitization. The script tries to create unique filename for each student for one certificate kind. Here the list of combinations

Students will get one kind of certificate
1. articpation certificate
    "Certificate of Participation in Music Composition"
2. Award of excellence certificate
    "Award of Excellence in Music Composition"
3. Honorable mentions certificate
    "Honorable Mention in Visual Arts"

Only for "Certificate of Participation" it will have one or more category
    1. Uunder one category
    "Certificate of Participation in Music Composition"

    1. Under two category
    "Certificate of Participation in Music Composition and Photography"

    1. Under multiple category
    "Certificate of Participation in Literature Visual Arts, Music Composition and Photography"

In this combination, one student may get one certificates for participation even if they have submitted on art on mutiple category. Like if student has submitted art work for 4 category. That student may get one "Certificate of Participation" mentioning all the category(Visual Arts, Film Production, Music Composition, Literature and Photography) or one "Certificate of Participation" and other kind of certificates.

This script uses a `userMap` for each certificate-kind but it accumulate the category for participation in `participationArray` to avoid multiple entries for the same student.

John-Doe-participating.docx
John-Doe-Visual-Arts.docx

Remember the script will overwrite if the file already exists, it is convenient to re-run the script multiple times.

It open `24-25_Reflections_Cerfiticate_Template.doc` replaces the `placeholder1` and `placeholder2` and stores the template inside `./certificates` directory. 

```
npm i
node ./scripts/populateTemplates.js
```