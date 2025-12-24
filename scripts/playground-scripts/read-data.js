const axios = require('axios');
const csv = require('csv-parser');
const streamifier = require('streamifier');

// Published Google Sheets CSV URL
const sheetURL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSL3lDUE-G8n0dAtcsW3awJUUNqvJ9m8XTrqPDZpfcA9t4KoJqzWFvMgPd6Bvr5nDgJ4TSujwIKxFbi/pub?output=csv";

const data = [];  // Array to store the processed data

// Fetch the CSV file from the URL
axios.get(sheetURL)
    .then(response => {
        console.log(response.data);
        // The data from Google Sheets is in response.data, now parse it
        streamifier.createReadStream(response.data)  // Turn CSV data into stream
            .pipe(csv())  // Parse the CSV
            .on('data', (row) => {
                // Clean up multi-line column names by removing line breaks and extra spaces
                const cleanedRow = {};
                Object.keys(row).forEach(key => {
                    const cleanedKey = key.replace(/[\r\n]+/g, ' ').trim(); // Replace line breaks and trim spaces
                    cleanedRow[cleanedKey] = row[key];
                });

                const selectedData = {
                    firstName: cleanedRow['Student First Name  - Nombre de Pila del Estudiante Aryan'], 
                    lastName: cleanedRow['Student Last Name - Apellido del Estudiante'],
                    grade: cleanedRow['Grade - Grado'],
                    artCategory: cleanedRow['Arts Category'],
                    awardType: cleanedRow['Award Type'],
                    // You can add more columns as needed
                };
                data.push(selectedData);
            })
            .on('end', () => {
                console.log('CSV file successfully processed.');
                console.log(data);  // Here, `data` will contain an array of selected columns

                // Pass the selected data to your docxtemplater code or other logic
            });
    })
    .catch(error => {
        console.error('Error fetching CSV:', error);
    });
