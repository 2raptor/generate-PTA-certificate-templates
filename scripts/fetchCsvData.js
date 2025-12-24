const axios = require('axios');
const csv = require('csv-parser');
const streamifier = require('streamifier');

const fetchCsvData = async (sheetURL) => {
    const data = [];
    try {
        const response = await axios.get(sheetURL);
        streamifier.createReadStream(response.data)
            .pipe(csv())
            .on('data', (row) => {
                // Clean up multi-line column names
                const cleanedRow = {};
                Object.keys(row).forEach(key => {
                    const cleanedKey = key.replace(/[\r\n]+/g, ' ').trim();
                    cleanedRow[cleanedKey] = row[key];
                });

                // Customize as per the needed columns
                const selectedData = {
                    firstName: cleanedRow['Student First Name  - Nombre de Pila del Estudiante Aryan'], 
                    lastName: cleanedRow['Student Last Name - Apellido del Estudiante'],
                    grade: cleanedRow['Grade - Grado'],
                    artCategory: cleanedRow['Arts Category (choose one) - Categoría Artística (Marcar solo una)'].split(' - ')[0],
                    awardArtCategory: cleanedRow['Arts Category'],
                    awardType: cleanedRow['Award Type'],
                    // You can add more columns as needed
                };
                data.push(selectedData);
            })
            .on('end', () => {
                console.log('CSV file successfully processed.', `Total data: ${data.length}`);
            });

        return new Promise((resolve, reject) => {
            setTimeout(() => {
                if (data.length > 0) {
                    resolve(data);
                } else {
                    reject(new Error("No data found"));
                }
            }, 1000);  // Wait for async processing of CSV
        });

    } catch (error) {
        console.error('Error fetching CSV:', error);
        throw error;
    }
};

module.exports = { fetchCsvData };