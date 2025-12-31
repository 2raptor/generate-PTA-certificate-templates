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
                    // console.log('Cleaned Key:', cleanedKey);
                    cleanedRow[cleanedKey] = row[key];
                });

                // Customize as per the needed columns
                const selectedData = {
                    fullName: cleanedRow['Name'], 
                    grade: cleanedRow['Grade'],
                    awardType: cleanedRow['Award'],
                    artCategory: cleanedRow['Category'],
                    school: cleanedRow['School'],
                    // You can add more columns as needed
                };
                // console.log('Selected Data:', selectedData);
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
