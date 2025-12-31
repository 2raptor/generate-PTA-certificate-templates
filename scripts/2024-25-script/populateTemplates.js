require('dotenv').config();
const { fetchCsvData } = require('./fetchCsvData');  // Import the CSV fetching function
const path = require('path');
const fs = require('fs');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');

function joinWithAnd(array) {
    if (array.length === 0) return '';
    if (array.length === 1) return array[0];
    if (array.length === 2) return array.join(' and ');
    return array.slice(0, -1).join(', ') + ' and ' + array[array.length - 1];
}

const updateTemplateWithData = async () => {
    // Published Google Sheets CSV URL
    const sheetURL = process.env.SHEET_URL;
    const certificateYear = process.env.CERTIFICATE_YEAR;

    try {
        // Get the current directory of the script
        const currentDir = __dirname;
        const templateDir = path.join(currentDir, `../../template/${certificateYear}`);
        const certificateDir = path.join(currentDir, `../../certificates/${certificateYear}`);
        // Create directory if needed
        if (!fs.existsSync(certificateDir)) {
            fs.mkdirSync(certificateDir);
            console.log('Directory created:', certificateDir);
        }

        // Load the Word template
        const templatePath = path.resolve(templateDir, 'template.docx');

        // Fetch CSV data
        const jsonData = await fetchCsvData(sheetURL);
        console.log('Fetched Data Count:', jsonData.length);

        // userhash for merging data
        const userMap = new Map();

        jsonData.forEach((data, index) => {
            const content = fs.readFileSync(templatePath, 'binary');

            // Load the template into Docxtemplater
            const zip = new PizZip(content);
            const doc = new Docxtemplater(zip, {
                paragraphLoop: true,
                linebreaks: true,
            });

            const fullName = `${data.firstName} ${data.lastName}`;
            const formattedFullName = fullName.replace(/\s+/g, '-');

            // use Map for optimzation
            let participationArray = []
            if (userMap.has(formattedFullName)) {
                const userData = userMap.get(formattedFullName);
                participationArray = userData.participation
            }
            if (!data.awardType && participationArray.indexOf(data.artCategory) === -1) {
                participationArray.push(data.artCategory);
            }
            userMap.set(formattedFullName, { participation: participationArray });


            const artcategoryAndAward = data.awardType ?
                `${data.awardType} in ${data.artCategory}` :
                `Certificate of Participation in ${joinWithAnd(participationArray)}`;
            let formattedGrade = "";
            switch (data.grade) {
                case "K": {
                    formattedGrade = "Kindergarten"
                    break;
                }
                case "PK": {
                    formattedGrade = "Pre-Kindergarten"
                    break;
                }
                default: {
                    formattedGrade = `Grade ${data.grade}`
                    break;
                }
            }
            const dataForTemplate = {
                name: fullName,
                grade: formattedGrade,
                artcategoryandaward: artcategoryAndAward
            };  // Adjust based on the structure of your data

            // Replace placeholders with data
            doc.render(dataForTemplate);

            // Generate the new Word document
            const outputBuffer = doc.getZip().generate({ type: 'nodebuffer' });

            // Save the new document
            const participationOrAward = data.awardType ? data.artCategory.replace(/\s+/g, '-') : 'participating';
            const fileNameToSave = `${formattedFullName}-${participationOrAward}.docx`;
            const outputPath = path.resolve(certificateDir, fileNameToSave);
            if (fs.existsSync(outputPath)) {
                console.log('       File already exists:', outputPath);
            }
            fs.writeFileSync(outputPath, outputBuffer);


            console.log(index, 'Document generated successfully at:', outputPath);
        });
    } catch (error) {
        console.error('Error updating template:', error);
    }
};

// Invoke the function to update the template
updateTemplateWithData();
