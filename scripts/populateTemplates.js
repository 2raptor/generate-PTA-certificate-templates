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
    const sheetURLParticipants = process.env.SHEET_URL_PARTICIPANTS;
    const sheetURLWinners = process.env.SHEET_URL_WINNERS;
    const certificateYear = process.env.CERTIFICATE_YEAR;
    const groupCertificateOfParticipationIntoOne = process.env.GROUP_CERTIFICATE_OF_PARTICIPATION_INTO_ONE === 'true';
    console.log('Group Certificate Of Participation into one certificate:', groupCertificateOfParticipationIntoOne);

    try {
        // Get the current directory of the script
        const currentDir = __dirname;
        const templateDir = path.join(currentDir, `../template/${certificateYear}`);
        const certificateDir = path.join(currentDir, `../certificates/${certificateYear}`);
        // Create directory if needed
        if (!fs.existsSync(certificateDir)) {
            fs.mkdirSync(certificateDir);
            console.log('Directory created:', certificateDir);
        }

        // Load the Word template
        const templatePath = path.resolve(templateDir, 'template.docx');

        // Fetch CSV data
        const jsonDataParticipants = await fetchCsvData(sheetURLParticipants);
        const jsonDataWinners = await fetchCsvData(sheetURLWinners);
        console.log('Fetched Data Count Participants:', jsonDataParticipants.length);
        console.log('Fetched Data Count Winners:', jsonDataWinners.length);

        const jsonData = [...jsonDataWinners, ...jsonDataParticipants];
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

            const fullName = `${data.fullName}`;
            const formattedFullName = fullName.replace(/\s+/g, '-');

            if (data.awardType == "Participating in ") {
                data.awardType = "";
            }

            // use Map for optimzation
            let participationArray = []
            if (groupCertificateOfParticipationIntoOne) {
                if (userMap.has(formattedFullName)) {
                    const userData = userMap.get(formattedFullName);
                    participationArray = userData.participation
                }
                if (!data.awardType && participationArray.indexOf(data.artCategory) === -1) {
                    participationArray.push(data.artCategory);
                }
                userMap.set(formattedFullName, { participation: participationArray });
            } else {
                 if (!data.awardType && participationArray.indexOf(data.artCategory) === -1) {
                    participationArray.push(data.artCategory);
                }
            }

            const artcategoryAndAward = data.awardType ?
                `${data.awardType} in ${data.artCategory}` :
                `Certificate of Participation in ${joinWithAnd(participationArray)}`;
            let formattedGrade = "";
            switch (data.grade) {
                case "K":
                case "Kindergarten":{
                    formattedGrade = "Kindergarten"
                    break;
                }
                case "PK": {
                    formattedGrade = "Pre-Kindergarten"
                    break;
                }
                case "TK": {
                    formattedGrade = "Transitional Kindergarten"
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
                artcategoryandaward: artcategoryAndAward,
                schoolname: data.school
            };  // Adjust based on the structure of your data

            // Replace placeholders with data
            doc.render(dataForTemplate);

            // Generate the new Word document
            const outputBuffer = doc.getZip().generate({ type: 'nodebuffer' });

            // Save the new document
            let participationOrAward = data.awardType ? data.artCategory.replace(/\s+/g, '-') : 'participating';
            if (participationOrAward === 'participating' && !groupCertificateOfParticipationIntoOne) {
                participationOrAward += `-${data.artCategory.replace(/\s+/g, '-')}`;
            }
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
