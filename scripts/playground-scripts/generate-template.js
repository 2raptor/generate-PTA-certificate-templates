const fs = require("fs");
const path = require("path");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");

// Load the Word template
const templatePath = path.resolve(__dirname, "template.docx");
const content = fs.readFileSync(templatePath, "binary");

// Load the template into Docxtemplater
const zip = new PizZip(content);
const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
});

// JSON data with placeholder replacements
const jsonData = {
    placeholder1: "John Doe",
    placeholder2: "2025-01-11",
};

try {
    // Replace placeholders with data
    doc.render(jsonData);
} catch (error) {
    console.error("Error rendering the document:", error);
    return;
}

// Generate the new Word document
const outputBuffer = doc.getZip().generate({ type: "nodebuffer" });

// Save the new document
const outputPath = path.resolve(__dirname, jsonData.placeholder1.replace(/\s+/g, '-')+".docx");
fs.writeFileSync(outputPath, outputBuffer);

console.log("Document generated successfully at:", outputPath);