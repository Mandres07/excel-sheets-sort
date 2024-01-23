const path = require('path');
const ExcelJS = require('exceljs');

async function sortSheetsAlphabetically(inputFilePath, outputFilePath) {
    const workbook = new ExcelJS.Workbook();
    const newWorkbook = new ExcelJS.Workbook();
    try {
        // Load the existing workbook
        await workbook.xlsx.readFile(inputFilePath);
        
        // Sort sheets alphabetically
        let sheetsNames = workbook.worksheets.map(s => s.name).sort();
        
        sheetsNames.forEach(name => {
            const originalSheet = workbook.getWorksheet(name);
            const sortedSheet = newWorkbook.addWorksheet(name);
            sortedSheet.model = originalSheet.model;
        })
        // Save the sorted workbook to a new file
        await newWorkbook.xlsx.writeFile(outputFilePath);

        console.log('Sheets sorted alphabetically successfully.');
    } catch (error) {
        console.error('Error:', error.message);
    }
}

// Example usage
const inputFilePath = 'input-file.xlsx';
const outputFilePath = 'output-file.xlsx';

sortSheetsAlphabetically(inputFilePath, outputFilePath);