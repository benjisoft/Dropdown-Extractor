import * as fs from 'fs';
import * as cheerio from 'cheerio';
import * as ExcelJS from 'exceljs';

// Read the HTML file
const html = fs.readFileSync('data.html', 'utf-8');

// Load the HTML into Cheerio
const $ = cheerio.load(html);

// Find all dropdown elements
const dropdowns = $('select');

// Create a new Excel workbook
const workbook = new ExcelJS.Workbook();

// Create a new worksheet
const worksheet = workbook.addWorksheet('Dropdown Options');

// Iterate over each dropdown
dropdowns.each((index, dropdown) => {
  const options = $(dropdown).find('option');

  // Iterate over each option
  options.each((index, option) => {
    const optionText = $(option).text();

    // Write the option to the worksheet
    worksheet.getCell(`A${index + 1}`).value = optionText;
  });
});

// Save the workbook to an Excel file
workbook.xlsx.writeFile('dropdown.xlsx')
  .then(() => {
    console.log('Dropdown options extracted and saved to ../dropdown.xlsx');
  })
  .catch((error) => {
    console.error('Error saving Excel file:', error);
  });
