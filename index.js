const XLSX = require('xlsx');
const moment = require('moment');

// Load the Excel file
const filePath = './account_statement.xlsx'; // Update with your file path
const workbook = XLSX.readFile(filePath);

// Select the first sheet
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

// Get data from the worksheet
const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

// Updated Excel to JS Date conversion function using moment
function excelDateToJSDate(serial) {
    if (!serial) return null;
    
    // Excel's epoch starts from 1900-01-01
    const utc_days = Math.floor(serial - 25569);
    const milliseconds = utc_days * 24 * 60 * 60 * 1000;
    return moment(milliseconds);
}

// Add classification to each row
let totalCredit = 0;
let totalDebit = 0;

// Update date range definitions using moment
const startDate = moment('2024-09-29');
const endDate = moment('2024-11-01');

const results = data.map((row, index) => {
    // Skip header row
    if (index === 0) return [...row, 'Classification'];

    const excelDate = row[3];
    const transactionDate = excelDateToJSDate(excelDate);
    const value = row[5];
    const classification = value < 0 ? 'Debit' : 'Credit';
    
    // Update date comparison using moment
    if (transactionDate && transactionDate.isBetween(startDate, endDate, 'day', '[]')) {
        if (classification === 'Credit') {
            totalCredit += value;
        } else {
            totalDebit += Math.abs(value);
        }
    }

    return [...row, classification];
});

// Separate transactions into credits and debits
const creditTransactions = [];
const debitTransactions = [];

results.forEach((row, index) => {
    if (index === 0) return; // Skip header

    const excelDate = row[3];
    const transactionDate = excelDateToJSDate(excelDate);
    const value = row[5];
    const description = row[4]; // Assuming description is in column 4
    
    const transaction = {
        date: transactionDate,
        amount: Math.abs(value),
        description: description
    };

    if (value < 0) {
        debitTransactions.push(transaction);
    } else {
        creditTransactions.push(transaction);
    }
});

// Sort transactions by amount (descending) and get top 10
const top10Credits = creditTransactions
    .sort((a, b) => b.amount - a.amount)
    .slice(0, 10);

const top10Debits = debitTransactions
    .sort((a, b) => b.amount - a.amount)
    .slice(0, 10);

// Log the final totals
console.log('\nFinal Totals:');
console.log('Date Range:', startDate.format('DD/MM/YYYY'), 'to', endDate.format('DD/MM/YYYY'));
console.log('Total Credit:', totalCredit.toFixed(2), 'BHD');
console.log('Total Debit:', totalDebit.toFixed(2), 'BHD');

// Log top 10 transactions
console.log('\nTop 10 Credit Transactions:');
top10Credits.forEach((trans, index) => {
    console.log(`${index + 1}. ${trans.date.format('DD/MM/YYYY')} - ${trans.description}: ${trans.amount.toFixed(2)} BHD`);
});

console.log('\nTop 10 Debit Transactions:');
top10Debits.forEach((trans, index) => {
    console.log(`${index + 1}. ${trans.date.format('DD/MM/YYYY')} - ${trans.description}: ${trans.amount.toFixed(2)} BHD`);
});
