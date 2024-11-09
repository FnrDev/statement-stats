const XLSX = require('xlsx');

// Load the Excel file
const filePath = './account_statement.xlsx'; // Update with your file path
const workbook = XLSX.readFile(filePath);

// Select the first sheet
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

// Get data from the worksheet
const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

// Function to convert Excel serial number to Date
function excelDateToJSDate(serial) {
    const utc_days  = Math.floor(serial - 25569);
    const utc_value = utc_days * 86400;                                        
    const date_info = new Date(utc_value * 1000);
    return new Date(date_info.getFullYear(), date_info.getMonth(), date_info.getDate());
}

// Add classification to each row
let totalCredit = 0;
let totalDebit = 0;

const startDate = new Date('2024-09-29');
const endDate = new Date('2024-11-01');

const results = data.map((row, index) => {
    // Skip header row
    if (index === 0) return [...row, 'Classification'];

    const excelDate = row[3];
    const transactionDate = excelDateToJSDate(excelDate);
    const value = parseFloat(row[5]);
    const classification = value < 0 ? 'Debit' : 'Credit';
    
    // Update totals only if date is within range
    if (transactionDate >= startDate && transactionDate <= endDate) {
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
console.log('Date Range:', '29/09/2024 to 01/11/2024');
console.log('Total Credit:', totalCredit.toFixed(2), 'BHD');
console.log('Total Debit:', totalDebit.toFixed(2), 'BHD');

// Log top 10 transactions
console.log('\nTop 10 Credit Transactions:');
top10Credits.forEach((trans, index) => {
    console.log(`${index + 1}. ${trans.date.toLocaleDateString()} - ${trans.description}: ${trans.amount.toFixed(2)} BHD`);
});

console.log('\nTop 10 Debit Transactions:');
top10Debits.forEach((trans, index) => {
    console.log(`${index + 1}. ${trans.date.toLocaleDateString()} - ${trans.description}: ${trans.amount.toFixed(2)} BHD`);
});
