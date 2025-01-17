import ExcelJS from 'exceljs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

// Create equivalent of `__dirname`
const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

const generateTestFiles = async () => {
    const transactions = [
        { Name: 'John Doe', Description: 'Camp Payment', Amount: 100.0 },
        { Name: 'Jane Smith', Description: 'Membership Fee', Amount: 50.0 },
        { Name: 'Alice Johnson', Description: 'Equipment Purchase', Amount: 75.0 },
        { Name: 'Bob Brown', Description: 'Donation', Amount: 200.0 },
    ];

    const codes = [
        { Code: 'CAMP', Keyword: 'Camp' },
        { Code: 'MEMBER', Keyword: 'Membership' },
        { Code: 'EQUIP', Keyword: 'Equipment' },
        { Code: 'DONATE', Keyword: 'Donation' },
    ];

    const debts = [
        { Name: 'John Doe', AmountOwed: 150.0 },
        { Name: 'Jane Smith', AmountOwed: 100.0 },
        { Name: 'Alice Johnson', AmountOwed: 50.0 },
        { Name: 'Bob Brown', AmountOwed: 0.0 },
    ];

    // Generate Transactions file
    const transactionsWorkbook = new ExcelJS.Workbook();
    const transactionsSheet = transactionsWorkbook.addWorksheet('Transactions');
    transactionsSheet.columns = Object.keys(transactions[0]).map(key => ({ header: key, key }));
    transactions.forEach(row => transactionsSheet.addRow(row));
    await transactionsWorkbook.xlsx.writeFile(join(__dirname, 'data/test_transactions.xlsx'));

    // Generate Codes file
    const codesWorkbook = new ExcelJS.Workbook();
    const codesSheet = codesWorkbook.addWorksheet('Codes');
    codesSheet.columns = Object.keys(codes[0]).map(key => ({ header: key, key }));
    codes.forEach(row => codesSheet.addRow(row));
    await codesWorkbook.xlsx.writeFile(join(__dirname, 'data/test_codes.xlsx'));

    // Generate Debts file
    const debtsWorkbook = new ExcelJS.Workbook();
    const debtsSheet = debtsWorkbook.addWorksheet('Debts');
    debtsSheet.columns = Object.keys(debts[0]).map(key => ({ header: key, key }));
    debts.forEach(row => debtsSheet.addRow(row));
    await debtsWorkbook.xlsx.writeFile(join(__dirname, 'data/test_debts.xlsx'));

    console.log('Test Excel files created successfully!');
};

generateTestFiles().catch(console.error);
