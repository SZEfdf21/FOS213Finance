// TEST 1 (VOOR EXCELHANDLER.JS)
// const ExcelHandler = require('./src/services/ExcelHandler');

// // Voorbeeld gebruik ExcelHandler
// (async () => {
//     try {
//         // Test read functionaliteit
//         const data = await ExcelHandler.readExcel('./data/input/sample.xlsx', 'Samenvatting');
//         console.log('Read Data:', data);

//         // Test write functionaliteit
//         const newData = [["Name", "Age"], ["John Doe", 30], ["Jane Smith", 25]];
//         await ExcelHandler.writeExcel('./data/output/output.xlsx', newData);
//         console.log('Write Successful!');
//     } catch (error) {
//         console.error('Error:', error.message);
//     }
// })();

// TEST 2 (VOOR TRANSACTIONPROCESSOR.JS)
// const ExcelHandler = require('./src/services/ExcelHandler');
// const TransactionProcessor = require('./src/services/TransactionProcessor');

// // Voorbeeld gebruik TransactionProcessor
// (async () => {
//     try {
//         // Test data
//         const sampleTransactions = [
//             { name: 'John Doe', description: 'Payment for Camp', amount: 100 },
//             { name: 'Jane Smith', description: 'Membership Fee', amount: 50 },
//         ];

//         const sampleCodes = [
//             { code: 'CAMP', keyword: 'Camp' },
//             { code: 'MEMBERSHIP', keyword: 'Membership' },
//         ];

//         // Match transacties met codes
//         const matchedTransactions = TransactionProcessor.matchTransactions(sampleTransactions, sampleCodes);
//         console.log('Matched Transactions:', matchedTransactions);

//         // Update schulden in Excel file
//         const debtFilePath = './data/input/sampleDebts.xlsx';
//         const debtSheetName = 'Debts';
//         await TransactionProcessor.updateDebts(matchedTransactions, debtFilePath, debtSheetName);
//         console.log('Debts updated successfully');

//         // Generate samenvatting van transacties
//         const transactionSummary = TransactionProcessor.generateSummary(matchedTransactions);
//         console.log('Transaction Summary:', transactionSummary);
//     } catch (error) {
//         console.error('Error:', error.message);
//     }
// })();

//TEST 3 (VOOR INVOICEGENERATOR.JS)
// const InvoiceGenerator = require('./src/services/InvoiceGenerator');
// const fs = require('fs');
// const path = require('path');

// (async () => {
//     try {
//         // Test data voor een invoice
//         const invoiceData = {
//             recipient: 'John Doe',
//             date: '2025-01-16',
//             items: [
//                 { name: 'Membership Fee', amount: '50.00' },
//                 { name: 'Camp Payment', amount: '100.00' },
//             ],
//             total: '150.00',
//         };

//         // Genereer HTML invoice
//         const htmlContent = InvoiceGenerator.generateInvoiceHTML(invoiceData);

//         // Save HTML file
//         const htmlFilePath = path.join(__dirname, 'data/output/invoice.html');
//         fs.writeFileSync(htmlFilePath, htmlContent);
//         console.log('HTML Invoice saved at:', htmlFilePath);

//         // Save PDF file
//         const pdfFilePath = path.join(__dirname, 'data/output/invoice.pdf');
//         await InvoiceGenerator.saveInvoicePDF(htmlContent, pdfFilePath);
//         console.log('PDF Invoice saved at:', pdfFilePath);

//     } catch (error) {
//         console.error('Error:', error.message);
//     }
// })();

// TEST 4 (VOOR EMAILHANDLER.JS)
// const EmailHandler = require('./src/services/EmailHandler');
// const InvoiceGenerator = require('./src/services/InvoiceGenerator');
// const fs = require('fs');
// const path = require('path');

// (async () => {
//     try {
//         // Test data voor invoice
//         const invoiceData = {
//             recipient: 'John Doe',
//             date: '2025-01-16',
//             items: [
//                 { name: 'Membership Fee', amount: '50.00' },
//                 { name: 'Camp Payment', amount: '100.00' },
//             ],
//             total: '150.00',
//         };

//         // Genereer HTML invoice
//         const htmlContent = InvoiceGenerator.generateInvoiceHTML(invoiceData);

//         // Save PDF file
//         const pdfFilePath = path.join(__dirname, 'data/output/invoice.pdf');
//         await InvoiceGenerator.saveInvoicePDF(htmlContent, pdfFilePath);
//         console.log('PDF Invoice saved at:', pdfFilePath);

//         // Stuur email met invoice
//         await EmailHandler.sendEmail({
//             to: 'jelle.swartebroekx@vdabcampus.be', // Vervang met email ontvanger(s)
//             subject: 'Your Invoice',
//             body: `<p>Dear ${invoiceData.recipient},</p><p>Please find your invoice attached.</p>`,
//             attachments: [
//                 {
//                     filename: 'invoice.pdf',
//                     path: pdfFilePath, // PDF toevoegen aan mail
//                 },
//             ],
//         });
//     } catch (error) {
//         console.error('Error:', error.message);
//     }
// })();

// TEST 5 (VOOR CLI.JS)
// const CLI = await import('./src/controllers/CLI.js');

// (async () => {
//     await CLI.default.runWorkflow();
// })();


// TEST 6 (VOOR AppController.js) => FINAL TEST SCRIPT
import AppController from './src/controllers/AppController.js';

// Entry point voor app
(async () => {
    try {
        console.log('De Hinde Accounting - Systeem Starten...');
        await AppController.run();
    } catch (error) {
        console.error('Onverwachte fout bij het uitvoeren van de applicatie: ', error.message);
    } finally {
        console.log('Accounting - Systeem handeling uitgevoerd.');
    }
})();
