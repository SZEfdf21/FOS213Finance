import inquirer from 'inquirer';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';
import TransactionProcessor from '../services/TransactionProcessor.js';
import ExcelHandler from '../services/ExcelHandler.js';
import InvoiceGenerator from '../services/InvoiceGenerator.js';

// Recreate __dirname for ES modules
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

class CLI {
    static async showMenu() {
        const { choice } = await inquirer.prompt([
            {
                type: 'list',
                name: 'choice',
                message: 'Kies een handeling',
                choices: [
                    '1. Verwerk Transacties(Excel File)',
                    '2. Maak een Factuur',
                    '3. Sluit',
                ],
            },
        ]);

        return choice;
    }

    static async promptUserForFile(type) {
        const { filePath } = await inquirer.prompt([
            {
                type: 'input',
                name: 'filePath',
                message: `Geef het pad naar het ${type} bestand in:`,
                validate: (input) => (fs.existsSync(input) ? true : 'Bestand niet gevonden, kies een geldig pad.'),
            },
        ]);

        return filePath;
    }

    static async runWorkflow() {
        while (true) {
            const choice = await this.showMenu();

            if (choice === '3. Sluit') {
                console.log('Tot ziens!');
                break;
            }

            switch (choice) {
                case '1. Verwerk Transacties(Excel File)': {
                    const excelFile = await this.promptUserForFile('Excel (Transacties)');
                    const sheetName = 'Transactions';
                    const transactions = await ExcelHandler.readExcel(excelFile, sheetName);
                    console.log('Transacties:', transactions);

                    const codesFile = await this.promptUserForFile('Excel (Codes)');
                    const codes = await ExcelHandler.readExcel(codesFile, 'Codes');
                    console.log('Codes:', codes);

                    const matchedTransactions = TransactionProcessor.matchTransactions(transactions, codes);
                    console.log('Overeenkomende transacties:', matchedTransactions);
                    break;
                }

                case '2. Maak een Factuur': {
                    const recipient = await inquirer.prompt([
                        {
                            type: 'input',
                            name: 'recipient',
                            message: 'Geef de naam van de ontvanger op:',
                        },
                    ]);

                    const invoiceData = {
                        recipient: recipient.recipient,
                        date: new Date().toLocaleDateString(),
                        items: [
                            { name: 'Membership Fee', amount: '50.00' },
                            { name: 'Camp Payment', amount: '100.00' },
                        ],
                        total: '150.00',
                    };

                    const htmlContent = InvoiceGenerator.generateInvoiceHTML(invoiceData);

                    const outputDir = path.resolve(__dirname, '../../data/output');
                    if (!fs.existsSync(outputDir)) {
                        fs.mkdirSync(outputDir, { recursive: true });
                    }
                    const pdfPath = path.resolve(outputDir, 'invoice.pdf');

                    console.log('Output Directory:', outputDir);
                    console.log('PDF Path:', pdfPath);

                    await InvoiceGenerator.saveInvoicePDF(htmlContent, pdfPath);
                    console.log('Factuur opgeslagen in:', pdfPath);
                    break;
                }
            }
        }
    }
}

export default CLI;
