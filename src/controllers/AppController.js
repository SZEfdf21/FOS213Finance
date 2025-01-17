import CLI from './CLI.js';
import ExcelHandler from '../services/ExcelHandler.js';
import TransactionProcessor from '../services/TransactionProcessor.js';
import InvoiceGenerator from '../services/InvoiceGenerator.js';

class AppController {
    static async run() {
        try {
            console.log("Welkom bij het 'De Hinde' - Accounting systeem");
            await CLI.runWorkflow();
        } catch (error) {
            console.error("Fout tijdens uitvoeren: ", error.message);
        } finally {
            console.log("Bedankt voor het systeem te gebruiken. Feedback & bugs of eventuele features mag je doorgeven aan Patou.");
        }
    }
}

export default AppController;