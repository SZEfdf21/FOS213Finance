import fs from 'fs';
import { PDFDocument } from 'pdf-lib';
import { fileURLToPath } from 'url';
import { dirname, join, resolve, isAbsolute } from 'path';

// dirname zelf maken (comptibility issue, future fix)
const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

class InvoiceGenerator {
    // Factuur in HTML format
    static generateInvoiceHTML(data) {
        const { recipient, items, total, date } = data;

        return `
            <html>
            <head>
                <style>
                    body { font-family: Arial, sans-serif; }
                    .invoice { max-width: 600px; margin: auto; padding: 20px; border: 1px solid #ddd; }
                    .header { font-size: 24px; font-weight: bold; }
                    .items { margin-top: 20px; }
                    .item { display: flex; justify-content: space-between; margin-bottom: 10px; }
                    .total { font-size: 20px; font-weight: bold; margin-top: 20px; }
                </style>
            </head>
            <body>
                <div class="invoice">
                    <div class="header">Invoice</div>
                    <div>Date: ${date}</div>
                    <div>Recipient: ${recipient}</div>
                    <div class="items">
                        ${items.map(item => `<div class="item"><span>${item.name}</span><span>${item.amount}</span></div>`).join('')}
                    </div>
                    <div class="total">Total: ${total}</div>
                </div>
            </body>
            </html>
        `;
    }

    // Save invoice in PDF format
    static async saveInvoicePDF(htmlContent, fileName = 'invoice.pdf') {
        // Resolve the output directory
        const outputDir = resolve(__dirname, '../../data/output');
        
        // Ensure the directory exists
        if (!fs.existsSync(outputDir)) {
            fs.mkdirSync(outputDir, { recursive: true });
        }

        // Ensure filePath is valid
        const filePath = isAbsolute(fileName)
            ? fileName
            : join(outputDir, fileName);

        // Debugging logs
        console.log('Output Directory:', outputDir);
        console.log('Full File Path:', filePath);

        // Create PDF
        const pdfDoc = await PDFDocument.create();
        const page = pdfDoc.addPage([600, 800]);
        page.drawText('PDF Generation Placeholder'); // Placeholder text

        // Write PDF to file
        const pdfBytes = await pdfDoc.save();
        fs.writeFileSync(filePath, pdfBytes);
        console.log(`Invoice saved at: ${filePath}`);
    }
}

export default InvoiceGenerator;