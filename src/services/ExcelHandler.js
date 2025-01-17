import ExcelJS from 'exceljs';

class ExcelHandler {
    // Leest data van de Excel sheet
    static async readExcel(filePath, sheetName) {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(filePath);
        const sheet = workbook.getWorksheet(sheetName);
        if (!sheet) {
            throw new Error(`Sheet ${sheetName} not found in ${filePath}`);
        }

        const data = [];
        sheet.eachRow((row) => {
            data.push(row.values);
        });

        return data;
    }

    // Update specifieke cellen in Excel sheet
    static async updateSheet(filePath, sheetName, updates) {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(filePath);
        const sheet = workbook.getWorksheet(sheetName);
        if (!sheet) {
            throw new Error(`Sheet ${sheetName} not found in ${filePath}`);
        }

        updates.forEach(({ row, column, value }) => {
            sheet.getRow(row).getCell(column).value = value;
        });

        await workbook.xlsx.writeFile(filePath);
    }

    // Schrijft data naar Excel sheet
    static async writeExcel(filePath, data, sheetName = "Sheet1") {
        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet(sheetName);

        data.forEach((rowData) => {
            sheet.addRow(rowData);
        });

        await workbook.xlsx.writeFile(filePath);
    }
}

export default ExcelHandler;