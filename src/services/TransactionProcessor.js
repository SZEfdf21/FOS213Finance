import ExcelHandler from './ExcelHandler.js';

class TransactionProcessor {
    // Matcht transacties met voorafbepaalde codes
    static matchTransactions(transactions, codes) {
        return transactions.map((transaction) => {
            const matchedCode = codes.find((code) => {
                return (
                    transaction.description &&
                    code.keyword &&
                    transaction.description.includes(code.keyword)
                );
            });
    
            return {
                ...transaction,
                code: matchedCode ? matchedCode.code : 'UNKNOWN',
            };
        });
    }
    

    // Update de schulden op basis van de transacties
    static async updateDebts(transactionData, debtFilePath, sheetName) {
        const debts = await ExcelHandler.readExcel(debtFilePath, sheetName);

        transactionData.forEach((transaction) => {
            const debtor = debts.find((debt) => debt.Name === transaction.name);
            if (debtor) {
                debtor.AmountOwed = (debtor.AmountOwed || 0) - transaction.amount;
            }
        });

        await ExcelHandler.writeExcel(debtFilePath, debts, sheetName);
    }

    // Genereer een samenvatting van de transacties
    static generateSummary(transactions) {
        return transactions.reduce((summary, transaction) => {
            const { code } = transaction;
            if (!summary[code]) {
                summary[code] = { count: 0, total: 0 };
            }
            summary[code].count += 1;
            summary[code].total += transaction.amount;
            return summary;
        }, {});
    }
}

export default TransactionProcessor;