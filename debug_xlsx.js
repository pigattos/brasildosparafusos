const XLSX = require('xlsx');
const path = require('path');

const filePath = 'C:\\Users\\Cassyano\\OneDrive - Brasil do Parafusos\\Comprasbrasil - Compras\\Base entrada de itens\\04.2026.xlsx';

try {
    const workbook = XLSX.readFile(filePath);
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

    console.log("--- DEBUG ESTRUTURA PLANILHA ---");
    console.log(`Total de linhas: ${rows.length}`);
    
    // Mostrar as primeiras 10 linhas para ver onde está o cabeçalho
    for (let i = 0; i < Math.min(rows.length, 10); i++) {
        console.log(`Linha ${i}:`, JSON.stringify(rows[i]));
    }

    // Tentar identificar o cabeçalho como no bridge.js
    let headerIndex = -1;
    for (let i = 0; i < Math.min(rows.length, 15); i++) {
        const rowStr = JSON.stringify(rows[i]).toLowerCase();
        if (rowStr.includes('data') || rowStr.includes('movto') || rowStr.includes('cod') || rowStr.includes('nf')) {
            headerIndex = i;
            break;
        }
    }
    console.log(`\nHeader Index detectado: ${headerIndex}`);
    
    if (headerIndex !== -1) {
        const headers = rows[headerIndex];
        console.log("Cabeçalhos detectados:", JSON.stringify(headers));
        console.log("Coluna L (index 11):", headers[11]);
    }

    // Amostra de dados da primeira linha após o cabeçalho
    if (headerIndex !== -1 && rows.length > headerIndex + 1) {
        console.log("\nAmostra de Dados (Linha " + (headerIndex + 1) + "):", JSON.stringify(rows[headerIndex + 1]));
        console.log("Valor na Coluna L (index 11):", rows[headerIndex + 1][11]);
    }

} catch (err) {
    console.error("Erro ao ler arquivo:", err.message);
}
