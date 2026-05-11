const XLSX = require('xlsx');
const LEADTIME_FILE = "C:\\Users\\Cassyano\\OneDrive - Brasil do Parafusos\\Compras\\Relatório Lead time\\Base Leadtime.xlsx";

try {
    const workbook = XLSX.readFile(LEADTIME_FILE);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    
    console.log("Headers found:");
    console.log(JSON.stringify(rows[0], null, 2));
    
    console.log("\nFirst row of data:");
    console.log(JSON.stringify(rows[1], null, 2));
    
    // Check columns E, F, G (indices 4, 5, 6)
    console.log("\nColumn E (index 4):", rows[0][4]);
    console.log("Column F (index 5):", rows[0][5]);
    console.log("Column G (index 6):", rows[0][6]);
    
} catch (e) {
    console.error("Error:", e.message);
}
