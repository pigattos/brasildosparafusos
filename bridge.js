const http = require('http');
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const PORT = 3000;
const SOURCE_DIR = "C:\\Users\\Cassyano\\OneDrive - Brasil do Parafusos\\Comprasbrasil - Compras\\Entradas Mensal";
const OUTPUT_FILE = "data-entradas.js";

/**
 * Normaliza nomes de colunas para busca flexível
 */
function normalizeKey(key) {
    return key.toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function findColumn(row, possibilities) {
    const keys = Object.keys(row);
    for (const p of possibilities) {
        const normP = normalizeKey(p);
        const match = keys.find(k => normalizeKey(k) === normP || normalizeKey(k).includes(normP));
        if (match) return match;
    }
    return null;
}

const server = http.createServer((req, res) => {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

    if (req.method === 'OPTIONS') {
        res.writeHead(204);
        res.end();
        return;
    }
    
    if (req.url === '/sync') {
        console.log("Recebido pedido de sincronização...");
        try {
            if (!fs.existsSync(SOURCE_DIR)) {
                throw new Error(`Diretório não encontrado: ${SOURCE_DIR}`);
            }

            const files = fs.readdirSync(SOURCE_DIR).filter(f => f.endsWith('.xlsx') || f.endsWith('.xls'));
            console.log(`Arquivos encontrados: ${files.length}`);
            
            let masterData = [];

            for (const file of files) {
                const fullPath = path.join(SOURCE_DIR, file);
                console.log(`Processando: ${file}`);
                
                const workbook = XLSX.readFile(fullPath, { cellDates: true });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const rawRows = XLSX.utils.sheet_to_json(firstSheet);
                
                if (rawRows.length === 0) continue;

                // Identificar colunas na primeira linha com dados
                const dKey = findColumn(rawRows[0], ['dt.movto', 'data', 'movimento']);
                const sKey = findColumn(rawRows[0], ['razao social', 'fornecedor', 'nome']);
                const vKey = findColumn(rawRows[0], ['vlr.cont', 'valor', 'total']);

                if (!dKey || !sKey || !vKey) {
                    console.warn(`Colunas não encontradas no arquivo ${file}. Pulando...`);
                    continue;
                }

                const processed = rawRows.filter(row => row[vKey] > 0).map(row => {
                    let dateVal = row[dKey];
                    let dateStr = dateVal;

                    // Normalizar para objeto Date e depois para YYYY-MM-DD
                    if (dateVal instanceof Date) {
                        dateStr = dateVal.toISOString().split('T')[0];
                    } else if (typeof dateVal === 'string') {
                        const parts = dateVal.split(/[\/\-]/);
                        if (parts.length === 3) {
                            if (parts[0].length === 4) {
                                dateStr = dateVal; // Já é YYYY-MM-DD
                            } else {
                                // Converter DD/MM/YYYY para YYYY-MM-DD
                                dateStr = `${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`;
                            }
                        }
                    } else if (typeof dateVal === 'number') {
                        // Excel serial date
                        const d = new Date(Math.round((dateVal - 25569) * 86400 * 1000));
                        dateStr = d.toISOString().split('T')[0];
                    }
                    
                    return {
                        date: dateStr,
                        supplier: (row[sKey] || 'NÃO IDENTIFICADO').toString().trim(),
                        value: parseFloat(row[vKey]) || 0
                    };
                });
                masterData = masterData.concat(processed);
            }

            const jsContent = `const PRE_LOADED_ENTRADAS = ${JSON.stringify(masterData, null, 4)};`;
            fs.writeFileSync(OUTPUT_FILE, jsContent);
            
            console.log(`Sincronização concluída! ${masterData.length} registros totais.`);
            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ 
                success: true, 
                count: masterData.length,
                data: masterData 
            }));
        } catch (e) {
            console.error("ERRO NA SINCRONIZAÇÃO:", e.message);
            res.writeHead(500, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ success: false, error: e.message }));
        }
    } else {
        res.writeHead(404);
        res.end();
    }
});

server.listen(PORT, () => {
    console.log(`\n🚀 PONTE DE DADOS ATIVA`);
    console.log(`----------------------------------`);
    console.log(`URL: http://localhost:${PORT}/sync`);
    console.log(`Pasta: ${SOURCE_DIR}`);
    console.log(`Destino: ${path.resolve(OUTPUT_FILE)}`);
    console.log(`----------------------------------\n`);
});
