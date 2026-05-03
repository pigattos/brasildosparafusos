const http = require('http');
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const PORT = 3000;
const SOURCE_DIR = "C:\\Users\\Cassyano\\OneDrive - Brasil do Parafusos\\Comprasbrasil - Compras\\Base entrada de itens";
const OUTPUT_FILE = "data-entradas.js";

/**
 * Normaliza nomes de colunas para busca flexível
 */
function normalizeKey(key) {
    if (!key) return "";
    return key.toString().toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
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
                const rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
                if (rows.length < 2) continue;

                // Encontrar a linha de cabeçalho (primeira que tenha 'data' ou 'movto')
                let headerIndex = -1;
                for (let i = 0; i < Math.min(rows.length, 15); i++) {
                    const rowStr = JSON.stringify(rows[i]).toLowerCase();
                    if (rowStr.includes('data') || rowStr.includes('movto') || rowStr.includes('cod') || rowStr.includes('nf')) {
                        headerIndex = i;
                        break;
                    }
                }

                if (headerIndex === -1) headerIndex = 0;

                const normalize = (s) => String(s || '').toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
                const headers = rows[headerIndex].map(h => normalize(h));
                const findIdx = (names) => headers.findIndex(h => names.some(n => h.includes(normalize(n))));

                const dIdx = findIdx(['dt.movto', 'data', 'movimento']);
                const cIdx = findIdx(['codigo', 'cod']);
                const desIdx = findIdx(['descricao', 'item', 'produto']);
                const qIdx = findIdx(['quantidade', 'qtd', 'unidades']);
                const nfIdx = findIdx(['nota fiscal', 'nf', 'doc']);
                const fIdx = findIdx(['finalid.item ordem com', 'finalidade']);
                const vIdx = findIdx(['vlr.cont.p/sped', 'vlr.cont', 'valor', 'total']);
                const sIdx = findIdx(['razao social', 'fornecedor', 'nome']);
                const gIdx = 11; // Coluna L (12ª coluna)

                const processed = rows.slice(headerIndex + 1).map(row => {
                    if (!row || row.length === 0) return null;
                    if (!row[vIdx] && !row[qIdx] && !row[cIdx]) return null;

                    let dateVal = row[dIdx];
                    let dateStr = "";

                    if (dateVal instanceof Date) {
                        dateStr = dateVal.toISOString().split('T')[0];
                    } else if (typeof dateVal === 'number') {
                        const d = new Date(Math.round((dateVal - 25569) * 86400 * 1000));
                        dateStr = d.toISOString().split('T')[0];
                    } else if (typeof dateVal === 'string') {
                        const parts = dateVal.split(/[\/\-]/);
                        if (parts.length === 3) {
                            if (parts[0].length === 4) dateStr = dateVal;
                            else dateStr = `${parts[2]}-${parts[1].padStart(2,'0')}-${parts[0].padStart(2,'0')}`;
                        }
                    }

                    return {
                        date: dateStr,
                        code: String(row[cIdx] || '').trim(),
                        description: String(row[desIdx] || '').trim(),
                        quantity: parseFloat(row[qIdx]) || 0,
                        group: String(row[gIdx] || 'DIVERSOS').trim(),
                        invoice: String(row[nfIdx] || '').trim(),
                        purpose: String(row[fIdx] || '').trim(),
                        supplier: String(row[sIdx] || 'NÃO IDENTIFICADO').trim(),
                        value: parseFloat(row[vIdx]) || 0
                    };
                }).filter(row => row && (row.quantity !== 0 || row.value !== 0));

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
