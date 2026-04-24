const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

const SOURCE_DIR = "C:\\Users\\Cassyano\\OneDrive - Brasil do Parafusos\\Comprasbrasil - Compras\\Entradas Mensal";
const OUTPUT_FILE = "data-entradas.js";
const TEMP_JSON = "temp_all_data.json";

function excelSerialToDate(serial) {
    const epoch = new Date(1899, 11, 30);
    const date = new Date(epoch.getTime() + serial * 86400 * 1000);
    return date.toISOString().split('T')[0];
}

async function runSync() {
    console.log("Iniciando Sincronização Global...");
    
    if (!fs.existsSync(SOURCE_DIR)) {
        console.error("Erro: Pasta não encontrada!");
        return;
    }

    const files = fs.readdirSync(SOURCE_DIR).filter(f => f.endsWith('.xlsx') || f.endsWith('.xls'));
    console.log(`Encontrados ${files.length} arquivos.`);

    let masterData = [];

    for (const file of files) {
        const fullPath = path.join(SOURCE_DIR, file);
        console.log(`Processando: ${file}...`);
        
        try {
            // Extrair JSON do Excel via npx
            const buffer = execSync(`npx -y xlsx-cli "${fullPath}" --json`, { encoding: 'utf8', maxBuffer: 10 * 1024 * 1024 });
            
            // Limpar saída (remover avisos de depreciação e BOM)
            const cleanJson = buffer.split('\n').filter(line => line.trim().startsWith('[') || line.trim().startsWith('{')).join('\n').replace(/^\uFEFF/, '');
            
            const rawRows = JSON.parse(cleanJson);
            
            // Processar linhas
            const processed = rawRows.filter(row => {
                const keys = Object.keys(row);
                const dateKey = keys.find(k => k.toLowerCase().includes('dt.movto') || k.toLowerCase().includes('data'));
                const supplierKey = keys.find(k => k.toLowerCase().includes('raz') || k.toLowerCase().includes('social') || k.toLowerCase().includes('fornecedor'));
                const valueKey = keys.find(k => k.toLowerCase().includes('vlr.cont') || k.toLowerCase().includes('valor'));
                return row[dateKey] && row[supplierKey] && row[valueKey] > 0;
            }).map(row => {
                const keys = Object.keys(row);
                const dateKey = keys.find(k => k.toLowerCase().includes('dt.movto') || k.toLowerCase().includes('data'));
                const supplierKey = keys.find(k => k.toLowerCase().includes('raz') || k.toLowerCase().includes('social') || k.toLowerCase().includes('fornecedor'));
                const valueKey = keys.find(k => k.toLowerCase().includes('vlr.cont') || k.toLowerCase().includes('valor'));

                return {
                    date: typeof row[dateKey] === 'number' ? excelSerialToDate(row[dateKey]) : row[dateKey],
                    supplier: row[supplierKey],
                    value: row[valueKey]
                };
            });
            
            masterData = masterData.concat(processed);
            console.log(`   + ${processed.length} registros adicionados.`);
            
        } catch (e) {
            console.error(`Erro ao processar ${file}:`, e.message);
        }
    }

    // Gerar o arquivo final
    const jsContent = `const PRE_LOADED_ENTRADAS = ${JSON.stringify(masterData, null, 4)};`;
    fs.writeFileSync(OUTPUT_FILE, jsContent);
    
    console.log("==========================================");
    console.log(`Sucesso! Total de ${masterData.length} registros salvos.`);
    console.log("O Dashboard será atualizado ao recarregar a página.");
}

runSync();
