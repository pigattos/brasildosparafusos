/**
 * Inteligência Evolutiva de Abastecimento v4.0
 * Baseado no novo Spec de Análise Evolutiva
 * 
 * Regra: Primeiro arquivo = BASE HISTÓRICA FIXA
 */
document.addEventListener('DOMContentLoaded', () => {
    console.log("Inteligência Evolutiva v4.0 Inicializada");

    // --- Configurações de Criticidade (Novo Spec) ---
    const CRITICIDADE = {
        'rupture': { label: 'Ruptura', weight: 3, color: '#fb7185' },
        'attention': { label: 'Atenção', weight: 2, color: '#f59e0b' },
        'suggest': { label: 'Sugestão', weight: 1, color: '#6366f1' },
        'ok': { label: 'Seguro', weight: 0, color: '#34d399' },
        'ignored': { label: 'Ignorado', weight: 0, color: '#9ca3af' }
    };

    const RECORRENCIA_MINIMA = 0.17; // > 17%

    // --- Estado Global ---
    let snapshotHistory = []; // { name, date, itemsMap, displayItems }
    let baseSnapshot = null;
    let currentSnapshot = null;
    let currentTimelineIdx = 0;
    let activeBuyer = 'all';
    let evolutionChart = null;
    let buyerMap = JSON.parse(localStorage.getItem('buyerMap') || '{}');

    // --- Elementos DOM ---
    const folderInputs = [document.getElementById('folder-upload'), document.getElementById('folder-upload-welcome')];
    const mainContent = document.getElementById('main-content');
    const welcomeState = document.getElementById('welcome-state');
    const loadingOverlay = document.getElementById('loading-overlay');
    const timelineRange = document.getElementById('timeline-range');
    const timelineTicks = document.getElementById('timeline-ticks');
    const tableBody = document.getElementById('evolution-table-body');
    const tableSearch = document.getElementById('table-search');
    const evolutionFilter = document.getElementById('evolution-filter');

    // --- Utilitários ---
    const formatCurrency = (v) => v.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
    
    function parseNumeric(val) {
        if (val === undefined || val === null || val === '') return 0;
        if (typeof val === 'number') return val;
        let str = val.toString().replace('R$', '').replace(/\s/g, '').trim();
        if (str.startsWith('.') || str.startsWith(',')) str = '0' + str;
        const hasComma = str.includes(',');
        const hasDot = str.includes('.');

        if (hasComma && hasDot) {
            if (str.lastIndexOf(',') > str.lastIndexOf('.')) {
                str = str.replace(/\./g, '').replace(',', '.');
            } else {
                str = str.replace(/,/g, '');
            }
        } else if (hasComma) {
            const parts = str.split(',');
            if (parts.length > 2 || parts[1].length === 3) str = str.replace(/,/g, '');
            else str = str.replace(',', '.');
        }
        const num = parseFloat(str);
        return isNaN(num) ? 0 : num;
    }

    function findColumn(headers, aliases) {
        const cleanHeaders = headers.map(h => String(h || '').toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, ""));
        const cleanAliases = aliases.map(a => a.toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, ""));
        for (let alias of cleanAliases) {
            const idx = cleanHeaders.indexOf(alias);
            if (idx !== -1) return headers[idx];
        }
        for (let alias of cleanAliases) {
            const idx = cleanHeaders.findIndex(h => h.includes(alias));
            if (idx !== -1) return headers[idx];
        }
        return null;
    }

    // --- Processamento de Arquivos ---
    async function processExcelFile(file) {
        return new Promise((resolve) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array', cellDates: true });
                    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                    const rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                    
                    // Achar cabeçalho
                    let headerIndex = -1;
                    for (let i = 0; i < Math.min(rawRows.length, 20); i++) {
                        if (rawRows[i].some(c => String(c||'').toLowerCase().includes('estoque') || String(c||'').toLowerCase().includes('produto'))) {
                            headerIndex = i;
                            break;
                        }
                    }
                    if (headerIndex === -1) return resolve(null);

                    const headers = rawRows[headerIndex];
                    const rows = XLSX.utils.sheet_to_json(worksheet, { range: headerIndex });
                    
                    const colMap = {
                        code: findColumn(headers, ['Código', 'Cód.', 'Produto', 'Cod']),
                        desc: findColumn(headers, ['Descrição', 'Desc', 'Produto']),
                        estoque: findColumn(headers, ['Estoque Atual', 'Saldo Disponível', 'Estoque', 'Est.']),
                        encomendas: findColumn(headers, ['Saldo O.C.', 'A Receber', 'Encomendas', 'O.C.']),
                        medVenda: findColumn(headers, ['Méd.Venda', 'Média Venda', 'Venda Média', 'Saída Média']),
                        fornecedor: findColumn(headers, ['Fornecedor', 'Principal Fornecedor', 'Fornec.']),
                        comprador: findColumn(headers, ['Comprador', 'Responsável', 'Buyer']),
                        totalSales: findColumn(headers, ['Vendas Total', 'Total Vendas', 'Saída Total'])
                    };

                    const monthCols = headers.filter(h => h && /^\d{1,2}\/\d{2,4}$/.test(String(h)));

                    const itemsMap = new Map();
                    rows.forEach(row => {
                        const code = String(row[colMap.code] || '').trim();
                        if (!code) return;

                        const estoque = parseNumeric(row[colMap.estoque]);
                        const encomendas = parseNumeric(row[colMap.encomendas]);
                        const salesTotal = parseNumeric(row[colMap.totalSales]);
                        let medVenda = parseNumeric(row[colMap.medVenda]);
                        let recorrencia = 0;

                        if (monthCols.length > 0) {
                            let activeMonths = 0;
                            let sum = 0;
                            monthCols.forEach(mCol => {
                                const v = parseNumeric(row[mCol]);
                                if (v > 0) { activeMonths++; sum += v; }
                            });
                            recorrencia = activeMonths / monthCols.length;
                            if (activeMonths > 0) medVenda = sum / activeMonths;
                        }

                        let status = 'ok';
                        const disponivel = estoque + encomendas;
                        if (recorrencia > RECORRENCIA_MINIMA) {
                            if (medVenda > disponivel) status = 'rupture';
                            else if ((medVenda * 2) > disponivel) status = 'attention';
                            else if ((medVenda * 3) > disponivel) status = 'suggest';
                        } else if (salesTotal === 0) {
                            status = 'ignored';
                        }

                        let comprador = row[colMap.comprador];
                        if (!comprador) {
                            const cleanCode = code.replace(/^0+/, '').replace(/[.]/g, '');
                            comprador = buyerMap[code] || buyerMap[cleanCode] || 'N/D';
                        }

                        itemsMap.set(code, {
                            code,
                            desc: row[colMap.desc] || 'S/D',
                            fornecedor: row[colMap.fornecedor] || 'N/D',
                            comprador,
                            status,
                            weight: CRITICIDADE[status].weight,
                            value: medVenda * (row['Custo Unitário'] || row['Custo'] || 1) // simplificado
                        });
                    });

                    // Tentar extrair data do nome do arquivo (07.05.2026.xlsx)
                    let fileDate = file.name.match(/(\d{2})[.\/](\d{2})[.\/](\d{4})/);
                    let dateStr = fileDate ? `${fileDate[3]}-${fileDate[2]}-${fileDate[1]}` : new Date().toISOString().split('T')[0];

                    resolve({ name: file.name, date: dateStr, itemsMap });
                } catch (err) {
                    console.error(err);
                    resolve(null);
                }
            };
            reader.readAsArrayBuffer(file);
        });
    }

    // --- Orquestração de Dados ---
    async function handleFiles(files) {
        loadingOverlay.style.display = 'flex';
        const filteredFiles = Array.from(files).filter(f => f.name.match(/\.(xlsx|xls)$/i) && !f.name.startsWith('~$'));
        
        const snaps = [];
        for (let f of filteredFiles) {
            const res = await processExcelFile(f);
            if (res) snaps.push(res);
        }

        if (snaps.length === 0) {
            loadingOverlay.style.display = 'none';
            alert("Nenhum arquivo válido encontrado.");
            return;
        }

        // 1. Ordenar por Data
        snaps.sort((a, b) => new Date(a.date) - new Date(b.date));
        snapshotHistory = snaps;
        
        // Regra de Ouro: O arquivo "07.05.2026" é a nossa BASE HISTÓRICA FIXA
        const explicitBase = snaps.find(s => s.name.includes('07.05.2026'));
        baseSnapshot = explicitBase || snaps[0];
        
        currentTimelineIdx = snaps.length - 1;

        updateDashboard();
        
        welcomeState.style.display = 'none';
        mainContent.style.display = 'block';
        loadingOverlay.style.display = 'none';
    }

    function updateDashboard() {
        const snap = snapshotHistory[currentTimelineIdx];
        if (!snap) return;

        currentSnapshot = snap;
        document.getElementById('base-date-display').textContent = baseSnapshot.date.split('-').reverse().join('/');
        document.getElementById('current-snapshot-date').textContent = snap.date.split('-').reverse().join('/');

        renderTimeline();
        calculateEvolution();
        renderCharts();
        renderTable();
        updatePerformance();
        updateBuyerFilter();
    }

    function renderTimeline() {
        timelineRange.max = snapshotHistory.length - 1;
        timelineRange.value = currentTimelineIdx;
        
        timelineTicks.innerHTML = snapshotHistory.map((s, idx) => `
            <span class="${idx === currentTimelineIdx ? 'active' : ''}" onclick="window.jumpTo(${idx})">
                ${s.date.split('-').slice(1).reverse().join('/')}
            </span>
        `).join('');
    }

    window.jumpTo = (idx) => {
        currentTimelineIdx = idx;
        updateDashboard();
    };

    function calculateEvolution() {
        const baseItems = Array.from(baseSnapshot.itemsMap.values());
        const currentItems = currentSnapshot.itemsMap;

        // Filtrar por Comprador se necessário
        const targetBaseItems = baseItems.filter(i => activeBuyer === 'all' || i.comprador === activeBuyer);
        
        let recovered = { rupture: 0, attention: 0, suggest: 0, total: 0 };
        let worsened = 0;
        let stable = 0;

        targetBaseItems.forEach(itemBase => {
            const itemCurrent = currentItems.get(itemBase.code);
            const currentWeight = itemCurrent ? itemCurrent.weight : 0; // se não existe no atual, assumimos que resolveu (seguro)
            
            const evolution = itemBase.weight - currentWeight;

            if (evolution > 0) {
                recovered.total++;
                if (itemBase.status === 'rupture') recovered.rupture++;
                else if (itemBase.status === 'attention') recovered.attention++;
                else if (itemBase.status === 'suggest') recovered.suggest++;
            } else if (evolution < 0) {
                worsened++;
            } else if (itemBase.weight > 0) {
                stable++;
            }
        });

        // KPI's
        document.getElementById('global-recovered-count').textContent = recovered.total;
        document.getElementById('rupture-recovered').textContent = recovered.rupture;
        document.getElementById('attention-recovered').textContent = recovered.attention;
        document.getElementById('worsened-count').textContent = worsened;

        const baseTotalRisk = targetBaseItems.filter(i => i.weight > 0).length;
        const efficiency = baseTotalRisk > 0 ? (recovered.total / baseTotalRisk * 100) : 0;
        document.getElementById('efficiency-percent').textContent = `${efficiency.toFixed(1)}%`;
        
        // Percentuais de cada nível
        const baseRuptures = targetBaseItems.filter(i => i.status === 'rupture').length;
        const baseAttentions = targetBaseItems.filter(i => i.status === 'attention').length;
        
        document.getElementById('rupture-recovered-percent').textContent = baseRuptures > 0 ? `${(recovered.rupture / baseRuptures * 100).toFixed(1)}% de sucesso` : 'N/A';
        document.getElementById('attention-recovered-percent').textContent = baseAttentions > 0 ? `${(recovered.attention / baseAttentions * 100).toFixed(1)}% de sucesso` : 'N/A';
    }

    function renderCharts() {
        const ctx = document.getElementById('evolution-chart').getContext('2d');
        if (evolutionChart) evolutionChart.destroy();

        const historyLabels = snapshotHistory.map(s => s.date.split('-').reverse().slice(0,2).join('/'));
        
        // Calcular Peso Total de Risco por snapshot para o gráfico
        const weightHistory = snapshotHistory.map(snap => {
            const items = Array.from(snap.itemsMap.values()).filter(i => activeBuyer === 'all' || i.comprador === activeBuyer);
            return items.reduce((acc, i) => acc + i.weight, 0);
        });

        evolutionChart = new Chart(ctx, {
            type: 'line',
            data: {
                labels: historyLabels,
                datasets: [{
                    label: 'Índice de Criticidade (Peso Total)',
                    data: weightHistory,
                    borderColor: '#6366f1',
                    backgroundColor: 'rgba(99, 102, 241, 0.1)',
                    fill: true,
                    tension: 0.4,
                    pointRadius: 6,
                    pointBackgroundColor: '#6366f1'
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { legend: { display: false } },
                scales: {
                    y: { 
                        beginAtZero: true, 
                        grid: { color: 'rgba(255,255,255,0.05)' },
                        ticks: { color: '#9ca3af' }
                    },
                    x: { 
                        grid: { display: false },
                        ticks: { color: '#9ca3af' }
                    }
                }
            }
        });
    }

    function renderTable() {
        const searchTerm = tableSearch.value.toLowerCase();
        const filterVal = evolutionFilter.value;
        
        const baseItems = Array.from(baseSnapshot.itemsMap.values());
        const currentItems = currentSnapshot.itemsMap;

        const rows = baseItems
            .filter(i => activeBuyer === 'all' || i.comprador === activeBuyer)
            .map(itemBase => {
                const itemCurrent = currentItems.get(itemBase.code) || { status: 'ok', weight: 0 };
                const evolution = itemBase.weight - itemCurrent.weight;
                return { itemBase, itemCurrent, evolution };
            })
            .filter(row => {
                const matchSearch = row.itemBase.code.toLowerCase().includes(searchTerm) || row.itemBase.desc.toLowerCase().includes(searchTerm);
                if (!matchSearch) return false;
                
                if (filterVal === 'improved') return row.evolution > 0;
                if (filterVal === 'worsened') return row.evolution < 0;
                if (filterVal === 'stable') return row.evolution === 0 && row.itemBase.weight > 0;
                return true;
            });

        tableBody.innerHTML = rows.slice(0, 100).map(row => `
            <tr>
                <td style="font-weight: 600;">${row.itemBase.code}</td>
                <td style="font-size: 0.85rem; opacity: 0.9;">${row.itemBase.desc.substring(0, 40)}...</td>
                <td style="font-size: 0.75rem; color: var(--text-muted);">${row.itemBase.fornecedor.substring(0, 20)}</td>
                <td style="text-align: center;"><span class="badge badge-${row.itemBase.status}">${row.itemBase.status}</span></td>
                <td style="text-align: center;"><span class="badge badge-${row.itemCurrent.status}">${row.itemCurrent.status}</span></td>
                <td style="text-align: center; font-weight: 800; color: ${row.evolution > 0 ? '#34d399' : (row.evolution < 0 ? '#fb7185' : '#9ca3af')}">
                    ${row.evolution > 0 ? '↑' : (row.evolution < 0 ? '↓' : '=')} ${Math.abs(row.evolution)}
                </td>
                <td>
                    <span class="evolution-status-tag ${row.evolution > 0 ? 'ev-good' : (row.evolution < 0 ? 'ev-bad' : 'ev-stable')}">
                        ${row.evolution > 0 ? 'Melhorou' : (row.evolution < 0 ? 'Piorou' : 'Estável')}
                    </span>
                </td>
            </tr>
        `).join('');
    }

    function updatePerformance() {
        const perfList = document.getElementById('performance-list');
        const currentItems = Array.from(currentSnapshot.itemsMap.values());
        
        // Agrupar por Comprador (Itens Recuperados)
        const buyers = {};
        const baseItems = Array.from(baseSnapshot.itemsMap.values());
        
        baseItems.forEach(iBase => {
            const iCurr = currentSnapshot.itemsMap.get(iBase.code) || { weight: 0 };
            const evol = iBase.weight - iCurr.weight;
            if (evol > 0) {
                buyers[iBase.comprador] = (buyers[iBase.comprador] || 0) + 1;
            }
        });

        const sortedBuyers = Object.entries(buyers).sort((a, b) => b[1] - a[1]);

        perfList.innerHTML = sortedBuyers.map(([name, count], idx) => `
            <div class="perf-item">
                <div class="perf-rank">${idx + 1}</div>
                <div class="perf-info">
                    <div class="perf-name">${name}</div>
                    <div class="perf-sub">${count} itens recuperados</div>
                </div>
                <div class="perf-bar-wrapper">
                    <div class="perf-bar" style="width: ${(count / sortedBuyers[0][1] * 100)}%"></div>
                </div>
            </div>
        `).join('');
    }

    function updateBuyerFilter() {
        const container = document.getElementById('buyer-filter-container');
        const buyers = new Set(snapshotHistory.flatMap(s => Array.from(s.itemsMap.values()).map(i => i.comprador)));
        
        const currentActive = activeBuyer;
        container.innerHTML = `<button class="buyer-btn ${activeBuyer === 'all' ? 'active' : ''}" data-buyer="all">👤 Todos</button>` + 
            Array.from(buyers).sort().map(b => `
                <button class="buyer-btn ${activeBuyer === b ? 'active' : ''}" data-buyer="${b}">${b}</button>
            `).join('');

        container.querySelectorAll('.buyer-btn').forEach(btn => {
            btn.onclick = () => {
                activeBuyer = btn.dataset.buyer;
                updateDashboard();
            };
        });
    }

    // --- Listeners ---
    folderInputs.forEach(input => input.onchange = (e) => handleFiles(e.target.files));
    timelineRange.oninput = (e) => {
        currentTimelineIdx = parseInt(e.target.value);
        updateDashboard();
    };
    tableSearch.oninput = renderTable;
    evolutionFilter.onchange = renderTable;

});
