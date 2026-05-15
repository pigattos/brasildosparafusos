/**
 * Rupturas Analítico v3.0 - Rebuilt from Scratch
 * Foco em Robustez de Dados e Comparativo Histórico
 */
document.addEventListener('DOMContentLoaded', () => {
    console.log("Rupturas Analítico v3.0 Inicializado");

    // --- Configurações ---
    const RECORRENCIA_MINIMA = 0.17; // > 17% (mínimo 2 meses em 6)
    
    // --- Estado Global ---
    let snapshotHistory = []; // Array de objetos { name, date, summary, items, rawData }
    let activeBuyer = 'all';
    let selectedSnapshots = new Set();
    let evolutionChart = null;
    let buyerMap = JSON.parse(localStorage.getItem('buyerMap') || '{}');
    let currentTimelineIdx = 0;

    // --- Elementos DOM ---
    const folderInput = document.getElementById('folder-upload');
    const historyTableBody = document.getElementById('history-table-body');
    const compareBtn = document.getElementById('compare-btn');
    const selectedCountEl = document.getElementById('selected-count');
    const modal = document.getElementById('comparison-modal');
    const modalBody = document.getElementById('modal-body');
    const closeModal = document.getElementById('close-modal');
    const loadingOverlay = document.getElementById('loading-overlay');

    // --- Utilitários de Formatação ---
    const formatCurrency = (v) => v.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
    const formatNumber = (v) => v.toLocaleString('pt-BR');

    // --- Sistema de Busca de Colunas ---
    function findColumn(headers, aliases) {
        const cleanHeaders = headers.map(h => String(h || '').toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, ""));
        const cleanAliases = aliases.map(a => a.toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, ""));

        // 1. Busca Exata
        for (let alias of cleanAliases) {
            const idx = cleanHeaders.indexOf(alias);
            if (idx !== -1) return headers[idx];
        }

        // 2. Busca Parcial
        for (let alias of cleanAliases) {
            const idx = cleanHeaders.findIndex(h => h.includes(alias));
            if (idx !== -1) return headers[idx];
        }

        return null;
    }

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
            if (parts.length > 2) {
                str = str.replace(/,/g, '');
            } else if (parts[1].length === 3) {
                str = str.replace(',', '');
            } else {
                str = str.replace(',', '.');
            }
        }
        
        const num = parseFloat(str);
        return isNaN(num) ? 0 : num;
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

                    if (rawRows.length === 0) return resolve(null);

                    // 1. Achar linha de cabeçalho
                    let headerIndex = -1;
                    for (let i = 0; i < Math.min(rawRows.length, 30); i++) {
                        const row = rawRows[i];
                        if (row.some(cell => {
                            const s = String(cell || '').toLowerCase();
                            return /^\d{1,2}\/\d{2,4}$/.test(s) || s.includes('estoque') || s.includes('produto') || s.includes('codigo');
                        })) {
                            headerIndex = i;
                            break;
                        }
                    }

                    if (headerIndex === -1) {
                        console.warn(`Cabeçalho não encontrado no arquivo ${file.name}`);
                        return resolve(null);
                    }

                    const headers = rawRows[headerIndex];
                    const rowsData = XLSX.utils.sheet_to_json(worksheet, { range: headerIndex });

                    // 2. Mapear Colunas Chave
                    const colMap = {
                        codigo: findColumn(headers, ['Produto', 'Código', 'Item', 'Cód.', 'ID', 'Referencia']),
                        desc: findColumn(headers, ['Descrição longa do produto', 'Descrição', 'Desc', 'Nome', 'Produto Descrição', 'Texto']),
                        estoque: findColumn(headers, ['Estoque', 'Saldo', 'Qtd. Estoque', 'Estoque Total', 'Saldo Atual', 'Saldo Disponível', 'Disp.', 'Qtd. Disponível', 'Estoque Atual']),
                        encomendas: findColumn(headers, [
                            'Encomendas', 'Qtd. Encomenda', 'Saldo Pedido Compra', 'Saldo Ped. Compra', 'Pedido Compra', 
                            'Qtd. em Pedido Compra', 'Qtd. no Pedido Compra', 'Saldo a Receber', 'A Receber', 'Pedidos', 
                            'Qtd. Pedida', 'Saldo Pedido', 'Compras', 'Qtd em Pedido', 'Qtd. Ped.', 'Saldo Ped.', 
                            'Pendência', 'Qtd. no Pedido', 'Encomenda', 'Pedido', 'Qtd Ped Compra', 'A Receber Total',
                            'A Entregar', 'Saldo a Entregar', 'Qtd. Pendente', 'Pendente', 'Saldo O.C.', 'Ord. Compra'
                        ]),
                        custo: findColumn(headers, ['Preço reposição', 'Custo aquisição', 'Custo Unitário', 'Custo', 'Preço Custo', 'Vlr. Custo', 'Custo Médio', 'Unitário']),
                        medVenda: findColumn(headers, ['Média Venda', 'Giro Mensal', 'Média Mensal', 'Giro Médio', 'Giro Dia', 'Media']),
                        vendasTotal: findColumn(headers, ['Vendas', 'Giro Total', 'Total Vendas', 'Saidas', 'Qtd. Vendida']),
                        comprador: findColumn(headers, ['Comprador', 'Responsável', 'Gestor'])
                    };

                    const monthCols = headers.filter(h => /^\d{1,2}\/\d{2,4}$/.test(String(h).trim()));

                    // 3. Processar Itens
                    const items = rowsData.map(row => {
                        const code = String(row[colMap.codigo] || '').trim();
                        const desc = String(row[colMap.desc] || '').trim();
                        if (!code && !desc) return null;

                        const estoque = parseNumeric(row[colMap.estoque]);
                        const encomendas = parseNumeric(row[colMap.encomendas]);
                        const custo = parseNumeric(row[colMap.custo]);
                        const vendasTotal = parseNumeric(row[colMap.vendasTotal]);
                        
                        let comprador = row[colMap.comprador];
                        if (!comprador && code) {
                            const cleanCode = code.replace(/^0+/, '').replace(/[.]/g, '');
                            comprador = buyerMap[code] || buyerMap[cleanCode] || 'N/D';
                        }

                        let medVenda = parseNumeric(row[colMap.medVenda]);
                        let recorrencia = 0;

                        // Se tiver colunas de meses, calcula recorrencia e media por elas
                        if (monthCols.length > 0) {
                            let activeMonths = 0;
                            let sumHistory = 0;
                            monthCols.forEach(mCol => {
                                const v = parseNumeric(row[mCol]);
                                if (v > 0) {
                                    activeMonths++;
                                    sumHistory += v;
                                }
                            });
                            recorrencia = activeMonths / monthCols.length;
                            // Sincronizar: Se tem histórico, a média é sempre baseada nele
                            medVenda = activeMonths > 0 ? sumHistory / activeMonths : 0;
                        } else {
                            // Se não tiver meses, tenta pegar recorrencia pronta se existir
                            const recRaw = parseNumeric(row[findColumn(headers, ['Recorrência', 'Frequência', 'Giro Freq'])]);
                            recorrencia = recRaw > 1 ? recRaw / 100 : recRaw;
                        }


                        // Classificação de Status
                        let status = 'ok';
                        const disponivel = estoque + encomendas;
                        if (recorrencia > RECORRENCIA_MINIMA) {
                            if (medVenda > disponivel) status = 'rupture';
                            else if ((medVenda * 2) > disponivel) status = 'attention';
                            else if ((medVenda * 3) > disponivel) status = 'suggest';
                        } else {
                            status = 'ignored';
                        }

                        return {
                            code, desc, comprador, status, 
                            medVenda, estoque, encomendas, custo,
                            value: medVenda * custo
                        };
                    }).filter(i => i !== null);

                    // 4. Extrair Data do Arquivo
                    let fileDate = new Date(file.lastModified).toISOString().split('T')[0];
                    const dateMatch = file.name.match(/(\d{4}[\.\-\/]\d{2}[\.\-\/]\d{2})|(\d{2}[\.\-\/]\d{2}[\.\-\/]\d{4})/);
                    if (dateMatch) {
                        let dm = dateMatch[0].replace(/[\.\/]/g, '-');
                        if (dm.split('-')[0].length === 2) {
                            const p = dm.split('-');
                            fileDate = `${p[2]}-${p[1]}-${p[0]}`;
                        } else {
                            fileDate = dm;
                        }
                    }

                    resolve({
                        name: file.name,
                        date: fileDate,
                        items: items
                    });
                } catch (err) {
                    console.error(`Erro ao ler arquivo ${file.name}:`, err);
                    resolve(null);
                }
            };
            reader.readAsArrayBuffer(file);
        });
    }

    // --- Funções de Interface ---
    function updateDashboard() {
        if (snapshotHistory.length === 0) return;

        // Filtrar por comprador para todos os snapshots
        const historyFiltered = snapshotHistory.map(snap => {
            const filteredItems = snap.items.filter(i => activeBuyer === 'all' || i.comprador === activeBuyer);
            
            const rupture = filteredItems.filter(i => i.status === 'rupture');
            const attention = filteredItems.filter(i => i.status === 'attention');
            const suggest = filteredItems.filter(i => i.status === 'suggest');

            return {
                ...snap,
                displayItems: filteredItems,
                summary: {
                    rupture: { count: rupture.length, value: rupture.reduce((acc, i) => acc + i.value, 0) },
                    attention: { count: attention.length, value: attention.reduce((acc, i) => acc + (i.value * 2), 0) },
                    suggest: { count: suggest.length, value: suggest.reduce((acc, i) => acc + (i.value * 3), 0) }
                }
            };
        });

        initTimeline(historyFiltered);
        renderCharts(historyFiltered);
        renderTable(historyFiltered);
        updateViewToSnapshot(historyFiltered, currentTimelineIdx);
    }

    function initTimeline(history) {
        const card = document.getElementById('timeline-card');
        const range = document.getElementById('timeline-range');
        const ticks = document.getElementById('timeline-ticks');
        if (!card || !range || history.length === 0) return;

        card.style.display = 'block';
        range.max = history.length - 1;
        
        // Se for o primeiro carregamento ou se o index estiver fora do novo range
        if (currentTimelineIdx >= history.length) currentTimelineIdx = history.length - 1;
        range.value = currentTimelineIdx;

        ticks.innerHTML = history.map((h, idx) => `
            <span style="cursor:pointer; opacity: ${idx === currentTimelineIdx ? '1' : '0.5'}" onclick="document.getElementById('timeline-range').value=${idx}; document.getElementById('timeline-range').dispatchEvent(new Event('input'));">
                ${h.date.split('-').slice(1).reverse().join('/')}
            </span>
        `).join('');
    }

    function updateViewToSnapshot(history, idx) {
        const snap = history[idx];
        if (!snap) return;

        document.getElementById('current-snapshot-date').textContent = snap.date.split('-').reverse().join('/');
        document.getElementById('current-snapshot-name').textContent = snap.name;

        renderSummary(history, idx);
        renderDailyDiff(history, idx);
        renderEfficiency(history, idx);
    }

    function renderEfficiency(history, idx) {
        const grid = document.getElementById('efficiency-grid');
        const card = document.getElementById('efficiency-card');
        const projectionEl = document.getElementById('efficiency-projection');
        if (!grid || history.length < 2) return;

        card.style.display = 'block';
        const snapBase = history[0]; // 07/05/2026
        const snapCurrent = history[idx];
        const dateBaseStr = snapBase.date.split('-').reverse().join('/');

        // Comparar item por item do base com o estado no current
        const baseItems = snapBase.displayItems.filter(i => i.status !== 'ok' && i.status !== 'ignored');
        const currentMap = new Map(snapCurrent.displayItems.map(i => [i.code, i.status]));

        const results = {
            rupture: { total: 0, attended: 0 },
            attention: { total: 0, attended: 0 },
            suggest: { total: 0, attended: 0 }
        };

        baseItems.forEach(item => {
            results[item.status].total++;
            const currentStatus = currentMap.get(item.code) || 'ok';
            
            // VOLTAR AO QUE ESTAVA: Lógica de melhoria de severidade
            // Considera "Atendido" se o status atual for melhor (menos severo) que o original
            const severity = { 'rupture': 3, 'attention': 2, 'suggest': 1, 'ok': 0, 'ignored': 0 };
            if (severity[currentStatus] < severity[item.status]) {
                results[item.status].attended++;
            }
        });

        const totalInitial = results.rupture.total + results.attention.total + results.suggest.total;
        const totalAttended = results.rupture.attended + results.attention.attended + results.suggest.attended;
        
        // Cálculo de Projeção
        const dateBase = new Date(snapBase.date);
        const dateCurrent = new Date(snapCurrent.date);
        const diffTime = Math.abs(dateCurrent - dateBase);
        const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) || 1;
        
        const ratePerDay = totalAttended / diffDays;
        const remaining = totalInitial - totalAttended;
        const daysToFinish = ratePerDay > 0 ? Math.ceil(remaining / ratePerDay) : '∞';

        projectionEl.innerHTML = `
            <div style="display: flex; justify-content: space-between; align-items: center; width: 100%;">
                <span>📅 Base: <strong>${dateBaseStr}</strong> vs <strong>${snapCurrent.date.split('-').reverse().join('/')}</strong></span>
                <span>🚀 Projeção: <strong>${daysToFinish} dias</strong> para zerar pendências (${ratePerDay.toFixed(1)} itens/dia)</span>
            </div>
        `;

        grid.innerHTML = `
            <div class="comp-box" style="border-top: 3px solid #fb7185;">
                <div style="font-size: 0.7rem; color: var(--text-muted); font-weight: 700; margin-bottom: 0.5rem;">RUPTURAS ATENDIDAS</div>
                <div style="font-size: 1.4rem; font-weight: 800;">${results.rupture.attended} / ${results.rupture.total} <small style="font-size: 0.7rem; opacity: 0.6;">Resolvidos</small></div>
                <div class="progress-bar-mini" style="background: rgba(251, 113, 133, 0.1);"><div style="width: ${(results.rupture.attended / results.rupture.total * 100) || 0}%; background: #fb7185;"></div></div>
                <div style="font-size: 0.65rem; margin-top: 0.5rem; color: #fb7185;">${((results.rupture.attended / results.rupture.total * 100) || 0).toFixed(1)}% de eficiência</div>
            </div>
            <div class="comp-box" style="border-top: 3px solid #f59e0b;">
                <div style="font-size: 0.7rem; color: var(--text-muted); font-weight: 700; margin-bottom: 0.5rem;">ATENÇÕES ATENDIDAS</div>
                <div style="font-size: 1.4rem; font-weight: 800;">${results.attention.attended} / ${results.attention.total} <small style="font-size: 0.7rem; opacity: 0.6;">Resolvidos</small></div>
                <div class="progress-bar-mini" style="background: rgba(245, 158, 11, 0.1);"><div style="width: ${(results.attention.attended / results.attention.total * 100) || 0}%; background: #f59e0b;"></div></div>
                <div style="font-size: 0.65rem; margin-top: 0.5rem; color: #f59e0b;">${((results.attention.attended / results.attention.total * 100) || 0).toFixed(1)}% de eficiência</div>
            </div>
            <div class="comp-box" style="border-top: 3px solid #34d399;">
                <div style="font-size: 0.7rem; color: var(--text-muted); font-weight: 700; margin-bottom: 0.5rem;">SUGESTÕES ATENDIDAS</div>
                <div style="font-size: 1.4rem; font-weight: 800;">${results.suggest.attended} / ${results.suggest.total} <small style="font-size: 0.7rem; opacity: 0.6;">Resolvidos</small></div>
                <div class="progress-bar-mini" style="background: rgba(52, 211, 153, 0.1);"><div style="width: ${(results.suggest.attended / results.suggest.total * 100) || 0}%; background: #34d399;"></div></div>
                <div style="font-size: 0.65rem; margin-top: 0.5rem; color: #34d399;">${((results.suggest.attended / results.suggest.total * 100) || 0).toFixed(1)}% de eficiência</div>
            </div>
            <div class="comp-box" style="background: rgba(52, 211, 153, 0.05); border-top: 3px solid #10b981;">
                <div style="font-size: 0.7rem; color: #10b981; font-weight: 700; margin-bottom: 0.5rem;">RESUMO GLOBAL</div>
                <div style="font-size: 1.4rem; font-weight: 800; color: #10b981;">${totalAttended} / ${totalInitial}</div>
                <div class="progress-bar-mini" style="background: rgba(16, 185, 129, 0.1);"><div style="width: ${(totalAttended / totalInitial * 100) || 0}%; background: #10b981;"></div></div>
                <div style="font-size: 0.65rem; margin-top: 0.5rem; color: #10b981;">${((totalAttended / totalInitial * 100) || 0).toFixed(1)}% da meta atingida</div>
            </div>
        `;
    }


    function renderSummary(history, idx) {
        const latest = history[idx];
        if (!latest) return;

        document.getElementById('rupture-value').textContent = formatCurrency(latest.summary.rupture.value);
        document.getElementById('rupture-count').textContent = `${latest.summary.rupture.count} itens em risco total`;
        
        document.getElementById('attention-value').textContent = formatCurrency(latest.summary.attention.value);
        document.getElementById('attention-count').textContent = `${latest.summary.attention.count} itens em monitoramento`;
        
        document.getElementById('suggest-value').textContent = formatCurrency(latest.summary.suggest.value);
        document.getElementById('suggest-count').textContent = `${latest.summary.suggest.count} itens para reposição`;

        // Tendências (Com relação ao snapshot anterior)
        if (idx > 0) {
            const prev = history[idx - 1];
            updateTrendBadge('rupture-trend', latest.summary.rupture.count, prev.summary.rupture.count, true);
            updateTrendBadge('attention-trend', latest.summary.attention.count, prev.summary.attention.count, true);
            updateTrendBadge('suggest-trend', latest.summary.suggest.count, prev.summary.suggest.count, true);
        } else {
            document.getElementById('rupture-trend').innerHTML = '';
            document.getElementById('attention-trend').innerHTML = '';
            document.getElementById('suggest-trend').innerHTML = '';
        }
    }

    function updateTrendBadge(id, curr, prev, inverse = false) {
        const el = document.getElementById(id);
        if (!el || prev === 0) return;
        const diff = ((curr - prev) / prev) * 100;
        const isUp = diff > 0;
        const color = (isUp ^ inverse) ? '#34d399' : '#fb7185';
        el.innerHTML = `<span style="color: ${color}">${isUp ? '▲' : '▼'} ${Math.abs(diff).toFixed(1)}%</span>`;
    }

    function renderTable(history) {
        historyTableBody.innerHTML = '';
        [...history].reverse().forEach(snap => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td><input type="checkbox" class="snap-check" data-id="${snap.name}" ${selectedSnapshots.has(snap.name) ? 'checked' : ''}></td>
                <td>${snap.name}</td>
                <td>${snap.date}</td>
                <td style="text-align: center;">${snap.items.length}</td>
                <td style="text-align: center;"><span class="badge badge-buy">${snap.summary.rupture.count}</span></td>
                <td style="text-align: right; font-weight: 600;">${formatCurrency(snap.summary.rupture.value)}</td>
                <td style="text-align: center;">
                    <button class="btn btn-secondary btn-sm preview-snap" data-id="${snap.name}">🔎 Ver</button>
                </td>
            `;
            historyTableBody.appendChild(tr);
        });

        // Listeners Checkboxes
        document.querySelectorAll('.snap-check').forEach(cb => {
            cb.addEventListener('change', () => {
                if (cb.checked) selectedSnapshots.add(cb.dataset.id);
                else selectedSnapshots.delete(cb.dataset.id);
                updateCompareButton();
            });
        });

        // Listener para o botão "Ver" individual
        document.querySelectorAll('.preview-snap').forEach(btn => {
            btn.addEventListener('click', () => {
                const snapName = btn.dataset.id;
                const snap = snapshotHistory.find(s => s.name === snapName);
                if (snap) {
                    showSingleSnapshot(snap);
                }
            });
        });
    }

    function showSingleSnapshot(snap) {
        const items = snap.items.filter(i => activeBuyer === 'all' || i.comprador === activeBuyer);
        const rupture = items.filter(i => i.status === 'rupture');
        const attention = items.filter(i => i.status === 'attention');
        const suggest = items.filter(i => i.status === 'suggest');

        modalBody.innerHTML = `
            <div style="margin-bottom: 2rem;">
                <h3>Relatório: ${snap.name}</h3>
                <p style="color: var(--text-muted)">Data: ${snap.date} | Itens Analisados: ${items.length}</p>
            </div>
            <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 1rem; margin-bottom: 2rem;">
                <div class="comp-box" style="border-left: 4px solid #fb7185;">
                    <div style="font-size: 0.8rem; color: #fb7185;">RUPTURAS</div>
                    <div style="font-size: 1.5rem; font-weight: 700;">${rupture.length}</div>
                </div>
                <div class="comp-box" style="border-left: 4px solid #f59e0b;">
                    <div style="font-size: 0.8rem; color: #f59e0b;">ATENÇÃO</div>
                    <div style="font-size: 1.5rem; font-weight: 700;">${attention.length}</div>
                </div>
                <div class="comp-box" style="border-left: 4px solid #34d399;">
                    <div style="font-size: 0.8rem; color: #34d399;">SUGESTÃO</div>
                    <div style="font-size: 1.5rem; font-weight: 700;">${suggest.length}</div>
                </div>
            </div>
            <div class="table-container" style="max-height: 400px;">
                <table>
                    <thead>
                        <tr><th>Código</th><th>Descrição</th><th>Status</th><th style="text-align:right">Valor</th></tr>
                    </thead>
                    <tbody>
                        ${items.filter(i => i.status !== 'ok' && i.status !== 'ignored').map(i => `
                            <tr>
                                <td>${i.code}</td>
                                <td>${i.desc}</td>
                                <td><span class="badge badge-${i.status}">${i.status}</span></td>
                                <td style="text-align:right">${formatCurrency(i.value)}</td>
                            </tr>
                        `).join('')}
                    </tbody>
                </table>
            </div>
        `;
        modal.style.display = 'flex';
    }

    function updateCompareButton() {
        selectedCountEl.textContent = selectedSnapshots.size;
        compareBtn.style.display = selectedSnapshots.size >= 2 ? 'block' : 'none';
    }

    function renderCharts(history) {
        const ctx = document.getElementById('evolution-chart');
        if (!ctx) return;

        if (evolutionChart) evolutionChart.destroy();

        evolutionChart = new Chart(ctx, {
            type: 'line',
            data: {
                labels: history.map(h => h.date),
                datasets: [
                    {
                        label: 'Rupturas',
                        data: history.map(h => h.summary.rupture.count),
                        borderColor: '#fb7185',
                        backgroundColor: 'rgba(251, 113, 133, 0.1)',
                        fill: true,
                        tension: 0.4
                    },
                    {
                        label: 'Atenção',
                        data: history.map(h => h.summary.attention.count),
                        borderColor: '#f59e0b',
                        backgroundColor: 'rgba(245, 158, 11, 0.1)',
                        fill: true,
                        tension: 0.4
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { legend: { display: false } },
                scales: {
                    y: { beginAtZero: true, grid: { color: 'rgba(255,255,255,0.05)' } },
                    x: { grid: { display: false } }
                }
            }
        });
    }

    function renderDailyDiff(history, idx) {
        const grid = document.getElementById('daily-diff-grid');
        const card = document.getElementById('daily-diff-card');
        if (!grid || history.length < 2) return;

        card.style.display = 'block';
        const snapOld = history[0]; // Sempre comparar com o início para ver o ganho acumulado
        const snapNew = history[idx]; // Ponto atual da linha do tempo

        // Se o slider estiver no primeiro item, compara com o próximo só para não ficar vazio, ou oculta
        if (idx === 0 && history.length > 1) {
            // No primeiro dia, mostra comparativo com o último disponível? Ou avisa que é o ponto inicial
            card.querySelector('h3').innerHTML = `<span class="header-icon">🔄</span> Ponto Inicial: ${snapOld.date}`;
            grid.innerHTML = `<div style="grid-column: 1/-1; text-align:center; padding: 1rem; color: var(--text-muted);">Este é o primeiro relatório do histórico. Arraste a linha do tempo para ver a evolução.</div>`;
            return;
        }

        // Atualizar título do card para refletir o período selecionado
        const titleEl = card.querySelector('h3');
        if (titleEl) {
            titleEl.innerHTML = `<span class="header-icon">🔄</span> Evolução Acumulada (${snapOld.date} à ${snapNew.date})`;
        }

        const res = calculateDiff(snapOld.displayItems, snapNew.displayItems);

        grid.innerHTML = `
            <div class="comp-box" style="border-left: 4px solid #34d399;">
                <div style="font-size: 0.8rem; color: #34d399; font-weight: 700;">✅ SANADOS NO PERÍODO</div>
                <div style="font-size: 1.5rem; font-weight: 800;">${res.solved.length} itens</div>
                <div style="font-size: 0.7rem; color: var(--text-muted);">Itens que saíram da ruptura desde o início.</div>
            </div>
            <div class="comp-box" style="border-left: 4px solid #fb7185;">
                <div style="font-size: 0.8rem; color: #fb7185; font-weight: 700;">🚨 NOVOS EM RISCO</div>
                <div style="font-size: 1.5rem; font-weight: 800;">${res.newRisks.length} itens</div>
                <div style="font-size: 0.7rem; color: var(--text-muted);">Novas rupturas detectadas no período.</div>
            </div>
            <div class="comp-box" style="border-left: 4px solid #f59e0b;">
                <div style="font-size: 0.8rem; color: #f59e0b; font-weight: 700;">📉 PIORARAM</div>
                <div style="font-size: 1.5rem; font-weight: 800;">${res.worsened.length} itens</div>
                <div style="font-size: 0.7rem; color: var(--text-muted);">Situação agravada entre o primeiro e último dia.</div>
            </div>
        `;
    }

    function calculateDiff(itemsA, itemsB) {
        const mapA = new Map(itemsA.filter(i => i.status !== 'ok' && i.status !== 'ignored').map(i => [i.code, i]));
        const mapB = new Map(itemsB.filter(i => i.status !== 'ok' && i.status !== 'ignored').map(i => [i.code, i]));

        const solved = [];
        const newRisks = [];
        const worsened = [];

        mapA.forEach((itemA, code) => {
            if (!mapB.has(code)) solved.push(itemA);
            else {
                const itemB = mapB.get(code);
                const weights = { rupture: 3, attention: 2, suggest: 1 };
                if (weights[itemB.status] > weights[itemA.status]) worsened.push({ from: itemA, to: itemB });
            }
        });

        mapB.forEach((itemB, code) => {
            if (!mapA.has(code)) newRisks.push(itemB);
        });

        return { solved, newRisks, worsened };
    }

    // --- Eventos ---
    folderInput.addEventListener('change', async (e) => {
        const files = Array.from(e.target.files).filter(f => f.name.match(/\.(xlsx|xls)$/i) && !f.name.startsWith('~$') && !f.name.toLowerCase().includes('leadtime'));
        if (files.length === 0) return;

        loadingOverlay.style.display = 'flex';
        snapshotHistory = [];
        selectedSnapshots.clear();
        currentTimelineIdx = 0; // Reset timeline

        for (let file of files) {
            const snap = await processExcelFile(file);
            if (snap) snapshotHistory.push(snap);
        }

        snapshotHistory.sort((a, b) => new Date(a.date) - new Date(b.date));
        currentTimelineIdx = snapshotHistory.length - 1; // Iniciar no snapshot mais recente
        updateDashboard();
        loadingOverlay.style.display = 'none';
    });

    document.querySelectorAll('.buyer-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('.buyer-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            activeBuyer = btn.dataset.buyer;
            updateDashboard();
        });
    });

    compareBtn.addEventListener('click', () => {
        console.log("Iniciando comparação...");
        const selected = snapshotHistory
            .filter(s => selectedSnapshots.has(s.name))
            .sort((a, b) => {
                const dateA = new Date(a.date.includes('-') ? a.date : a.date.split('.').reverse().join('-'));
                const dateB = new Date(b.date.includes('-') ? b.date : b.date.split('.').reverse().join('-'));
                return dateA - dateB;
            });

        if (selected.length < 2) {
            alert("Selecione pelo menos 2 arquivos para comparar.");
            return;
        }

        const snapOld = selected[0];
        const snapNew = selected[selected.length - 1];
        
        // Aplicar filtro de comprador nos itens da comparação também
        const itemsOld = snapOld.items.filter(i => activeBuyer === 'all' || i.comprador === activeBuyer);
        const itemsNew = snapNew.items.filter(i => activeBuyer === 'all' || i.comprador === activeBuyer);
        
        const res = calculateDiff(itemsOld, itemsNew);

        modalBody.innerHTML = `
            <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 2rem; margin-bottom: 2rem; background: rgba(0,0,0,0.2); padding: 1.5rem; border-radius: 8px;">
                <div>
                    <div style="color: var(--text-muted); font-size: 0.8rem;">Snapshot Inicial</div>
                    <div style="font-weight: 700; font-size: 1.1rem;">${snapOld.date}</div>
                    <div style="font-size: 0.8rem; opacity: 0.7;">${snapOld.name}</div>
                </div>
                <div style="text-align: right;">
                    <div style="color: var(--text-muted); font-size: 0.8rem;">Snapshot Final</div>
                    <div style="font-weight: 700; font-size: 1.1rem;">${snapNew.date}</div>
                    <div style="font-size: 0.8rem; opacity: 0.7;">${snapNew.name}</div>
                </div>
            </div>

            <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 1.5rem;">
                <div class="comp-box">
                    <h4 style="color: #34d399; margin-bottom: 1rem;">✅ Resolvidos (${res.solved.length})</h4>
                    <div style="max-height: 400px; overflow-y: auto;">
                        <table style="width: 100%; font-size: 0.8rem;">
                            ${res.solved.map(i => `<tr><td>${i.code}</td><td>${i.desc.substring(0,25)}...</td><td style="text-align:right">${formatCurrency(i.value)}</td></tr>`).join('')}
                        </table>
                    </div>
                </div>
                <div class="comp-box">
                    <h4 style="color: #fb7185; margin-bottom: 1rem;">🚨 Novos em Risco (${res.newRisks.length})</h4>
                    <div style="max-height: 400px; overflow-y: auto;">
                        <table style="width: 100%; font-size: 0.8rem;">
                            ${res.newRisks.map(i => `<tr><td>${i.code}</td><td>${i.desc.substring(0,25)}...</td><td style="text-align:right; color:#fb7185">${formatCurrency(i.value)}</td></tr>`).join('')}
                        </table>
                    </div>
                </div>
                <div class="comp-box">
                    <h4 style="color: #f59e0b; margin-bottom: 1rem;">📉 Pioraram (${res.worsened.length})</h4>
                    <div style="max-height: 400px; overflow-y: auto;">
                        <table style="width: 100%; font-size: 0.8rem;">
                            ${res.worsened.map(w => `<tr><td>${w.to.code}</td><td><span style="font-size:0.7rem">${w.from.status}</span> → <span style="color:#fb7185; font-weight:700">${w.to.status}</span></td></tr>`).join('')}
                        </table>
                    </div>
                </div>
            </div>
        `;
        modal.style.display = 'flex';
    });

    closeModal.addEventListener('click', () => modal.style.display = 'none');
    window.addEventListener('click', (e) => { if (e.target === modal) modal.style.display = 'none'; });

    // Listener para o Slider de Linha do Tempo
    const timelineRange = document.getElementById('timeline-range');
    if (timelineRange) {
        timelineRange.addEventListener('input', (e) => {
            currentTimelineIdx = parseInt(e.target.value);
            
            // Recalcular historyFiltered localmente para atualizar a view
            // (Poderia ser otimizado guardando historyFiltered globalmente)
            const historyFiltered = snapshotHistory.map(snap => {
                const filteredItems = snap.items.filter(i => activeBuyer === 'all' || i.comprador === activeBuyer);
                const rupture = filteredItems.filter(i => i.status === 'rupture');
                const attention = filteredItems.filter(i => i.status === 'attention');
                const suggest = filteredItems.filter(i => i.status === 'suggest');
                return {
                    ...snap,
                    displayItems: filteredItems,
                    summary: {
                        rupture: { count: rupture.length, value: rupture.reduce((acc, i) => acc + i.value, 0) },
                        attention: { count: attention.length, value: attention.reduce((acc, i) => acc + (i.value * 2), 0) },
                        suggest: { count: suggest.length, value: suggest.reduce((acc, i) => acc + (i.value * 3), 0) }
                    }
                };
            });
            
            updateViewToSnapshot(historyFiltered, currentTimelineIdx);
            
            // Atualizar opacidade dos labels (ticks)
            const ticks = document.querySelectorAll('#timeline-ticks span');
            ticks.forEach((t, idx) => t.style.opacity = idx === currentTimelineIdx ? '1' : '0.5');
        });
    }
});
