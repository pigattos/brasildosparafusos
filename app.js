document.addEventListener('DOMContentLoaded', () => {
    console.log("Gestor de Estoque v2.1 - Filtro de meses ativos (até 04/2026)");
    const fileUpload = document.getElementById('file-upload');
    const dropZone = document.getElementById('drop-zone');
    const statsSection = document.getElementById('stats-section');
    const tableContainer = document.getElementById('table-container');
    const tableBody = document.getElementById('table-body');
    const tableSearch = document.getElementById('table-search');
    const supplierSearch = document.getElementById('supplier-search');
    const filterBtns = document.querySelectorAll('.filter-btn');
    const clearFiltersBtn = document.getElementById('clear-filters');
    const chartSection = document.getElementById('chart-section');

    // Register ChartJS Plugin
    Chart.register(ChartDataLabels);

    let currentData = [];
    let filteredData = [];
    let activeFilters = ['all'];
    let sortRecorrenciaDir = 'none';
    let myChart = null;
    let supplierChart = null;
    let groupChart = null;

    // --- Core Logic ---

    /**
     * Normalizes Unit of Measure to 'Pç'
     */
    function normalizeUnit(un) {
        if (!un) return 'Pç';
        const u = un.toString().toUpperCase().trim();
        if (u === 'PC' || u === 'UN' || u === 'UND' || u === 'PÇ' || u === 'PEÇA') return 'Pç';
        return u;
    }

    /**
     * Tenta encontrar o valor em uma linha usando múltiplos nomes de coluna possíveis (Case-Insensitive)
     */
    function getValue(row, keys) {
        const rowKeys = Object.keys(row);
        const cleanStr = (s) => s.toString().toLowerCase().trim()
            .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
            .replace(/[.\-_/\\()]/g, ' ')
            .replace(/\s+/g, ' ')
            .trim();

        for (let key of keys) {
            const target = cleanStr(key);
            
            // 1. Exact match in raw keys
            if (row[key] !== undefined) return { value: row[key], col: key };
            
            // 2. Exact match in normalized keys
            const normKey = rowKeys.find(rk => cleanStr(rk) === target);
            if (normKey) return { value: row[normKey], col: normKey };
        }

        // 3. Fallback: word match or partial match
        for (let key of keys) {
            const target = cleanStr(key);
            if (target.length < 3) continue;
            
            const fuzzyKey = rowKeys.find(rk => {
                const cRK = cleanStr(rk);
                const words = cRK.split(' ');
                
                // Matches if any word in the column name starts with our target or vice versa
                return words.some(word => {
                    if (word === target) return true;
                    if (word.length >= 4 && target.startsWith(word)) return true;
                    if (target.length >= 4 && word.startsWith(target)) return true;
                    return false;
                });
            });
            if (fuzzyKey) return { value: row[fuzzyKey], col: fuzzyKey };
        }
        
        return { value: undefined, col: 'N/A' };
    }

    /**
     * Converte valor para número de forma segura, tratando formatos brasileiros (vírgula/ponto)
     */
    function parseNum(val) {
        if (val === undefined || val === null || val === '') return 0;
        if (typeof val === 'number') return val;
        let str = val.toString().replace('R$', '').replace(/\s/g, '').trim();
        
        // Handle BR/INT formatting automatically
        if (str.startsWith('.') || str.startsWith(',')) str = '0' + str;

        const hasComma = str.includes(',');
        const hasDot = str.includes('.');

        if (hasComma && hasDot) {
            if (str.lastIndexOf(',') > str.lastIndexOf('.')) {
                str = str.replace(/\./g, '').replace(',', '.'); // BR
            } else {
                str = str.replace(/,/g, ''); // INT
            }
        } else if (hasComma) {
            str = str.replace(',', '.');
        }
        
        const num = parseFloat(str);
        // For inventory/sales, we generally treat negative/NaN as 0 to avoid visual artifacts
        return isNaN(num) ? 0 : num;
    }

    /**
     * Calculates Recurrence % (Months with sales > 0)
     */
    function calculateRecurrence(row, monthCols) {
        let count = 0;
        if (monthCols && monthCols.length > 0) {
            monthCols.forEach(col => {
                const val = parseNum(row[col]);
                if (val > 0) count++;
            });
        }
        const percent = monthCols.length > 0 ? (count / monthCols.length) * 100 : 0;
        return { percent, count };
    }

    /**
     * Processes raw JSON from Excel
     */
    function processData(rawData) {
        if (!rawData || rawData.length === 0) return [];

        // IMPROVED: Gather all unique keys across all rows to avoid missing columns
        // if the first row is incomplete.
        const allKeysSet = new Set();
        rawData.forEach(row => {
            Object.keys(row).forEach(k => allKeysSet.add(k));
        });
        const allKeys = Array.from(allKeysSet);

        // Identify Month Columns (format MM/YYYY) - Filtered to exclude future months
        const monthCols = allKeys.filter(k => {
            if (!/^\d{2}\/\d{4}$/.test(k)) return false;
            const [m, y] = k.split('/').map(Number);
            // Limit to months <= 04/2026 (requested range)
            const monthVal = y * 12 + m;
            const limitVal = 2026 * 12 + 4;
            return monthVal <= limitVal;
        }).sort((a, b) => {
            const [mA, yA] = a.split('/');
            const [mB, yB] = b.split('/');
            return (yA + mA).localeCompare(yB + mB);
        }).slice(-6); // Ensure only the last 6 months are considered
        
        console.log(`Dados Brutos: ${rawData.length} linhas | Colunas Detectadas: ${allKeys.length}`);
        console.log('Meses identificados:', monthCols);

        const processed = [];
        
        rawData.forEach((row, index) => {
            const prodObj = getValue(row, ['Produto', 'Código', 'Item', 'Cód.', 'ID']);
            const descObj = getValue(row, ['Descrição longa do produto', 'Descrição', 'Desc', 'Nome', 'Produto Descrição', 'Texto']);
            
            // VALIDATION: Skip row if it has no product ID AND no description (likely empty or junk row)
            if (!prodObj.value && !descObj.value) {
                return; 
            }

            const produto = prodObj.value ? prodObj.value.toString().trim() : 'N/A';
            const descricao = descObj.value ? descObj.value.toString().trim() : 'Sem descrição';

            const unObj = getValue(row, ['UN', 'Unidade', 'Medida', 'U', 'Med']);
            const grupoObj = getValue(row, ['Grupo', 'Categoria', 'Família', 'Linha', 'Grupo de Produto', 'Cód. Grupo', 'Subgrupo']);
            const fornecObj = getValue(row, ['Razão social do fornecedor', 'Fornecedor', 'Fornec', 'Fabricante', 'Último Fornecedor', 'Fornecedor Principal', 'Nome Fornecedor']);
            const vendasObj = getValue(row, ['Vendas', 'Qtd. Vendida', 'Venda Total', 'Total Vendas', 'Venda', 'Saídas', 'Giro']);
            const estoqueObj = getValue(row, ['Estoque', 'Saldo', 'Qtd. Estoque', 'Estoque Total', 'Saldo Atual', 'Saldo Disponível', 'Disp.', 'Qtd. Disponível', 'Estoque Atual']);
            const encomObj = getValue(row, [
                'Encomendas', 'Qtd. Encomenda', 'Saldo Pedido Compra', 'Saldo Ped. Compra', 'Pedido Compra', 
                'Qtd. em Pedido Compra', 'Qtd. no Pedido Compra', 'Saldo a Receber', 'A Receber', 'Pedidos', 
                'Qtd. Pedida', 'Saldo Pedido', 'Compras', 'Qtd em Pedido', 'Qtd. Ped.', 'Saldo Ped.', 
                'Pendência', 'Qtd. no Pedido', 'Encomenda', 'Pedido', 'Qtd Ped Compra', 'A Receber Total',
                'A Entregar', 'Saldo a Entregar', 'Qtd. Pendente', 'Pendente', 'Saldo O.C.', 'Ord. Compra'
            ]);
            const custoObj = getValue(row, ['Custo aquisição', 'Custo Unitário', 'Custo', 'Preço Custo', 'Vlr. Custo', 'Custo Médio', 'Unitário']);

            let vendas = parseNum(vendasObj.value);
            const recData = calculateRecurrence(row, monthCols);
            
            // FALLBACK: If "Vendas" column is 0 or missing, sum the month columns
            if (vendas === 0 && monthCols.length > 0) {
                monthCols.forEach(col => {
                    const v = parseNum(row[col]);
                    if (v > 0) vendas += v;
                });
            }

            const medVenda = recData.count > 0 ? (vendas / recData.count) : 0;
            
            const estoque = parseNum(estoqueObj.value);
            const encomendas = parseNum(encomObj.value);
            const custo = parseNum(custoObj.value);
            
            const recorrencia = recData.percent;
            
            const mappingInfo = `Linha: ${index + 2} | Cód: "${prodObj.col}" | Estoque: "${estoqueObj.col}" | Encomendas: "${encomObj.col}"`;
            
            // Logic for risk classification (Refactored for maximum clarity)
            let emRisco = false;
            let emAtencao = false;
            let emSugestao = false;
            let situacao = 'seguro';

            // Only evaluate for risks if recurrence is > 33%
            if (recorrencia > 33) {
                const totalDisponivel = estoque + encomendas;
                
                if (medVenda > totalDisponivel) {
                    emRisco = true;
                    situacao = 'ruptura';
                } else if ((medVenda * 2) > totalDisponivel) {
                    emAtencao = true;
                    situacao = 'atencao';
                } else if ((medVenda * 3) > totalDisponivel) {
                    emSugestao = true;
                    situacao = 'sugestao';
                }
            }

            let tendencia = 'stable';
            if (monthCols.length >= 2) {
                const lastVal = parseNum(row[monthCols[monthCols.length - 1]]);
                const prevVal = parseNum(row[monthCols[monthCols.length - 2]]);
                if (lastVal > prevVal) tendencia = 'up';
                else if (lastVal < prevVal) tendencia = 'down';
            }

            processed.push({
                produto,
                descricao,
                un: normalizeUnit(unObj.value),
                grupo: grupoObj.value || 'Geral',
                fornecedor: fornecObj.value || 'N/D',
                estoque,
                encomendas,
                vendas,
                medVenda: medVenda.toFixed(3),
                tendencia,
                custo: custo.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' }),
                recorrencia: recorrencia.toFixed(0),
                historico: monthCols.map(col => Math.max(0, parseNum(row[col]))), // Ensure non-negative history for chart
                monthLabels: monthCols,
                situacao,
                emRisco,
                emAtencao,
                emSugestao,
                temEncomenda: encomendas > 0,
                mappingInfo
            });
        });

        console.log(`Itens Processados: ${processed.length}`);
        return processed;
    }

    /**
     * Renders processed data to the table
     */
    function renderTable(data) {
        tableBody.innerHTML = '';
        data.forEach(item => {
            const mainTr = document.createElement('tr');
            mainTr.className = 'main-row';
            
            let badgeClass = 'badge-ok';
            let badgeText = '✅ SEGURO';
            if (item.situacao === 'ruptura') {
                badgeClass = 'badge-buy';
                badgeText = '🚨 RUPTURA';
            } else if (item.situacao === 'atencao') {
                badgeClass = 'badge-warn';
                badgeText = '⚡ ATENÇÃO';
            } else if (item.situacao === 'sugestao') {
                badgeClass = 'badge-caution';
                badgeText = '⚠️ COMPRAR';
            }
            
            // Highlight if there's a pending order
            if (item.temEncomenda) {
                badgeText += ' 📦';
            }
            
            mainTr.innerHTML = `
                <td>
                    <span class="cell-label">Produto</span>
                    <div class="cell-value" style="font-weight: 600;">${item.produto}</div>
                </td>
                <td>
                    <span class="cell-label">Descrição</span>
                    <div class="cell-value td-desc" title="${item.descricao}">${item.descricao}</div>
                </td>
                <td style="text-align: center;">
                    <span class="cell-label">UN</span>
                    <div class="cell-value">${item.un}</div>
                </td>

                <td title="${item.fornecedor}">
                    <span class="cell-label">Fornecedor</span>
                    <div class="cell-value td-supplier">${item.fornecedor}</div>
                </td>
                <td style="text-align: center;">
                    <span class="cell-label">Estoque</span>
                    <div class="cell-value" style="display: flex; flex-direction: column; align-items: center;">
                        <span style="font-weight: 600;">${item.estoque}</span>
                        ${item.temEncomenda ? `<span style="font-size: 0.75rem; color: var(--info); font-weight: 700;" title="Saldo em Pedido">+ ${item.encomendas} ped.</span>` : ''}
                    </div>
                </td>
                <td style="text-align: center;">
                    <span class="cell-label">Vendas (6m)</span>
                    <div class="cell-value" style="display: flex; align-items: center; justify-content: center;">
                        <span>${item.vendas}</span>
                        ${item.tendencia === 'up' ? '<span class="trend-indicator trend-up" title="Vendas subindo em relação ao mês anterior">▲</span>' : 
                          item.tendencia === 'down' ? '<span class="trend-indicator trend-down" title="Vendas caindo em relação ao mês anterior">▼</span>' : ''}
                    </div>
                </td>
                <td style="text-align: center;">
                    <span class="cell-label">Recorrência</span>
                    <div class="cell-value" style="display: flex; flex-direction: column; align-items: center; gap: 4px;">
                        <div class="rec-bar-container" style="width: 60px;">
                            <div class="rec-bar" style="width: ${item.recorrencia}%"></div>
                        </div>
                        <span style="font-size: 0.75rem; font-weight: 600;">${item.recorrencia}%</span>
                    </div>
                </td>
                <td style="text-align: center;">
                    <span class="cell-label">Estado</span>
                    <div class="cell-value">
                        <span class="badge ${badgeClass}">${badgeText}</span>
                    </div>
                </td>
            `;
            
            const detailTr = document.createElement('tr');
            detailTr.className = 'detail-row hidden';
            detailTr.innerHTML = `
                <td colspan="8">
                    <div class="row-details">
                        <div class="detail-item">
                            <span class="detail-label">Encomendas</span>
                            <span class="detail-value">${item.encomendas}</span>
                        </div>
                        <div class="detail-item">
                            <span class="detail-label">Média Venda</span>
                            <span class="detail-value">
                                ${item.medVenda}
                                ${item.tendencia === 'up' ? '<span class="trend-indicator trend-up" style="margin-left:5px;">▲</span>' : 
                                  item.tendencia === 'down' ? '<span class="trend-indicator trend-down" style="margin-left:5px;">▼</span>' : ''}
                            </span>
                        </div>
                        <div class="detail-item">
                            <span class="detail-label">Custo Unit.</span>
                            <span class="detail-value">${item.custo}</span>
                        </div>
                        <div class="detail-item sparkline-detail">
                            <span class="detail-label">Histórico Vendas (Mensal)</span>
                            <div class="sparkline-container">
                                ${generateSparkline(item.historico, item.monthLabels)}
                            </div>
                        </div>
                        <div class="detail-item" style="border-top:1px solid var(--border-color); padding-top:10px; width:100%;">
                            <span class="detail-label">Mapeamento de Dados (Debug)</span>
                            <span class="detail-value" style="font-size:0.75rem; color:var(--text-muted); font-weight:normal;">
                                ${item.mappingInfo}
                            </span>
                        </div>
                    </div>
                </td>
            `;

            mainTr.addEventListener('click', () => {
                const isHidden = detailTr.classList.contains('hidden');
                // Optional: Close other expanded rows
                // document.querySelectorAll('.detail-row').forEach(dr => dr.classList.add('hidden'));
                // document.querySelectorAll('.main-row').forEach(mr => mr.classList.remove('expanded'));

                if (isHidden) {
                    detailTr.classList.remove('hidden');
                    mainTr.classList.add('expanded');
                } else {
                    detailTr.classList.add('hidden');
                    mainTr.classList.remove('expanded');
                }
            });

            tableBody.appendChild(mainTr);
            tableBody.appendChild(detailTr);
        });

        // Update stats (using current filtered data)
        const totalItems = data.length;
        const buyCount = data.filter(i => i.situacao === 'ruptura').length;
        const attentionCount = data.filter(i => i.situacao === 'atencao').length;
        const suggestCount = data.filter(i => i.situacao === 'sugestao').length;
        const ordersCount = data.filter(i => i.temEncomenda).length;

        document.getElementById('total-items').textContent = totalItems;
        document.getElementById('to-buy-count').textContent = buyCount;
        document.getElementById('attention-count').textContent = attentionCount;
        document.getElementById('suggestion-count').textContent = suggestCount;
        const ordersCountEl = document.getElementById('orders-count');
        if (ordersCountEl) ordersCountEl.textContent = ordersCount;

        // Update Charts with filtered data
        updateChart(data);
    }

    // --- File Handling ---

    /**
     * Generates an enlarged SVG sparkline for history visualization
     */
    function generateSparkline(history, labels) {
        if (!history || history.length === 0) return '<span style="color:var(--text-muted); font-size:0.7rem;">Sem histórico</span>';
        
        const width = 700; // Large width to use right space
        const height = 90; // Balanced height
        const horizontalPadding = 30;
        const topPadding = 25; // For value labels
        const bottomPadding = 20; // For month labels
        
        const max = Math.max(...history, 5);
        const drawableWidth = width - (horizontalPadding * 2);
        const drawableHeight = height - topPadding - bottomPadding;
        
        const stepX = drawableWidth / (history.length - 1 || 1);
        
        const points = history.map((val, i) => {
            const x = horizontalPadding + i * stepX;
            const y = height - bottomPadding - (val / max * drawableHeight);
            return { x, y, val, label: labels[i] };
        });

        const linePath = points.map((p, i) => (i === 0 ? 'M' : 'L') + `${p.x},${p.y}`).join(' ');
        const areaPath = linePath + ` L${points[points.length-1].x},${height - bottomPadding} L${points[0].x},${height - bottomPadding} Z`;

        return `
            <svg width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" style="overflow: visible; display: block;">
                <defs>
                    <linearGradient id="sparklineGradientDetail" x1="0%" y1="0%" x2="0%" y2="100%">
                        <stop offset="0%" style="stop-color:var(--info); stop-opacity:0.4" />
                        <stop offset="100%" style="stop-color:var(--info); stop-opacity:0.01" />
                    </linearGradient>
                </defs>
                <path d="${areaPath}" fill="url(#sparklineGradientDetail)" />
                <path d="${linePath}" fill="none" stroke="var(--info)" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" />
                ${points.map(p => `
                    <circle cx="${p.x}" cy="${p.y}" r="3.5" fill="var(--info)" class="spark-dot">
                        <title>${p.label}: ${p.val}</title>
                    </circle>
                    <text x="${p.x}" y="${p.y - 12}" text-anchor="middle" font-size="10" font-weight="bold" fill="white" style="font-family: 'Outfit', sans-serif;">
                        ${p.val}
                    </text>
                    <text x="${p.x}" y="${height - 2}" text-anchor="middle" font-size="9" fill="var(--text-muted)" style="font-family: 'Outfit', sans-serif;">
                        ${p.label}
                    </text>
                `).join('')}
            </svg>
        `;
    }

    function handleFile(file) {
        if (!file) return;

        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheet = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheet];
                
                // Use defval: null to ensure all columns exist in all objects
                const json = XLSX.utils.sheet_to_json(worksheet, { 
                    defval: null,
                    blankrows: false 
                });

                console.log('JSON Bruto carregado:', json.length, 'linhas.');
                
                if (json.length === 0) {
                    alert('A planilha parece estar vazia ou mal formatada.');
                    return;
                }

                currentData = processData(json);
                filteredData = [...currentData];
                
                showDashboard();
                renderTable(filteredData);
            } catch (err) {
                console.error('Erro ao processar arquivo:', err);
                alert('Ocorreu um erro ao ler a planilha. Verifique se o formato está correto.');
            }
        };
        reader.readAsArrayBuffer(file);
    }

    function showDashboard() {
        dropZone.style.display = 'none';
        chartSection.style.display = 'flex';
        statsSection.style.display = 'grid';
        document.getElementById('criteria-legend').style.display = 'flex';
        tableContainer.style.display = 'block';
    }

    function updateChart(data) {
        const buyCount = data.filter(i => i.situacao === 'ruptura').length;
        const attentionCount = data.filter(i => i.situacao === 'atencao').length;
        const suggestCount = data.filter(i => i.situacao === 'sugestao').length;
        const okCount = data.length - buyCount - attentionCount - suggestCount;

        const ctx = document.getElementById('purchase-pie-chart').getContext('2d');
        
        const chartData = {
            labels: ['🚨 Ruptura', '⚡ Atenção', '⚠️ Sugestão', '✅ Seguro'],
            datasets: [{
                data: [buyCount, attentionCount, suggestCount, okCount],
                backgroundColor: ['#ef4444', '#f97316', '#eab308', '#10b981'],
                borderColor: ['rgba(239, 68, 68, 0.2)', 'rgba(249, 115, 22, 0.2)', 'rgba(234, 179, 8, 0.2)', 'rgba(16, 185, 129, 0.2)'],
                borderWidth: 1,
                hoverOffset: 15
            }]
        };

        if (myChart) {
            myChart.data = chartData;
            myChart.update();
        } else {
            myChart = new Chart(ctx, {
                type: 'pie',
                data: chartData,
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    onClick: (e, elements) => {
                        if (elements.length > 0) {
                            const index = elements[0].index;
                            let filter = 'all';
                            if (index === 0) filter = 'buy';
                            else if (index === 1) filter = 'attention';
                            else if (index === 2) filter = 'suggest';
                            
                            toggleFilter(filter);
                        }
                    },
                    onHover: (event, chartElement) => {
                        event.native.target.style.cursor = chartElement[0] ? 'pointer' : 'default';
                    },
                    plugins: {
                        legend: {
                            position: 'bottom',
                            labels: {
                                color: '#f3f4f6',
                                font: { family: 'Outfit', size: 14 }
                            }
                        },
                        title: {
                            display: true,
                            text: 'Indicador de Risco (Filtrado)',
                            color: '#f3f4f6',
                            font: { family: 'Outfit', size: 18, weight: '600' }
                        },
                        datalabels: {
                            color: '#fff',
                            font: { weight: 'bold', size: 14 },
                            formatter: (value, ctx) => {
                                let sum = ctx.chart.data.datasets[0].data.reduce((acc, val) => acc + val, 0);
                                if (sum === 0) return '';
                                let percentage = (value * 100 / sum).toFixed(1) + "%";
                                return (value > 0) ? percentage : '';
                            }
                        }
                    }
                }
            });
        }

        // --- Supplier Line Chart Logic ---
        const supplierCtx = document.getElementById('supplier-risk-chart').getContext('2d');
        
        // Group and count suggestions by supplier (items in Ruptura OR Suggestion)
        const supplierStats = {};
        data.filter(i => i.situacao !== 'seguro').forEach(i => {
            supplierStats[i.fornecedor] = (supplierStats[i.fornecedor] || 0) + 1;
        });

        // Top 15 Suppliers by count
        const sortedSuppliers = Object.entries(supplierStats)
            .sort((a, b) => b[1] - a[1])
            .slice(0, 15);

        const supplierChartData = {
            labels: sortedSuppliers.map(s => s[0]),
            datasets: [{
                label: 'Sugerido/Risco',
                data: sortedSuppliers.map(s => s[1]),
                backgroundColor: 'rgba(99, 102, 241, 0.4)',
                borderColor: '#6366f1',
                borderWidth: 1,
                borderRadius: 5,
            }]
        };

        if (supplierChart) {
            supplierChart.data = supplierChartData;
            supplierChart.update();
        } else {
            supplierChart = new Chart(supplierCtx, {
                type: 'bar',
                data: supplierChartData,
                options: {
                    indexAxis: 'y',
                    responsive: true,
                    maintainAspectRatio: false,
                    onClick: (e, elements) => {
                        if (elements.length > 0) {
                            const index = elements[0].index;
                            const label = supplierChart.data.labels[index];
                            supplierSearch.value = label;
                            supplierSearch.dispatchEvent(new Event('input'));
                        }
                    },
                    onHover: (event, chartElement) => {
                        event.native.target.style.cursor = chartElement[0] ? 'pointer' : 'default';
                    },
                    plugins: {
                        legend: { display: false },
                        title: {
                            display: true,
                            text: 'Risco por Fornecedor (Filtrado)',
                            color: '#f3f4f6',
                            font: { family: 'Outfit', size: 18, weight: '600' }
                        },
                        datalabels: { display: false }
                    },
                    scales: {
                        x: {
                            beginAtZero: true,
                            grid: { color: 'rgba(255, 255, 255, 0.05)' },
                            ticks: { color: '#9ca3af' }
                        },
                        y: {
                            grid: { display: false },
                            ticks: { 
                                color: '#9ca3af',
                                font: { size: 11 }
                            }
                        }
                    }
                }
            });
        }

        // --- Group Bar Chart (Horizontal) ---
        const groupCtx = document.getElementById('group-bar-chart').getContext('2d');
        
        const groupStats = {};
        data.filter(i => i.situacao !== 'seguro').forEach(i => {
            groupStats[i.grupo] = (groupStats[i.grupo] || 0) + 1;
        });

        const sortedGroups = Object.entries(groupStats)
            .sort((a, b) => b[1] - a[1])
            .slice(0, 15);

        const groupDataDetails = {
            labels: sortedGroups.map(g => g[0]),
            datasets: [{
                label: 'Sugerido/Risco',
                data: sortedGroups.map(g => g[1]),
                backgroundColor: 'rgba(245, 158, 11, 0.4)',
                borderColor: '#f59e0b',
                borderWidth: 1,
                borderRadius: 5,
            }]
        };

        if (groupChart) {
            groupChart.data = groupDataDetails;
            groupChart.update();
        } else {
            groupChart = new Chart(groupCtx, {
                type: 'bar',
                data: groupDataDetails,
                options: {
                    indexAxis: 'y',
                    responsive: true,
                    maintainAspectRatio: false,
                    onClick: (e, elements) => {
                        if (elements.length > 0) {
                            const index = elements[0].index;
                            const label = groupChart.data.labels[index];
                            tableSearch.value = label;
                            // Trigger the input event to update everything
                            tableSearch.dispatchEvent(new Event('input'));
                        }
                    },
                    onHover: (event, chartElement) => {
                        event.native.target.style.cursor = chartElement[0] ? 'pointer' : 'default';
                    },
                    plugins: {
                        legend: { display: false },
                        title: {
                            display: true,
                            text: 'Risco por Grupo (Filtrado)',
                            color: '#f3f4f6',
                            font: { family: 'Outfit', size: 18, weight: '600' }
                        },
                        datalabels: { display: false }
                    },
                    scales: {
                        x: {
                            beginAtZero: true,
                            grid: { color: 'rgba(255, 255, 255, 0.05)' },
                            ticks: { color: '#9ca3af' }
                        },
                        y: {
                            grid: { display: false },
                            ticks: { color: '#9ca3af' }
                        }
                    }
                }
            });
        }
    }

    function applyAllFilters() {
        const prodTerm = tableSearch.value.toLowerCase();
        const suppTerm = supplierSearch.value.toLowerCase();

        filteredData = currentData.filter(i => {
            // 1. Filter by Product search (optimized to ignore symbols like dots)
            const cleanProd = i.produto.toString().replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
            const cleanTerm = prodTerm.replace(/[^a-zA-Z0-9]/g, '');
            
            const matchProd = cleanProd.includes(cleanTerm) || 
                              i.grupo.toString().toLowerCase().includes(prodTerm) ||
                              i.descricao.toString().toLowerCase().includes(prodTerm);
            
            // 2. Filter by Supplier search
            const matchSupp = i.fornecedor.toString().toLowerCase().includes(suppTerm);
            
            // 3. Filter by Buttons (Status) - Intelligent multi-select logic
            let matchStatus = true;
            if (!activeFilters.includes('all')) {
                const statusFilters = activeFilters.filter(f => ['buy', 'attention', 'suggest'].includes(f));
                const flagFilters = activeFilters.filter(f => ['has-order', 'high-rec'].includes(f));
                
                // If any status filter is selected, item must match one of them
                let passStatus = statusFilters.length === 0; // true if no status filter is selected
                if (statusFilters.length > 0) {
                    if (statusFilters.includes('buy') && i.situacao === 'ruptura') passStatus = true;
                    if (statusFilters.includes('attention') && i.situacao === 'atencao') passStatus = true;
                    if (statusFilters.includes('suggest') && i.situacao === 'sugestao') passStatus = true;
                }
                
                // Item must also match ALL selected flags (AND logic for flags)
                let passFlags = true;
                if (flagFilters.includes('has-order') && !i.temEncomenda) passFlags = false;
                if (flagFilters.includes('high-rec') && parseFloat(i.recorrencia) <= 50) passFlags = false;
                
                matchStatus = passStatus && passFlags;
            }

            return matchProd && matchSupp && matchStatus;
        });

        // 4. Sorting by Recurrence
        if (sortRecorrenciaDir === 'desc') {
            filteredData.sort((a, b) => parseFloat(b.recorrencia) - parseFloat(a.recorrencia));
        } else if (sortRecorrenciaDir === 'asc') {
            filteredData.sort((a, b) => parseFloat(a.recorrencia) - parseFloat(b.recorrencia));
        }

        renderTable(filteredData);
    }

    function toggleFilter(filter) {
        if (filter === 'all') {
            activeFilters = ['all'];
        } else {
            // Remove 'all' if it exists
            activeFilters = activeFilters.filter(f => f !== 'all');
            
            if (activeFilters.includes(filter)) {
                // Toggle off
                activeFilters = activeFilters.filter(f => f !== filter);
            } else {
                // Toggle on
                activeFilters.push(filter);
            }

            // If nothing is selected, default to 'all'
            if (activeFilters.length === 0) {
                activeFilters = ['all'];
            }
        }

        // Update UI
        filterBtns.forEach(btn => {
            const btnFilter = btn.getAttribute('data-filter');
            if (activeFilters.includes(btnFilter)) {
                btn.classList.add('active');
            } else {
                btn.classList.remove('active');
            }
        });

        applyAllFilters();
    }

    // --- Events ---

    fileUpload.addEventListener('change', (e) => {
        handleFile(e.target.files[0]);
    });

    // Drag and Drop implementation
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('drag-over');
    });

    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('drag-over');
    });

    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
        handleFile(e.dataTransfer.files[0]);
    });

    dropZone.addEventListener('click', () => {
        fileUpload.click();
    });

    // Combined search logic
    tableSearch.addEventListener('input', applyAllFilters);
    supplierSearch.addEventListener('input', applyAllFilters);

    // Sort Recorrência
    const sortRecHeader = document.getElementById('sort-recorrencia');
    if (sortRecHeader) {
        sortRecHeader.addEventListener('click', () => {
            if (sortRecorrenciaDir === 'none' || sortRecorrenciaDir === 'asc') {
                sortRecorrenciaDir = 'desc';
                sortRecHeader.querySelector('.sort-icon').textContent = '🔽';
            } else {
                sortRecorrenciaDir = 'asc';
                sortRecHeader.querySelector('.sort-icon').textContent = '🔼';
            }
            applyAllFilters();
        });
    }

    // Multi-select status filters
    filterBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            const filter = btn.getAttribute('data-filter');
            toggleFilter(filter);
        });
    });

    // Clear filters
    clearFiltersBtn.addEventListener('click', () => {
        tableSearch.value = '';
        supplierSearch.value = '';
        activeFilters = ['all'];
        filterBtns.forEach(btn => {
            const btnFilter = btn.getAttribute('data-filter');
            if (btnFilter === 'all') btn.classList.add('active');
            else btn.classList.remove('active');
        });
        applyAllFilters();
    });

    // --- Export Logic ---
    window.exportToExcel = (meses = 1) => {
        if (filteredData.length === 0) return;
        
        const todayLocale = new Date().toLocaleDateString('pt-BR');
        const fileSuffix = meses === 1 ? 'risco_ruptura_1m' : 'sugestao_compra_3m';
        const sheetName = meses === 1 ? 'Risco Ruptura (1 Mês)' : 'Sugestão Compra (3 Meses)';

        // Filter data based on the export type - now strictly exclusive
        const itemsToExport = filteredData.filter(i => {
            if (meses === 1) return i.situacao === 'ruptura';
            if (meses === 2) return i.situacao === 'atencao';
            if (meses === 3) return i.situacao === 'sugestao';
            return true;
        });

        if (itemsToExport.length === 0) {
            alert(`Nenhum item encontrado para exportação de ${meses === 1 ? 'Risco' : 'Sugestão'}.`);
            return;
        }

        // Prepare data with the specific columns requested
        const exportData = itemsToExport.map(i => {
            // Quantity strictly based on monthly average as requested (1x for Risk, 4x for Suggestion)
            const qtdSugerida = Math.ceil(parseFloat(i.medVenda) * meses);
            
            return {
                "Produto": i.produto,
                "Quantidade": qtdSugerida,
                "Centro de Custo": 1,
                "Finalidade": 1,
                "Data da Necessidade": todayLocale
            };
        });

        const worksheet = XLSX.utils.json_to_sheet(exportData);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);

        // Format column widths for better readability
        const wscols = [
            {wch: 15}, // Produto
            {wch: 40}, // Descrição
            {wch: 12}, // Quantidade
            {wch: 15}, // Centro de Custo
            {wch: 12}, // Finalidade
            {wch: 18}  // Data Necessidade
        ];
        worksheet['!cols'] = wscols;

        XLSX.writeFile(workbook, `relatorio_${fileSuffix}_${todayLocale.replace(/\//g, '-')}.xlsx`);
    };

    window.exportAttention = () => exportToExcel(2); // Helper for Attention export if needed

});
