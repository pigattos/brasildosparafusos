document.addEventListener('DOMContentLoaded', () => {
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
        return u; // Mantém original caso seja caixa, rolo, etc.
    }

    /**
     * Calculates Recurrence % (Months with sales > 0)
     */
    function calculateRecurrence(row, monthCols) {
        if (!monthCols || monthCols.length === 0) return 0;
        let monthsWithSales = 0;
        monthCols.forEach(col => {
            const val = parseFloat(row[col]) || 0;
            if (val > 0) monthsWithSales++;
        });
        return (monthsWithSales / monthCols.length) * 100;
    }

    /**
     * Processes raw JSON from Excel
     */
    function processData(rawData) {
        // Identify Month Columns (format MM/YYYY)
        const allKeys = Object.keys(rawData[0]);
        const monthCols = allKeys.filter(k => /^\d{2}\/\d{4}$/.test(k)).sort((a, b) => {
            const [mA, yA] = a.split('/');
            const [mB, yB] = b.split('/');
            return (yA + mA).localeCompare(yB + mB); // Sorts by Year then Month
        });
        
        console.log('Identifying months:', monthCols);

        return rawData.map(row => {
            const produto = row['Produto'] || row['Código'] || row['Item'] || 'N/A';
            const descricao = row['Descrição longa do produto'] || row['Descrição'] || 'Sem descrição';
            const un = normalizeUnit(row['UN'] || row['Unidade']);
            const grupo = row['Grupo'] || row['Categoria'] || 'Geral';
            const fornecedor = row['Razão social do fornecedor'] || row['Fornecedor'] || 'N/D';
            const vendas = parseFloat(row['Vendas']) || 0;
            const medVenda = parseFloat(row['Méd.Venda']) || (vendas / 6);
            const estoque = parseFloat(row['Estoque']) || 0;
            const encomendas = parseFloat(row['Encomendas']) || 0;
            const custo = parseFloat(row['Custo aquisição']) || parseFloat(row['Custo']) || 0;
            const recorrencia = calculateRecurrence(row, monthCols);
            
            // Formula 1: Stock < 1 Month Sales AND Recurrence > 33%
            const emRisco = medVenda > (estoque + encomendas) && recorrencia > 33;
            // Formula 2: Stock < 2 Months Sales
            const emAtencao = (medVenda * 2) > (estoque + encomendas);
            // Formula 3: Stock < 3 Months Sales
            const emSugestao = (medVenda * 3) > (estoque + encomendas);

            // Trend logic: compare last month with previous month
            let tendencia = 'stable';
            if (monthCols.length >= 2) {
                const lastVal = parseFloat(row[monthCols[monthCols.length - 1]]) || 0;
                const prevVal = parseFloat(row[monthCols[monthCols.length - 2]]) || 0;
                if (lastVal > prevVal) tendencia = 'up';
                else if (lastVal < prevVal) tendencia = 'down';
            }

            let situacao = 'seguro';
            if (emRisco) situacao = 'ruptura';
            else if (emAtencao) situacao = 'atencao';
            else if (emSugestao) situacao = 'sugestao';

            return {
                produto,
                descricao,
                un,
                grupo,
                fornecedor,
                estoque,
                encomendas,
                vendas,
                medVenda: medVenda.toFixed(2),
                tendencia,
                custo: custo.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' }),
                recorrencia: recorrencia.toFixed(0),
                historico: monthCols.map(col => parseFloat(row[col]) || 0),
                monthLabels: monthCols,
                situacao,
                emRisco,
                emAtencao,
                emSugestao
            };
        });
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
                    <div class="cell-value">${item.estoque}</div>
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

        document.getElementById('total-items').textContent = totalItems;
        document.getElementById('to-buy-count').textContent = buyCount;
        document.getElementById('attention-count').textContent = attentionCount;
        document.getElementById('suggestion-count').textContent = suggestCount;

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
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheet = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheet];
            const json = XLSX.utils.sheet_to_json(worksheet);

            console.log('Parsed JSON:', json);
            
            if (json.length === 0) {
                alert('A planilha parece estar vazia.');
                return;
            }

            currentData = processData(json);
            filteredData = [...currentData];
            
            showDashboard();
            renderTable(filteredData);
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
            // 1. Filter by Product search (now also searches by Group and Description)
            const matchProd = i.produto.toString().toLowerCase().includes(prodTerm) || 
                              i.grupo.toString().toLowerCase().includes(prodTerm) ||
                              i.descricao.toString().toLowerCase().includes(prodTerm);
            
            // 2. Filter by Supplier search
            const matchSupp = i.fornecedor.toString().toLowerCase().includes(suppTerm);
            
            // 3. Filter by Buttons (Status) - Multi-select logic
            let matchStatus = true;
            if (!activeFilters.includes('all')) {
                matchStatus = activeFilters.some(filter => {
                    if (filter === 'buy') return i.situacao === 'ruptura';
                    if (filter === 'attention') return i.situacao === 'atencao';
                    if (filter === 'suggest') return i.situacao === 'sugestao';
                    if (filter === 'high-rec') return parseFloat(i.recorrencia) > 50;
                    return false;
                });
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
