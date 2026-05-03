/**
 * Dashboard de Entradas Mensais
 * Desenvolvido para Brasil dos Parafusos
 */

document.addEventListener('DOMContentLoaded', () => {
    const nfUpload = document.getElementById('nf-upload');
    const loadingOverlay = document.getElementById('loading-overlay');
    const welcomeView = document.getElementById('welcome-view');
    const dashboardView = document.getElementById('dashboard-view');
    const nfSearch = document.getElementById('nf-search');
    
    let nfData = [];
    let selectedMonth = 'all'; // Formato: YYYY-MM ou 'all'
    let suppliersChart = null;
    let volumeChart = null;
    let averageChart = null;
    let dailyChart = null;
    let monthlyChart = null;
    let showAllSuppliers = false; // State for Top 10 vs All toggle

    // --- Event Listeners ---
    if (nfUpload) nfUpload.addEventListener('change', handleFileUpload);
    if (nfSearch) nfSearch.addEventListener('input', () => updateDashboard());

    const exportBtn = document.getElementById('export-entradas-btn');
    if (exportBtn) exportBtn.addEventListener('click', exportToExcel);

    const folderUpload = document.getElementById('folder-upload');
    if (folderUpload) folderUpload.addEventListener('change', handleFolderUpload);

    const toggleSuppliersBtn = document.getElementById('toggle-suppliers-view');
    if (toggleSuppliersBtn) {
        toggleSuppliersBtn.addEventListener('click', () => {
            showAllSuppliers = !showAllSuppliers;
            toggleSuppliersBtn.textContent = showAllSuppliers ? 'Ver Top 10' : 'Ver Todos';
            toggleSuppliersBtn.classList.toggle('active', showAllSuppliers);
            updateDashboard();
        });
    }

    // --- Auto Load Logic ---
    if (typeof PRE_LOADED_ENTRADAS !== 'undefined' && PRE_LOADED_ENTRADAS.length > 0) {
        nfData = PRE_LOADED_ENTRADAS.map(item => {
            const dateObj = parseNFDate(item.date);
            return {
                ...item,
                date: formatDate(dateObj),
                rawDate: dateObj
            };
        }).filter(item => !isNaN(item.rawDate.getTime()));
        
        if (welcomeView) welcomeView.style.display = 'none';
        if (dashboardView) dashboardView.style.display = 'block';
        setTimeout(() => updateDashboard(), 100);
    }

    // ----------------------

    // --- Logic Functions ---

    async function handleFolderUpload(e) {
        const files = Array.from(e.target.files);
        if (!files || files.length === 0) return;

        showLoading(true);

        try {
            let allData = [];
            for (let i = 0; i < files.length; i++) {
                const file = files[i];
                // Filtra apenas arquivos excel/csv e ignora arquivos temporários
                if (file.name.match(/\.(xlsx|xls|csv)$/i) && !file.name.startsWith('~$')) {
                    try {
                        const data = await readExcel(file);
                        const processed = processIndividualFile(data);
                        allData = allData.concat(processed);
                    } catch (err) {
                        console.warn(`Erro ao ler o arquivo ${file.name}:`, err);
                    }
                }
            }

            if (allData.length > 0) {
                nfData = allData;
                welcomeView.style.display = 'none';
                dashboardView.style.display = 'block';
                updateDashboard();
            } else {
                alert('Nenhum dado válido encontrado nos arquivos da pasta.');
            }
        } catch (error) {
            console.error('Erro ao processar pasta:', error);
            alert('Erro ao processar a pasta.');
        } finally {
            showLoading(false);
            e.target.value = ''; // Reset input to allow re-upload
        }
    }

    function processIndividualFile(rows) {
        if (!rows || rows.length < 2) return [];

        // Encontrar a linha de cabeçalho
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

        const colDate = findIdx(['dt.movto', 'data', 'movimento']);
        const colCode = findIdx(['codigo', 'cod']);
        const colDesc = findIdx(['descricao', 'item', 'produto']);
        const colQty = findIdx(['quantidade', 'qtd', 'unidades']);
        const colNF = findIdx(['nota fiscal', 'nf', 'doc']);
        const colPurpose = findIdx(['finalid.item ordem com', 'finalidade']);
        const colValue = findIdx(['vlr.cont.p/sped', 'vlr.cont', 'valor', 'total']);
        const colSupplier = findIdx(['razao social', 'fornecedor', 'nome']);
        const colGroup = 11; // Coluna L fixa

        return rows.slice(headerIndex + 1).map(row => {
            if (!row || row.length === 0) return null;
            if (!row[colValue] && !row[colQty] && !row[colCode]) return null;

            const rawDateValue = row[colDate];
            const dateObj = parseExcelDate(rawDateValue);

            return {
                date: formatDate(dateObj),
                rawDate: dateObj,
                code: String(row[colCode] || '').trim(),
                description: String(row[colDesc] || '').trim(),
                quantity: parseFloat(row[colQty]) || 0,
                group: String(row[colGroup] || 'DIVERSOS').trim(),
                invoice: String(row[colNF] || '').trim(),
                purpose: String(row[colPurpose] || '').trim(),
                supplier: String(row[colSupplier] || 'NÃO IDENTIFICADO').trim(),
                value: parseFloat(row[colValue]) || 0
            };
        }).filter(item => item && (item.quantity !== 0 || item.value !== 0) && !isNaN(item.rawDate.getTime()));
    }

    async function handleFileUpload(e) {
        const file = e.target.files[0];
        if (!file) return;

        showLoading(true);

        try {
            const data = await readExcel(file);
            processData(data);
            
            welcomeView.style.display = 'none';
            dashboardView.style.display = 'block';
            
            updateDashboard();
        } catch (error) {
            console.error('Erro ao processar arquivo:', error);
            alert('Erro ao ler o arquivo Excel. Verifique se as colunas "Dt.movto", "Razão social" e "Vlr.cont" estão presentes.');
        } finally {
            showLoading(false);
        }
    }

    function readExcel(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array', cellDates: true });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                resolve(json);
            };
            reader.onerror = reject;
            reader.readAsArrayBuffer(file);
        });
    }

    function processData(data) {
        nfData = processIndividualFile(data);
    }

    function findColumn(row, possibilities) {
        const keys = Object.keys(row);
        for (const p of possibilities) {
            const match = keys.find(k => k.trim().toLowerCase() === p.toLowerCase());
            if (match) return match;
        }
        // Fallback: primeira coluna que contém parte do nome
        return keys.find(k => possibilities.some(p => k.toLowerCase().includes(p.toLowerCase()))) || null;
    }

    function parseNFDate(val) {
        if (!val) return new Date(NaN);
        if (val instanceof Date) return val;
        
        // Se for string YYYY-MM-DD (formato do bridge)
        if (typeof val === 'string' && val.includes('-')) {
            const parts = val.split('-');
            if (parts.length === 3 && parts[0].length === 4) {
                return new Date(parts[0], parts[1] - 1, parts[2]);
            }
        }
        
        // Se for string DD/MM/YYYY
        if (typeof val === 'string' && val.includes('/')) {
            const parts = val.split('/');
            if (parts.length === 3) {
                return new Date(parts[2], parts[1] - 1, parts[0]);
            }
        }
        
        return new Date(val);
    }

    function parseExcelDate(val) {
        if (val instanceof Date) return val;
        if (typeof val === 'number') {
            // Excel serial date
            return new Date(Math.round((val - 25569) * 86400 * 1000));
        }
        if (typeof val === 'string') {
            // Tentar converter string DD/MM/YYYY
            const parts = val.split(/[\/\-]/);
            if (parts.length === 3) {
                if (parts[0].length === 4) return new Date(val); // YYYY-MM-DD
                return new Date(`${parts[2]}-${parts[1]}-${parts[0]}`); // DD-MM-YYYY
            }
        }
        return new Date(val);
    }

    function updateDashboard() {
        if (!nfData || nfData.length === 0) return;
        
        const searchTerm = nfSearch.value.toLowerCase();
        const filteredBySearch = nfData.filter(item => 
            item.description.toLowerCase().includes(searchTerm) ||
            item.code.toLowerCase().includes(searchTerm) ||
            item.group.toLowerCase().includes(searchTerm) ||
            item.supplier.toLowerCase().includes(searchTerm) ||
            item.invoice.toLowerCase().includes(searchTerm)
        );
        const filtered = filterByPeriodAndSearch();
        
        // 1. Stats
        const totalValue = filtered.reduce((acc, item) => acc + item.value, 0);
        const count = filtered.length;

        const valReposicao = filtered.reduce((acc, item) => {
            const p = (item.purpose || '').toLowerCase();
            return (p.includes('reposicao') || p.includes('estoque')) ? acc + item.value : acc;
        }, 0);

        const valVendaCasada = filtered.reduce((acc, item) => {
            const p = (item.purpose || '').toLowerCase();
            return (p.includes('venda') && p.includes('casada')) ? acc + item.value : acc;
        }, 0);

        document.getElementById('total-received').textContent = formatCurrency(totalValue);
        document.getElementById('nf-count').textContent = `${count} registros`;
        document.getElementById('qty-reposicao').textContent = formatCurrency(valReposicao);
        document.getElementById('qty-venda-casada').textContent = formatCurrency(valVendaCasada);

        // 2. Trend Logic: Indicators only for specific months
        let comparisonMonth = selectedMonth;

        let prevSupplierLookup = {};
        let currentSupplierLookup = {};
        let prevGroupLookup = {};
        let currentGroupLookup = {};
        
        let prevTotals = { total: 0, reposicao: 0, vendaCasada: 0 };

        if (comparisonMonth !== 'all') {
            const [yr, mo] = comparisonMonth.split('-').map(Number);
            const prevMo = mo === 1 ? 12 : mo - 1;
            const prevYr = mo === 1 ? yr - 1 : yr;
            const prevKey = `${prevYr}-${String(prevMo).padStart(2, '0')}`;
            
            // Current month data (for trends when in "All" view)
            const currData = nfData.filter(item => {
                const d = item.rawDate;
                return `${d.getFullYear()}-${(d.getMonth()+1).toString().padStart(2,'0')}` === comparisonMonth;
            });
            const currSuppliers = aggregateBySupplier(currData);
            currentSupplierLookup = Object.fromEntries(currSuppliers.map(s => [s.supplier.toLowerCase().trim(), s.total]));
            const currGroups = aggregateByGroup(currData);
            currentGroupLookup = Object.fromEntries(currGroups.map(g => [g.group.toLowerCase().trim(), g.total]));

            // Previous month data
            const prevData = nfData.filter(item => {
                const d = item.rawDate;
                return `${d.getFullYear()}-${(d.getMonth()+1).toString().padStart(2,'0')}` === prevKey;
            });
            
            // Calculate totals for cards
            prevTotals.total = prevData.reduce((acc, item) => acc + item.value, 0);
            prevTotals.reposicao = prevData.reduce((acc, item) => {
                const p = (item.purpose || '').toLowerCase();
                return (p.includes('reposicao') || p.includes('estoque')) ? acc + item.value : acc;
            }, 0);
            prevTotals.vendaCasada = prevData.reduce((acc, item) => {
                const p = (item.purpose || '').toLowerCase();
                return (p.includes('venda') && p.includes('casada')) ? acc + item.value : acc;
            }, 0);

            const prevSuppliers = aggregateBySupplier(prevData);
            prevSupplierLookup = Object.fromEntries(prevSuppliers.map(s => [s.supplier.toLowerCase().trim(), s.total]));
            const prevGroups = aggregateByGroup(prevData);
            prevGroupLookup = Object.fromEntries(prevGroups.map(g => [g.group.toLowerCase().trim(), g.total]));
            
            // Update Card Badges
            renderCardTrend('trend-total', totalValue, prevTotals.total);
            renderCardTrend('trend-reposicao', valReposicao, prevTotals.reposicao);
            renderCardTrend('trend-venda-casada', valVendaCasada, prevTotals.vendaCasada);
        } else {
            // Hide trends if no comparison
            ['trend-total', 'trend-reposicao', 'trend-venda-casada'].forEach(id => {
                document.getElementById(id).style.display = 'none';
            });
        }
        
        // Define supplierTotals for the ranking chart
        const supplierTotals = aggregateBySupplier(filteredBySearch);
        renderSuppliersChart(supplierTotals.slice(0, 10), currentSupplierLookup, prevSupplierLookup);

        // 3. Aggregate groups for charts
        const groupTotals = aggregateByGroup(filtered);

        // 4. Charts
        renderVolumeChart(groupTotals.slice(0, 20), currentGroupLookup, prevGroupLookup);
        renderDailyChart(aggregateByDate(filtered));
        renderMonthlyChart(aggregateByMonth(filteredBySearch)); // Monthly chart should keep full history


        // 4. Table
        renderTable(filteredBySearch);
        
        // 5. Update Timeline
        renderTimeline();
    }

    function filterByPeriodAndSearch() {
        const searchTerm = nfSearch.value.toLowerCase();
        return nfData.filter(item => {
            const matchesSearch = 
                item.description.toLowerCase().includes(searchTerm) ||
                item.code.toLowerCase().includes(searchTerm) ||
                item.group.toLowerCase().includes(searchTerm) ||
                item.supplier.toLowerCase().includes(searchTerm) ||
                item.invoice.toLowerCase().includes(searchTerm);
            
            // Extrair YYYY-MM do item
            const d = item.rawDate;
            const itemMonth = `${d.getFullYear()}-${(d.getMonth() + 1).toString().padStart(2, '0')}`;
            
            const matchesPeriod = selectedMonth === 'all' || itemMonth === selectedMonth;
            return matchesSearch && matchesPeriod;
        });
    }

    function renderTimeline() {
        const timelineContainer = document.getElementById('timeline-container');
        if (!timelineContainer) return;

        // Extrair meses únicos
        const monthsMap = new Map();
        nfData.forEach(item => {
            const d = item.rawDate;
            if (!d || isNaN(d.getTime())) return;
            
            const year = d.getFullYear();
            const month = d.getMonth() + 1;
            const key = `${year}-${month.toString().padStart(2, '0')}`;
            
            if (!monthsMap.has(key)) {
                const monthName = d.toLocaleString('pt-br', { month: 'long' });
                monthsMap.set(key, { key, monthName, year });
            }
        });

        const sortedMonths = Array.from(monthsMap.values()).sort((a, b) => b.key.localeCompare(a.key));

        let html = `
            <div class="timeline-item ${selectedMonth === 'all' ? 'active' : ''}" data-period="all">
                <span class="timeline-month">Tudo</span>
                <span class="timeline-year">Histórico Total</span>
            </div>
        `;

        sortedMonths.forEach(m => {
            html += `
                <div class="timeline-item ${selectedMonth === m.key ? 'active' : ''}" data-period="${m.key}">
                    <span class="timeline-month">${m.monthName}</span>
                    <span class="timeline-year">${m.year}</span>
                </div>
            `;
        });

        timelineContainer.innerHTML = html;

        // Adicionar eventos
        timelineContainer.querySelectorAll('.timeline-item').forEach(item => {
            item.addEventListener('click', () => {
                selectedMonth = item.getAttribute('data-period');
                updateDashboard();
            });
        });
    }

    function aggregateBySupplier(data) {
        const map = {};
        data.forEach(item => {
            const supplier = item.supplier || 'N/A';
            if (!map[supplier]) {
                map[supplier] = { 
                    supplier, 
                    total: 0, 
                    reposicao: 0, 
                    vendaCasada: 0 
                };
            }
            map[supplier].total += item.value;
            
            const p = (item.purpose || '').toLowerCase();
            if (p.includes('reposicao') || p.includes('estoque')) {
                map[supplier].reposicao += item.value;
            } else if (p.includes('venda') && p.includes('casada')) {
                map[supplier].vendaCasada += item.value;
            }
        });

        return Object.values(map).sort((a, b) => b.total - a.total);
    }

    function aggregateByGroup(data) {
        const stats = {};
        data.forEach(item => {
            const key = item.group || 'DIVERSOS';
            if (!stats[key]) {
                stats[key] = { 
                    total: 0, 
                    count: 0,
                    reposicao: 0,
                    vendaCasada: 0
                };
            }
            stats[key].total += item.value;
            stats[key].count += item.quantity;

            const p = (item.purpose || '').toLowerCase();
            if (p.includes('reposicao') || p.includes('estoque')) {
                stats[key].reposicao += item.value;
            } else if (p.includes('venda') && p.includes('casada')) {
                stats[key].vendaCasada += item.value;
            }
        });
        return Object.keys(stats).map(key => ({
            group: key,
            total: stats[key].total,
            count: stats[key].count,
            reposicao: stats[key].reposicao,
            vendaCasada: stats[key].vendaCasada
        })).sort((a, b) => b.total - a.total);
    }

    function aggregateByPurpose(data) {
        const stats = {};
        data.forEach(item => {
            const key = item.purpose || 'NÃO INFORMADO';
            if (!stats[key]) stats[key] = { total: 0, count: 0 };
            stats[key].total += item.value;
            stats[key].count += item.quantity;
        });
        return Object.keys(stats).map(key => ({
            supplier: key, // Label para o gráfico
            total: stats[key].total,
            count: stats[key].count,
            average: stats[key].total
        })).sort((a, b) => b.total - a.total);
    }

    function aggregateByDate(data) {
        const map = {};
        data.forEach(item => {
            const dateStr = item.date;
            if (!map[dateStr]) map[dateStr] = { total: 0, count: 0, raw: item.rawDate };
            map[dateStr].total += item.value;
            map[dateStr].count += 1;
        });

        // Ordenar por data
        return Object.entries(map)
            .map(([date, stats]) => ({ 
                date, 
                total: stats.total, 
                count: stats.count,
                raw: stats.raw 
            }))
            .sort((a, b) => a.raw - b.raw);
    }
    
    function aggregateByMonth(data) {
        const map = {};
        data.forEach(item => {
            const d = item.rawDate;
            if (!d || isNaN(d.getTime())) return;
            
            const monthKey = `${d.getFullYear()}-${(d.getMonth() + 1).toString().padStart(2, '0')}`;
            if (!map[monthKey]) {
                map[monthKey] = { 
                    total: 0, 
                    reposicao: 0,
                    vendaCasada: 0,
                    label: d.toLocaleString('pt-br', { month: 'short', year: '2-digit' }),
                    key: monthKey
                };
            }
            map[monthKey].total += item.value;
            
            const p = (item.purpose || '').toLowerCase();
            if (p.includes('reposicao') || p.includes('estoque')) {
                map[monthKey].reposicao += item.value;
            } else if (p.includes('venda') && p.includes('casada')) {
                map[monthKey].vendaCasada += item.value;
            }
        });

        // Ordenar por chave de mês
        return Object.values(map).sort((a, b) => a.key.localeCompare(b.key));
    }

    // --- UI Rendering ---

    function renderTable(data) {
        const tbody = document.getElementById('entradas-table-body');
        tbody.innerHTML = '';

        data.forEach(item => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${item.date}</td>
                <td>${item.code}</td>
                <td title="${item.description}">${item.description.length > 30 ? item.description.substring(0, 30) + '...' : item.description}</td>
                <td style="text-align: center;">${item.quantity.toLocaleString('pt-br')}</td>
                <td>${item.group}</td>
                <td>${item.invoice}</td>
                <td>${item.purpose}</td>
                <td style="text-align: right; font-weight: 500;">${formatCurrency(item.value)}</td>
            `;
            tbody.appendChild(tr);
        });
    }

    function filterData() {
        const query = nfSearch.value.toLowerCase();
        return nfData.filter(item => 
            item.supplier.toLowerCase().includes(query) || 
            item.date.includes(query)
        );
    }

    function renderSuppliersChart(items, currentLookup, prevLookup) {
        const ctx = document.getElementById('suppliers-chart').getContext('2d');
        if (suppliersChart) suppliersChart.destroy();

        suppliersChart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: items.map(d => d.supplier),
                datasets: [
                    {
                        label: 'Reposição de Estoque',
                        data: items.map(d => d.reposicao),
                        backgroundColor: 'rgba(99, 102, 241, 0.85)',
                        borderColor: '#6366f1',
                        borderWidth: 1,
                        borderRadius: 4,
                        stack: 'stack1'
                    },
                    {
                        label: 'Venda Casada',
                        data: items.map(d => d.vendaCasada),
                        backgroundColor: 'rgba(16, 185, 129, 0.85)',
                        borderColor: '#10b981',
                        borderWidth: 1,
                        borderRadius: 4,
                        stack: 'stack1'
                    },
                    {
                        label: 'Outros',
                        data: items.map(d => Math.max(0, d.total - d.reposicao - d.vendaCasada)),
                        backgroundColor: 'rgba(148, 163, 184, 0.3)',
                        borderColor: 'rgba(148, 163, 184, 0.5)',
                        borderWidth: 1,
                        borderRadius: 4,
                        stack: 'stack1'
                    }
                ]
            },
            plugins: [
                ChartDataLabels,
                makeHorizPercentPlugin(items, 'supplier', currentLookup, prevLookup, '#e2e8f0', 'rgba(99,102,241,0.1)', 'rgba(99,102,241,0.3)', 'rgba(99,102,241,0.15)')
            ],
            options: {
                indexAxis: 'y',
                responsive: true,
                maintainAspectRatio: false,
                layout: { padding: { right: 120, top: 10 } },
                plugins: {
                    legend: {
                        display: true,
                        position: 'top',
                        labels: { color: '#e2e8f0', font: { family: 'Outfit', size: 10 } }
                    },
                    datalabels: {
                        display: (context) => {
                            const val = context.dataset.data[context.dataIndex];
                            const total = context.chart.data.datasets.reduce((acc, ds) => acc + (ds.data[context.dataIndex] || 0), 0);
                            return total > 0 && (val / total) > 0.2; // Mostra % se for > 20% do fornecedor
                        },
                        color: '#fff',
                        font: { weight: 'bold', size: 9 },
                        formatter: (value, context) => {
                            const total = context.chart.data.datasets.reduce((acc, ds) => acc + (ds.data[context.dataIndex] || 0), 0);
                            const pct = total > 0 ? ((value / total) * 100).toFixed(0) : 0;
                            return pct + '%';
                        }
                    },
                    tooltip: {
                        mode: 'index',
                        intersect: false,
                        backgroundColor: 'rgba(15, 23, 42, 0.95)',
                        padding: 12,
                        callbacks: {
                            label: function(context) {
                                const val = context.raw;
                                const total = context.chart.data.datasets.reduce((acc, ds) => acc + (ds.data[context.dataIndex] || 0), 0);
                                const pct = total > 0 ? ((val / total) * 100).toFixed(1) : 0;
                                return `${context.dataset.label}: ${formatCurrency(val)} (${pct}%)`;
                            },
                            footer: (items) => {
                                const total = items.reduce((acc, it) => acc + it.raw, 0);
                                return `TOTAL: ${formatCurrency(total)}`;
                            }
                        }
                    }
                },
                scales: {
                    x: {
                        stacked: true,
                        grid: { color: 'rgba(255,255,255,0.05)' },
                        ticks: { color: '#94a3b8', callback: (v) => formatCurrency(v) }
                    },
                    y: {
                        stacked: true,
                        grid: { display: false },
                        ticks: { 
                            color: '#e2e8f0', 
                            font: { family: 'Outfit', size: 10 },
                            callback: function(value) {
                                const label = this.getLabelForValue(value);
                                return label.length > 15 ? label.substr(0, 15) + '...' : label;
                            }
                        }
                    }
                }
            }
        });
    }

    function makeHorizPercentPlugin(items, valueKey, prevLookup,
        neutralColor, neutralFill, neutralStroke, neutralGlow) {

        const PILL_BG = '#0d1526'; 
        const total = items.reduce((s, d) => s + (d[valueKey] || 0), 0);

        return {
            id: 'horizPercentBadge',
            afterDatasetsDraw(chart) {
                const { ctx: c } = chart;
                const meta = chart.getDatasetMeta(0);

                items.forEach((d, i) => {
                    const barEl = meta.data[i];
                    if (!barEl) return;

                    const pct = total > 0 ? (d[valueKey] / total) * 100 : 0;
                    if (pct < 0.1) return;

                    let arrow = '';
                    let txtColor  = neutralColor;
                    let bdColor   = neutralStroke;
                    let glowColor = neutralGlow;

                    if (prevLookup) {
                        const prevVal = prevLookup[d.supplier.toLowerCase()];
                        if (prevVal !== undefined) {
                            const curr = d[valueKey];
                            if (curr > prevVal) {
                                arrow    = '↑';
                                txtColor  = '#34d399';
                                bdColor   = 'rgba(52,211,153,0.45)';
                                glowColor = 'rgba(52,211,153,0.25)';
                            } else if (curr < prevVal) {
                                arrow    = '↓';
                                txtColor  = '#fb7185';
                                bdColor   = 'rgba(251,113,133,0.45)';
                                glowColor = 'rgba(251,113,133,0.25)';
                            }
                        }
                    }

                    const text = arrow ? `${arrow} ${pct.toFixed(1)}%` : `${pct.toFixed(1)}%`;

                    const barRight   = barEl.x;
                    const barCenterY = barEl.y;
                    const barHalfH   = (barEl.height || 28) / 2;

                    c.save();
                    c.font = '600 10px Outfit, system-ui, sans-serif';

                    const tw    = c.measureText(text).width;
                    const padX  = 9;
                    const pillW = tw + padX * 2;
                    const pillH = 18;
                    const pillR = pillH / 2;
                    const pillX = barRight + 10; // Position to the right of the bar
                    const pillY = barCenterY - pillH / 2;

                    c.shadowColor   = glowColor;
                    c.shadowBlur    = 8;
                    c.beginPath();
                    c.roundRect(pillX, pillY, pillW, pillH, pillR);
                    c.fillStyle = PILL_BG;
                    c.fill();

                    c.shadowBlur  = 0;
                    c.strokeStyle = bdColor;
                    c.lineWidth   = 1;
                    c.stroke();

                    c.fillStyle    = txtColor;
                    c.textAlign    = 'center';
                    c.textBaseline = 'middle';
                    c.fillText(text, pillX + pillW / 2, pillY + pillH / 2);

                    c.restore();
                });
            }
        };
    }
    function renderVolumeChart(topData, currentLookup, prevLookup) {
        const ctx = document.getElementById('volume-chart').getContext('2d');
        if (volumeChart) volumeChart.destroy();

        // Ordenar por valor (total) para este gráfico
        const volumeData = [...topData].sort((a, b) => b.total - a.total);

        volumeChart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: volumeData.map(d => d.group),
                datasets: [
                    {
                        label: 'Reposição de Estoque',
                        data: volumeData.map(d => d.reposicao),
                        backgroundColor: 'rgba(99, 102, 241, 0.85)',
                        borderColor: '#6366f1',
                        borderWidth: 1,
                        borderRadius: 4,
                        stack: 'stack1'
                    },
                    {
                        label: 'Venda Casada',
                        data: volumeData.map(d => d.vendaCasada),
                        backgroundColor: 'rgba(16, 185, 129, 0.85)',
                        borderColor: '#10b981',
                        borderWidth: 1,
                        borderRadius: 4,
                        stack: 'stack1'
                    },
                    {
                        label: 'Outros',
                        data: volumeData.map(d => Math.max(0, d.total - d.reposicao - d.vendaCasada)),
                        backgroundColor: 'rgba(148, 163, 184, 0.3)',
                        borderColor: 'rgba(148, 163, 184, 0.5)',
                        borderWidth: 1,
                        borderRadius: 4,
                        stack: 'stack1'
                    }
                ]
            },
            plugins: [
                ChartDataLabels,
                makeHorizPercentPlugin(volumeData, 'group', currentLookup, prevLookup, '#e2e8f0', 'rgba(99,102,241,0.1)', 'rgba(99,102,241,0.3)', 'rgba(99,102,241,0.15)')
            ],
            options: {
                indexAxis: 'y',
                responsive: true,
                maintainAspectRatio: false,
                layout: { padding: { right: 120, top: 10 } },
                plugins: {
                    legend: {
                        display: true,
                        position: 'top',
                        labels: { color: '#e2e8f0', font: { family: 'Outfit', size: 10 } }
                    },
                    datalabels: {
                        display: (context) => {
                            const val = context.dataset.data[context.dataIndex];
                            const total = context.chart.data.datasets.reduce((acc, ds) => acc + (ds.data[context.dataIndex] || 0), 0);
                            return total > 0 && (val / total) > 0.15; // Mostra % se for > 15% do grupo
                        },
                        color: '#fff',
                        font: { weight: 'bold', size: 9 },
                        formatter: (value, context) => {
                            const total = context.chart.data.datasets.reduce((acc, ds) => acc + (ds.data[context.dataIndex] || 0), 0);
                            const pct = total > 0 ? ((value / total) * 100).toFixed(0) : 0;
                            return pct + '%';
                        }
                    },
                    tooltip: {
                        mode: 'index',
                        intersect: false,
                        backgroundColor: 'rgba(15, 23, 42, 0.95)',
                        padding: 12,
                        callbacks: {
                            label: function(context) {
                                const val = context.raw;
                                const total = context.chart.data.datasets.reduce((acc, ds) => acc + (ds.data[context.dataIndex] || 0), 0);
                                const pct = total > 0 ? ((val / total) * 100).toFixed(1) : 0;
                                return `${context.dataset.label}: ${formatCurrency(val)} (${pct}%)`;
                            },
                            footer: (items) => {
                                const total = items.reduce((acc, it) => acc + it.raw, 0);
                                return `TOTAL: ${formatCurrency(total)}`;
                            }
                        }
                    }
                },
                scales: {
                    x: { 
                        stacked: true,
                        display: true,
                        grid: { color: 'rgba(255,255,255,0.05)' }, 
                        ticks: { color: '#94a3b8', callback: (v) => formatCurrency(v) }
                    },
                    y: { 
                        stacked: true,
                        grid: { display: false }, 
                        ticks: { 
                            color: '#e2e8f0', 
                            font: { size: 10 },
                            callback: function(value) {
                                const label = this.getLabelForValue(value);
                                return label.length > 20 ? label.substr(0, 20) + '...' : label;
                            }
                        } 
                    }
                }
            }
        });
    }


    function renderDailyChart(dailyData) {
        const ctx = document.getElementById('daily-chart').getContext('2d');
        
        if (dailyChart) dailyChart.destroy();

        const gradient = ctx.createLinearGradient(0, 0, 0, 400);
        gradient.addColorStop(0, 'rgba(16, 185, 129, 0.3)');
        gradient.addColorStop(1, 'rgba(16, 185, 129, 0)');

        dailyChart = new Chart(ctx, {
            type: 'line',
            data: {
                labels: dailyData.map(d => d.date),
                datasets: [{
                    label: 'Valor Total Recebido (R$)',
                    data: dailyData.map(d => d.total),
                    borderColor: '#10b981',
                    borderWidth: 3,
                    backgroundColor: gradient,
                    fill: true,
                    tension: 0.4,
                    pointRadius: 6,
                    pointHoverRadius: 9,
                    pointBackgroundColor: '#10b981',
                    pointBorderColor: '#fff',
                    pointBorderWidth: 2
                }]
            },
            plugins: [ChartDataLabels],
            options: {
                responsive: true,
                maintainAspectRatio: false,
                onClick: (e, elements) => {
                    if (elements.length > 0) {
                        const index = elements[0].index;
                        const date = dailyData[index].date;
                        nfSearch.value = date; // Filtra pela data (formatada DD/MM/YYYY)
                        updateDashboard();
                    } else {
                        if (nfSearch.value !== "") {
                            nfSearch.value = "";
                            updateDashboard();
                        }
                    }
                },
                onHover: (event, chartElement) => {
                    event.native.target.style.cursor = chartElement[0] ? 'pointer' : 'default';
                },
                plugins: {
                    legend: { display: false },
                    datalabels: {
                        anchor: 'end',
                        align: 'top',
                        color: '#10b981',
                        offset: 4,
                        font: { weight: 'bold', size: 9, family: 'Outfit' },
                        formatter: (value) => formatCurrency(value),
                        display: 'auto' // Evita sobreposição
                    },
                    tooltip: {
                        backgroundColor: 'rgba(15, 23, 42, 0.9)',
                        padding: 12,
                        cornerRadius: 8,
                        callbacks: {
                            title: (items) => `Data: ${items[0].label}`,
                            label: (context) => ` Total: ${formatCurrency(context.raw)}`
                        }
                    }
                },
                scales: {
                    x: { 
                        grid: { display: false },
                        ticks: { color: '#94a3b8' }
                    },
                    y: { 
                        grid: { color: 'rgba(255,255,255,0.05)' },
                        beginAtZero: true,
                        ticks: { 
                            color: '#94a3b8',
                            callback: (value) => formatCurrency(value)
                        }
                    }
                }
            }
        });
    }

    // --- Helpers ---

    function formatDate(val) {
        if (!val) return 'N/A';
        const d = val instanceof Date ? val : new Date(val);
        if (isNaN(d.getTime())) return val.toString();
        
        const day = d.getDate().toString().padStart(2, '0');
        const month = (d.getMonth() + 1).toString().padStart(2, '0');
        const year = d.getFullYear();
        return `${day}/${month}/${year}`;
    }

    function formatCurrency(val) {
        return new Intl.NumberFormat('pt-BR', {
            style: 'currency',
            currency: 'BRL'
        }).format(val);
    }

    function showLoading(show) {
        loadingOverlay.style.display = show ? 'flex' : 'none';
    }

    /**
     * Plugin de badges premium para gráficos horizontais.
     * Exibe a variação percentual (crescimento/queda) em relação ao mês anterior.
     *
     * @param {Array}       items         - Dados do gráfico
     * @param {string}      labelKey      - Chave do label ('supplier'|'group')
     * @param {Object}      currentLookup - { [key]: currentValue }
     * @param {Object}      prevLookup    - { [key]: prevValue }
     */
    function makeHorizPercentPlugin(items, labelKey, currentLookup, prevLookup,
        neutralColor, neutralFill, neutralStroke, neutralGlow) {

        const PILL_BG = '#0d1526'; 

        return {
            id: 'horizPercentBadge',
            afterDatasetsDraw(chart) {
                const { ctx: c } = chart;
                const meta = chart.getDatasetMeta(0);

                items.forEach((d, i) => {
                    const barEl = meta.data[i];
                    if (!barEl) return;

                    const key = d[labelKey].toLowerCase().trim();
                    const curr = currentLookup ? (currentLookup[key] || 0) : 0;
                    const prev = prevLookup ? (prevLookup[key] || 0) : 0;

                    let growth = 0;
                    let text = '';
                    let txtColor  = neutralColor;
                    let bdColor   = neutralStroke;
                    let glowColor = neutralGlow;

                    if (prev > 0) {
                        growth = ((curr - prev) / prev) * 100;
                        if (growth > 0) {
                            text = `+${growth.toFixed(1)}%`;
                            txtColor  = '#34d399';
                            bdColor   = 'rgba(52,211,153,0.45)';
                            glowColor = 'rgba(52,211,153,0.25)';
                        } else if (growth < 0) {
                            text = `${growth.toFixed(1)}%`;
                            txtColor  = '#fb7185';
                            bdColor   = 'rgba(251,113,133,0.45)';
                            glowColor = 'rgba(251,113,133,0.25)';
                        } else {
                            text = '0.0%';
                        }
                    } else if (curr > 0) {
                        text = 'Novo';
                        txtColor  = '#60a5fa';
                        bdColor   = 'rgba(96,165,250,0.45)';
                        glowColor = 'rgba(96,165,250,0.25)';
                    }

                    if (!text) return;

                    const barRight   = barEl.x;
                    const barCenterY = barEl.y;

                    c.save();
                    c.font = '700 11px Outfit, system-ui, sans-serif';

                    const tw    = c.measureText(text).width;
                    const padX  = 10;
                    const pillW = tw + padX * 2;
                    const pillH = 20;
                    const pillR = pillH / 2;
                    const pillX = barRight + 12; // Mover para fora da barra
                    const pillY = barCenterY - pillH / 2; // Centralizar verticalmente

                    // Subtle glow
                    c.shadowColor   = glowColor;
                    c.shadowBlur    = 12;
                    c.shadowOffsetX = 0;
                    c.shadowOffsetY = 0;

                    // Solid dark fill — crisp against bars
                    c.beginPath();
                    c.roundRect(pillX, pillY, pillW, pillH, pillR);
                    c.fillStyle = PILL_BG;
                    c.fill();

                    // Thin colored border
                    c.shadowBlur  = 0;
                    c.strokeStyle = bdColor;
                    c.lineWidth   = 1.5;
                    c.stroke();

                    // Text
                    c.fillStyle    = txtColor;
                    c.textAlign    = 'center';
                    c.textBaseline = 'middle';
                    c.fillText(text, pillX + pillW / 2, pillY + pillH / 2);

                    c.restore();
                });
            }
        };
    }

    function calcMonthlyVariation(monthlyData) {
        return monthlyData.map((d, i) => {
            if (i === 0) return null;
            const prev = monthlyData[i - 1].total;
            if (!prev || prev === 0) return null;
            return ((d.total - prev) / prev) * 100;
        });
    }

    function renderMonthlyChart(monthlyData) {
        const ctx = document.getElementById('monthly-chart').getContext('2d');
        if (monthlyChart) monthlyChart.destroy();

        monthlyChart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: monthlyData.map(d => d.label),
                datasets: [
                    {
                        label: 'Reposição de Estoque',
                        data: monthlyData.map(d => d.reposicao),
                        backgroundColor: 'rgba(99, 102, 241, 0.85)',
                        borderColor: '#6366f1',
                        borderWidth: 1,
                        borderRadius: 4,
                        stack: 'stack1'
                    },
                    {
                        label: 'Venda Casada',
                        data: monthlyData.map(d => d.vendaCasada),
                        backgroundColor: 'rgba(16, 185, 129, 0.85)',
                        borderColor: '#10b981',
                        borderWidth: 1,
                        borderRadius: 4,
                        stack: 'stack1'
                    },
                    {
                        label: 'Outros',
                        data: monthlyData.map(d => Math.max(0, d.total - d.reposicao - d.vendaCasada)),
                        backgroundColor: 'rgba(148, 163, 184, 0.3)',
                        borderColor: 'rgba(148, 163, 184, 0.5)',
                        borderWidth: 1,
                        borderRadius: 4,
                        stack: 'stack1'
                    }
                ]
            },
            plugins: [ChartDataLabels],
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        display: true,
                        position: 'top',
                        labels: { color: '#e2e8f0', font: { family: 'Outfit', size: 11 } }
                    },
                    tooltip: {
                        mode: 'index',
                        intersect: false,
                        backgroundColor: 'rgba(15, 23, 42, 0.95)',
                        titleColor: '#6366f1',
                        padding: 12,
                        callbacks: {
                            label: function(context) {
                                const val = context.raw;
                                const total = context.chart.data.datasets.reduce((acc, ds) => acc + ds.data[context.dataIndex], 0);
                                const pct = total > 0 ? ((val / total) * 100).toFixed(1) : 0;
                                return `${context.dataset.label}: ${formatCurrency(val)} (${pct}%)`;
                            },
                            footer: (items) => {
                                const total = items.reduce((acc, it) => acc + it.raw, 0);
                                return `TOTAL: ${formatCurrency(total)}`;
                            }
                        }
                    },
                    datalabels: {
                        display: (context) => {
                            const val = context.dataset.data[context.dataIndex];
                            const total = context.chart.data.datasets.reduce((acc, ds) => acc + (ds.data[context.dataIndex] || 0), 0);
                            return total > 0 && (val / total) > 0.15; // Only show if > 15% to avoid clutter
                        },
                        color: '#fff',
                        font: { weight: 'bold', size: 10 },
                        formatter: (value, context) => {
                            const total = context.chart.data.datasets.reduce((acc, ds) => acc + (ds.data[context.dataIndex] || 0), 0);
                            const pct = total > 0 ? ((value / total) * 100).toFixed(0) : 0;
                            return pct + '%';
                        }
                    }
                },
                scales: {
                    x: {
                        stacked: true,
                        grid: { display: false },
                        ticks: { color: '#94a3b8' }
                    },
                    y: {
                        stacked: true,
                        grid: { color: 'rgba(255,255,255,0.05)' },
                        ticks: { 
                            color: '#94a3b8',
                            callback: (value) => formatCurrency(value)
                        }
                    }
                },
                onClick: (e, elements) => {
                    if (elements.length > 0) {
                        const index = elements[0].index;
                        const monthKey = monthlyData[index].key;
                        selectedMonth = selectedMonth === monthKey ? 'all' : monthKey;
                        updateDashboard();
                    } else {
                        selectedMonth = 'all';
                        updateDashboard();
                    }
                }
            }
        });
    }

    function renderCardTrend(elementId, current, previous) {
        const el = document.getElementById(elementId);
        if (!el) return;
        
        if (!previous || previous === 0) {
            el.style.display = 'none';
            return;
        }

        el.style.display = 'inline-block';
        const variation = ((current - previous) / previous) * 100;
        const text = variation > 0 ? `↑ +${variation.toFixed(1)}%` : `↓ ${variation.toFixed(1)}%`;
        
        el.textContent = text;
        el.className = 'card-trend-badge ' + (variation > 0 ? 'up' : 'down');
    }

    function exportToExcel() {
        if (nfData.length === 0) return;
        
        const worksheet = XLSX.utils.json_to_sheet(nfData.map(d => ({
            'Data Movto': d.date,
            'Fornecedor': d.supplier,
            'Valor (R$)': d.value
        })));
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Entradas");
        XLSX.writeFile(workbook, "Relatorio_Entradas.xlsx");
    }
});
