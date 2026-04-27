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

    // --- Event Listeners ---
    if (nfUpload) nfUpload.addEventListener('change', handleFileUpload);
    if (nfSearch) nfSearch.addEventListener('input', () => updateDashboard());

    const exportBtn = document.getElementById('export-entradas-btn');
    if (exportBtn) exportBtn.addEventListener('click', exportToExcel);

    const syncFolderBtn = document.getElementById('sync-folder-btn');
    if (syncFolderBtn) syncFolderBtn.addEventListener('click', handleFolderSync);

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

    async function handleFolderSync() {
        showLoading(true);
        
        try {
            // Tenta se comunicar com a ponte local (bridge.js)
            console.log("Tentando conexão com a ponte local...");
            const response = await fetch('http://localhost:3000/sync', { 
                mode: 'cors',
                cache: 'no-cache'
            });
            
            if (!response.ok) throw new Error('Servidor ponte retornou erro.');
            
            const result = await response.json();
            console.log("Resposta da ponte:", result);
            
            if (result.success) {
                // Recarregar a página para pegar o novo data-entradas.js
                alert(`Sincronização concluída! ${result.count} registros processados.`);
                location.reload();
                return;
            }
        } catch (e) {
            console.warn("Ponte local não detectada ou erro:", e.message);
        }

        // Fallback para o Directory Picker (Segurança do Navegador)
        if (!window.showDirectoryPicker) {
            alert('Ponte local não detectada. Por favor, execute "node bridge.js" ou use um navegador moderno (Chrome/Edge).');
            showLoading(false);
            return;
        }

        try {
            const dirHandle = await window.showDirectoryPicker();
            let allMergedData = [];
            
            for await (const entry of dirHandle.values()) {
                if (entry.kind === 'file' && (entry.name.endsWith('.xlsx') || entry.name.endsWith('.xls'))) {
                    console.log(`Lendo arquivo: ${entry.name}`);
                    const file = await entry.getFile();
                    const rawData = await readExcel(file);
                    const processed = processIndividualFile(rawData);
                    allMergedData = allMergedData.concat(processed);
                }
            }
            
            if (allMergedData.length > 0) {
                nfData = allMergedData;
                welcomeView.style.display = 'none';
                dashboardView.style.display = 'block';
                updateDashboard();
                alert(`Sucesso! ${allMergedData.length} registros carregados da pasta selecionada.`);
            } else {
                alert('Nenhum arquivo Excel válido encontrado.');
            }
            
        } catch (error) {
            console.error('Erro na sincronização manual:', error);
            if (error.name !== 'AbortError') alert('Erro ao acessar a pasta.');
        } finally {
            showLoading(false);
        }
    }

    function processIndividualFile(data) {
        if (!data || data.length === 0) return [];

        const firstRow = data[0];
        const colDate = findColumn(firstRow, ['Dt.movto', 'Data', 'Movimento']);
        const colSupplier = findColumn(firstRow, ['Razão social', 'Fornecedor', 'Nome']);
        const colValue = findColumn(firstRow, ['Vlr.cont', 'Valor', 'Total']);

        return data.map(row => {
            const rawDateValue = row[colDate];
            const supplier = row[colSupplier] || 'NÃO IDENTIFICADO';
            const value = parseFloat(row[colValue]) || 0;
            const dateObj = parseExcelDate(rawDateValue);

            return {
                date: formatDate(dateObj),
                rawDate: dateObj,
                supplier: supplier.toString().trim(),
                value: value
            };
        }).filter(item => item.value > 0 && !isNaN(item.rawDate.getTime()));
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
                const json = XLSX.utils.sheet_to_json(worksheet);
                resolve(json);
            };
            reader.onerror = reject;
            reader.readAsArrayBuffer(file);
        });
    }

    function processData(data) {
        if (!data || data.length === 0) return;

        // Identificar colunas dinamicamente (case-insensitive e trim)
        const firstRow = data[0];
        const colDate = findColumn(firstRow, ['Dt.movto', 'Data', 'Movimento']);
        const colSupplier = findColumn(firstRow, ['Razão social', 'Fornecedor', 'Nome']);
        const colValue = findColumn(firstRow, ['Vlr.cont', 'Valor', 'Total']);

        // Mapear e limpar dados
        nfData = data.map(row => {
            const rawDateValue = row[colDate];
            const supplier = row[colSupplier] || 'NÃO IDENTIFICADO';
            const value = parseFloat(row[colValue]) || 0;

            const dateObj = parseExcelDate(rawDateValue);

            return {
                date: formatDate(dateObj),
                rawDate: dateObj,
                supplier: supplier.toString().trim(),
                value: value
            };
        }).filter(item => item.value > 0 && !isNaN(item.rawDate.getTime()));
    }

    function findColumn(row, possibilities) {
        const keys = Object.keys(row);
        for (const p of possibilities) {
            const match = keys.find(k => k.trim().toLowerCase() === p.toLowerCase());
            if (match) return match;
        }
        // Fallback: primeira coluna que contém parte do nome
        return keys.find(k => possibilities.some(p => k.toLowerCase().includes(p.toLowerCase()))) || keys[0];
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
        
        const filteredBySearch = nfData.filter(item => 
            item.supplier.toLowerCase().includes(nfSearch.value.toLowerCase())
        );
        const filtered = filterByPeriodAndSearch();
        
        // 1. Stats
        const total = filtered.reduce((acc, item) => acc + item.value, 0);
        const count = filtered.length;
        const avg = count > 0 ? total / count : 0;

        document.getElementById('total-received').textContent = formatCurrency(total);
        document.getElementById('nf-count').textContent = `${count} notas fiscais`;
        document.getElementById('avg-nf-value').textContent = formatCurrency(avg);

        // 2. Top Supplier
        const supplierTotals = aggregateBySupplier(filtered);
        const topSupplierEl = document.getElementById('top-supplier');
        const topSupplierValueEl = document.getElementById('top-supplier-value');
        
        if (supplierTotals.length > 0) {
            topSupplierEl.textContent = supplierTotals[0].supplier;
            topSupplierValueEl.textContent = formatCurrency(supplierTotals[0].total);
        } else {
            topSupplierEl.textContent = "-";
            topSupplierValueEl.textContent = "R$ 0,00";
        }

        // 3. Charts
        renderSuppliersChart(supplierTotals.slice(0, 10));
        renderVolumeChart(supplierTotals.slice(0, 10));
        renderAverageChart(supplierTotals.slice(0, 10));
        renderDailyChart(aggregateByDate(filtered));
        renderMonthlyChart(aggregateByMonth(filteredBySearch));


        // 4. Table
        renderTable(filtered);
        
        // 5. Update Timeline
        renderTimeline();
    }

    function filterByPeriodAndSearch() {
        const searchTerm = nfSearch.value.toLowerCase();
        return nfData.filter(item => {
            const matchesSearch = item.supplier.toLowerCase().includes(searchTerm);
            
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
        const stats = {};
        data.forEach(item => {
            if (!stats[item.supplier]) {
                stats[item.supplier] = { total: 0, count: 0 };
            }
            stats[item.supplier].total += item.value;
            stats[item.supplier].count += 1;
        });

        return Object.keys(stats).map(supplier => {
            const total = stats[supplier].total;
            const count = stats[supplier].count;
            return {
                supplier,
                total: total,
                count: count,
                average: total / count
            };
        }).sort((a, b) => b.total - a.total);
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
                    label: d.toLocaleString('pt-br', { month: 'short', year: '2-digit' }),
                    key: monthKey
                };
            }
            map[monthKey].total += item.value;
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
                <td>${item.supplier}</td>
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

    function renderSuppliersChart(topData) {
        const ctx = document.getElementById('suppliers-chart').getContext('2d');
        
        if (suppliersChart) suppliersChart.destroy();

        // Criar gradiente
        const gradient = ctx.createLinearGradient(0, 0, 400, 0);
        gradient.addColorStop(0, 'rgba(99, 102, 241, 0.8)');
        gradient.addColorStop(1, 'rgba(16, 185, 129, 0.8)');

        suppliersChart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: topData.map(d => d.supplier),
                datasets: [{
                    label: 'Total Recebido (R$)',
                    data: topData.map(d => d.total),
                    backgroundColor: gradient,
                    borderColor: 'rgba(255, 255, 255, 0.2)',
                    borderWidth: 1,
                    borderRadius: 6,
                    maxBarThickness: 40,
                    hoverBackgroundColor: 'rgba(99, 102, 241, 1)',
                }]
            },
            plugins: [ChartDataLabels],
            options: {
                indexAxis: 'y',
                responsive: true,
                maintainAspectRatio: false,
                animation: { duration: 1000 },
                layout: { padding: { right: 50 } },
                onClick: (e, elements) => {
                    if (elements.length > 0) {
                        const index = elements[0].index;
                        const supplier = topData[index].supplier;
                        nfSearch.value = supplier;
                        updateDashboard();
                    } else {
                        // Limpar filtro se clicar fora das barras
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
                        align: 'right',
                        color: '#f8fafc',
                        font: { weight: 'bold', size: 10 },
                        formatter: (value) => formatCurrency(value),
                        clip: false
                    },
                    tooltip: {
                        backgroundColor: 'rgba(15, 23, 42, 0.9)',
                        titleColor: '#6366f1',
                        bodyColor: '#fff',
                        padding: 12,
                        cornerRadius: 8,
                        displayColors: false,
                        callbacks: {
                            label: (context) => `Total: ${formatCurrency(context.raw)}`
                        }
                    }
                },
                scales: {
                    x: { 
                        display: false,
                        grid: { display: false },
                        grace: '10%'
                    },
                    y: { 
                        grid: { display: false },
                        beginAtZero: true,
                        ticks: { 
                            color: '#e2e8f0', 
                            font: { size: 11, family: 'Outfit' },
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
    function renderVolumeChart(topData) {
        const ctx = document.getElementById('volume-chart').getContext('2d');
        if (volumeChart) volumeChart.destroy();

        // Ordenar por volume para este gráfico
        const volumeData = [...topData].sort((a, b) => b.count - a.count);

        const gradient = ctx.createLinearGradient(0, 0, 400, 0);
        gradient.addColorStop(0, 'rgba(16, 185, 129, 0.8)');
        gradient.addColorStop(1, 'rgba(34, 197, 94, 0.8)');

        volumeChart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: volumeData.map(d => d.supplier),
                datasets: [{
                    label: 'Quantidade de Notas',
                    data: volumeData.map(d => d.count),
                    backgroundColor: gradient,
                    borderColor: 'rgba(255, 255, 255, 0.2)',
                    borderWidth: 1,
                    borderRadius: 6,
                    maxBarThickness: 40,
                    hoverBackgroundColor: 'rgba(16, 185, 129, 1)',
                }]
            },
            plugins: [ChartDataLabels],
            options: {
                indexAxis: 'y',
                responsive: true,
                maintainAspectRatio: false,
                onClick: (e, elements) => {
                    if (elements.length > 0) {
                        const index = elements[0].index;
                        const supplier = volumeData[index].supplier;
                        nfSearch.value = supplier;
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
                        align: 'right',
                        color: '#f8fafc',
                        font: { weight: 'bold', size: 10 },
                        formatter: (value) => `${value} NFs`,
                        clip: false
                    },
                    tooltip: {
                        backgroundColor: 'rgba(15, 23, 42, 0.9)',
                        titleColor: '#10b981',
                        bodyColor: '#fff',
                        padding: 12,
                        cornerRadius: 8,
                        displayColors: false,
                        callbacks: {
                            label: (context) => `Volume: ${context.raw} notas`
                        }
                    }
                },
                scales: {
                    x: { display: false, grid: { display: false }, grace: '10%' },
                    y: { grid: { display: false }, ticks: { color: '#e2e8f0', font: { size: 11, family: 'Outfit' } } }
                }
            }
        });
    }

    function renderAverageChart(topData) {
        const ctx = document.getElementById('average-chart').getContext('2d');
        if (averageChart) averageChart.destroy();

        const avgData = [...topData].sort((a, b) => b.average - a.average);

        const gradient = ctx.createLinearGradient(0, 0, 400, 0);
        gradient.addColorStop(0, 'rgba(245, 158, 11, 0.8)');
        gradient.addColorStop(1, 'rgba(251, 191, 36, 0.8)');

        averageChart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: avgData.map(d => d.supplier),
                datasets: [{
                    label: 'Média por Nota (R$)',
                    data: avgData.map(d => d.average),
                    backgroundColor: gradient,
                    borderColor: 'rgba(255, 255, 255, 0.2)',
                    borderWidth: 1,
                    borderRadius: 6,
                    maxBarThickness: 40,
                    hoverBackgroundColor: 'rgba(245, 158, 11, 1)',
                }]
            },
            plugins: [ChartDataLabels],
            options: {
                indexAxis: 'y',
                responsive: true,
                maintainAspectRatio: false,
                onClick: (e, elements) => {
                    if (elements.length > 0) {
                        const index = elements[0].index;
                        const supplier = avgData[index].supplier;
                        nfSearch.value = supplier;
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
                        align: 'right',
                        color: '#f8fafc',
                        font: { weight: 'bold', size: 10 },
                        formatter: (value) => formatCurrency(value),
                        clip: false
                    },
                    tooltip: {
                        backgroundColor: 'rgba(15, 23, 42, 0.9)',
                        titleColor: '#f59e0b',
                        bodyColor: '#fff',
                        padding: 12,
                        cornerRadius: 8,
                        displayColors: false,
                        callbacks: {
                            label: (context) => `Média: ${formatCurrency(context.raw)}`
                        }
                    }
                },
                scales: {
                    x: { display: false, grid: { display: false }, grace: '15%' },
                    y: { grid: { display: false }, ticks: { color: '#e2e8f0', font: { size: 11, family: 'Outfit' } } }
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

    function renderMonthlyChart(monthlyData) {
        const ctx = document.getElementById('monthly-chart').getContext('2d');
        if (monthlyChart) monthlyChart.destroy();

        const gradient = ctx.createLinearGradient(0, 0, 0, 300);
        gradient.addColorStop(0, 'rgba(99, 102, 241, 0.4)');
        gradient.addColorStop(1, 'rgba(99, 102, 241, 0)');

        monthlyChart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: monthlyData.map(d => d.label),
                datasets: [{
                    label: 'Total Recebido Mensal (R$)',
                    data: monthlyData.map(d => d.total),
                    backgroundColor: monthlyData.map(d => 
                        d.key === selectedMonth ? 'rgba(16, 185, 129, 0.8)' : 'rgba(99, 102, 241, 0.6)'
                    ),
                    borderColor: monthlyData.map(d => 
                        d.key === selectedMonth ? '#10b981' : '#6366f1'
                    ),
                    borderWidth: 2,
                    borderRadius: 8,
                    hoverBackgroundColor: 'rgba(99, 102, 241, 0.9)',
                }]
            },
            plugins: [ChartDataLabels],
            options: {
                responsive: true,
                maintainAspectRatio: false,
                onClick: (e, elements) => {
                    if (elements.length > 0) {
                        const index = elements[0].index;
                        const monthKey = monthlyData[index].key;
                        
                        if (selectedMonth === monthKey) {
                            selectedMonth = 'all'; // Toggle off if clicking same month
                        } else {
                            selectedMonth = monthKey;
                        }
                        updateDashboard();
                    } else {
                        selectedMonth = 'all';
                        updateDashboard();
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
                        color: '#f8fafc',
                        offset: 4,
                        font: { weight: 'bold', size: 10, family: 'Outfit' },
                        formatter: (value) => formatCurrency(value),
                        display: (context) => context.dataset.data[context.dataIndex] > 0
                    },
                    tooltip: {
                        backgroundColor: 'rgba(15, 23, 42, 0.9)',
                        padding: 12,
                        cornerRadius: 8,
                        callbacks: {
                            label: (context) => ` Total: ${formatCurrency(context.raw)}`
                        }
                    }
                },
                scales: {
                    x: { 
                        grid: { display: false },
                        ticks: { 
                            color: '#e2e8f0',
                            font: { family: 'Outfit', size: 12 }
                        }
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
