document.addEventListener('DOMContentLoaded', () => {
    console.log("Módulo de Análise de Rupturas Carregado");
    
    const loadingOverlay = document.getElementById('loading-overlay');
    const folderUpload = document.getElementById('folder-upload');
    const historyTableBody = document.getElementById('history-table-body');
    
    let evolutionChart = null;
    let valueChart = null;

    function showLoading() {
        if (loadingOverlay) loadingOverlay.style.display = 'flex';
    }

    function hideLoading() {
        if (loadingOverlay) loadingOverlay.style.display = 'none';
    }

    /**
     * Tenta encontrar o valor em uma linha usando múltiplos nomes de coluna possíveis (Case-Insensitive)
     * Mesma lógica do app.js para garantir paridade total.
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
            if (row[key] !== undefined) return row[key];
            const normKey = rowKeys.find(rk => cleanStr(rk) === target);
            if (normKey) return row[normKey];
        }

        // Fallback: word match
        for (let key of keys) {
            const target = cleanStr(key);
            if (target.length < 3) continue;
            const fuzzyKey = rowKeys.find(rk => {
                const words = cleanStr(rk).split(' ');
                return words.some(word => word === target || (word.length >= 4 && target.startsWith(word)));
            });
            if (fuzzyKey) return row[fuzzyKey];
        }
        return undefined;
    }

    /**
     * Converte valor para número de forma segura (paridade com app.js)
     */
    function parseNum(val) {
        if (val === undefined || val === null || val === '') return 0;
        if (typeof val === 'number') return val;
        let str = val.toString().replace('R$', '').replace(/\s/g, '').trim();
        if (str.startsWith('.') || str.startsWith(',')) str = '0' + str;
        const hasComma = str.includes(',');
        const hasDot = str.includes('.');
        if (hasComma && hasDot) {
            if (str.lastIndexOf(',') > str.lastIndexOf('.')) str = str.replace(/\./g, '').replace(',', '.');
            else str = str.replace(/,/g, '');
        } else if (hasComma) str = str.replace(',', '.');
        const num = parseFloat(str);
        return isNaN(num) ? 0 : num;
    }

    async function handleFolderUpload(e) {
        const files = Array.from(e.target.files).filter(f => f.name.match(/\.(xlsx|xls)$/i) && !f.name.startsWith('~$'));
        if (files.length === 0) return;

        showLoading();
        
        try {
            const historyData = [];

            for (const file of files) {
                const data = await readExcel(file);
                if (!data || data.length === 0) continue;

                let fileDate = new Date(file.lastModified).toISOString().split('T')[0];
                const dateMatch = file.name.match(/(\d{4}-\d{2}-\d{2})|(\d{2}-\d{2}-\d{4})/);
                if (dateMatch) {
                    fileDate = dateMatch[0];
                    if (fileDate.includes('-') && fileDate.split('-')[0].length === 2) {
                        const parts = fileDate.split('-');
                        fileDate = `${parts[2]}-${parts[1]}-${parts[0]}`;
                    }
                }

                const rows = data;
                let ruptureCount = 0; let ruptureValue = 0;
                let attentionCount = 0; let attentionValue = 0;
                let suggestCount = 0; let suggestValue = 0;
                let totalItems = rows.length;

                const allKeys = Object.keys(rows[0] || {});
                const validMonthInfo = allKeys.filter(k => /^\d{2}\/\d{4}$/.test(k)).filter(k => {
                    const [m, y] = k.split('/').map(Number);
                    return (y < 2026) || (y === 2026 && m <= 5);
                });

                rows.forEach(row => {
                    // Aliases sincronizados com app.js
                    const estoque = parseNum(getValue(row, ['Estoque', 'Saldo', 'Qtd. Estoque', 'Estoque Total', 'Saldo Atual', 'Saldo Disponível', 'Disp.', 'Qtd. Disponível', 'Estoque Atual']));
                    const encomendas = parseNum(getValue(row, [
                        'Encomendas', 'Qtd. Encomenda', 'Saldo Pedido Compra', 'Saldo Ped. Compra', 'Pedido Compra', 
                        'Qtd. em Pedido Compra', 'Qtd. no Pedido Compra', 'Saldo a Receber', 'A Receber', 'Pedidos', 
                        'Qtd. Pedida', 'Saldo Pedido', 'Compras', 'Qtd em Pedido', 'Qtd. Ped.', 'Saldo Ped.', 
                        'Pendência', 'Qtd. no Pedido', 'Encomenda', 'Pedido', 'Qtd Ped Compra', 'A Receber Total',
                        'A Entregar', 'Saldo a Entregar', 'Qtd. Pendente', 'Pendente', 'Saldo O.C.', 'Ord. Compra'
                    ]));
                    const custo = parseNum(getValue(row, ['Preço reposição', 'Custo aquisição', 'Custo Unitário', 'Custo', 'Preço Custo', 'Vlr. Custo', 'Custo Médio', 'Unitário']));
                    const vendasCol = parseNum(getValue(row, ['Vendas', 'Qtd. Vendida', 'Venda Total', 'Total Vendas', 'Venda', 'Saídas', 'Giro']));
                    
                    let totalVendas = vendasCol;
                    let activeMonths = 0;
                    let recorrencia = 0;
                    let medVenda = 0;

                    if (validMonthInfo.length > 0) {
                        let sumMonths = 0;
                        validMonthInfo.forEach(mKey => {
                            const val = parseNum(row[mKey]);
                            if (val > 0) {
                                sumMonths += val;
                                activeMonths++;
                            }
                        });
                        if (totalVendas === 0) totalVendas = sumMonths;
                        recorrencia = activeMonths / validMonthInfo.length;
                        medVenda = activeMonths > 0 ? (totalVendas / activeMonths) : 0;
                    } else {
                        medVenda = parseNum(getValue(row, ['med.venda', 'media', 'giro', 'venda mensal']));
                        const recRaw = getValue(row, ['recorrencia', 'giro freq', 'frequencia']);
                        recorrencia = parseNum(recRaw);
                        if (recorrencia > 1) recorrencia = recorrencia / 100;
                        else if (!recRaw) recorrencia = 0;
                    }
                    
                    const totalDisponivel = estoque + encomendas;
                    const passesRecurrence = (recorrencia > 0.33);

                    if (passesRecurrence) {
                        if (medVenda > totalDisponivel) {
                            ruptureCount++;
                            ruptureValue += (medVenda * 1 * custo);
                        } else if ((medVenda * 2) > totalDisponivel) {
                            attentionCount++;
                            attentionValue += (medVenda * 2 * custo);
                        } else if ((medVenda * 3) > totalDisponivel) {
                            suggestCount++;
                            suggestValue += (medVenda * 3 * custo);
                        }
                    }
                });


                historyData.push({
                    file: file.name,
                    date: fileDate,
                    totalItems,
                    rupture: { count: ruptureCount, value: ruptureValue },
                    attention: { count: attentionCount, value: attentionValue },
                    suggest: { count: suggestCount, value: suggestValue }
                });
            }

            // Ordenar por data
            historyData.sort((a, b) => new Date(a.date) - new Date(b.date));
            renderDashboard(historyData);

        } catch (e) {
            console.error("ERRO NA ANÁLISE DE RUPTURAS:", e);
            alert("Erro ao processar arquivos: " + e.message);
        } finally {
            hideLoading();
            e.target.value = ''; // Reset input
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

    function renderDashboard(data) {
        if (!data || data.length === 0) {
            historyTableBody.innerHTML = '<tr><td colspan="7" style="text-align:center; padding: 2rem; color: var(--text-muted);">Nenhum arquivo encontrado na pasta selecionada.</td></tr>';
            return;
        }

        // 1. Update Summary Cards (Latest snapshot)
        const latest = data[data.length - 1];
        const previous = data.length > 1 ? data[data.length - 2] : null;
        
        const formatCurrency = (val) => val.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
        
        document.getElementById('current-rupture-value').textContent = formatCurrency(latest.rupture.value);
        document.getElementById('current-rupture-count').textContent = `${latest.rupture.count} itens em risco total`;
        
        document.getElementById('current-attention-value').textContent = formatCurrency(latest.attention.value);
        document.getElementById('current-attention-count').textContent = `${latest.attention.count} itens em monitoramento`;
        
        document.getElementById('current-suggest-value').textContent = formatCurrency(latest.suggest.value);
        document.getElementById('current-suggest-count').textContent = `${latest.suggest.count} itens para reposição base`;

        // Calculate and render trends
        if (previous) {
            renderCardTrend('trend-rupture', latest.rupture.value, previous.rupture.value, true); // true = higher is worse
            renderCardTrend('trend-attention', latest.attention.value, previous.attention.value, true);
            renderCardTrend('trend-suggest', latest.suggest.value, previous.suggest.value, true);
        }

        // 2. Render Table
        historyTableBody.innerHTML = '';
        data.slice().reverse().forEach(entry => {
            const row = document.createElement('tr');
            row.innerHTML = `
                <td>${formatDate(entry.date)}</td>
                <td style="font-size: 0.8rem; color: var(--text-muted);">${entry.file}</td>
                <td style="text-align: center;">${entry.totalItems}</td>
                <td style="text-align: center; font-weight: 700; color: #fb7185;">${entry.rupture.count}</td>
                <td style="text-align: center; font-weight: 700; color: #f59e0b;">${entry.attention.count}</td>
                <td style="text-align: center; font-weight: 700; color: #818cf8;">${entry.suggest.count}</td>
                <td style="text-align: right; font-weight: 700;">${formatCurrency(entry.rupture.value)}</td>
            `;
            historyTableBody.appendChild(row);
        });

        // 3. Render Evolution Charts
        renderEvolutionChart(data);
        renderValueChart(data);
        renderVariationChart(data);
    }

    function renderCardTrend(elementId, current, previous, inverse = false) {
        const el = document.getElementById(elementId);
        if (!el) return;

        if (!previous || previous === 0) {
            el.style.display = 'none';
            return;
        }

        const diff = current - previous;
        const pct = (diff / previous) * 100;
        
        if (Math.abs(pct) < 0.1) {
            el.textContent = '=';
            el.className = 'card-trend-badge';
            el.style.display = 'inline-block';
            return;
        }

        const isUp = diff > 0;
        const label = (isUp ? '↑ ' : '↓ ') + Math.abs(pct).toFixed(1) + '%';
        
        el.textContent = label;
        el.style.display = 'inline-block';
        
        // Se inverse for true, subida (vermelho) é ruim, descida (verde) é bom
        if (inverse) {
            el.className = 'card-trend-badge ' + (isUp ? 'down' : 'up'); // 'down' class is red, 'up' is green
        } else {
            el.className = 'card-trend-badge ' + (isUp ? 'up' : 'down');
        }
    }

    function formatDate(dateStr) {
        const parts = dateStr.split('-');
        if (parts.length === 3) {
            return `${parts[2]}/${parts[1]}/${parts[0]}`;
        }
        return dateStr;
    }

    function renderEvolutionChart(data) {
        const ctx = document.getElementById('rupture-evolution-chart').getContext('2d');
        const labels = data.map(d => formatDate(d.date));
        
        const chartData = {
            labels: labels,
            datasets: [
                {
                    label: 'Ruptura (Crítico)',
                    data: data.map(d => d.rupture.count),
                    borderColor: '#fb7185',
                    backgroundColor: 'rgba(251, 113, 133, 0.1)',
                    tension: 0.4,
                    fill: true,
                    pointRadius: 4,
                    pointHoverRadius: 6
                },
                {
                    label: 'Atenção (2m)',
                    data: data.map(d => d.attention.count),
                    borderColor: '#f59e0b',
                    backgroundColor: 'rgba(245, 158, 11, 0.1)',
                    tension: 0.4,
                    fill: true,
                    pointRadius: 4,
                    pointHoverRadius: 6
                },
                {
                    label: 'Sugestão (3m)',
                    data: data.map(d => d.suggest.count),
                    borderColor: '#818cf8',
                    backgroundColor: 'rgba(129, 140, 248, 0.1)',
                    tension: 0.4,
                    fill: true,
                    pointRadius: 4,
                    pointHoverRadius: 6
                }
            ]
        };

        if (evolutionChart) evolutionChart.destroy();
        
        evolutionChart = new Chart(ctx, {
            type: 'line',
            data: chartData,
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        position: 'top',
                        labels: { color: '#94a3b8', font: { family: 'Outfit', size: 12 } }
                    },
                    datalabels: { display: false }
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        grid: { color: 'rgba(255, 255, 255, 0.05)' },
                        ticks: { color: '#94a3b8' }
                    },
                    x: {
                        grid: { display: false },
                        ticks: { color: '#94a3b8' }
                    }
                }
            }
        });
    }

    function renderValueChart(data) {
        const ctx = document.getElementById('value-evolution-chart').getContext('2d');
        const labels = data.map(d => formatDate(d.date));
        
        const chartData = {
            labels: labels,
            datasets: [
                {
                    label: 'Valor em Ruptura (Capital Parado)',
                    data: data.map(d => d.rupture.value),
                    borderColor: '#fb7185',
                    backgroundColor: 'rgba(251, 113, 133, 0.2)',
                    tension: 0.3,
                    fill: true,
                    borderWidth: 3
                }
            ]
        };

        if (valueChart) valueChart.destroy();
        
        valueChart = new Chart(ctx, {
            type: 'line',
            data: chartData,
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        callbacks: {
                            label: function(context) {
                                return new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(context.parsed.y);
                            }
                        }
                    }
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        grid: { color: 'rgba(255, 255, 255, 0.05)' },
                        ticks: { 
                            color: '#94a3b8',
                            callback: function(value) {
                                if (value >= 1000) return 'R$ ' + (value/1000).toFixed(1) + 'k';
                                return 'R$ ' + value;
                            }
                        }
                    },
                    x: {
                        grid: { display: false },
                        ticks: { color: '#94a3b8' }
                    }
                }
            }
        });
    }

    let variationChart = null;
    function renderVariationChart(data) {
        // Find or create container for the new chart
        let container = document.getElementById('variation-container');
        if (!container) {
            const row = document.createElement('div');
            row.className = 'entradas-row-full';
            row.style.marginBottom = '2rem';
            row.innerHTML = `
                <div class="chart-container-nf" id="variation-container" style="min-height: 400px;">
                    <h3 style="margin-bottom: 1.5rem; color: var(--info);">Variação Diária da Ruptura (Melhora vs Piora)</h3>
                    <div style="height: 300px; position: relative;">
                        <canvas id="rupture-variation-chart"></canvas>
                    </div>
                </div>
            `;
            // Insert before the table
            const tableSection = document.querySelector('.data-section');
            tableSection.parentNode.insertBefore(row, tableSection);
            container = document.getElementById('variation-container');
        }

        const ctx = document.getElementById('rupture-variation-chart').getContext('2d');
        
        // Skip first element as there's no variation
        const labels = data.slice(1).map(d => formatDate(d.date));
        const variations = data.slice(1).map((d, i) => {
            const prev = data[i];
            return d.rupture.value - prev.rupture.value;
        });

        const chartData = {
            labels: labels,
            datasets: [{
                label: 'Variação Financeira',
                data: variations,
                backgroundColor: variations.map(v => v > 0 ? 'rgba(251, 113, 133, 0.6)' : 'rgba(52, 211, 153, 0.6)'),
                borderColor: variations.map(v => v > 0 ? '#fb7185' : '#34d399'),
                borderWidth: 1,
                borderRadius: 4
            }]
        };

        if (variationChart) variationChart.destroy();
        
        variationChart = new Chart(ctx, {
            type: 'bar',
            data: chartData,
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        callbacks: {
                            label: function(context) {
                                const val = context.parsed.y;
                                const sign = val > 0 ? '+' : '';
                                return 'Delta: ' + sign + new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(val);
                            }
                        }
                    }
                },
                scales: {
                    y: {
                        grid: { color: 'rgba(255, 255, 255, 0.05)' },
                        ticks: { color: '#94a3b8' }
                    },
                    x: {
                        grid: { display: false },
                        ticks: { color: '#94a3b8' }
                    }
                }
            }
        });
    }

    if (folderUpload) folderUpload.addEventListener('change', handleFolderUpload);

    // We no longer call fetchRuptureData() automatically since we need user to select a folder
    // But we can check if there's cached data (optional)
});
