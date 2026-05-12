/**
 * Evolução e Projeção de Investimento v1.1
 * Processamento local de pastas selecionadas pelo usuário
 */
document.addEventListener('DOMContentLoaded', () => {
    let historyData = [];
    let chart = null;

    const formatCurrency = (v) => v.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
    const formatNumber = (v) => v.toLocaleString('pt-BR');

    // Helpers para processamento de Excel (baseado na Bridge)
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

    function getValueInRow(row, possibilities) {
        const col = findColumn(row, possibilities);
        return col ? row[col] : null;
    }

    /**
     * Processa um arquivo individualmente
     */
    async function processFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array', cellDates: true });
                    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    const rows = XLSX.utils.sheet_to_json(firstSheet);
                    
                    // Extração de data do nome do arquivo
                    let fileDate = new Date(file.lastModified).toISOString().split('T')[0];
                    const dateMatch = file.name.match(/(\d{4}-\d{2}-\d{2})|(\d{2}-\d{2}-\d{4})/);
                    if (dateMatch) {
                        fileDate = dateMatch[0];
                        if (fileDate.includes('-') && fileDate.split('-')[0].length === 2) {
                            // Converte DD-MM-YYYY para YYYY-MM-DD se necessário
                            const p = fileDate.split('-');
                            fileDate = `${p[2]}-${p[1]}-${p[0]}`;
                        }
                    }

                    let ruptureCount = 0, ruptureValue = 0, ruptureInvest = 0;
                    let attentionCount = 0, attentionValue = 0, attentionInvest = 0;
                    let suggestCount = 0, suggestValue = 0, suggestInvest = 0;

                    // Identificar colunas de meses (mesma lógica da Bridge)
                    const rawRows = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
                    const headerRow = rawRows[0] || [];
                    const monthInfo = headerRow.map((val, idx) => {
                        let label = "";
                        if (val instanceof Date) {
                            label = `${(val.getMonth() + 1).toString().padStart(2, '0')}/${val.getFullYear()}`;
                        } else {
                            label = String(val || '');
                        }
                        return { label, original: val };
                    }).filter(h => /^\d{2}\/\d{4}$/.test(h.label));

                    // Limite de data do projeto
                    const validMonthInfo = monthInfo.filter(h => {
                        const [m, y] = h.label.split('/').map(Number);
                        return (y < 2026) || (y === 2026 && m <= 5);
                    });

                    rows.forEach(row => {
                        const estoque = parseFloat(getValueInRow(row, ['estoque', 'saldo', 'atual'])) || 0;
                        const encomendas = parseFloat(getValueInRow(row, ['encomendas', 'pedido', 'transito', 'receber'])) || 0;
                        const custo = parseFloat(getValueInRow(row, ['preco reposicao', 'custo', 'unitario'])) || 0;
                        
                        const vendasRaw = parseFloat(getValueInRow(row, ['vendas', 'qtd. vendida', 'venda total', 'total vendas'])) || 0;
                        let totalVendas = vendasRaw;
                        let activeMonths = 0;
                        let recorrencia = 0;
                        let medVenda = 0;

                        if (validMonthInfo.length > 0) {
                            let sumMonths = 0;
                            validMonthInfo.forEach(m => {
                                const val = parseFloat(row[m.label] || row[m.original] || 0);
                                if (val > 0) {
                                    sumMonths += val;
                                    activeMonths++;
                                }
                            });
                            if (totalVendas === 0) totalVendas = sumMonths;
                            recorrencia = activeMonths / validMonthInfo.length;
                            medVenda = activeMonths > 0 ? (totalVendas / activeMonths) : 0;
                        } else {
                            medVenda = parseFloat(getValueInRow(row, ['med.venda', 'media', 'giro', 'venda mensal'])) || 0;
                            recorrencia = 1; // Default se não houver meses
                        }

                        const totalDisponivel = estoque + encomendas;
                        const passesRecurrence = (recorrencia > 0.33);

                        if (passesRecurrence) {
                            if (medVenda > totalDisponivel) {
                                ruptureCount++;
                                ruptureValue += (medVenda * custo);
                                ruptureInvest += Math.max(0, (medVenda - totalDisponivel) * custo);
                            } else if ((medVenda * 2) > totalDisponivel) {
                                attentionCount++;
                                attentionValue += (medVenda * 2 * custo);
                                attentionInvest += Math.max(0, (medVenda * 2 - totalDisponivel) * custo);
                            } else if ((medVenda * 3) > totalDisponivel) {
                                suggestCount++;
                                suggestValue += (medVenda * 3 * custo);
                                suggestInvest += Math.max(0, (medVenda * 3 - totalDisponivel) * custo);
                            }
                        }
                    });

                    resolve({
                        file: file.name,
                        date: fileDate,
                        totalItems: rows.length,
                        rupture: { count: ruptureCount, value: ruptureValue, invest: ruptureInvest },
                        attention: { count: attentionCount, value: attentionValue, invest: attentionInvest },
                        suggest: { count: suggestCount, value: suggestValue, invest: suggestInvest }
                    });
                } catch (err) {
                    console.error("Erro ao ler arquivo:", file.name, err);
                    resolve(null);
                }
            };
            reader.onerror = () => resolve(null);
            reader.readAsArrayBuffer(file);
        });
    }

    /**
     * Função disparada ao selecionar a pasta
     */
    async function handleFolderSelect(event) {
        const files = Array.from(event.target.files).filter(f => f.name.endsWith('.xlsx') || f.name.endsWith('.xls'));
        if (files.length === 0) {
            alert("Nenhum arquivo Excel encontrado na pasta selecionada.");
            return;
        }

        document.getElementById('loading-overlay').style.display = 'flex';
        historyData = [];

        try {
            const results = await Promise.all(files.map(f => processFile(f)));
            historyData = results.filter(r => r !== null);
            
            // Ordenar por data
            historyData.sort((a, b) => new Date(a.date) - new Date(b.date));

            if (historyData.length > 0) {
                renderDashboard();
            } else {
                alert("Não foi possível extrair dados válidos dos arquivos.");
            }
        } catch (error) {
            console.error("Erro no processamento global:", error);
            alert("Ocorreu um erro ao processar os arquivos.");
        } finally {
            document.getElementById('loading-overlay').style.display = 'none';
        }
    }

    function renderDashboard() {
        if (historyData.length === 0) return;

        const latest = historyData[historyData.length - 1];
        
        // Render Cards
        const cardsContainer = document.getElementById('projection-cards');
        cardsContainer.innerHTML = `
            <div class="projection-card">
                <span class="level-badge level-rupture">Ruptura (1 mês)</span>
                <div class="invest-label">Itens Críticos</div>
                <div style="font-size: 1.5rem; font-weight: 700; color: #fff;">${latest.rupture.count} itens</div>
                <div class="invest-label" style="margin-top: 1rem;">Capital em Risco</div>
                <div class="invest-value" style="color: #f87171;">${formatCurrency(latest.rupture.value)}</div>
                <div class="invest-label">Investimento de Cobertura</div>
                <div style="font-size: 1.1rem; color: #fff;">${formatCurrency(latest.rupture.invest)}</div>
            </div>
            <div class="projection-card">
                <span class="level-badge level-attention">Atenção (2 meses)</span>
                <div class="invest-label">Itens em Monitoramento</div>
                <div style="font-size: 1.5rem; font-weight: 700; color: #fff;">${latest.attention.count} itens</div>
                <div class="invest-label" style="margin-top: 1rem;">Capital em Risco</div>
                <div class="invest-value" style="color: #fbbf24;">${formatCurrency(latest.attention.value)}</div>
                <div class="invest-label">Investimento de Cobertura</div>
                <div style="font-size: 1.1rem; color: #fff;">${formatCurrency(latest.attention.invest)}</div>
            </div>
            <div class="projection-card">
                <span class="level-badge level-suggest">Sugestão (3 meses)</span>
                <div class="invest-label">Sugestões de Compra</div>
                <div style="font-size: 1.5rem; font-weight: 700; color: #fff;">${latest.suggest.count} itens</div>
                <div class="invest-label" style="margin-top: 1rem;">Capital em Risco</div>
                <div class="invest-value" style="color: #34d399;">${formatCurrency(latest.suggest.value)}</div>
                <div class="invest-label">Investimento de Cobertura</div>
                <div style="font-size: 1.1rem; color: #fff;">${formatCurrency(latest.suggest.invest)}</div>
            </div>
        `;

        renderChart();
        updateTotal();
    }

    function renderChart() {
        const ctx = document.getElementById('evolution-chart').getContext('2d');
        if (chart) chart.destroy();

        const labels = historyData.map(d => {
            const date = new Date(d.date);
            return date.toLocaleDateString('pt-BR', { day: '2-digit', month: '2-digit' });
        });

        chart = new Chart(ctx, {
            type: 'line',
            data: {
                labels: labels,
                datasets: [
                    {
                        label: 'Ruptura (Investimento)',
                        data: historyData.map(d => d.rupture.invest),
                        borderColor: '#f87171',
                        backgroundColor: 'rgba(239, 68, 68, 0.1)',
                        fill: true,
                        tension: 0.4
                    },
                    {
                        label: 'Atenção (Investimento)',
                        data: historyData.map(d => d.attention.invest),
                        borderColor: '#fbbf24',
                        backgroundColor: 'rgba(245, 158, 11, 0.1)',
                        fill: true,
                        tension: 0.4
                    },
                    {
                        label: 'Sugestão (Investimento)',
                        data: historyData.map(d => d.suggest.invest),
                        borderColor: '#34d399',
                        backgroundColor: 'rgba(16, 185, 129, 0.1)',
                        fill: true,
                        tension: 0.4
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        position: 'top',
                        labels: { color: '#94a3b8', font: { family: 'Inter' } }
                    },
                    tooltip: {
                        mode: 'index',
                        intersect: false,
                        callbacks: {
                            label: function(context) {
                                return context.dataset.label + ': ' + formatCurrency(context.raw);
                            }
                        }
                    }
                },
                scales: {
                    y: {
                        grid: { color: 'rgba(255, 255, 255, 0.05)' },
                        ticks: { color: '#94a3b8', callback: (v) => formatCurrency(v).split(',')[0] }
                    },
                    x: {
                        grid: { display: false },
                        ticks: { color: '#94a3b8' }
                    }
                }
            }
        });
    }

    function updateTotal() {
        if (historyData.length === 0) return;
        const latest = historyData[historyData.length - 1];
        let total = 0;

        document.querySelectorAll('.sim-controls input:checked').forEach(input => {
            const level = input.value;
            total += latest[level].invest;
        });

        const grandTotalEl = document.getElementById('grand-total');
        grandTotalEl.textContent = formatCurrency(total);
        
        // Animate total
        grandTotalEl.style.transform = 'scale(1.05)';
        setTimeout(() => grandTotalEl.style.transform = 'scale(1)', 200);
    }

    // Event Listeners
    const syncBtn = document.getElementById('sync-btn');
    const folderInput = document.getElementById('folder-input');

    syncBtn.addEventListener('click', () => folderInput.click());
    folderInput.addEventListener('change', handleFolderSelect);
    
    document.querySelectorAll('.sim-controls .checkbox-wrapper').forEach(wrapper => {
        wrapper.addEventListener('click', (e) => {
            const input = wrapper.querySelector('input');
            if (e.target !== input) {
                input.checked = !input.checked;
            }
            wrapper.classList.toggle('active', input.checked);
            updateTotal();
        });
    });

    // Removido fetchData automático inicial para forçar a seleção da pasta
});
