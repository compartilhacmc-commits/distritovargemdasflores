// ===================================
// CONFIGURAÇÃO DA PLANILHA (DUAS ABAS)
// ===================================
const SHEET_ID = '1r6NLcVkVLD5vp4UxPEa7TcreBpOd0qeNt-QREOG4Xr4';

// ✅ CONFIGURAÇÃO DAS DUAS ABAS
const SHEETS = [
    {
        name: 'PENDÊNCIAS ELDORADO',
        url: `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent('PENDÊNCIAS ELDORADO')}`
    },
    {
        name: 'RESOLVIDOS',
        url: `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent('RESOLVIDOS')}`
    }
];

// ===================================
// VARIÁVEIS GLOBAIS
// ===================================
let allData = [];
let filteredData = [];
let chartUnidades = null;
let chartEspecialidades = null;
let chartStatus = null;
let chartPizzaStatus = null;

// ✅ NOVOS GRÁFICOS
let chartPendenciasPrestador = null;
let chartPendenciasMes = null;

// ===================================
// FUNÇÃO AUXILIAR PARA BUSCAR VALOR DE COLUNA
// ===================================
function getColumnValue(item, possibleNames, defaultValue = '-') {
    for (let name of possibleNames) {
        if (item.hasOwnProperty(name) && item[name]) {
            return item[name];
        }
    }
    return defaultValue;
}

// ===================================
// ✅ REGRA DE PENDÊNCIA: COLUNA "USUÁRIO" PREENCHIDA
// ===================================
function isPendenciaByUsuario(item) {
    const usuario = getColumnValue(item, ['Usuário', 'Usuario', 'USUÁRIO', 'USUARIO'], '');
    return !!(usuario && String(usuario).trim() !== '');
}

// ===================================
// MULTISELECT (CHECKBOX) HELPERS
// ===================================
function toggleMultiSelect(id) {
    document.getElementById(id).classList.toggle('open');
}

// fecha dropdown ao clicar fora
document.addEventListener('click', (e) => {
    document.querySelectorAll('.multi-select').forEach(ms => {
        if (!ms.contains(e.target)) ms.classList.remove('open');
    });
});

function escapeHtml(str) {
    return String(str)
        .replaceAll('&', '&amp;')
        .replaceAll('<', '&lt;')
        .replaceAll('>', '&gt;')
        .replaceAll('"', '&quot;')
        .replaceAll("'", '&#039;');
}

function renderMultiSelect(panelId, values, onChange) {
    const panel = document.getElementById(panelId);
    panel.innerHTML = '';

    const actions = document.createElement('div');
    actions.className = 'ms-actions';
    actions.innerHTML = `
      <button type="button" class="ms-all">Marcar todos</button>
      <button type="button" class="ms-none">Limpar</button>
    `;
    panel.appendChild(actions);

    const btnAll = actions.querySelector('.ms-all');
    const btnNone = actions.querySelector('.ms-none');

    btnAll.addEventListener('click', () => {
        panel.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = true);
        onChange();
    });

    btnNone.addEventListener('click', () => {
        panel.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = false);
        onChange();
    });

    values.forEach(v => {
        const item = document.createElement('label');
        item.className = 'ms-item';
        item.innerHTML = `
          <input type="checkbox" value="${escapeHtml(v)}">
          <span>${escapeHtml(v)}</span>
        `;
        item.querySelector('input').addEventListener('change', onChange);
        panel.appendChild(item);
    });
}

function getSelectedFromPanel(panelId) {
    const panel = document.getElementById(panelId);
    return [...panel.querySelectorAll('input[type="checkbox"]:checked')].map(cb => cb.value);
}

function setMultiSelectText(textId, selected, fallbackLabel) {
    const el = document.getElementById(textId);
    if (!selected || selected.length === 0) el.textContent = fallbackLabel;
    else if (selected.length === 1) el.textContent = selected[0];
    else el.textContent = `${selected.length} selecionados`;
}

// ===================================
// INICIALIZAÇÃO
// ===================================
document.addEventListener('DOMContentLoaded', function() {
    console.log('Iniciando carregamento de dados...');
    loadData();
});

// ===================================
// ✅ CARREGAR DADOS DAS DUAS ABAS
// ===================================
async function loadData() {
    showLoading(true);
    allData = [];
    
    try {
        console.log('Carregando dados das duas abas...');

        // ✅ CARREGAR AS DUAS ABAS EM PARALELO
        const promises = SHEETS.map(sheet => 
            fetch(sheet.url)
                .then(response => {
                    if (!response.ok) {
                        throw new Error(`Erro HTTP na aba "${sheet.name}": ${response.status}`);
                    }
                    return response.text();
                })
                .then(csvText => {
                    console.log(`Dados CSV da aba "${sheet.name}" recebidos`);
                    return { name: sheet.name, csv: csvText };
                })
        );

        const results = await Promise.all(promises);

        // ✅ PROCESSAR CADA ABA
        results.forEach(result => {
            const rows = parseCSV(result.csv);

            if (rows.length < 2) {
                console.warn(`Aba "${result.name}" está vazia ou sem dados`);
                return;
            }

            const headers = rows[0];
            console.log(`Cabeçalhos da aba "${result.name}":`, headers);

            const sheetData = rows.slice(1)
                .filter(row => row.length > 1 && row[0])
                .map(row => {
                    const obj = { _origem: result.name }; // ✅ Marca a origem dos dados
                    headers.forEach((header, index) => {
                        obj[header.trim()] = (row[index] || '').trim();
                    });
                    return obj;
                });

            console.log(`${sheetData.length} registros carregados da aba "${result.name}"`);
            allData.push(...sheetData);
        });

        console.log(`Total de registros carregados (ambas as abas): ${allData.length}`);
        console.log('Primeiro registro completo:', allData[0]);

        if (allData.length === 0) {
            throw new Error('Nenhum dado foi carregado das planilhas');
        }

        filteredData = [...allData];

        populateFilters();
        updateDashboard();

        console.log('Dados carregados com sucesso!');

    } catch (error) {
        console.error('Erro ao carregar dados:', error);
        alert(`Erro ao carregar dados da planilha: ${error.message}\n\nVerifique:\n1. As planilhas estão públicas?\n2. Os nomes das abas estão corretos?\n3. Há dados nas planilhas?`);
    } finally {
        showLoading(false);
    }
}

// ===================================
// PARSE CSV (COM SUPORTE A ASPAS)
// ===================================
function parseCSV(text) {
    const rows = [];
    let currentRow = [];
    let currentCell = '';
    let insideQuotes = false;

    for (let i = 0; i < text.length; i++) {
        const char = text[i];
        const nextChar = text[i + 1];

        if (char === '"') {
            if (insideQuotes && nextChar === '"') {
                currentCell += '"';
                i++;
            } else {
                insideQuotes = !insideQuotes;
            }
        } else if (char === ',' && !insideQuotes) {
            currentRow.push(currentCell.trim());
            currentCell = '';
        } else if ((char === '\n' || char === '\r') && !insideQuotes) {
            if (currentCell || currentRow.length > 0) {
                currentRow.push(currentCell.trim());
                rows.push(currentRow);
                currentRow = [];
                currentCell = '';
            }
            if (char === '\r' && nextChar === '\n') {
                i++;
            }
        } else {
            currentCell += char;
        }
    }

    if (currentCell || currentRow.length > 0) {
        currentRow.push(currentCell.trim());
        rows.push(currentRow);
    }

    return rows;
}

// ===================================
// MOSTRAR/OCULTAR LOADING
// ===================================
function showLoading(show) {
    const overlay = document.getElementById('loadingOverlay');
    if (show) overlay.classList.add('active');
    else overlay.classList.remove('active');
}

// ===================================
// ✅ POPULAR FILTROS (MULTISELECT + MÊS)
// ===================================
function populateFilters() {
    const statusList = [...new Set(allData.map(item => item['Status']))].filter(Boolean).sort();
    renderMultiSelect('msStatusPanel', statusList, applyFilters);

    const unidades = [...new Set(allData.map(item => item['Unidade Solicitante']))].filter(Boolean).sort();
    renderMultiSelect('msUnidadePanel', unidades, applyFilters);

    const especialidades = [...new Set(allData.map(item => item['Cbo Especialidade']))].filter(Boolean).sort();
    renderMultiSelect('msEspecialidadePanel', especialidades, applyFilters);

    const prestadores = [...new Set(allData.map(item => item['Prestador']))].filter(Boolean).sort();
    renderMultiSelect('msPrestadorPanel', prestadores, applyFilters);

    setMultiSelectText('msStatusText', [], 'Todos');
    setMultiSelectText('msUnidadeText', [], 'Todas');
    setMultiSelectText('msEspecialidadeText', [], 'Todas');
    setMultiSelectText('msPrestadorText', [], 'Todos');

    // ✅ POPULAR FILTRO DE MÊS
    populateMonthFilter();
}

// ===================================
// ✅ POPULAR FILTRO DE MÊS (BASEADO EM "Data Início da Pendência")
// ===================================
function populateMonthFilter() {
    const selectMes = document.getElementById('filterMes');
    const mesesSet = new Set();

    allData.forEach(item => {
        const dataInicio = parseDate(getColumnValue(item, [
            'Data Início da Pendência',
            'Data Inicio da Pendencia',
            'Data Início Pendência',
            'Data Inicio Pendencia'
        ]));

        if (dataInicio) {
            const mesAno = `${dataInicio.getFullYear()}-${String(dataInicio.getMonth() + 1).padStart(2, '0')}`;
            mesesSet.add(mesAno);
        }
    });

    const mesesOrdenados = Array.from(mesesSet).sort().reverse();

    selectMes.innerHTML = '<option value="">Todos os Meses</option>';
    
    mesesOrdenados.forEach(mesAno => {
        const [ano, mes] = mesAno.split('-');
        const nomeMes = new Date(ano, mes - 1).toLocaleDateString('pt-BR', { month: 'long', year: 'numeric' });
        const option = document.createElement('option');
        option.value = mesAno;
        option.textContent = nomeMes.charAt(0).toUpperCase() + nomeMes.slice(1);
        selectMes.appendChild(option);
    });
}

// ===================================
// ✅ APLICAR FILTROS (MULTISELECT + MÊS)
// ===================================
function applyFilters() {
    const statusSel = getSelectedFromPanel('msStatusPanel');
    const unidadeSel = getSelectedFromPanel('msUnidadePanel');
    const especialidadeSel = getSelectedFromPanel('msEspecialidadePanel');
    const prestadorSel = getSelectedFromPanel('msPrestadorPanel');
    const mesSel = document.getElementById('filterMes').value;

    setMultiSelectText('msStatusText', statusSel, 'Todos');
    setMultiSelectText('msUnidadeText', unidadeSel, 'Todas');
    setMultiSelectText('msEspecialidadeText', especialidadeSel, 'Todas');
    setMultiSelectText('msPrestadorText', prestadorSel, 'Todos');

    filteredData = allData.filter(item => {
        const okStatus = (statusSel.length === 0) || statusSel.includes(item['Status'] || '');
        const okUnidade = (unidadeSel.length === 0) || unidadeSel.includes(item['Unidade Solicitante'] || '');
        const okEsp = (especialidadeSel.length === 0) || especialidadeSel.includes(item['Cbo Especialidade'] || '');
        const okPrest = (prestadorSel.length === 0) || prestadorSel.includes(item['Prestador'] || '');

        // ✅ FILTRO POR MÊS
        let okMes = true;
        if (mesSel) {
            const dataInicio = parseDate(getColumnValue(item, [
                'Data Início da Pendência',
                'Data Inicio da Pendencia',
                'Data Início Pendência',
                'Data Inicio Pendencia'
            ]));
            if (dataInicio) {
                const mesAnoItem = `${dataInicio.getFullYear()}-${String(dataInicio.getMonth() + 1).padStart(2, '0')}`;
                okMes = (mesAnoItem === mesSel);
            } else {
                okMes = false;
            }
        }

        return okStatus && okUnidade && okEsp && okPrest && okMes;
    });

    updateDashboard();
}

// ===================================
// ✅ LIMPAR FILTROS (MULTISELECT + MÊS)
// ===================================
function clearFilters() {
    ['msStatusPanel','msUnidadePanel','msEspecialidadePanel','msPrestadorPanel'].forEach(panelId => {
        const panel = document.getElementById(panelId);
        if (!panel) return;
        panel.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = false);
    });

    setMultiSelectText('msStatusText', [], 'Todos');
    setMultiSelectText('msUnidadeText', [], 'Todas');
    setMultiSelectText('msEspecialidadeText', [], 'Todas');
    setMultiSelectText('msPrestadorText', [], 'Todos');

    document.getElementById('searchInput').value = '';
    document.getElementById('filterMes').value = '';

    filteredData = [...allData];
    updateDashboard();
}

// ===================================
// PESQUISAR NA TABELA
// ===================================
function searchTable() {
    const searchValue = document.getElementById('searchInput').value.toLowerCase();
    const tbody = document.getElementById('tableBody');
    const rows = tbody.getElementsByTagName('tr');

    let visibleCount = 0;

    for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        const cells = row.getElementsByTagName('td');
        let found = false;

        for (let j = 0; j < cells.length; j++) {
            const cellText = cells[j].textContent.toLowerCase();
            if (cellText.includes(searchValue)) {
                found = true;
                break;
            }
        }

        if (found) {
            row.style.display = '';
            visibleCount++;
        } else {
            row.style.display = 'none';
        }
    }

    const footer = document.getElementById('tableFooter');
    footer.textContent = `Mostrando ${visibleCount} de ${filteredData.length} registros`;
}

// ===================================
// ATUALIZAR DASHBOARD
// ===================================
function updateDashboard() {
    updateCards();
    updateCharts();
    updateTable();
}

// ===================================
// ✅ ATUALIZAR CARDS (CONTANDO POR "USUÁRIO" PREENCHIDO)
// ===================================
function updateCards() {
    const total = allData.length;
    const filtrado = filteredData.length;

    const hoje = new Date();
    let pendencias15 = 0;
    let pendencias30 = 0;

    filteredData.forEach(item => {
        if (!isPendenciaByUsuario(item)) return;

        const dataInicio = parseDate(getColumnValue(item, [
            'Data Início da Pendência',
            'Data Inicio da Pendencia',
            'Data Início Pendência',
            'Data Inicio Pendencia'
        ]));

        if (dataInicio) {
            const diasDecorridos = Math.floor((hoje - dataInicio) / (1000 * 60 * 60 * 24));
            if (diasDecorridos >= 15 && diasDecorridos < 30) pendencias15++;
            if (diasDecorridos >= 30) pendencias30++;
        }
    });

    document.getElementById('totalPendencias').textContent = total;
    document.getElementById('pendencias15').textContent = pendencias15;
    document.getElementById('pendencias30').textContent = pendencias30;

    const percentFiltrados = total > 0 ? ((filtrado / total) * 100).toFixed(1) : '100.0';
    document.getElementById('percentFiltrados').textContent = percentFiltrados + '%';
}

// ===================================
// ✅ ATUALIZAR GRÁFICOS
// ===================================
function updateCharts() {
    // ✅ Gráfico de Unidades (SOMENTE COM "USUÁRIO" PREENCHIDO)
    const unidadesCount = {};
    filteredData.forEach(item => {
        if (!isPendenciaByUsuario(item)) return;
        const unidade = item['Unidade Solicitante'] || 'Não informado';
        unidadesCount[unidade] = (unidadesCount[unidade] || 0) + 1;
    });

    const unidadesLabels = Object.keys(unidadesCount)
        .sort((a, b) => unidadesCount[b] - unidadesCount[a])
        .slice(0, 50);
    const unidadesValues = unidadesLabels.map(label => unidadesCount[label]);

    createHorizontalBarChart('chartUnidades', unidadesLabels, unidadesValues, '#48bb78');

    // ✅ Gráfico de Especialidades (SOMENTE COM "USUÁRIO" PREENCHIDO)
    const especialidadesCount = {};
    filteredData.forEach(item => {
        if (!isPendenciaByUsuario(item)) return;
        const especialidade = item['Cbo Especialidade'] || 'Não informado';
        especialidadesCount[especialidade] = (especialidadesCount[especialidade] || 0) + 1;
    });

    const especialidadesLabels = Object.keys(especialidadesCount)
        .sort((a, b) => especialidadesCount[b] - especialidadesCount[a])
        .slice(0, 50);
    const especialidadesValues = especialidadesLabels.map(label => especialidadesCount[label]);

    createHorizontalBarChart('chartEspecialidades', especialidadesLabels, especialidadesValues, '#ef4444');

    // ✅ GRÁFICO DE STATUS (mantido)
    const statusCount = {};
    filteredData.forEach(item => {
        const status = item['Status'] || 'Não informado';
        statusCount[status] = (statusCount[status] || 0) + 1;
    });

    const statusLabels = Object.keys(statusCount)
        .sort((a, b) => statusCount[b] - statusCount[a]);
    const statusValues = statusLabels.map(label => statusCount[label]);

    createVerticalBarChart('chartStatus', statusLabels, statusValues, '#f97316');

    // ✅ GRÁFICO DE PIZZA (mantido)
    createPieChart('chartPizzaStatus', statusLabels, statusValues);

    // ✅ NOVO: PENDÊNCIAS POR PRESTADOR (VERTICAL, USUÁRIO PREENCHIDO)
    const prestadorCount = {};
    filteredData.forEach(item => {
        if (!isPendenciaByUsuario(item)) return;
        const prest = item['Prestador'] || 'Não informado';
        prestadorCount[prest] = (prestadorCount[prest] || 0) + 1;
    });

    const prestLabels = Object.keys(prestadorCount)
        .sort((a, b) => prestadorCount[b] - prestadorCount[a])
        .slice(0, 50);
    const prestValues = prestLabels.map(l => prestadorCount[l]);

    createVerticalBarChartCenteredValue(
        'chartPendenciasPrestador',
        prestLabels,
        prestValues,
        '#4c1d95' // roxo escuro
    );

    // ✅ NOVO: PENDÊNCIAS POR MÊS (VERTICAL, USUÁRIO PREENCHIDO)
    const mesCount = {};
    filteredData.forEach(item => {
        if (!isPendenciaByUsuario(item)) return;

        const dataInicio = parseDate(getColumnValue(item, [
            'Data Início da Pendência',
            'Data Inicio da Pendencia',
            'Data Início Pendência',
            'Data Inicio Pendencia'
        ]));

        let chave = 'Não informado';
        if (dataInicio) {
            const mesAno = `${dataInicio.getFullYear()}-${String(dataInicio.getMonth() + 1).padStart(2, '0')}`;
            const [ano, mes] = mesAno.split('-');
            const nomeMes = new Date(ano, mes - 1).toLocaleDateString('pt-BR', { month: 'long', year: 'numeric' });
            chave = nomeMes.charAt(0).toUpperCase() + nomeMes.slice(1);
        }

        mesCount[chave] = (mesCount[chave] || 0) + 1;
    });

    const mesLabels = Object.keys(mesCount)
        .sort((a, b) => mesCount[b] - mesCount[a])
        .slice(0, 50);
    const mesValues = mesLabels.map(l => mesCount[l]);

    createVerticalBarChartCenteredValue(
        'chartPendenciasMes',
        mesLabels,
        mesValues,
        '#0b2a6f' // azul escuro
    );
}

// ===================================
// CRIAR GRÁFICO DE BARRAS HORIZONTAIS
// ===================================
function createHorizontalBarChart(canvasId, labels, data, color) {
    const ctx = document.getElementById(canvasId);

    if (canvasId === 'chartUnidades' && chartUnidades) chartUnidades.destroy();
    if (canvasId === 'chartEspecialidades' && chartEspecialidades) chartEspecialidades.destroy();

    const chart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                label: 'Quantidade',
                data: data,
                backgroundColor: color,
                borderWidth: 0,
                borderRadius: 4,
                barPercentage: 0.75,
                categoryPercentage: 0.85
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                tooltip: {
                    enabled: true,
                    backgroundColor: 'rgba(0, 0, 0, 0.8)',
                    titleFont: { size: 14, weight: 'bold' },
                    bodyFont: { size: 13 },
                    padding: 12,
                    cornerRadius: 8
                }
            },
            scales: {
                x: { display: false, grid: { display: false } },
                y: {
                    ticks: {
                        font: { size: 12, weight: '500' },
                        color: '#4a5568',
                        padding: 8
                    },
                    grid: { display: false }
                }
            },
            layout: { padding: { right: 50 } }
        },
        plugins: [{
            id: 'customLabels',
            afterDatasetsDraw: function(chart) {
                const ctx = chart.ctx;
                chart.data.datasets.forEach(function(dataset, i) {
                    const meta = chart.getDatasetMeta(i);
                    if (!meta.hidden) {
                        meta.data.forEach(function(element, index) {
                            ctx.fillStyle = '#000000';
                            ctx.font = 'bold 14px Arial';
                            ctx.textAlign = 'left';
                            ctx.textBaseline = 'middle';

                            const dataString = dataset.data[index].toString();
                            const xPos = element.x + 10;
                            const yPos = element.y;

                            ctx.fillText(dataString, xPos, yPos);
                        });
                    }
                });
            }
        }]
    });

    if (canvasId === 'chartUnidades') chartUnidades = chart;
    if (canvasId === 'chartEspecialidades') chartEspecialidades = chart;
}

// ===================================
// ✅ NOVO: GRÁFICO VERTICAL COM VALOR NO MEIO DA BARRA (BRANCO E NEGRITO)
// ===================================
function createVerticalBarChartCenteredValue(canvasId, labels, data, color) {
    const ctx = document.getElementById(canvasId);

    if (canvasId === 'chartPendenciasPrestador' && chartPendenciasPrestador) chartPendenciasPrestador.destroy();
    if (canvasId === 'chartPendenciasMes' && chartPendenciasMes) chartPendenciasMes.destroy();

    const chart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels,
            datasets: [{
                label: 'Quantidade',
                data,
                backgroundColor: color,
                borderWidth: 0,
                borderRadius: 6,
                barPercentage: 0.70,
                categoryPercentage: 0.75,
                maxBarThickness: 40
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                tooltip: {
                    enabled: true,
                    backgroundColor: 'rgba(0,0,0,0.85)',
                    titleFont: { size: 14, weight: 'bold' },
                    bodyFont: { size: 13 },
                    padding: 12,
                    cornerRadius: 8
                }
            },
            scales: {
                x: {
                    ticks: {
                        font: { size: 12, weight: '600' },
                        color: '#4a5568',
                        maxRotation: 45,
                        minRotation: 0
                    },
                    grid: { display: false }
                },
                y: {
                    beginAtZero: true,
                    ticks: {
                        font: { size: 12, weight: '600' },
                        color: '#4a5568'
                    },
                    grid: { color: 'rgba(0,0,0,0.06)' }
                }
            }
        },
        plugins: [{
            id: 'centerValueInsideVerticalBar',
            afterDatasetsDraw(chart) {
                const { ctx } = chart;
                const meta = chart.getDatasetMeta(0);
                const dataset = chart.data.datasets[0];

                ctx.save();
                ctx.fillStyle = '#FFFFFF';
                ctx.font = 'bold 14px Arial';
                ctx.textAlign = 'center';
                ctx.textBaseline = 'middle';

                meta.data.forEach((bar, i) => {
                    const value = dataset.data[i];
                    const centerX = bar.x;
                    const centerY = bar.y + (bar.height / 2);
                    ctx.fillText(String(value), centerX, centerY);
                });

                ctx.restore();
            }
        }]
    });

    if (canvasId === 'chartPendenciasPrestador') chartPendenciasPrestador = chart;
    if (canvasId === 'chartPendenciasMes') chartPendenciasMes = chart;
}

// ===================================
// ✅ CRIAR GRÁFICO DE BARRAS VERTICAIS (STATUS) COM VALORES NO MEIO DAS BARRAS
// ===================================
function createVerticalBarChart(canvasId, labels, data, color) {
    const ctx = document.getElementById(canvasId);

    if (chartStatus) chartStatus.destroy();

    const chart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels,
            datasets: [{
                label: 'Quantidade',
                data,
                backgroundColor: color,
                borderWidth: 0,
                borderRadius: 6,
                barPercentage: 0.55,
                categoryPercentage: 0.70,
                maxBarThickness: 28
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                tooltip: {
                    enabled: true,
                    backgroundColor: 'rgba(0,0,0,0.85)',
                    titleFont: { size: 14, weight: 'bold' },
                    bodyFont: { size: 13 },
                    padding: 12,
                    cornerRadius: 8
                }
            },
            scales: {
                x: {
                    ticks: {
                        font: { size: 12, weight: '600' },
                        color: '#4a5568',
                        maxRotation: 45,
                        minRotation: 0
                    },
                    grid: { display: false }
                },
                y: {
                    beginAtZero: true,
                    ticks: {
                        font: { size: 12, weight: '600' },
                        color: '#4a5568'
                    },
                    grid: { color: 'rgba(0,0,0,0.06)' }
                }
            }
        },
        plugins: [{
            id: 'statusValueLabelsInsideBar',
            afterDatasetsDraw(chart) {
                const { ctx } = chart;
                const meta = chart.getDatasetMeta(0);
                const dataset = chart.data.datasets[0];

                ctx.save();
                ctx.fillStyle = '#FFFFFF';
                ctx.font = 'bold 16px Arial';
                ctx.textAlign = 'center';
                ctx.textBaseline = 'middle';

                meta.data.forEach((bar, i) => {
                    const value = dataset.data[i];
                    const yPos = bar.y + (bar.height / 2);
                    ctx.fillText(String(value), bar.x, yPos);
                });

                ctx.restore();
            }
        }]
    });

    chartStatus = chart;
}

// ===================================
// ✅ CRIAR GRÁFICO DE PIZZA COM LEGENDA COMPLETA (TEXTO + BOLINHA)
// ===================================
function createPieChart(canvasId, labels, data) {
    const ctx = document.getElementById(canvasId);

    if (chartPizzaStatus) chartPizzaStatus.destroy();

    const colors = [
        '#3b82f6', '#ef4444', '#10b981', '#f59e0b', '#8b5cf6',
        '#ec4899', '#06b6d4', '#f97316', '#6366f1', '#84cc16'
    ];

    const total = data.reduce((sum, val) => sum + val, 0);

    chartPizzaStatus = new Chart(ctx, {
        type: 'pie',
        data: {
            labels: labels,
            datasets: [{
                data: data,
                backgroundColor: colors.slice(0, labels.length),
                borderWidth: 3,
                borderColor: '#ffffff'
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    display: true,
                    position: 'right',
                    labels: {
                        font: { 
                            size: 14,
                            weight: 'bold',
                            family: 'Arial, sans-serif'
                        },
                        color: '#000000',
                        padding: 15,
                        usePointStyle: true,
                        pointStyle: 'circle',
                        boxWidth: 20,
                        boxHeight: 20,
                        generateLabels: function(chart) {
                            const datasets = chart.data.datasets;
                            const labels = chart.data.labels;
                            
                            return labels.map((label, i) => {
                                const value = datasets[0].data[i];
                                const percentage = total > 0 ? ((value / total) * 100).toFixed(1) : '0.0';
                                
                                return {
                                    text: `${label} (${percentage}%)`,
                                    fillStyle: datasets[0].backgroundColor[i],
                                    strokeStyle: datasets[0].backgroundColor[i],
                                    lineWidth: 2,
                                    hidden: false,
                                    index: i
                                };
                            });
                        }
                    }
                },
                tooltip: {
                    enabled: true,
                    backgroundColor: 'rgba(0, 0, 0, 0.8)',
                    titleFont: { size: 14, weight: 'bold' },
                    bodyFont: { size: 13 },
                    padding: 12,
                    cornerRadius: 8,
                    callbacks: {
                        label: function(context) {
                            const value = context.parsed;
                            const percentage = total > 0 ? ((value / total) * 100).toFixed(1) : '0.0';
                            return `${context.label}: ${percentage}% (${value} registros)`;
                        }
                    }
                }
            }
        },
        plugins: [{
            id: 'customPieLabelsInside',
            afterDatasetsDraw: function(chart) {
                const ctx = chart.ctx;
                const dataset = chart.data.datasets[0];
                const meta = chart.getDatasetMeta(0);
                
                ctx.save();
                ctx.font = 'bold 14px Arial';
                ctx.textAlign = 'center';
                ctx.textBaseline = 'middle';
                
                meta.data.forEach(function(element, index) {
                    const value = dataset.data[index];
                    const percentage = total > 0 ? ((value / total) * 100).toFixed(1) : '0.0';
                    
                    if (parseFloat(percentage) > 5) {
                        ctx.fillStyle = '#ffffff';
                        const position = element.tooltipPosition();
                        ctx.fillText(`${percentage}%`, position.x, position.y);
                    }
                });
                
                ctx.restore();
            }
        }]
    });
}

// ===================================
// ✅ ATUALIZAR TABELA SEM A COLUNA "SOLICITAÇÃO"
// ===================================
function updateTable() {
    const tbody = document.getElementById('tableBody');
    const footer = document.getElementById('tableFooter');
    tbody.innerHTML = '';

    if (filteredData.length === 0) {
        tbody.innerHTML = '<tr><td colspan="12" class="loading-message"><i class="fas fa-inbox"></i> Nenhum registro encontrado</td></tr>';
        footer.textContent = 'Mostrando 0 registros';
        return;
    }

    const hoje = new Date();

    filteredData.forEach(item => {
        const row = document.createElement('tr');

        const origem = item['_origem'] || '-';

        const dataSolicitacao = getColumnValue(item, [
            'Data da Solicitação',
            'Data Solicitação',
            'Data da Solicitacao',
            'Data Solicitacao'
        ]);

        const prontuario = getColumnValue(item, [
            'Nº Prontuário',
            'N° Prontuário',
            'Numero Prontuário',
            'Prontuário',
            'Prontuario'
        ]);

        const dataInicioStr = getColumnValue(item, [
            'Data Início da Pendência',
            'Data Inicio da Pendencia',
            'Data Início Pendência',
            'Data Inicio Pendencia'
        ]);

        const prazo15 = getColumnValue(item, [
            'Data Final do Prazo (Pendência com 15 dias)',
            'Data Final do Prazo (Pendencia com 15 dias)',
            'Data Final Prazo 15d',
            'Prazo 15 dias'
        ]);

        const email15 = getColumnValue(item, [
            'Data do envio do Email (Prazo: Pendência com 15 dias)',
            'Data do envio do Email (Prazo: Pendencia com 15 dias)',
            'Data Envio Email 15d',
            'Email 15 dias'
        ]);

        const prazo30 = getColumnValue(item, [
            'Data Final do Prazo (Pendência com 30 dias)',
            'Data Final do Prazo (Pendencia com 30 dias)',
            'Data Final Prazo 30d',
            'Prazo 30 dias'
        ]);

        const email30 = getColumnValue(item, [
            'Data do envio do Email (Prazo: Pendência com 30 dias)',
            'Data do envio do Email (Prazo: Pendencia com 30 dias)',
            'Data Envio Email 30d',
            'Email 30 dias'
        ]);

        // ✅ VERIFICAR SE ESTÁ VENCENDO EM 15 DIAS (entre 15 e 30 dias)
        const dataInicio = parseDate(dataInicioStr);
        let isVencendo15 = false;
        if (dataInicio) {
            const diasDecorridos = Math.floor((hoje - dataInicio) / (1000 * 60 * 60 * 24));
            if (diasDecorridos >= 15 && diasDecorridos < 30) {
                isVencendo15 = true;
            }
        }

        row.innerHTML = `
            <td>${origem}</td>
            <td>${formatDate(dataSolicitacao)}</td>
            <td>${prontuario}</td>
            <td>${item['Telefone'] || '-'}</td>
            <td>${item['Unidade Solicitante'] || '-'}</td>
            <td>${item['Cbo Especialidade'] || '-'}</td>
            <td>${formatDate(dataInicioStr)}</td>
            <td>${item['Status'] || '-'}</td>
            <td>${formatDate(prazo15)}</td>
            <td>${formatDate(email15)}</td>
            <td>${formatDate(prazo30)}</td>
            <td>${formatDate(email30)}</td>
        `;

        // ✅ APLICAR DESTAQUE AMARELO SE VENCENDO EM 15 DIAS
        if (isVencendo15) {
            row.classList.add('row-vencendo-15');
        }

        tbody.appendChild(row);
    });

    const total = allData.length;
    const showing = filteredData.length;
    footer.textContent = `Mostrando de 1 até ${showing} de ${total} registros`;
}

// ===================================
// FUNÇÕES AUXILIARES
// ===================================
function parseDate(dateString) {
    if (!dateString || dateString === '-') return null;

    let match = dateString.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (match) return new Date(match[3], match[2] - 1, match[1]);

    match = dateString.match(/(\d{4})-(\d{2})-(\d{2})/);
    if (match) return new Date(match[1], match[2] - 1, match[3]);

    return null;
}

function formatDate(dateString) {
    if (!dateString || dateString === '-') return '-';

    const date = parseDate(dateString);
    if (!date || isNaN(date.getTime())) return dateString;

    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();

    return `${day}/${month}/${year}`;
}

// ===================================
// ATUALIZAR DADOS
// ===================================
function refreshData() {
    loadData();
}

// ===================================
// DOWNLOAD EXCEL (SEM COLUNA "SOLICITAÇÃO")
// ===================================
function downloadExcel() {
    if (filteredData.length === 0) {
        alert('Não há dados para exportar.');
        return;
    }

    const exportData = filteredData.map(item => ({
        'Origem': item['_origem'] || '',
        'Data Solicitação': getColumnValue(item, ['Data da Solicitação', 'Data Solicitação', 'Data da Solicitacao', 'Data Solicitacao'], ''),
        'Nº Prontuário': getColumnValue(item, ['Nº Prontuário', 'N° Prontuário', 'Numero Prontuário', 'Prontuário', 'Prontuario'], ''),
        'Telefone': item['Telefone'] || '',
        'Unidade Solicitante': item['Unidade Solicitante'] || '',
        'CBO Especialidade': item['Cbo Especialidade'] || '',
        'Data Início Pendência': getColumnValue(item, ['Data Início da Pendência','Data Início Pendência','Data Inicio da Pendencia','Data Inicio Pendencia'], ''),
        'Status': item['Status'] || '',
        'Prestador': item['Prestador'] || '',
        'Data Final Prazo 15d': getColumnValue(item, ['Data Final do Prazo (Pendência com 15 dias)','Data Final do Prazo (Pendencia com 15 dias)','Data Final Prazo 15d','Prazo 15 dias'], ''),
        'Data Envio Email 15d': getColumnValue(item, ['Data do envio do Email (Prazo: Pendência com 15 dias)','Data do envio do Email (Prazo: Pendencia com 15 dias)','Data Envio Email 15d','Email 15 dias'], ''),
        'Data Final Prazo 30d': getColumnValue(item, ['Data Final do Prazo (Pendência com 30 dias)','Data Final do Prazo (Pendencia com 30 dias)','Data Final Prazo 30d','Prazo 30 dias'], ''),
        'Data Envio Email 30d': getColumnValue(item, ['Data do envio do Email (Prazo: Pendência com 30 dias)','Data do envio do Email (Prazo: Pendencia com 30 dias)','Data Envio Email 30d','Email 30 dias'], '')
    }));

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Dados Completos');

    ws['!cols'] = [
        { wch: 20 }, { wch: 18 }, { wch: 15 }, { wch: 15 },
        { wch: 30 }, { wch: 30 }, { wch: 18 }, { wch: 20 },
        { wch: 25 }, { wch: 18 }, { wch: 20 }, { wch: 18 }, { wch: 20 }
    ];

    const hoje = new Date().toISOString().split('T')[0];
    XLSX.writeFile(wb, `Dados_Eldorado_${hoje}.xlsx`);
}
