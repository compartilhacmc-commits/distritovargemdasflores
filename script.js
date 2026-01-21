// ===================================
// CONFIGURAÇÃO DA PLANILHA (DUAS ABAS)
// ===================================

// ID da planilha (mantido)
const SHEET_ID = '1lMGO9Hh_qL9OKI270fPL7lxadr-BZN9x_ZtmQeX6OcA';

// ✅ IMPORTANTE:
// Para carregar dados via fetch no GitHub Pages, use SEMPRE o endpoint de exportação CSV.
// Formato:
// https://docs.google.com/spreadsheets/d/ID/export?format=csv&gid=GID

const SPREADSHEET_ID_REAL = '1IHknmxe3xAnfy5Bju_23B5ivIL-qMaaE6q_HuPaLBpk';

const SHEETS = [
    {
        name: 'PENDÊNCIAS VARGEM DAS FLORES',
        gid: '278071504',
        url: `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID_REAL}/export?format=csv&gid=278071504`
    },
    {
        name: 'RESOLVIDOS',
        gid: '451254610',
        url: `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID_REAL}/export?format=csv&gid=451254610`
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
// FUNÇÕES DE NORMALIZAÇÃO (para não quebrar com variação de cabeçalho)
// ===================================
function normalizeKey(str) {
    return String(str || '')
        .trim()
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '') // remove acentos
        .replace(/\s+/g, ' ');
}

function getValueByHeaderAliases(item, aliases, defaultValue = '-') {
    // Tenta direto primeiro (rápido)
    for (const a of aliases) {
        if (Object.prototype.hasOwnProperty.call(item, a) && item[a] !== '' && item[a] != null) return item[a];
    }

    // Tenta por normalização (robusto)
    const wanted = aliases.map(normalizeKey);
    for (const key of Object.keys(item)) {
        const nk = normalizeKey(key);
        if (wanted.includes(nk)) {
            const v = item[key];
            if (v !== '' && v != null) return v;
        }
    }

    return defaultValue;
}

// ===================================
// ✅ REGRA DE PENDÊNCIA: COLUNA "USUÁRIO" PREENCHIDA
// ===================================
function isPendenciaByUsuario(item) {
    const usuario = getValueByHeaderAliases(item, ['Usuário', 'Usuario', 'USUÁRIO', 'USUARIO'], '');
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
        console.log('Carregando dados das duas abas (CSV export)...');
        console.log('Planilhas:', SHEETS.map(s => s.url));

        const promises = SHEETS.map(sheet =>
            fetch(sheet.url, { cache: 'no-store' })
                .then(response => {
                    if (!response.ok) {
                        throw new Error(`Erro HTTP na aba "${sheet.name}": ${response.status}`);
                    }
                    return response.text();
                })
                .then(csvText => {
                    // Se por algum motivo vier HTML (planilha não pública), acusar cedo
                    const head = csvText.slice(0, 200).toLowerCase();
                    if (head.includes('<html') || head.includes('<!doctype')) {
                        throw new Error(`A aba "${sheet.name}" não retornou CSV. Verifique se a planilha está pública e o link é export?format=csv&gid=...`);
                    }
                    console.log(`CSV recebido da aba "${sheet.name}" (${csvText.length} chars)`);
                    return { name: sheet.name, csv: csvText };
                })
        );

        const results = await Promise.all(promises);

        results.forEach(result => {
            const rows = parseCSV(result.csv);

            if (!rows || rows.length < 2) {
                console.warn(`Aba "${result.name}" está vazia ou sem dados`);
                return;
            }

            const headers = rows[0].map(h => String(h || '').trim());
            console.log(`Cabeçalhos da aba "${result.name}":`, headers);

            const sheetData = rows.slice(1)
                .filter(row => row && row.length > 0 && String(row.join('')).trim() !== '')
                .map(row => {
                    const obj = { _origem: result.name };
                    headers.forEach((header, index) => {
                        obj[header] = (row[index] || '').trim();
                    });
                    return obj;
                });

            console.log(`${sheetData.length} registros carregados da aba "${result.name}"`);
            allData.push(...sheetData);
        });

        console.log(`Total de registros carregados (ambas as abas): ${allData.length}`);
        console.log('Primeiro registro:', allData[0]);

        if (allData.length === 0) {
            throw new Error('Nenhum dado foi carregado das planilhas (CSV).');
        }

        filteredData = [...allData];

        populateFilters();
        updateDashboard();

        console.log('Dados carregados e painel atualizado com sucesso!');

    } catch (error) {
        console.error('Erro ao carregar dados:', error);
        alert(
            `Erro ao carregar dados da planilha: ${error.message}\n\n` +
            `Verifique:\n` +
            `1) A planilha está "Publicar na web" / visível para qualquer pessoa com o link?\n` +
            `2) Os GIDs estão corretos?\n` +
            `3) As abas têm dados (linha de cabeçalho + linhas abaixo)?`
        );
        // mantém tabela com mensagem amigável
        const tbody = document.getElementById('tableBody');
        if (tbody) {
            tbody.innerHTML = '<tr><td colspan="12" class="loading-message"><i class="fas fa-exclamation-triangle"></i> Erro ao carregar dados</td></tr>';
        }
    } finally {
        showLoading(false);
    }
}

// ===================================
// PARSE CSV (COM SUPORTE A ASPAS) + AUTO-DETECT ; ou ,
// ===================================
function detectDelimiter(text) {
    const sample = text.slice(0, 2000);
    const commas = (sample.match(/,/g) || []).length;
    const semis = (sample.match(/;/g) || []).length;
    // Google Sheets pt-BR frequentemente usa ;, então prioriza o maior
    return semis > commas ? ';' : ',';
}

function parseCSV(text) {
    const delimiter = detectDelimiter(text);
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
        } else if (char === delimiter && !insideQuotes) {
            currentRow.push(currentCell);
            currentCell = '';
        } else if ((char === '\n' || char === '\r') && !insideQuotes) {
            if (currentCell !== '' || currentRow.length > 0) {
                currentRow.push(currentCell);
                rows.push(currentRow.map(c => String(c ?? '').trim()));
                currentRow = [];
                currentCell = '';
            }
            if (char === '\r' && nextChar === '\n') i++;
        } else {
            currentCell += char;
        }
    }

    if (currentCell !== '' || currentRow.length > 0) {
        currentRow.push(currentCell);
        rows.push(currentRow.map(c => String(c ?? '').trim()));
    }

    // remove linhas totalmente vazias
    return rows.filter(r => r.some(cell => String(cell).trim() !== ''));
}

// ===================================
// MOSTRAR/OCULTAR LOADING
// ===================================
function showLoading(show) {
    const overlay = document.getElementById('loadingOverlay');
    if (!overlay) return;
    if (show) overlay.classList.add('active');
    else overlay.classList.remove('active');
}

// ===================================
// ✅ POPULAR FILTROS (MULTISELECT + MÊS)
// ===================================
function populateFilters() {
    const statusList = [...new Set(allData.map(item => getValueByHeaderAliases(item, ['Status'], '')))]
        .filter(Boolean).sort();
    renderMultiSelect('msStatusPanel', statusList, applyFilters);

    const unidades = [...new Set(allData.map(item => getValueByHeaderAliases(item, ['Unidade Solicitante'], '')))]
        .filter(Boolean).sort();
    renderMultiSelect('msUnidadePanel', unidades, applyFilters);

    const especialidades = [...new Set(allData.map(item => getValueByHeaderAliases(item, ['Cbo Especialidade', 'CBO Especialidade'], '')))]
        .filter(Boolean).sort();
    renderMultiSelect('msEspecialidadePanel', especialidades, applyFilters);

    const prestadores = [...new Set(allData.map(item => getValueByHeaderAliases(item, ['Prestador'], '')))]
        .filter(Boolean).sort();
    renderMultiSelect('msPrestadorPanel', prestadores, applyFilters);

    setMultiSelectText('msStatusText', [], 'Todos');
    setMultiSelectText('msUnidadeText', [], 'Todas');
    setMultiSelectText('msEspecialidadeText', [], 'Todas');
    setMultiSelectText('msPrestadorText', [], 'Todos');

    populateMonthFilter();
}

// ===================================
// ✅ POPULAR FILTRO DE MÊS (BASEADO EM "Data Início da Pendência")
// ===================================
function populateMonthFilter() {
    const selectMes = document.getElementById('filterMes');
    if (!selectMes) return;

    const mesesSet = new Set();

    allData.forEach(item => {
        const dataInicio = parseDate(getValueByHeaderAliases(item, [
            'Data Início da Pendência',
            'Data Inicio da Pendencia',
            'Data Início Pendência',
            'Data Inicio Pendencia'
        ], ''));

        if (dataInicio) {
            const mesAno = `${dataInicio.getFullYear()}-${String(dataInicio.getMonth() + 1).padStart(2, '0')}`;
            mesesSet.add(mesAno);
        }
    });

    const mesesOrdenados = Array.from(mesesSet).sort().reverse();

    selectMes.innerHTML = '<option value="">Todos os Meses</option>';

    mesesOrdenados.forEach(mesAno => {
        const [ano, mes] = mesAno.split('-');
        const nomeMes = new Date(Number(ano), Number(mes) - 1).toLocaleDateString('pt-BR', { month: 'long', year: 'numeric' });
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
        const status = getValueByHeaderAliases(item, ['Status'], '');
        const unidade = getValueByHeaderAliases(item, ['Unidade Solicitante'], '');
        const esp = getValueByHeaderAliases(item, ['Cbo Especialidade', 'CBO Especialidade'], '');
        const prest = getValueByHeaderAliases(item, ['Prestador'], '');

        const okStatus = (statusSel.length === 0) || statusSel.includes(status);
        const okUnidade = (unidadeSel.length === 0) || unidadeSel.includes(unidade);
        const okEsp = (especialidadeSel.length === 0) || especialidadeSel.includes(esp);
        const okPrest = (prestadorSel.length === 0) || prestadorSel.includes(prest);

        let okMes = true;
        if (mesSel) {
            const dataInicio = parseDate(getValueByHeaderAliases(item, [
                'Data Início da Pendência',
                'Data Inicio da Pendencia',
                'Data Início Pendência',
                'Data Inicio Pendencia'
            ], ''));
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

        const dataInicio = parseDate(getValueByHeaderAliases(item, [
            'Data Início da Pendência',
            'Data Inicio da Pendencia',
            'Data Início Pendência',
            'Data Inicio Pendencia'
        ], ''));

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
    // Unidades (somente usuário preenchido)
    const unidadesCount = {};
    filteredData.forEach(item => {
        if (!isPendenciaByUsuario(item)) return;
        const unidade = getValueByHeaderAliases(item, ['Unidade Solicitante'], 'Não informado') || 'Não informado';
        unidadesCount[unidade] = (unidadesCount[unidade] || 0) + 1;
    });

    const unidadesLabels = Object.keys(unidadesCount)
        .sort((a, b) => unidadesCount[b] - unidadesCount[a])
        .slice(0, 50);
    const unidadesValues = unidadesLabels.map(label => unidadesCount[label]);

    createHorizontalBarChart('chartUnidades', unidadesLabels, unidadesValues, '#48bb78');

    // Especialidades (somente usuário preenchido)
    const especialidadesCount = {};
    filteredData.forEach(item => {
        if (!isPendenciaByUsuario(item)) return;
        const especialidade = getValueByHeaderAliases(item, ['Cbo Especialidade', 'CBO Especialidade'], 'Não informado') || 'Não informado';
        especialidadesCount[especialidade] = (especialidadesCount[especialidade] || 0) + 1;
    });

    const especialidadesLabels = Object.keys(especialidadesCount)
        .sort((a, b) => especialidadesCount[b] - especialidadesCount[a])
        .slice(0, 50);
    const especialidadesValues = especialidadesLabels.map(label => especialidadesCount[label]);

    createHorizontalBarChart('chartEspecialidades', especialidadesLabels, especialidadesValues, '#ef4444');

    // Status
    const statusCount = {};
    filteredData.forEach(item => {
        const status = getValueByHeaderAliases(item, ['Status'], 'Não informado') || 'Não informado';
        statusCount[status] = (statusCount[status] || 0) + 1;
    });

    const statusLabels = Object.keys(statusCount)
        .sort((a, b) => statusCount[b] - statusCount[a]);
    const statusValues = statusLabels.map(label => statusCount[label]);

    createVerticalBarChart('chartStatus', statusLabels, statusValues, '#f97316');

    // Pizza
    createPieChart('chartPizzaStatus', statusLabels, statusValues);

    // Prestador (vertical, usuário preenchido)
    const prestadorCount = {};
    filteredData.forEach(item => {
        if (!isPendenciaByUsuario(item)) return;
        const prest = getValueByHeaderAliases(item, ['Prestador'], 'Não informado') || 'Não informado';
        prestadorCount[prest] = (prestadorCount[prest] || 0) + 1;
    });

    const prestLabels = Object.keys(prestadorCount)
        .sort((a, b) => prestadorCount[b] - prestadorCount[a])
        .slice(0, 50);
    const prestValues = prestLabels.map(l => prestadorCount[l]);

    createVerticalBarChartCenteredValue('chartPendenciasPrestador', prestLabels, prestValues, '#4c1d95');

    // Mês (vertical, usuário preenchido)
    const mesCount = {};
    filteredData.forEach(item => {
        if (!isPendenciaByUsuario(item)) return;

        const dataInicio = parseDate(getValueByHeaderAliases(item, [
            'Data Início da Pendência',
            'Data Inicio da Pendencia',
            'Data Início Pendência',
            'Data Inicio Pendencia'
        ], ''));

        let chave = 'Não informado';
        if (dataInicio) {
            const ano = dataInicio.getFullYear();
            const mes = dataInicio.getMonth();
            const nomeMes = new Date(ano, mes, 1).toLocaleDateString('pt-BR', { month: 'long', year: 'numeric' });
            chave = nomeMes.charAt(0).toUpperCase() + nomeMes.slice(1);
        }

        mesCount[chave] = (mesCount[chave] || 0) + 1;
    });

    const mesLabels = Object.keys(mesCount)
        .sort((a, b) => mesCount[b] - mesCount[a])
        .slice(0, 50);
    const mesValues = mesLabels.map(l => mesCount[l]);

    createVerticalBarChartCenteredValue('chartPendenciasMes', mesLabels, mesValues, '#0b2a6f');
}

// ===================================
// CRIAR GRÁFICO DE BARRAS HORIZONTAIS
// ===================================
function createHorizontalBarChart(canvasId, labels, data, color) {
    const canvas = document.getElementById(canvasId);
    if (!canvas) return;

    if (canvasId === 'chartUnidades' && chartUnidades) chartUnidades.destroy();
    if (canvasId === 'chartEspecialidades' && chartEspecialidades) chartEspecialidades.destroy();

    const chart = new Chart(canvas, {
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

                            const dataString = String(dataset.data[index]);
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
    const canvas = document.getElementById(canvasId);
    if (!canvas) return;

    if (canvasId === 'chartPendenciasPrestador' && chartPendenciasPrestador) chartPendenciasPrestador.destroy();
    if (canvasId === 'chartPendenciasMes' && chartPendenciasMes) chartPendenciasMes.destroy();

    const chart = new Chart(canvas, {
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
    const canvas = document.getElementById(canvasId);
    if (!canvas) return;

    if (chartStatus) chartStatus.destroy();

    const chart = new Chart(canvas, {
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
    const canvas = document.getElementById(canvasId);
    if (!canvas) return;

    if (chartPizzaStatus) chartPizzaStatus.destroy();

    const colors = [
        '#3b82f6', '#ef4444', '#10b981', '#f59e0b', '#8b5cf6',
        '#ec4899', '#06b6d4', '#f97316', '#6366f1', '#84cc16'
    ];

    const total = data.reduce((sum, val) => sum + val, 0);

    chartPizzaStatus = new Chart(canvas, {
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
// ✅ ATUALIZAR TABELA
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

        const dataSolicitacao = getValueByHeaderAliases(item, [
            'Data da Solicitação',
            'Data Solicitação',
            'Data da Solicitacao',
            'Data Solicitacao'
        ], '-');

        const prontuario = getValueByHeaderAliases(item, [
            'Nº Prontuário',
            'N° Prontuário',
            'Numero Prontuário',
            'Prontuário',
            'Prontuario'
        ], '-');

        const dataInicioStr = getValueByHeaderAliases(item, [
            'Data Início da Pendência',
            'Data Inicio da Pendencia',
            'Data Início Pendência',
            'Data Inicio Pendencia'
        ], '-');

        const prazo15 = getValueByHeaderAliases(item, [
            'Data Final do Prazo (Pendência com 15 dias)',
            'Data Final do Prazo (Pendencia com 15 dias)',
            'Data Final Prazo 15d',
            'Prazo 15 dias'
        ], '-');

        const email15 = getValueByHeaderAliases(item, [
            'Data do envio do Email (Prazo: Pendência com 15 dias)',
            'Data do envio do Email (Prazo: Pendencia com 15 dias)',
            'Data Envio Email 15d',
            'Email 15 dias'
        ], '-');

        const prazo30 = getValueByHeaderAliases(item, [
            'Data Final do Prazo (Pendência com 30 dias)',
            'Data Final do Prazo (Pendencia com 30 dias)',
            'Data Final Prazo 30d',
            'Prazo 30 dias'
        ], '-');

        const email30 = getValueByHeaderAliases(item, [
            'Data do envio do Email (Prazo: Pendência com 30 dias)',
            'Data do envio do Email (Prazo: Pendencia com 30 dias)',
            'Data Envio Email 30d',
            'Email 30 dias'
        ], '-');

        // OBS: Mantive sua regra do destaque, mas com o nome certo da sua aba atual.
        // Se você realmente quiser destacar outra aba, troque a string abaixo.
        const dataInicio = parseDate(dataInicioStr);
        let isVencendo15 = false;

        if (dataInicio && origem === 'PENDÊNCIAS VARGEM DAS FLORES') {
            const diasDecorridos = Math.floor((hoje - dataInicio) / (1000 * 60 * 60 * 24));
            if (diasDecorridos >= 15 && diasDecorridos < 30) isVencendo15 = true;
        }

        row.innerHTML = `
            <td>${escapeHtml(origem)}</td>
            <td>${escapeHtml(formatDate(dataSolicitacao))}</td>
            <td>${escapeHtml(prontuario)}</td>
            <td>${escapeHtml(getValueByHeaderAliases(item, ['Telefone'], '-'))}</td>
            <td>${escapeHtml(getValueByHeaderAliases(item, ['Unidade Solicitante'], '-'))}</td>
            <td>${escapeHtml(getValueByHeaderAliases(item, ['Cbo Especialidade', 'CBO Especialidade'], '-'))}</td>
            <td>${escapeHtml(formatDate(dataInicioStr))}</td>
            <td>${escapeHtml(getValueByHeaderAliases(item, ['Status'], '-'))}</td>
            <td>${escapeHtml(formatDate(prazo15))}</td>
            <td>${escapeHtml(formatDate(email15))}</td>
            <td>${escapeHtml(formatDate(prazo30))}</td>
            <td>${escapeHtml(formatDate(email30))}</td>
        `;

        if (isVencendo15) row.classList.add('row-vencendo-15');
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

    const s = String(dateString).trim();

    let match = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (match) return new Date(Number(match[3]), Number(match[2]) - 1, Number(match[1]));

    match = s.match(/(\d{4})-(\d{2})-(\d{2})/);
    if (match) return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));

    return null;
}

function formatDate(dateString) {
    if (!dateString || dateString === '-') return '-';

    const date = parseDate(dateString);
    if (!date || isNaN(date.getTime())) return String(dateString);

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
        'Data Solicitação': getValueByHeaderAliases(item, ['Data da Solicitação', 'Data Solicitação', 'Data da Solicitacao', 'Data Solicitacao'], ''),
        'Nº Prontuário': getValueByHeaderAliases(item, ['Nº Prontuário', 'N° Prontuário', 'Numero Prontuário', 'Prontuário', 'Prontuario'], ''),
        'Telefone': getValueByHeaderAliases(item, ['Telefone'], ''),
        'Unidade Solicitante': getValueByHeaderAliases(item, ['Unidade Solicitante'], ''),
        'CBO Especialidade': getValueByHeaderAliases(item, ['Cbo Especialidade', 'CBO Especialidade'], ''),
        'Data Início Pendência': getValueByHeaderAliases(item, ['Data Início da Pendência','Data Início Pendência','Data Inicio da Pendencia','Data Inicio Pendencia'], ''),
        'Status': getValueByHeaderAliases(item, ['Status'], ''),
        'Prestador': getValueByHeaderAliases(item, ['Prestador'], ''),
        'Data Final Prazo 15d': getValueByHeaderAliases(item, ['Data Final do Prazo (Pendência com 15 dias)','Data Final do Prazo (Pendencia com 15 dias)','Data Final Prazo 15d','Prazo 15 dias'], ''),
        'Data Envio Email 15d': getValueByHeaderAliases(item, ['Data do envio do Email (Prazo: Pendência com 15 dias)','Data do envio do Email (Prazo: Pendencia com 15 dias)','Data Envio Email 15d','Email 15 dias'], ''),
        'Data Final Prazo 30d': getValueByHeaderAliases(item, ['Data Final do Prazo (Pendência com 30 dias)','Data Final do Prazo (Pendencia com 30 dias)','Data Final Prazo 30d','Prazo 30 dias'], ''),
        'Data Envio Email 30d': getValueByHeaderAliases(item, ['Data do envio do Email (Prazo: Pendência com 30 dias)','Data do envio do Email (Prazo: Pendencia com 30 dias)','Data Envio Email 30d','Email 30 dias'], '')
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
