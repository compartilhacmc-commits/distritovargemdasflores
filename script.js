// ===================================
// FUNÇÃO: URL CSV (Google Sheets gviz) + ANTI-CACHE
// ===================================
function gvizCsvUrl(sheetId, gid) {
  const cacheBust = Date.now();
  return `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&gid=${gid}&_=${cacheBust}`;
}

// ===================================
// CONFIGURAÇÃO DA PLANILHA (DUAS ABAS)
// ===================================
// ✅ PLANILHA "VARGEM DAS FLORES"
const SHEET_ID = '1IHknmxe3xAnfy5Bju_23B5ivIL-qMaaE6q_HuPaLBpk';

const SHEETS = [
  {
    name: 'PENDÊNCIAS VARGEM DAS FLORES',
    url: gvizCsvUrl(SHEET_ID, '278071504'),
    distrito: 'VARGEM DAS FLORES',
    tipo: 'PENDENTE'
  },
  {
    name: 'RESOLVIDOS VARGEM DAS FLORES',
    url: gvizCsvUrl(SHEET_ID, '451254610'),
    distrito: 'VARGEM DAS FLORES',
    tipo: 'RESOLVIDO'
  }
];

// ===================================
// VARIÁVEIS GLOBAIS
// ===================================
let allData = [];
let filteredData = [];
let chartPendenciasNaoResolvidasUnidade = null;
let chartUnidades = null;
let chartEspecialidades = null;
let chartStatus = null;
let chartPizzaStatus = null;
let chartPendenciasPrestador = null;
let chartPendenciasMes = null;
let chartResolutividadeUnidade = null;
let chartResolutividadePrestador = null;

// ===================================
// FUNÇÃO AUXILIAR PARA BUSCAR VALOR DE COLUNA
// ===================================
function getColumnValue(item, possibleNames, defaultValue = '-') {
  for (let name of possibleNames) {
    if (Object.prototype.hasOwnProperty.call(item, name) && item[name]) {
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
// ✅ HELPERS DE ORIGEM (CORREÇÃO PARA NÃO DEPENDER DE NOME FIXO)
// ===================================
function isOrigemPendencias(item) {
  const origem = String(item?._origem || '').toUpperCase();
  return origem.includes('PEND'); // pega "PENDÊNCIAS ..." (qualquer variação)
}

function isOrigemResolvidos(item) {
  const origem = String(item?._origem || '').toUpperCase();
  return origem.includes('RESOLV'); // pega "RESOLVIDOS ..." (qualquer variação)
}

// ===================================
// MULTISELECT (CHECKBOX) HELPERS
// ===================================
function toggleMultiSelect(id) {
  document.getElementById(id).classList.toggle('open');
}

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
document.addEventListener('DOMContentLoaded', function () {
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

    const promises = SHEETS.map(sheet =>
      fetch(sheet.url, { cache: 'no-store' })
        .then(response => {
          if (!response.ok) {
            throw new Error(`Erro HTTP na aba "${sheet.name}": ${response.status}`);
          }
          return response.text();
        })
        .then(csvText => {
          csvText = csvText.replace(/^\uFEFF/, '');

          if (csvText.includes('<html') || csvText.includes('<!DOCTYPE')) {
            throw new Error(
              `Aba "${sheet.name}" retornou HTML (provável falta de permissão ou planilha não pública).`
            );
          }

          console.log(`Dados CSV da aba "${sheet.name}" recebidos`);
          return { name: sheet.name, csv: csvText };
        })
    );

    const results = await Promise.all(promises);

    results.forEach(result => {
      const rows = parseCSV(result.csv);

      if (rows.length < 2) {
        console.warn(`Aba "${result.name}" está vazia ou sem dados`);
        return;
      }

      const headers = rows[0].map(h => (h || '').trim());
      console.log(`Cabeçalhos da aba "${result.name}":`, headers);

      const sheetData = rows.slice(1)
        .filter(row => row.length > 1 && (row[0] || '').trim() !== '')
        .map(row => {
          const obj = { _origem: result.name };
          headers.forEach((header, index) => {
            if (!header) return;
            obj[header] = (row[index] || '').trim();
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
    alert(
      `Erro ao carregar dados da planilha: ${error.message}\n\n` +
      `Verifique:\n` +
      `1. A planilha está com acesso "Qualquer pessoa com o link pode visualizar"? \n` +
      `2. Os GIDs estão corretos (aba certa)?\n` +
      `3. Há dados nas abas?\n`
    );
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
      if (char === '\r' && nextChar === '\n') i++;
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
  if (!overlay) return;
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

  populateMonthFilter();
}

// ===================================
// ✅ POPULAR FILTRO DE MÊS (MULTISELECT STYLE)
// ===================================
function populateMonthFilter() {
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
  const mesesFormatados = mesesOrdenados.map(mesAno => {
    const [ano, mes] = mesAno.split('-');
    const nomeMes = new Date(ano, mes - 1).toLocaleDateString('pt-BR', { month: 'long', year: 'numeric' });
    return nomeMes.charAt(0).toUpperCase() + nomeMes.slice(1);
  });

  renderMultiSelect('msMesPanel', mesesFormatados, applyFilters);
  setMultiSelectText('msMesText', [], 'Todos os Meses');
}

// ===================================
// ✅ APLICAR FILTROS (MULTISELECT + MÊS)
// ===================================
function applyFilters() {
  const statusSel = getSelectedFromPanel('msStatusPanel');
  const unidadeSel = getSelectedFromPanel('msUnidadePanel');
  const especialidadeSel = getSelectedFromPanel('msEspecialidadePanel');
  const prestadorSel = getSelectedFromPanel('msPrestadorPanel');
  const mesSel = getSelectedFromPanel('msMesPanel');

  setMultiSelectText('msStatusText', statusSel, 'Todos');
  setMultiSelectText('msUnidadeText', unidadeSel, 'Todas');
  setMultiSelectText('msEspecialidadeText', especialidadeSel, 'Todas');
  setMultiSelectText('msPrestadorText', prestadorSel, 'Todos');
  setMultiSelectText('msMesText', mesSel, 'Todos os Meses');

  filteredData = allData.filter(item => {
    const okStatus = (statusSel.length === 0) || statusSel.includes(item['Status'] || '');
    const okUnidade = (unidadeSel.length === 0) || unidadeSel.includes(item['Unidade Solicitante'] || '');
    const okEsp = (especialidadeSel.length === 0) || especialidadeSel.includes(item['Cbo Especialidade'] || '');
    const okPrest = (prestadorSel.length === 0) || prestadorSel.includes(item['Prestador'] || '');

    let okMes = true;
    if (mesSel.length > 0) {
      const dataInicio = parseDate(getColumnValue(item, [
        'Data Início da Pendência',
        'Data Inicio da Pendencia',
        'Data Início Pendência',
        'Data Inicio Pendencia'
      ]));

      if (dataInicio) {
        const mesAno = `${dataInicio.getFullYear()}-${String(dataInicio.getMonth() + 1).padStart(2, '0')}`;
        const [ano, mes] = mesAno.split('-');
        const nomeMes = new Date(ano, mes - 1).toLocaleDateString('pt-BR', { month: 'long', year: 'numeric' });
        const mesFormatado = nomeMes.charAt(0).toUpperCase() + nomeMes.slice(1);
        okMes = mesSel.includes(mesFormatado);
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
  ['msStatusPanel', 'msUnidadePanel', 'msEspecialidadePanel', 'msPrestadorPanel', 'msMesPanel'].forEach(panelId => {
    const panel = document.getElementById(panelId);
    if (!panel) return;
    panel.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = false);
  });

  setMultiSelectText('msStatusText', [], 'Todos');
  setMultiSelectText('msUnidadeText', [], 'Todas');
  setMultiSelectText('msEspecialidadeText', [], 'Todas');
  setMultiSelectText('msPrestadorText', [], 'Todos');
  setMultiSelectText('msMesText', [], 'Todos os Meses');

  const si = document.getElementById('searchInput');
  if (si) si.value = '';

  filteredData = [...allData];
  updateDashboard();
}

// ===================================
// PESQUISAR NA TABELA
// ===================================
function searchTable() {
  const searchValue = (document.getElementById('searchInput')?.value || '').toLowerCase();
  const tbody = document.getElementById('tableBody');
  if (!tbody) return;

  const rows = tbody.getElementsByTagName('tr');
  let visibleCount = 0;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const cells = row.getElementsByTagName('td');
    let found = false;

    for (let j = 0; j < cells.length; j++) {
      const cellText = (cells[j].textContent || '').toLowerCase();
      if (cellText.includes(searchValue)) {
        found = true;
        break;
      }
    }

    row.style.display = found ? '' : 'none';
    if (found) visibleCount++;
  }

  const footer = document.getElementById('tableFooter');
  if (footer) footer.textContent = `Mostrando ${visibleCount} de ${filteredData.length} registros`;
}

// ===================================
// DASHBOARD
// ===================================
function updateDashboard() {
  updateCards();
  updateCharts();
  updateTable();
}

// ===================================
// ✅ CARDS (CONTANDO POR "USUÁRIO" PREENCHIDO)
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
// ✅ GRÁFICOS
// ===================================
function updateCharts() {
  // ✅ PENDÊNCIAS NÃO RESOLVIDAS POR UNIDADE - VERMELHO (#dc2626)
  const pendenciasNaoResolvidasUnidade = {};
  filteredData.forEach(item => {
    if (!isOrigemPendencias(item)) return;
    if (!isPendenciaByUsuario(item)) return;

    const unidade = item['Unidade Solicitante'] || 'Não informado';
    pendenciasNaoResolvidasUnidade[unidade] = (pendenciasNaoResolvidasUnidade[unidade] || 0) + 1;
  });

  const pendenciasNRLabels = Object.keys(pendenciasNaoResolvidasUnidade)
    .sort((a, b) => pendenciasNaoResolvidasUnidade[b] - pendenciasNaoResolvidasUnidade[a])
    .slice(0, 50);
  const pendenciasNRValues = pendenciasNRLabels.map(label => pendenciasNaoResolvidasUnidade[label]);

  createHorizontalBarChart('chartPendenciasNaoResolvidasUnidade', pendenciasNRLabels, pendenciasNRValues, '#dc2626');

  // Gráfico de Unidades (GERAL)
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

  // ✅ GRÁFICO DE ESPECIALIDADES - VERDE ESCURO (#065f46)
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

  createHorizontalBarChart('chartEspecialidades', especialidadesLabels, especialidadesValues, '#065f46');

  // Gráfico de Status
  const statusCount = {};
  filteredData.forEach(item => {
    const status = item['Status'] || 'Não informado';
    statusCount[status] = (statusCount[status] || 0) + 1;
  });

  const statusLabels = Object.keys(statusCount)
    .sort((a, b) => statusCount[b] - statusCount[a]);
  const statusValues = statusLabels.map(label => statusCount[label]);

  createVerticalBarChart('chartStatus', statusLabels, statusValues, '#f97316');

  // Gráfico de Pizza
  createPieChart('chartPizzaStatus', statusLabels, statusValues);

  // Pendências por Prestador
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

  createVerticalBarChartCenteredValue('chartPendenciasPrestador', prestLabels, prestValues, '#4c1d95');

  // Pendências por Mês
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
    .sort((a, b) => mesCount
