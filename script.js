// ===================================
// FUNÇÃO: URL CSV (Google Sheets gviz) + ANTI-CACHE
// ===================================
function gvizCsvUrl(sheetId, gid) {
  const cacheBust = Date.now();
  return `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&gid=${gid}&_=${cacheBust}`;
}

// ===================================
// ✅ CONFIGURAÇÃO DA PLANILHA (NOVAS ABAS VARGEM DAS FLORES)
// ===================================
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

    const results = await Promise<span class="cursor">█</span>
