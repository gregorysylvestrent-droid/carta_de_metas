// ============================================================
//  DASHBOARD KPI — LOGÍSTICA · NORTE TECH
//  Arquivo: Dashboard_Logistica.gs
//
//  INSTRUÇÕES DE USO:
//  1. Abra a planilha "CARTA_DE_METAS_-_LOGÍSTICA" no Google Sheets
//  2. Acesse Extensões → Apps Script
//  3. Cole TODO o conteúdo deste arquivo no editor
//  4. Crie um segundo arquivo HTML no Apps Script:
//     · Clique em "+" → HTML
//     · Nomeie como: dashboard_logistica_nortetech
//     · Cole o conteúdo do arquivo .html fornecido separadamente
//  5. Salve (Ctrl+S) e execute openDashboard() uma vez para
//     conceder permissões
//  6. Depois disso, use o menu "📊 Dashboard KPI" que aparece
//     na planilha
// ============================================================

// ============================================================
//  CONFIGURAÇÃO — ajuste se os nomes das abas mudarem
// ============================================================
var CONFIG = {
  // Nome exato da aba com a descrição dos KPIs (sem resultados mensais)
  ABA_DESCRICAO: 'DESCRIÇÃO',

  // Mapeamento: número da aba (1–12) → nome da coluna de resultado
  // As abas mensais são nomeadas "1", "2", ..., "12" na planilha
  // A coluna de resultado é a 16ª (índice 15, "P" em A=0)
  COL_RESULTADO: 15,   // índice 0-based da coluna com o valor realizado no mês

  // Índices 0-based das colunas na aba DESCRIÇÃO
  COLS: {
    ORDEM: 0,   // A - Ordem
    MATRICULA: 1,   // B - Matrícula
    NOME: 2,   // C - Nome Completo
    NUM_CARTA: 3,   // D - Nº da Carta
    NUM_IND: 4,   // E - Nº Ind
    COD_IND: 5,   // F - Cód Indicador
    INDICADOR: 6,   // G - Indicador
    CRITERIO: 7,   // H - Critério de Apuração
    BASE_DADOS: 8,   // I - Base de Dados
    RESPONSAVEL: 9,   // J - Responsável
    MELHOR: 10,   // K - Melhor (↑ ou ↓)
    RANGE: 11,   // L - Range
    PESO: 12,   // M - Peso
    TIPO_META: 13,   // N - Tipo de Meta
    META: 14,   // O - Meta
  },

  // Total de meses na planilha
  NUM_MESES: 12,

  // Nomes dos meses em pt-BR (usados no dashboard)
  NOMES_MESES: [
    'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
    'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'
  ]
};


// ============================================================
//  doGet — abre o dashboard como WebApp (deploy como web app)
// ============================================================
function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('dashboard_logistica_nortetech')
    .setTitle('Dashboard KPI — Logística · Norte Tech')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}


// ============================================================
//  openDashboard — abre o dashboard num diálogo modal (sidebar não)
//  Chame esta função pelo menu ou botão na planilha
// ============================================================
function openDashboard() {
  var html = HtmlService
    .createHtmlOutputFromFile('dashboard_logistica_nortetech')
    .setTitle('Dashboard KPI Logística')
    .setWidth(1500)
    .setHeight(1000);

  SpreadsheetApp.getUi().showModalDialog(html, '📊 Dashboard KPI — Logística Norte Tech');
}


// ============================================================
//  openDashboardSidebar — versão painel lateral (mais estreito)
// ============================================================
function openDashboardSidebar() {
  var html = HtmlService
    .createHtmlOutputFromFile('dashboard_logistica_nortetech')
    .setTitle('Dashboard KPI')
    .setWidth(800);

  SpreadsheetApp.getUi().showSidebar(html);
}


// ============================================================
//  getDashboardData — chamada pelo HTML via google.script.run
//  Retorna JSON com todos os KPIs e resultados mensais
// ============================================================
function getDashboardData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var kpis = lerKpisDescricao_(ss);
    var monthly = lerResultadosMensais_(ss, kpis);

    var payload = {
      kpis: kpis,
      monthly: monthly,
      months: CONFIG.NOMES_MESES.slice(0, CONFIG.NUM_MESES),
      geradoEm: new Date().toLocaleString('pt-BR', { timeZone: 'America/Manaus' })
    };

    return payload;

  } catch (e) {
    return {
      erro: e.message,
      stack: e.stack
    };
  }
}


// ============================================================
//  lerKpisDescricao_  (privado)
//  Lê a aba DESCRIÇÃO e retorna array de objetos KPI
// ============================================================
function lerKpisDescricao_(ss) {
  var aba = ss.getSheetByName(CONFIG.ABA_DESCRICAO);
  if (!aba) {
    throw new Error(
      'Aba "' + CONFIG.ABA_DESCRICAO + '" não encontrada. ' +
      'Verifique o nome exato na planilha.'
    );
  }

  var dados = aba.getDataRange().getValues();
  // Remove cabeçalho (linha 0)
  var linhas = dados.slice(1);
  var c = CONFIG.COLS;
  var kpis = [];

  linhas.forEach(function (row) {
    // Ignora linhas vazias ou de totais (sem código de indicador)
    if (!row[c.COD_IND] || String(row[c.COD_IND]).trim() === '') return;
    // Ignora linha de soma final (sem nome)
    if (!row[c.NOME] || String(row[c.NOME]).trim() === '') return;

    var meta = parseFloat(row[c.META]) || 0;
    var peso = parseFloat(row[c.PESO]) || 0;
    var range = parseFloat(row[c.RANGE]) || 0;

    kpis.push({
      ordem: parseInt(row[c.ORDEM]) || 0,
      matricula: String(row[c.MATRICULA] || '').trim(),
      nome: String(row[c.NOME] || '').trim(),
      numCarta: String(row[c.NUM_CARTA] || '').trim(),
      numInd: String(row[c.NUM_IND] || '').trim(),
      codInd: String(row[c.COD_IND] || '').trim(),
      indicador: String(row[c.INDICADOR] || '').trim(),
      criterio: String(row[c.CRITERIO] || '').trim(),
      baseDados: String(row[c.BASE_DADOS] || '').trim(),
      responsavel: String(row[c.RESPONSAVEL] || '').trim(),
      melhor: String(row[c.MELHOR] || '↑').trim(),
      range: range,
      peso: peso,
      tipo: String(row[c.TIPO_META] || 'PORCENTAGEM').trim(),
      meta: meta
    });
  });

  if (kpis.length === 0) {
    throw new Error(
      'Nenhum KPI encontrado na aba "' + CONFIG.ABA_DESCRICAO + '". ' +
      'Verifique se a planilha tem dados a partir da linha 2.'
    );
  }

  return kpis;
}


// ============================================================
//  lerResultadosMensais_  (privado)
//  Lê cada aba mensal (1, 2, ... 12) e retorna objeto:
//  { "1": [val0, val1, ...], "2": [...], ... }
//  Os índices do array correspondem à ordem dos KPIs em kpis[]
// ============================================================
function lerResultadosMensais_(ss, kpis) {
  var monthly = {};
  var c = CONFIG.COLS;

  for (var mes = 1; mes <= CONFIG.NUM_MESES; mes++) {
    var nomAba = String(mes);
    var aba = ss.getSheetByName(nomAba);

    if (!aba) {
      monthly[nomAba] = new Array(kpis.length).fill(null);
      continue;
    }

    var dados = aba.getDataRange().getValues();
    var linhas = dados.slice(1);

    var mapaCodigos = {};

    linhas.forEach(function (row) {
      var cod = String(row[c.COD_IND] || '').trim();

      if (!cod) return;

      var val = row[CONFIG.COL_RESULTADO];

      if (val === '' || val === null || val === undefined) {
        mapaCodigos[cod] = null;
      } else {
        var n = parseFloat(val);
        mapaCodigos[cod] = isNaN(n) ? null : n;
      }
    });

    monthly[nomAba] = kpis.map(function (kpi) {
      return mapaCodigos[kpi.codInd] ?? null;
    });
  }

  return monthly;
}


// ============================================================
//  onOpen — cria menu personalizado na planilha
// ============================================================
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 Dashboard KPI')
    .addItem('🚀 Abrir Dashboard (Tela Cheia)', 'openDashboard')
    .addItem('📌 Abrir no Painel Lateral', 'openDashboardSidebar')
    .addSeparator()
    .addItem('🔄 Recarregar Dados (teste)', 'testarDados')
    .addItem('ℹ️ Ajuda / Instruções', 'mostrarAjuda')
    .addToUi();
}


// ============================================================
//  testarDados — executa getDashboardData e exibe resumo
//  Use para verificar se a leitura está correta antes de abrir
// ============================================================
function testarDados() {
  var json = getDashboardData();
  var dados = JSON.parse(json);

  if (dados.erro) {
    SpreadsheetApp.getUi().alert(
      '❌ Erro ao ler dados:\n\n' + dados.erro
    );
    return;
  }

  var totalKpis = dados.kpis.length;
  var resumoMeses = Object.keys(dados.monthly).map(function (m) {
    var vals = dados.monthly[m];
    var comDados = vals.filter(function (v) { return v !== null; }).length;
    return CONFIG.NOMES_MESES[parseInt(m) - 1].slice(0, 3) +
      ': ' + comDados + '/' + totalKpis;
  }).join('\n  ');

  SpreadsheetApp.getUi().alert(
    '✅ Dados lidos com sucesso!\n\n' +
    '📋 KPIs encontrados: ' + totalKpis + '\n' +
    '📅 Dados por mês:\n  ' + resumoMeses + '\n\n' +
    '⏱️ Gerado em: ' + dados.geradoEm
  );
}


// ============================================================
//  mostrarAjuda — exibe instruções de configuração
// ============================================================
function mostrarAjuda() {
  var msg =
    '📊 DASHBOARD KPI — LOGÍSTICA NORTE TECH\n' +
    '─────────────────────────────────────\n\n' +
    'ESTRUTURA ESPERADA DA PLANILHA:\n' +
    '• Aba "DESCRIÇÃO" → lista mestre de todos os KPIs\n' +
    '• Abas "1", "2", ... "12" → resultados mensais\n' +
    '  (1=Janeiro, 2=Fevereiro, etc.)\n\n' +
    'COLUNA DE RESULTADO (abas mensais):\n' +
    '• Coluna P (16ª coluna, índice 15) — valor realizado\n' +
    '• Para alterar, edite CONFIG.COL_RESULTADO no script\n\n' +
    'COMO USAR:\n' +
    '1. Menu "📊 Dashboard KPI" → "🚀 Abrir Dashboard"\n' +
    '2. O dashboard carrega os dados em tempo real\n' +
    '3. Use "🔄 Recarregar Dados" para testar a leitura\n\n' +
    'DEPLOY COMO WEB APP (opcional):\n' +
    '• Apps Script → Implantar → Novo Implantação\n' +
    '• Tipo: Aplicativo da Web\n' +
    '• Acesso: "Qualquer pessoa na organização"\n' +
    '• Isso gera uma URL pública do dashboard\n\n' +
    'SUPORTE:\n' +
    '• Verifique se o nome das abas está exato\n' +
    '• Use "🔄 Recarregar Dados" para diagnóstico\n';

  SpreadsheetApp.getUi().alert(msg);
}


// ============================================================
//  getEmployeeList — utilitário opcional
//  Retorna lista única de colaboradores (para debug)
// ============================================================
function getEmployeeList() {
  var json = getDashboardData();
  var dados = JSON.parse(json);
  if (dados.erro) return [];

  var seen = {};
  var lista = [];
  dados.kpis.forEach(function (kpi) {
    var mats = kpi.matricula.split(/\s+/);
    mats.forEach(function (mat) {
      mat = mat.trim();
      if (mat && !seen[mat]) {
        seen[mat] = true;
        lista.push({ id: mat, nome: kpi.nome });
      }
    });
  });
  return lista;
}