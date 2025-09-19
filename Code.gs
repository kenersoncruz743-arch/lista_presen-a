/** =========================
 *  LOGIN / MENU / DOGET
 *  ========================= */

/**
 * doGet - app web (abre a tela de login)
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile("login")
    .evaluate()
    .setTitle("Login - Controle de Presença")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Valida usuário/senha na aba "Usuarios".
 * Retorna { ok: true, usuario, abas: [...] } ou { ok: false }
 */
function validarLogin(usuario, senha) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Usuarios");
  if (!sh) return { ok: false, msg: "Aba 'Usuarios' não encontrada" };

  const dados = sh.getDataRange().getValues();
  const abas = [];

  usuario = String(usuario || "").trim();
  senha = String(senha || "").trim();

  for (let i = 1; i < dados.length; i++) {
    const u = String(dados[i][0] || "").trim();
    const s = String(dados[i][1] || "").trim();
    const aba = String(dados[i][2] || "").trim();
    if (u === usuario && s === senha && aba) {
      abas.push(aba);
    }
  }

  // dedup e ordena
  const unicas = [...new Set(abas)].sort((a,b)=>a.localeCompare(b,'pt-BR'));
  if (unicas.length === 0) return { ok: false };
  return { ok: true, usuario: usuario, abas: unicas };
}

/**
 * Retorna HTML do menu (avaliado) para ser injetado no cliente.
 * Usa template "menu" com variables usuario e abas.
 */
function abrirMenu(usuario) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Usuarios");
  if (!sh) return "<p>Aba 'Usuarios' não encontrada.</p>";

  // monta lista de abas (garantia): pesquisar todas as linhas onde o usuario coincide
  const dados = sh.getDataRange().getValues();
  const abas = [];
  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][0] || "").trim() === String(usuario || "").trim()) {
      if (dados[i][2]) abas.push(String(dados[i][2]).trim());
    }
  }

  const template = HtmlService.createTemplateFromFile("menu");
  template.usuario = usuario || "";
  template.abas = [...new Set(abas)].sort((a,b)=>a.localeCompare(b,'pt-BR'));
  return template.evaluate().getContent();
}

/**
 * Retorna HTML do painel (avaliado) para a aba selecionada.
 * Usa template "painel" com variáveis supervisor e aba.
 */
function abrirItemMenu(aba, usuario) {
  const template = HtmlService.createTemplateFromFile("painel");
  template.supervisor = usuario || "";
  template.aba = aba || "";
  return template.evaluate().getContent();
}

/**
 * Auxiliar (retorna abas para um usuário) caso queira usar diretamente do cliente
 */
function getAbasUsuario(usuario) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Usuarios");
  if (!sh) return [];
  const dados = sh.getDataRange().getValues();
  const u = String(usuario || "").trim();
  const abas = [];
  for (let i = 1; i < dados.length; i++) {
    const user = String(dados[i][0] || "").trim();
    const aba  = String(dados[i][2] || "").trim();
    if (user === u && aba) abas.push(aba);
  }
  return [...new Set(abas)].sort((a,b)=>a.localeCompare(b,'pt-BR'));
}

// =========================
// Buscar colaboradores na aba Quadro
// =========================
function buscarNoQuadro(filtro) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Quadro");
  if (!sh) return [];

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const dados = sh.getRange(2, 1, lastRow - 1, 3).getValues(); // Matricula | Nome | Funcao
  const f = (filtro || "").toString().trim().toLowerCase();

  const lista = [];
  for (let i = 0; i < dados.length; i++) {
    const mat = String(dados[i][0] || "");
    const nome = String(dados[i][1] || "");
    const funcao = String(dados[i][2] || "");
    if (!f || mat.toLowerCase().includes(f) || nome.toLowerCase().includes(f)) {
      lista.push({ matricula: mat, nome: nome, funcao: funcao });
    }
  }

  // Ordena
  lista.sort((a, b) => {
    const fa = (a.funcao || "").toLowerCase();
    const fb = (b.funcao || "").toLowerCase();
    if (fa < fb) return -1;
    if (fa > fb) return 1;
    return (a.nome || "").toLowerCase().localeCompare((b.nome || "").toLowerCase(), 'pt-BR');
  });
  return lista;
}

/**
 * Buffer temporário em "Lista"
 * Colunas: [Supervisor, Aba, Matrícula, Nome, Função, Status]
 */
function getBuffer(supervisor, aba) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Lista");
  if (!sh) return [];

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const dados = sh.getRange(2, 1, lastRow - 1, 6).getValues();
  const buffer = [];

  const supLower = (supervisor || "").toString().trim().toLowerCase();
  const abaLower = (aba || "").toString().trim().toLowerCase();

  dados.forEach(r => {
    const supOk = (r[0] || "").toString().trim().toLowerCase() === supLower;
    const abaOk = !aba ? true : ((r[1] || "").toString().trim().toLowerCase() === abaLower);
    if (supOk && abaOk) {
      buffer.push({
        matricula: r[2],
        nome: r[3],
        funcao: r[4],
        status: r[5] || ""
      });
    }
  });

  return buffer;
}

function adicionarBuffer(supervisor, aba, colaborador) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("Lista");
  if (!sh) sh = ss.insertSheet("Lista");

  if (sh.getLastRow() === 0) {
    sh.appendRow(["Supervisor", "Aba", "Matrícula", "Nome", "Função", "Status"]);
  }

  const lastRow = sh.getLastRow();
  const dados = lastRow > 1 ? sh.getRange(2, 1, lastRow - 1, 6).getValues() : [];
  const existe = dados.some(r => r[0] === supervisor && String(r[2]) === String(colaborador.matricula));
  if (existe) return { ok: false, msg: "Já existe no buffer" };

  sh.appendRow([supervisor, aba, colaborador.matricula, colaborador.nome, colaborador.funcao, ""]);
  return { ok: true };
}

function removerBuffer(supervisor, matricula) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Lista");
  if (!sh) return { ok: false };

  const dados = sh.getDataRange().getValues(); // inclui cabeçalho
  matricula = matricula.toString().trim();
  supervisor = supervisor.toString().trim();

  for (let i = 1; i < dados.length; i++) {
    const sup = (dados[i][0] || "").toString().trim();
    const mat = (dados[i][2] || "").toString().trim();
    if (sup === supervisor && mat === matricula) {
      sh.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false };
}

function atualizarStatusBuffer(supervisor, matricula, status) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Lista");
  if (!sh) return { ok: false };

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { ok: false };

  const dados = sh.getRange(2, 1, lastRow - 1, 6).getValues();
  for (let i = 0; i < dados.length; i++) {
    if (dados[i][0] === supervisor && String(dados[i][2]) === String(matricula)) {
      sh.getRange(i + 2, 6).setValue(status); // Coluna 6 = Status
      return { ok: true };
    }
  }
  return { ok: false };
}

function salvarBufferNaBase(dadosDoHtml) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shBase = ss.getSheetByName("Base");
  if (!shBase) return { ok: false, msg: "Aba Base não encontrada" };

  const supervisorLogado = dadosDoHtml && dadosDoHtml.length > 0 ? dadosDoHtml[0][0] : null;
  if (!supervisorLogado) return { ok: false, msg: "Nenhum dado recebido" };

  // Garante coluna Data
  if (shBase.getLastColumn() < 7) {
    shBase.insertColumnAfter(6);
    shBase.getRange(1, 7).setValue("Data");
  }

  // Limpa conteúdo existente do supervisor
  const todas = shBase.getDataRange().getValues(); // inclui cabeçalho
  for (let i = 1; i < todas.length; i++) {
    if ((todas[i][0] || "") === supervisorLogado) {
      shBase.getRange(i + 1, 1, 1, shBase.getLastColumn()).clearContent();
    }
  }

  // Insere snapshot
  const hoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
  let salvos = 0;

  dadosDoHtml.forEach(l => {
    const [sup, aba, matricula, nome, funcao, status] = l;
    if (!matricula && !nome && !funcao) return;
    shBase.appendRow([sup, aba, matricula, nome, funcao, status, hoje]);
    salvos++;
  });

  return { ok: true, msg: `${salvos} registros do supervisor ${supervisorLogado} salvos na Base` };
}
