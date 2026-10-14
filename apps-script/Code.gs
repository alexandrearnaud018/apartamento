// Ponte entre o site e a planilha.
// A senha fica em Configurações do projeto → Propriedades do script → SENHA (nunca no código).
// Toda chamada é um POST com corpo JSON: { senha, acao, ... }

const PLANILHA_ID = '1lwzfOEZgWn2aDMayUpgKv-3pqWSwZ10Uqrmg2_rvRFs';
const ABA = 'Itens';
const COLUNAS = ['id', 'tipo', 'nome', 'status', 'valor', 'obs', 'ordem'];

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (_) { return resposta({ erro: 'requisicao_invalida' }); }

  const senha = PropertiesService.getScriptProperties().getProperty('SENHA');
  if (!senha || req.senha !== senha) return resposta({ erro: 'senha' });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (req.acao === 'listar') return resposta({ itens: lerItens() });
    if (req.acao === 'salvar') { salvar(req.item); return resposta({ ok: true }); }
    if (req.acao === 'excluir') { excluir(req.id); return resposta({ ok: true }); }
    if (req.acao === 'ordenar') { ordenar(req.ordens); return resposta({ ok: true }); }
    return resposta({ erro: 'acao_desconhecida' });
  } catch (err) {
    return resposta({ erro: 'falha', detalhe: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function resposta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function aba() {
  return SpreadsheetApp.openById(PLANILHA_ID).getSheetByName(ABA);
}

function lerItens() {
  const linhas = aba().getDataRange().getValues();
  linhas.shift(); // cabeçalho
  return linhas
    .filter((l) => l[0] !== '')
    .map((l) => {
      const it = {};
      COLUNAS.forEach((c, i) => { it[c] = l[i] === '' ? null : l[i]; });
      it.id = String(it.id);
      it.obs = it.obs || '';
      return it;
    });
}

// Linha (1-based) de cada id
function indice() {
  const ids = aba().getRange(1, 1, aba().getLastRow(), 1).getValues();
  const map = {};
  ids.forEach((r, i) => { if (i > 0 && r[0] !== '') map[String(r[0])] = i + 1; });
  return map;
}

// Cria ou atualiza; só os campos enviados mudam
function salvar(item) {
  if (!item || !item.id) throw new Error('item sem id');
  const sh = aba();
  const linha = indice()[item.id];
  if (linha) {
    const atual = sh.getRange(linha, 1, 1, COLUNAS.length).getValues()[0];
    const nova = COLUNAS.map((c, i) => (c in item ? (item[c] == null ? '' : item[c]) : atual[i]));
    sh.getRange(linha, 1, 1, COLUNAS.length).setValues([nova]);
  } else {
    sh.appendRow(COLUNAS.map((c) => (item[c] == null ? '' : item[c])));
  }
}

function excluir(id) {
  const linha = indice()[id];
  if (linha) aba().deleteRow(linha);
}

// ordens: { id: numero, ... }
function ordenar(ordens) {
  const sh = aba(), idx = indice(), col = COLUNAS.indexOf('ordem') + 1;
  Object.keys(ordens || {}).forEach((id) => {
    if (idx[id]) sh.getRange(idx[id], col).setValue(ordens[id]);
  });
}
