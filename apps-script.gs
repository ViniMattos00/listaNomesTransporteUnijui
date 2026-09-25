/**
 * Recebe as inscrições da página index.html e grava na planilha.
 * Cole este código em: Planilha > Extensões > Apps Script.
 */
const NOME_ABA = 'Inscrições';
const CABECALHO = ['Enviado em', 'Data da viagem', 'Nome completo', 'CPF'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000); // evita que envios simultâneos se atropelem

  try {
    const dados = JSON.parse(e.postData.contents);
    const aba = obterAba();

    aba.appendRow([
      dados.enviadoEm,
      dados.dataViagem,
      dados.nome,
      "'" + dados.cpf // apóstrofo força texto (mantém zeros à esquerda)
    ]);

    return resposta({ ok: true });
  } catch (err) {
    return resposta({ ok: false, erro: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Permite testar abrindo a URL do Web App no navegador.
function doGet() {
  return resposta({ ok: true, mensagem: 'Web App de inscrições no ar.' });
}

function obterAba() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  let aba = planilha.getSheetByName(NOME_ABA);
  if (!aba) aba = planilha.insertSheet(NOME_ABA);
  if (aba.getLastRow() === 0) {
    aba.appendRow(CABECALHO);
    aba.getRange(1, 1, 1, CABECALHO.length).setFontWeight('bold');
    aba.setFrozenRows(1);
  }
  return aba;
}

function resposta(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
