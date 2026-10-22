/**
 * Recebe da página index.html os dias em que cada aluno VAI precisar
 * do transporte e grava na planilha (uma linha por CPF, uma coluna por dia).
 * Cole este código em: Planilha > Extensões > Apps Script e implante como
 * App da Web (Executar como: eu · Quem pode acessar: qualquer pessoa).
 */
const NOME_ABA = 'Dias com transporte';
const CABECALHO = ['Atualizado em', 'Nome completo', 'CPF'];
const COLUNAS_FIXAS = CABECALHO.length;
const MARCA = 'Vai';

// GET ?acao=buscar&cpf=00000000000&periodo=27/10/2026,28/10/2026
// Retorna os dias do período que o CPF já marcou como "vai".
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.acao !== 'buscar') {
    return resposta({ ok: true, mensagem: 'Web App de dias com transporte no ar.' });
  }

  try {
    const cpf = String(p.cpf || '').replace(/\D/g, '');
    const periodo = String(p.periodo || '').split(',').filter(String);
    const aba = obterAba();
    const linha = acharLinha(aba, cpf);
    if (!linha) return resposta({ ok: true, encontrado: false, vai: [] });

    const cabecalho = lerCabecalho(aba);
    const valores = aba.getRange(linha, 1, 1, cabecalho.length).getValues()[0];
    const vai = periodo.filter(dia => {
      const col = cabecalho.indexOf(dia);
      return col >= 0 && valores[col] === MARCA;
    });
    return resposta({ ok: true, encontrado: true, vai: vai });
  } catch (err) {
    return resposta({ ok: false, erro: String(err) });
  }
}

// POST { nome, cpf, periodo: [dias], vai: [dias], enviadoEm }
// Cria ou atualiza a linha do CPF.
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000); // evita que envios simultâneos se atropelem

  try {
    const dados = JSON.parse(e.postData.contents);
    const cpf = String(dados.cpf || '').replace(/\D/g, '');
    if (cpf.length !== 11) throw new Error('CPF inválido');

    const aba = obterAba();
    const cabecalho = garantirColunas(aba, dados.periodo || []);
    let linha = acharLinha(aba, cpf);
    if (!linha) linha = aba.getLastRow() + 1;

    const range = aba.getRange(linha, 1, 1, cabecalho.length);
    const valores = range.getValues()[0];
    valores[0] = dados.enviadoEm;
    valores[1] = dados.nome;
    valores[2] = "'" + formatarCPF(cpf); // apóstrofo força texto (mantém zeros à esquerda)
    (dados.periodo || []).forEach(dia => {
      valores[cabecalho.indexOf(dia)] = (dados.vai || []).indexOf(dia) >= 0 ? MARCA : '';
    });
    range.setValues([valores]);

    return resposta({ ok: true });
  } catch (err) {
    return resposta({ ok: false, erro: String(err) });
  } finally {
    lock.releaseLock();
  }
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

function lerCabecalho(aba) {
  return aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0].map(String);
}

// Cria no cabeçalho as colunas dos dias que ainda não existem.
function garantirColunas(aba, dias) {
  const cabecalho = lerCabecalho(aba);
  dias.forEach(dia => {
    if (cabecalho.indexOf(dia) >= 0) return;
    const col = cabecalho.length + 1;
    aba.getRange(1, col).setNumberFormat('@').setValue(dia).setFontWeight('bold');
    cabecalho.push(dia);
  });
  return cabecalho;
}

// Número da linha do CPF (comparando só os dígitos), ou 0 se não existir.
function acharLinha(aba, cpf) {
  const ultima = aba.getLastRow();
  if (ultima < 2) return 0;
  const cpfs = aba.getRange(2, COLUNAS_FIXAS, ultima - 1, 1).getValues();
  for (let i = 0; i < cpfs.length; i++) {
    if (String(cpfs[i][0]).replace(/\D/g, '') === cpf) return i + 2;
  }
  return 0;
}

function formatarCPF(c) {
  return c.replace(/(\d{3})(\d{3})(\d{3})(\d{2})/, '$1.$2.$3-$4');
}

function resposta(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
