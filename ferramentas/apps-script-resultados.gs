/**
 * Fazenda de Projetos — recebe os resultados das avaliações do site.
 * Cada aluno que termina vira uma linha na aba "Resultados" desta planilha,
 * e o relatório em PDF é salvo na pasta "Avaliações - Fazenda de Projetos" do seu Google Drive.
 *
 * Como instalar (uma vez só, no computador):
 * 1. Abra esta planilha > Extensões > Apps Script. Apague o que estiver lá e cole este código. Salve.
 * 2. Implantar > Nova implantação > tipo "App da Web".
 *    Executar como: Eu.   Quem pode acessar: Qualquer pessoa.   > Implantar > Autorizar acesso.
 * 3. Copie o "URL do app da Web" (termina em /exec) e mande para o Claude colocar no site.
 */
const PASTA = 'Avaliações - Fazenda de Projetos';
const ABA = 'Resultados';
const CAB = ['Data e hora', 'Avaliação', 'Turma', 'Nome', 'Tentativa', 'Nota (0 a 10)', 'Acertos', 'Erros', 'Total',
             'Aproveitamento (%)', 'Tempo total', 'Questões erradas', 'PDF', 'Detalhe das respostas'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (!d.nome || !d.pdf) throw new Error('dados incompletos');
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    let aba = planilha.getSheetByName(ABA);
    if (!aba) {
      aba = planilha.insertSheet(ABA);
      aba.appendRow(CAB);
      aba.getRange(1, 1, 1, CAB.length).setFontWeight('bold').setBackground('#1D3B5C').setFontColor('#ffffff');
      aba.setFrozenRows(1);
    }
    const pastas = DriveApp.getFoldersByName(PASTA);
    const pasta = pastas.hasNext() ? pastas.next() : DriveApp.createFolder(PASTA);
    const sub = String(d.avaliacao || 'Avaliação');
    const subs = pasta.getFoldersByName(sub);
    const destino = subs.hasNext() ? subs.next() : pasta.createFolder(sub);
    const blob = Utilities.newBlob(Utilities.base64Decode(d.pdf), 'application/pdf', String(d.arquivo || 'relatorio.pdf'));
    const arq = destino.createFile(blob);
    aba.appendRow([new Date(), sub, d.turma, d.nome, d.tentativa, d.nota, d.acertos, d.erros, d.total,
                   d.pct, d.tempo, d.errou, arq.getUrl(), d.detalhe]);
    return saida({ ok: true });
  } catch (err) {
    return saida({ ok: false, erro: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() { return saida({ ok: true, msg: 'Recebedor de resultados da Fazenda de Projetos funcionando.' }); }

function saida(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
