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
const CAB = ['Data e hora', 'Avaliação', 'Nº de chamada', 'Nome', 'Tentativa', 'Nota (0 a 10)', 'Acertos', 'Erros', 'Total',
             'Aproveitamento (%)', 'Tempo total', 'Questões erradas', 'PDF', 'Detalhe das respostas'];

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return saida({ ok: false, erro: 'json' }); }
  if (d.acao) return aula(d);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
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

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.acao === 'estado') return saida(estadoAula());
  return saida({ ok: true, msg: 'Recebedor de resultados da Fazenda de Projetos funcionando.' });
}

/* ===================== Aula ao vivo =====================
 * O professor controla a questão que aparece para todos (código do professor abaixo).
 * Os alunos consultam o estado e mandam a resposta da questão aberta.
 * As respostas ficam guardadas por 6 horas e, ao revelar cada questão, vão para a aba "Aula ao vivo".
 */
const CODIGO_PROFESSOR = 'PROF-E412';
const SEIS_HORAS = 21600;
const MAX_CHAMADA = 80;

function estadoAula() {
  const c = CacheService.getScriptCache();
  let s = c.get('aula_estado');
  if (!s) {
    s = PropertiesService.getScriptProperties().getProperty('aula_estado');
    if (s) c.put('aula_estado', s, SEIS_HORAS);
  }
  const o = s ? JSON.parse(s) : { fase: 'fechada' };
  o.ok = true;
  return o;
}

function aula(d) {
  const c = CacheService.getScriptCache();
  const est = estadoAula();
  const ch = parseInt(d.chamada, 10);
  if (d.acao === 'entrar') {
    if (!(ch >= 1 && ch <= MAX_CHAMADA) || !d.nome || d.sessao !== est.sessao) return saida({ ok: false, erro: 'sessao', estado: est });
    c.put('al_' + d.sessao + '_' + ch, String(d.nome).slice(0, 80), SEIS_HORAS);
    return saida({ ok: true, estado: est });
  }
  if (d.acao === 'responder') {
    if (d.sessao !== est.sessao || +d.q !== +est.q || est.fase !== 'aberta') return saida({ ok: false, erro: 'fechada', estado: est });
    if (!(ch >= 1 && ch <= MAX_CHAMADA)) return saida({ ok: false, erro: 'chamada' });
    c.put('r_' + d.sessao + '_' + d.q + '_' + ch, JSON.stringify({ nome: String(d.nome || '').slice(0, 80), resp: d.resp, t: Date.now() }), SEIS_HORAS);
    c.put('al_' + d.sessao + '_' + ch, String(d.nome || '').slice(0, 80), SEIS_HORAS);
    return saida({ ok: true });
  }
  if (d.chave !== CODIGO_PROFESSOR) return saida({ ok: false, erro: 'codigo' });
  if (d.acao === 'testar') return saida({ ok: true });
  if (d.acao === 'controle') {
    const novo = JSON.stringify({ sessao: String(d.sessao), q: +d.q, fase: String(d.fase), total: +d.total, t: Date.now() });
    PropertiesService.getScriptProperties().setProperty('aula_estado', novo);
    c.put('aula_estado', novo, SEIS_HORAS);
    return saida({ ok: true });
  }
  if (d.acao === 'respostas') {
    const chaves = [], alunos = [];
    for (let i = 1; i <= MAX_CHAMADA; i++) { chaves.push('r_' + d.sessao + '_' + d.q + '_' + i); alunos.push('al_' + d.sessao + '_' + i); }
    const r = c.getAll(chaves), a = c.getAll(alunos), resp = {}, turma = {};
    Object.keys(r).forEach(function (k) { resp[k.split('_').pop()] = JSON.parse(r[k]); });
    Object.keys(a).forEach(function (k) { turma[k.split('_').pop()] = a[k]; });
    return saida({ ok: true, respostas: resp, alunos: turma });
  }
  if (d.acao === 'salvar') {
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const planilha = SpreadsheetApp.getActiveSpreadsheet();
      let aba = planilha.getSheetByName('Aula ao vivo');
      if (!aba) {
        aba = planilha.insertSheet('Aula ao vivo');
        const cab = ['Data e hora', 'Aula', 'Questão', 'Leitura certa', 'Nº de chamada', 'Nome', 'Resposta', 'Resultado'];
        aba.appendRow(cab);
        aba.getRange(1, 1, 1, cab.length).setFontWeight('bold').setBackground('#1D3B5C').setFontColor('#ffffff');
        aba.setFrozenRows(1);
      }
      const linhas = (d.linhas || []).map(function (l) { return [new Date(), String(d.aula || ''), l.q, l.certa, l.chamada, l.nome, l.resp, l.resultado]; });
      if (linhas.length) aba.getRange(aba.getLastRow() + 1, 1, linhas.length, 8).setValues(linhas);
      return saida({ ok: true, gravadas: linhas.length });
    } finally { lock.releaseLock(); }
  }
  return saida({ ok: false, erro: 'acao' });
}

function saida(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
