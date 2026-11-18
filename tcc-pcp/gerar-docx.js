/* Assistente do TCC – Técnico em Planejamento e Controle da Produção (CEEP Erotides Ângelo Nichele)
   Monta o arquivo Word (.docx) formatado em ABNT a partir do que o grupo preencheu.
   Funciona no navegador (com ../vendor/docx.iife.js) e no Node (require('docx')). */
(function (root) {
  'use strict';

  const MESES_ABNT = ['jan.', 'fev.', 'mar.', 'abr.', 'maio', 'jun.', 'jul.', 'ago.', 'set.', 'out.', 'nov.', 'dez.'];

  function construir(st, imgs, D) {
    const {
      Document, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell, WidthType,
      ShadingType, BorderStyle, Header, PageNumber, TableOfContents, LevelFormat, TabStopType, Tab,
      VerticalAlign, HighlightColor, ImageRun, Bookmark, PageReference
    } = D;

    const FONT = 'Times New Roman';
    const W = 9071;
    const L15 = 360;
    const ano = st.ano || '2026';
    imgs = imgs || {};

    // ----------------------------------------------------------- utilitários
    const limpa = s => String(s == null ? '' : s).replace(/\r/g, '').trim();
    const tem = s => limpa(s) !== '';
    const paragrafos = s => limpa(s).split(/\n\s*\n|\n/).map(x => x.trim()).filter(Boolean);

    function tr(text, o = {}) {
      return new TextRun({ text, font: FONT, size: o.size || 24, bold: o.bold, italics: o.italics,
        highlight: o.falta ? HighlightColor.YELLOW : undefined });
    }
    const falta = (o = {}) => (oque) => tr(`[FALTA: ${oque}]`, Object.assign({}, o, { falta: true }));
    function txt(valor, oque, o = {}) { return tem(valor) ? tr(limpa(valor), o) : falta(o)(oque); }

    function P(children, o = {}) {
      return new Paragraph({ alignment: o.align || AlignmentType.JUSTIFIED,
        indent: o.noIndent ? undefined : { firstLine: 709 },
        spacing: { line: o.line || L15, before: o.before || 0, after: o.after || 0 },
        keepNext: o.keepNext, pageBreakBefore: o.pageBreak, children: Array.isArray(children) ? children : [children] });
    }
    function textoCorrido(valor, oque, o = {}) {
      const ps = paragrafos(valor);
      if (!ps.length) return [P([falta()(oque)], o)];
      return ps.map(t => P([tr(t)], o));
    }
    const C = (children, o = {}) => new Paragraph({ alignment: AlignmentType.CENTER,
      spacing: { line: o.line || L15, before: o.before || 0, after: o.after || 0 }, pageBreakBefore: o.pageBreak,
      children: Array.isArray(children) ? children : [children] });
    const vazio = (n = 1) => Array.from({ length: n }, () => new Paragraph({ spacing: { line: L15 }, children: [] }));

    function H1(t, o = {}) {
      return new Paragraph({ heading: HeadingLevel.HEADING_1, pageBreakBefore: o.first ? false : true,
        alignment: o.center ? AlignmentType.CENTER : AlignmentType.LEFT, children: [tr(t, { bold: true })] });
    }
    const H2 = t => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [tr(t)] });
    const H3 = t => new Paragraph({ heading: HeadingLevel.HEADING_3, children: [tr(t, { bold: true })] });
    const TIT = t => C([tr(t, { bold: true })], { pageBreak: true, after: 360 });

    const toc = [];
    const regH = (level, title) => toc.push({ title, level });
    const h1 = (t, o) => { regH(1, t); return H1(t, o); };
    const h2 = t => { regH(2, t); return H2(t); };
    const h3 = t => { regH(3, t); return H3(t); };

    let listaN = 0;
    function lista(itens, ref) {
      listaN++;
      return itens.map(t => new Paragraph({ numbering: { reference: ref, level: 0, instance: listaN },
        alignment: AlignmentType.JUSTIFIED, spacing: { line: L15 }, children: [tr(t)] }));
    }

    // ----------------------------------------------- quadros, tabelas e figuras
    const cont = { Figura: 0, Quadro: 0, Tabela: 0, eq: 0 };
    const ilustracoes = []; // {tipo, n, legenda, bm}
    const bord = { style: BorderStyle.SINGLE, size: 4, color: '000000' };

    function cel(t, w, o = {}) {
      const vazioCel = !tem(t);
      return new TableCell({ width: { size: w, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
        shading: o.head ? { type: ShadingType.CLEAR, fill: 'E7E6E6', color: 'auto' } : undefined,
        margins: { top: 40, bottom: 40, left: 80, right: 80 },
        borders: { top: bord, bottom: bord, left: bord, right: bord },
        children: [new Paragraph({ alignment: o.head ? AlignmentType.CENTER : (o.right ? AlignmentType.RIGHT : AlignmentType.LEFT),
          spacing: { line: 240 },
          children: [vazioCel && !o.podeVazio ? tr('—', { size: 20 }) : tr(limpa(t), { size: 20, bold: o.head || o.bold })] })] });
    }
    function tabela(cab, linhas, pesos, o = {}) {
      const tot = pesos.reduce((a, b) => a + b, 0);
      const ws = pesos.map(p => Math.floor(W * p / tot)); ws[ws.length - 1] += W - ws.reduce((a, b) => a + b, 0);
      return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: ws,
        rows: [new TableRow({ tableHeader: true, children: cab.map((h, i) => cel(h, ws[i], { head: true })) }),
          ...linhas.map((l, li) => new TableRow({ cantSplit: true, children: l.map((t, i) => cel(t, ws[i], {
            bold: o.negritoUltima && li === linhas.length - 1, right: o.direita && o.direita.includes(i), podeVazio: o.podeVazio })) }))] });
    }
    function legenda(tipo, texto) {
      cont[tipo]++;
      const n = cont[tipo];
      const bm = `${tipo.toLowerCase()}_${n}`;
      const leg = `${tipo} ${n} – ${texto}`;
      ilustracoes.push({ tipo, n, legenda: leg, bm });
      return { n, par: new Paragraph({ keepNext: true, keepLines: true, alignment: AlignmentType.CENTER,
        spacing: { before: 240, after: 60, line: 240 },
        children: [new Bookmark({ id: bm, children: [tr(leg, { size: 20, bold: true })] })] }) };
    }
    const fonte = (t) => new Paragraph({ spacing: { before: 60, after: 240, line: 240 },
      children: [tr(tem(t) ? limpa(t) : `Fonte: os autores (${ano}).`, { size: 20 })] });
    function quadro(tipo, texto, cab, linhas, pesos, fon, o) {
      const lg = legenda(tipo, texto);
      return { n: lg.n, blocos: [lg.par, tabela(cab, linhas, pesos, o), fonte(fon)] };
    }

    function imagem(fig) {
      const im = fig && fig.img && imgs[fig.img];
      if (!im) return null;
      const maxW = 560, maxH = 380;
      let w = im.w || 800, h = im.h || 600;
      const k = Math.min(maxW / w, maxH / h, 1);
      w = Math.round(w * k); h = Math.round(h * k);
      return new ImageRun({ type: im.type === 'png' ? 'png' : 'jpg', data: im.data, transformation: { width: w, height: h } });
    }
    // retorna {n, blocos}
    function figura(fig, legPadrao, oqueFalta) {
      const leg = tem(fig && fig.legenda) ? limpa(fig.legenda) : legPadrao;
      const lg = legenda('Figura', leg);
      const run = imagem(fig);
      const corpo = run
        ? new Paragraph({ keepNext: true, keepLines: true, alignment: AlignmentType.CENTER, spacing: { line: 240, lineRule: 'auto', before: 0, after: 0 }, children: [run] })
        : new Paragraph({ keepNext: true, alignment: AlignmentType.CENTER, spacing: { line: 240 },
            border: { top: bord, bottom: bord, left: bord, right: bord },
            children: [new TextRun({ text: '', break: 2 }), falta({ size: 20 })(oqueFalta || 'foto ou imagem'), new TextRun({ text: '', break: 2 })] });
      return { n: lg.n, blocos: [lg.par, corpo, fonte(fig && fig.fonte)] };
    }

    function equacao(expr) {
      cont.eq++;
      return new Paragraph({ tabStops: [{ type: TabStopType.CENTER, position: Math.floor(W / 2) }, { type: TabStopType.RIGHT, position: W }],
        spacing: { line: L15, before: 120, after: 120 }, keepNext: false,
        children: [new TextRun({ children: [new Tab()] }), tr(limpa(expr)), new TextRun({ children: [new Tab(), `(${cont.eq})`], font: FONT, size: 24 })] });
    }

    const rows = (arr, n) => (arr || []).filter(r => r && r.some(c => tem(c))).map(r => Array.from({ length: n }, (_, i) => limpa(r[i])));
    const num = v => { const x = parseFloat(String(v || '').replace(/[^\d,.\-]/g, '').replace(/\.(?=\d{3}(\D|$))/g, '').replace(',', '.')); return isNaN(x) ? 0 : x; };
    const real = v => v.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

    const integrantes = (st.integrantes || []).map(limpa).filter(Boolean);
    const titulo = limpa(st.titulo).toUpperCase();
    const subtitulo = limpa(st.subtitulo);
    const objEsp = (st.objEsp || []).map(limpa).filter(Boolean);
    const org = limpa(st.organizacao) || 'organização estudada';
    const simul = st.ficticia ? ', com dados simulados' : '';
    const fonteAut = `Fonte: os autores (${ano})${simul}.`;
    const ESCOLA = 'CENTRO ESTADUAL DE EDUCAÇÃO PROFISSIONAL DE FAZENDA RIO GRANDE - EROTIDES ÂNGELO NICHELE';

    // ================================================================ PRÉ-TEXTUAIS
    const pre = [];
    const tituloRuns = () => {
      const r = [titulo ? tr(titulo, { bold: true }) : falta({ bold: true })('título do trabalho')];
      if (subtitulo) r.push(tr(': ' + subtitulo, { bold: false }));
      return r;
    };
    const nomes = (pb) => (integrantes.length ? integrantes : ['']).map((n, i) =>
      C([n ? tr(n.toUpperCase(), { bold: true }) : falta({ bold: true })('nomes dos integrantes')], { pageBreak: pb && i === 0 }));
    // capa
    pre.push(
      C([tr(ESCOLA, { bold: true })]), ...vazio(1),
      C([tr('CURSO TÉCNICO EM PLANEJAMENTO E CONTROLE DA PRODUÇÃO', { bold: true })]),
      ...vazio(3), ...nomes(false), ...vazio(Math.max(2, 8 - integrantes.length)),
      C(tituloRuns()), ...vazio(9),
      C([tr('FAZENDA RIO GRANDE', { bold: true })]), C([tr(ano, { bold: true })]),
    );
    // folha de rosto
    const orientador = limpa(st.orientador) || 'Prof. Rodrigo Leandro Martins';
    const natureza = `Trabalho de Conclusão de Curso apresentado ao Curso Técnico em Planejamento e Controle da Produção do ${ESCOLA}, como requisito parcial para a obtenção do título de Técnico em Planejamento e Controle da Produção.`;
    pre.push(
      ...nomes(true), ...vazio(Math.max(2, 8 - integrantes.length)), C(tituloRuns()), ...vazio(3),
      new Paragraph({ indent: { left: 4535 }, alignment: AlignmentType.JUSTIFIED, spacing: { line: 240, after: 240 }, children: [tr(natureza, { size: 22 })] }),
      new Paragraph({ indent: { left: 4535 }, spacing: { line: 240 }, children: [tr('Orientador(a): ' + orientador, { size: 22 })] }),
      ...vazio(7), C([tr('FAZENDA RIO GRANDE', { bold: true })]), C([tr(ano, { bold: true })]),
    );
    if (tem(st.dedicatoria)) {
      pre.push(new Paragraph({ pageBreakBefore: true, children: [] }), ...vazio(20),
        ...paragrafos(st.dedicatoria).map(t => new Paragraph({ indent: { left: 4535 }, alignment: AlignmentType.JUSTIFIED, spacing: { line: L15 }, children: [tr(t, { italics: true })] })));
    }
    if (tem(st.agradecimentos)) pre.push(TIT('AGRADECIMENTOS'), ...textoCorrido(st.agradecimentos));
    pre.push(TIT('RESUMO'),
      new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { line: 240, after: 240 },
        children: [txt(paragrafos(st.resumo).join(' '), 'resumo')] }),
      new Paragraph({ spacing: { line: 240 }, children: [tr('Palavras-chave: ', { bold: true }), txt(st.palavras, 'palavras-chave')] }));

    // ===================================================== TEXTUAIS
    const T = [];
    const intro = (frase, q) => P([tr(frase)], { keepNext: true });
    const n = v => num(v);
    const fmt = (v, d = 1) => v.toLocaleString('pt-BR', { maximumFractionDigits: d });
    const temN = v => /\d/.test(String(v || ''));

    // 1 INTRODUÇÃO
    T.push(h1('1 INTRODUÇÃO', { first: true }), ...textoCorrido(st.tema, 'texto de abertura sobre o tema, com citação'),
      h2('1.1 CONTEXTUALIZAÇÃO DO PROBLEMA'), ...textoCorrido(st.contexto, 'contextualização do problema, com números'),
      h2('1.2 PROBLEMA DE PESQUISA'), ...textoCorrido(st.problema, 'pergunta do problema'),
      h2('1.3 JUSTIFICATIVA'), ...textoCorrido(st.justificativa, 'justificativa'),
      h2('1.4 OBJETIVOS'));
    const og = limpa(st.objGeral).replace(/^(o\s+)?objetivo geral( deste trabalho)?\s*(é|e|:)\s*/i, '');
    T.push(og ? P([tr('O objetivo geral deste trabalho é ' + og.charAt(0).toLowerCase() + og.slice(1).replace(/\.?$/, '.'))]) : P([falta()('objetivo geral')]));
    if (objEsp.length) {
      T.push(P([tr('Para alcançá-lo, foram definidos os seguintes objetivos específicos:')], { keepNext: true }));
      T.push(...lista(objEsp.map((o, i) => o.replace(/[.;]+$/, '') + (i === objEsp.length - 1 ? '.' : ';')).map(o => o.charAt(0).toLowerCase() + o.slice(1)), 'letras'));
    } else T.push(P([falta()('objetivos específicos')]));
    T.push(h2('1.5 ESTRUTURA DO TRABALHO'),
      P([tr(`Este trabalho está organizado em nove seções. A primeira apresenta o tema, o problema e os objetivos. A segunda caracteriza ${/^(a|o)\s/i.test(org) ? org.charAt(0).toLowerCase() + org.slice(1) : 'a organização ' + org}. A terceira reúne a fundamentação teórica. A quarta descreve a metodologia. A quinta apresenta o diagnóstico do processo atual. A sexta detalha a proposta de melhoria e a sétima, a aplicação das ferramentas de PCP e qualidade. A oitava apresenta e discute os resultados e a nona apresenta as conclusões.`)]));

    // 2 CARACTERIZAÇÃO
    T.push(h1('2 CARACTERIZAÇÃO DA ORGANIZAÇÃO'), h2('2.1 HISTÓRICO E SEGMENTO DE ATUAÇÃO'), ...textoCorrido(st.historico, 'histórico e segmento de atuação'));
    T.push(h2('2.2 PRODUTOS'));
    const prods = rows(st.produtos, 4);
    if (prods.length) {
      const pctOk = prods.every(r => temN(r[3]));
      const linhas = prods.map(r => [r[0], r[1], r[2], r[3]]);
      if (pctOk) linhas.push(['Total', '', '–', fmt(prods.reduce((a, r) => a + n(r[3]), 0))]);
      const q = quadro('Tabela', 'Produtos ou serviços e produção média', ['Família de produtos', 'Exemplos', 'Produção por período', 'Faturamento (%)'], linhas, [2, 2.8, 1.8, 1.5], fonteAut, { negritoUltima: pctOk, podeVazio: true });
      T.push(P([tr(`A Tabela ${q.n} apresenta os produtos ou serviços da organização.`)]), ...q.blocos);
    } else T.push(P([falta()('tabela de produtos ou serviços')]));
    if (tem(st.produtosTexto)) T.push(...textoCorrido(st.produtosTexto));
    T.push(h2('2.3 ESTRUTURA ORGANIZACIONAL'));
    const est = rows(st.estrutura, 3);
    if (est.length) {
      const tot = est.reduce((a, r) => a + n(r[2]), 0);
      const linhas = est.concat([['Total', '', fmt(tot, 0)]]);
      const q = quadro('Quadro', 'Estrutura organizacional', ['Setor', 'Função', 'Quantidade'], linhas, [2.2, 3.6, 1.4], fonteAut, { negritoUltima: true, direita: [2], podeVazio: true });
      T.push(P([tr(`A organização conta com ${fmt(tot, 0)} colaboradores, distribuídos conforme o Quadro ${q.n}.`)]), ...q.blocos);
    } else T.push(P([falta()('quadro da estrutura organizacional')]));
    T.push(h2('2.4 ESTRUTURA FÍSICA E CAPACIDADE PRODUTIVA'), ...textoCorrido(st.capacidade, 'estrutura física e capacidade produtiva'));
    T.push(h2('2.5 PROCESSO PRODUTIVO'), ...textoCorrido(st.processo, 'processo produtivo'));

    // 3 FUNDAMENTAÇÃO
    T.push(h1('3 FUNDAMENTAÇÃO TEÓRICA'));
    const fund = (st.fund || []).filter(f => f && (tem(f.titulo) || tem(f.texto)));
    if (!fund.length) T.push(P([falta()('tópicos da fundamentação teórica')]));
    fund.forEach((f, i) => { T.push(h2(`3.${i + 1} ${(limpa(f.titulo) || 'TÓPICO').toUpperCase()}`), ...textoCorrido(f.texto, 'texto do tópico')); });

    // 4 METODOLOGIA
    T.push(h1('4 METODOLOGIA'), h2('4.1 TIPO DE PESQUISA'), ...textoCorrido(st.tipoPesquisa, 'tipo de pesquisa'),
      h2('4.2 OBJETO DE ESTUDO'), ...textoCorrido(st.objeto, 'objeto de estudo'),
      h2('4.3 COLETA DE DADOS'), ...textoCorrido(st.coleta, 'coleta de dados'),
      h2('4.4 ANÁLISE DOS DADOS'), ...textoCorrido(st.analise, 'análise dos dados'),
      h2('4.5 ETAPAS DA PESQUISA'));
    const etapas = rows(st.etapas, 3);
    if (etapas.length) {
      const q = quadro('Quadro', 'Etapas da pesquisa', ['Etapa', 'Atividade', 'Ferramenta', 'Período'], etapas.map((e, i) => [String(i + 1), e[0], e[1], e[2]]), [0.8, 3, 3, 1.6], `Fonte: os autores (${ano}).`);
      T.push(P([tr(`O Quadro ${q.n} apresenta a sequência de etapas do trabalho, as ferramentas empregadas e o período de realização.`)]), ...q.blocos);
    } else T.push(P([falta()('quadro de etapas da pesquisa')]));

    // 5 DIAGNÓSTICO
    T.push(h1('5 DIAGNÓSTICO DO PROCESSO'), h2('5.1 DESCRIÇÃO DO PROCESSO ATUAL'), ...textoCorrido(st.procAtual, 'descrição do processo atual'));
    const fF = figura(st.figFluxo, 'Fluxograma do processo atual', 'fluxograma do processo atual');
    T.push(P([tr(`A Figura ${fF.n} apresenta o fluxograma desse processo.`)]), ...fF.blocos);
    if (tem(st.procAnalise)) T.push(...textoCorrido(st.procAnalise));
    T.push(h2('5.2 LEVANTAMENTO DE DADOS DO PROBLEMA'), ...textoCorrido(st.dadosTexto, 'como os dados foram levantados'));
    const dados = rows(st.dados, 2).filter(r => tem(r[0]) || temN(r[1]));
    if (dados.length) {
      const tot = dados.reduce((a, r) => a + n(r[1]), 0);
      const ord = dados.slice().sort((x, y) => n(y[1]) - n(x[1]));
      let ac = 0;
      const linhas = ord.map(r => { const pc = tot ? n(r[1]) / tot * 100 : 0; ac += pc; return [r[0], r[1], fmt(pc, 0), fmt(Math.min(ac, 100), 0)]; });
      linhas.push(['Total', fmt(tot, 2), '100', '–']);
      const leg = limpa(st.dadosUnid) ? limpa(st.dadosUnid).charAt(0).toUpperCase() + limpa(st.dadosUnid).slice(1) : 'Ocorrências por item';
      const q = quadro('Tabela', leg, ['Item ou causa', limpa(st.dadosUnid) ? 'Quantidade' : 'Ocorrências', '%', '% acumulado'], linhas, [3.6, 1.4, 1.1, 1.4], fonteAut, { negritoUltima: true, direita: [1, 2, 3], podeVazio: true });
      T.push(P([tr(`A Tabela ${q.n} organiza as ocorrências em ordem decrescente de frequência (análise de Pareto).`)]), ...q.blocos);
    } else T.push(P([falta()('tabela de ocorrências por item ou causa')]));
    if (tem(st.dadosAnalise)) T.push(...textoCorrido(st.dadosAnalise));
    const indA = rows(st.indAtuais, 3);
    if (indA.length) {
      const q = quadro('Tabela', 'Indicadores do processo atual', ['Indicador', 'Valor atual', 'Como foi obtido'], indA, [3.4, 1.4, 2.8], fonteAut);
      T.push(P([tr(`A Tabela ${q.n} reúne os indicadores que descrevem a situação atual.`)]), ...q.blocos);
    } else T.push(P([falta()('tabela de indicadores do processo atual')]));
    T.push(h2('5.3 ANÁLISE DAS CAUSAS'), ...textoCorrido(st.causasTexto, 'como as causas foram levantadas'));
    const fI = figura(st.figIshikawa, 'Diagrama de Ishikawa', 'diagrama de Ishikawa');
    T.push(P([tr(`As possíveis causas foram organizadas no diagrama de Ishikawa da Figura ${fI.n}.`)]), ...fI.blocos);
    const pq = rows(st.porques, 2);
    if (pq.length) {
      const q = quadro('Quadro', 'Aplicação dos 5 Porquês', ['Pergunta', 'Resposta'], pq, [3, 4], `Fonte: os autores (${ano}).`);
      T.push(P([tr(`Para aprofundar a causa principal, aplicou-se a técnica dos 5 Porquês, apresentada no Quadro ${q.n}.`)]), ...q.blocos);
    } else T.push(P([falta()('quadro dos 5 Porquês')]));
    if (tem(st.causasAnalise)) T.push(...textoCorrido(st.causasAnalise));
    T.push(h2('5.4 PRIORIZAÇÃO DAS CAUSAS'));
    const gut = rows(st.gut, 4).map(r => ({ r, p: temN(r[1]) && temN(r[2]) && temN(r[3]) ? n(r[1]) * n(r[2]) * n(r[3]) : 0 })).sort((a, b) => b.p - a.p);
    if (gut.length) {
      const q = quadro('Tabela', 'Priorização das causas pela matriz GUT', ['Causa', 'G', 'U', 'T', 'G × U × T', 'Prioridade'],
        gut.map((g, i) => [g.r[0], g.r[1], g.r[2], g.r[3], g.p ? String(g.p) : '', `${i + 1}ª`]), [3.6, 0.7, 0.7, 0.7, 1.3, 1.2], `Fonte: os autores (${ano}).`, { direita: [1, 2, 3, 4] });
      T.push(P([tr(`As causas identificadas foram priorizadas pela matriz GUT, com notas de 1 a 5 para gravidade (G), urgência (U) e tendência (T), conforme a Tabela ${q.n}.`)]), ...q.blocos);
    } else T.push(P([falta()('matriz GUT')]));
    T.push(h2('5.5 SÍNTESE DO DIAGNÓSTICO'), ...textoCorrido(st.sintese, 'síntese do diagnóstico'));

    // 6 PROPOSTA
    T.push(h1('6 PROPOSTA DE MELHORIA'), h2('6.1 RELAÇÃO ENTRE CAUSAS E AÇÕES'));
    const acoes = rows(st.acoes, 3);
    if (acoes.length) {
      const q = quadro('Quadro', 'Relação entre causas, ações e ferramentas', ['Causa (Capítulo 5)', 'Ação proposta', 'Ferramenta'], acoes, [2.4, 3.4, 1.8], `Fonte: os autores (${ano}).`);
      T.push(P([tr(`O Quadro ${q.n} relaciona cada causa identificada no diagnóstico à ação proposta e à ferramenta correspondente.`)]), ...q.blocos);
    } else T.push(P([falta()('quadro de causas e ações')]));
    T.push(h2('6.2 PLANO DE AÇÃO'));
    const custo = num(st.investimento);
    if (tem(st.planoTexto)) T.push(...textoCorrido(st.planoTexto));
    if (st.figPlano && st.figPlano.img) { const fg = figura(st.figPlano, 'Plano de ação'); T.push(P([tr(`O plano de ação está apresentado na Figura ${fg.n}.`)]), ...fg.blocos); }
    if (!tem(st.planoTexto) && !(st.figPlano && st.figPlano.img)) T.push(P([falta()('plano de ação')]));
    if (/\d/.test(String(st.investimento || ''))) T.push(P([tr(`O custo total estimado da proposta é de R$ ${real(custo)}.`)]));
    else T.push(P([falta()('custo total da proposta')]));
    T.push(h2('6.3 INDICADORES DE DESEMPENHO'));
    const inds = rows(st.indicadores, 4);
    if (inds.length) {
      const q = quadro('Quadro', 'Indicadores de desempenho', ['Indicador', 'Fórmula', 'Valor atual', 'Meta'], inds, [2.2, 3.6, 1.2, 1.2], `Fonte: os autores (${ano}).`);
      T.push(P([tr(`Para avaliar a melhoria, foram definidos os indicadores do Quadro ${q.n}, cada um com fórmula, valor atual e meta.`)]), ...q.blocos);
    } else T.push(P([falta()('quadro de indicadores')]));

    // 7 APLICAÇÃO
    T.push(h1('7 APLICAÇÃO DAS FERRAMENTAS DE PCP E QUALIDADE'));
    const fs = (st.ferramentas || []).filter(f => f && ([f.titulo, f.dados, f.formula, f.substituicao, f.significado].some(tem) || (f.fig && f.fig.img)));
    if (!fs.length) T.push(P([falta()('aplicação das ferramentas')]));
    fs.forEach((f, i) => {
      T.push(h2(`7.${i + 1} ${(limpa(f.titulo) || 'FERRAMENTA').toUpperCase()}`));
      T.push(...textoCorrido(f.dados, 'dados de entrada'));
      if (tem(f.formula)) {
        const ne = cont.eq + 1;
        T.push(P([tr(`O cálculo foi feito pela Equação ${ne}:`)], { keepNext: true }), equacao(f.formula));
        if (tem(f.variaveis)) T.push(P([tr('Onde ' + limpa(f.variaveis).replace(/^(onde|em que)\s*/i, '').replace(/\.$/, '') + '.')], { noIndent: true }));
      }
      if (tem(f.substituicao)) T.push(P([tr(tem(f.formula) ? 'Substituindo os valores:' : 'O cálculo resultou em:')], { keepNext: true }), equacao(f.substituicao));
      if (f.fig && f.fig.img) { const fg = figura(f.fig, limpa(f.fig.legenda) || limpa(f.titulo)); T.push(P([tr(`O resultado está apresentado na Figura ${fg.n}.`)]), ...fg.blocos); }
      T.push(...textoCorrido(f.significado, 'interpretação do resultado'));
    });

    // 8 RESULTADOS
    T.push(h1('8 RESULTADOS E DISCUSSÃO'), h2('8.1 APRESENTAÇÃO DOS RESULTADOS'));
    const res = rows(st.resultados, 5);
    const variacao = (a, d) => {
      if (!temN(a) || !temN(d)) return '–';
      const x = n(a), y = n(d);
      if (/%/.test(a + d)) { const pp = y - x; return (pp > 0 ? '+' : pp < 0 ? '−' : '') + fmt(Math.abs(pp)) + ' p.p.'; }
      if (!x) return '–';
      const v = (y - x) / x * 100; return (v > 0 ? '+' : v < 0 ? '−' : '') + fmt(Math.abs(v), 0) + '%';
    };
    if (res.length) {
      const q = quadro('Tabela', 'Indicadores antes e depois da melhoria', ['Indicador', 'Antes', 'Depois', 'Variação', 'Meta', 'Situação'],
        res.map(r => [r[0], r[1] || '–', r[2], variacao(r[1], r[2]), r[3], r[4]]), [2.6, 1, 1, 1.2, 1, 1.3], fonteAut, { podeVazio: true });
      T.push(P([tr(`A Tabela ${q.n} compara os indicadores antes e depois da melhoria.`)]), ...q.blocos);
    } else T.push(P([falta()('tabela de indicadores antes e depois')]));
    T.push(...textoCorrido(st.resultadosTexto, 'como os resultados foram obtidos e o que mostram'));
    if (st.figResultado && st.figResultado.img) { const fg = figura(st.figResultado, 'Indicadores antes e depois da melhoria'); T.push(P([tr(`A Figura ${fg.n} ilustra a comparação.`)]), ...fg.blocos); }
    T.push(h2('8.2 ANÁLISE FINANCEIRA'), ...textoCorrido(st.economiaTexto, 'cálculo da economia mensal'));
    const eco = n(st.economiaMensal);
    if (custo && eco) {
      const pb = custo / eco;
      T.push(P([tr(`Como o investimento total foi de R$ ${real(custo)} e a economia estimada é de R$ ${real(eco)} por mês, o tempo de retorno é obtido pela Equação ${cont.eq + 1}:`)], { keepNext: true }),
        equacao(`Payback = Investimento ÷ Economia mensal = ${real(custo)} ÷ ${real(eco)} ≈ ${fmt(pb, 1)} ${pb < 2 ? 'mês' : 'meses'}`));
    } else T.push(P([falta()('custo total da proposta e economia mensal para o payback')]));
    if (tem(st.financeiroTexto)) T.push(...textoCorrido(st.financeiroTexto));
    T.push(h2('8.3 DISCUSSÃO'));
    if (objEsp.length) {
      const disc = st.discussao || [];
      const linhas = objEsp.map((o, i) => [o, limpa((disc[i] || {}).resultado), limpa((disc[i] || {}).situacao)]);
      const q = quadro('Quadro', 'Objetivos específicos e resultados obtidos', ['Objetivo específico', 'Resultado obtido', 'Situação'], linhas, [3, 3.5, 1.8], `Fonte: os autores (${ano}).`);
      T.push(P([tr(`O Quadro ${q.n} relaciona os objetivos específicos aos resultados obtidos.`)]), ...q.blocos);
    }
    T.push(...textoCorrido(st.discussaoTexto, 'discussão dos resultados'));

    // 9 CONCLUSÃO
    T.push(h1('9 CONCLUSÃO'), ...textoCorrido(st.conclusao, 'conclusão'));

    // REFERÊNCIAS
    T.push(h1('REFERÊNCIAS', { center: true }));
    const refs = (st.referencias || []).filter(r => r && (tem(r.titulo) || tem(r.livre))).map(formatarRef).sort((a, b) => a.chave.localeCompare(b.chave, 'pt-BR'));
    if (!refs.length) T.push(P([falta()('referências')], { noIndent: true }));
    refs.forEach(r => T.push(new Paragraph({ alignment: AlignmentType.LEFT, spacing: { line: 240, after: 240 }, children: r.partes.map(p => tr(p.t, { bold: p.b })) })));

    // APÊNDICES
    let letra = 0;
    const apx = t => h1(`APÊNDICE ${'ABCDEFG'[letra++]} – ${t}`, { center: true });
    const perg = paragrafos(st.roteiro);
    if (perg.length) { T.push(apx('ROTEIRO DE ENTREVISTA')); T.push(...lista(perg.map(p => p.replace(/^\d+[.)-]\s*/, '')), 'numeros')); }
    T.push(apx('REGISTRO DO USO DE INTELIGÊNCIA ARTIFICIAL'));
    const ia = rows(st.ia, 5);
    if (st.naoUsouIA) T.push(P([tr('A equipe não utilizou ferramentas de inteligência artificial na elaboração deste trabalho.')]));
    else if (ia.length) T.push(...quadro('Quadro', 'Registro do uso de inteligência artificial', ['Ferramenta', 'Data', 'Finalidade', 'Pedido feito (prompt)', 'Como foi verificado e usado'], ia, [1.3, 1.4, 1.7, 2.5, 2.3], `Fonte: os autores (${ano}).`).blocos);
    else T.push(P([falta()('registro do uso de IA (ou a declaração de que não usaram)')]));
    T.push(apx('DIVISÃO DE TAREFAS'));
    const tarefas = integrantes.map((nm, i) => [nm, limpa((st.tarefas || [])[i])]);
    if (tarefas.length) T.push(...quadro('Quadro', 'Divisão de tarefas', ['Integrante', 'Tarefas principais'], tarefas, [2.5, 6], `Fonte: os autores (${ano}).`).blocos);

    // ---------------------------------------------------- listas e sumário
    const linhaLista = (it) => new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: W, leader: 'dot' }], spacing: { line: L15 },
      children: [tr(it.legenda), new TextRun({ children: [new Tab()], font: FONT, size: 24 }), new PageReference(it.bm)] });
    const figs = ilustracoes.filter(i => i.tipo === 'Figura');
    const quads = ilustracoes.filter(i => i.tipo === 'Quadro'), tabs = ilustracoes.filter(i => i.tipo === 'Tabela');
    if (figs.length) pre.push(TIT('LISTA DE FIGURAS'), ...figs.map(linhaLista));
    if (quads.length) pre.push(TIT('LISTA DE QUADROS'), ...quads.map(linhaLista));
    if (tabs.length) pre.push(TIT('LISTA DE TABELAS'), ...tabs.map(linhaLista));
    const siglas = rows(st.siglas, 2).sort((a, b) => a[0].localeCompare(b[0], 'pt-BR'));
    if (siglas.length) pre.push(TIT('LISTA DE ABREVIATURAS E SIGLAS'), ...siglas.map(([s, t]) =>
      new Paragraph({ tabStops: [{ type: TabStopType.LEFT, position: 1700 }], spacing: { line: L15 }, children: [tr(s), new TextRun({ children: [new Tab()] }), tr(t)] })));
    pre.push(TIT('SUMÁRIO'), new TableOfContents('Sumário', { hyperlink: true, headingStyleRange: '1-3', cachedEntries: toc }));

    // ------------------------------------------------------------- documento
    const margem = { top: 1701, left: 1701, bottom: 1134, right: 1134, header: 709, footer: 709 };
    const pagina = { size: { width: 11906, height: 16838 }, margin: margem };
    const lvl = (ref, format, text) => ({ reference: ref, levels: [{ level: 0, format, text, alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] });
    return new Document({
      features: { updateFields: true },
      creator: integrantes.join(', ') || 'Grupo do TCC',
      title: titulo || 'TCC – Técnico em PCP',
      description: 'Gerado pelo Assistente do TCC – Técnico em PCP (CEEP Fazenda Rio Grande)',
      styles: {
        default: { document: { run: { font: FONT, size: 24 }, paragraph: { spacing: { line: L15 } } } },
        paragraphStyles: [
          { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
            run: { font: FONT, size: 24, bold: true, allCaps: true, color: '000000' },
            paragraph: { spacing: { before: 0, after: 360, line: L15 }, keepNext: true, outlineLevel: 0 } },
          { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
            run: { font: FONT, size: 24, bold: false, allCaps: true, color: '000000' },
            paragraph: { spacing: { before: 360, after: 240, line: L15 }, keepNext: true, outlineLevel: 1 } },
          { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
            run: { font: FONT, size: 24, bold: true, color: '000000' },
            paragraph: { spacing: { before: 240, after: 240, line: L15 }, keepNext: true, outlineLevel: 2 } },
          { id: 'TOC1', name: 'toc 1', basedOn: 'Normal', next: 'Normal', run: { font: FONT, size: 24, bold: true }, paragraph: { spacing: { line: L15, before: 0, after: 0 } } },
          { id: 'TOC2', name: 'toc 2', basedOn: 'Normal', next: 'Normal', run: { font: FONT, size: 24 }, paragraph: { indent: { left: 284 }, spacing: { line: L15, before: 0, after: 0 } } },
          { id: 'TOC3', name: 'toc 3', basedOn: 'Normal', next: 'Normal', run: { font: FONT, size: 24 }, paragraph: { indent: { left: 567 }, spacing: { line: L15, before: 0, after: 0 } } },
        ],
      },
      numbering: { config: [lvl('numeros', LevelFormat.DECIMAL, '%1.'), lvl('letras', LevelFormat.LOWER_LETTER, '%1)')] },
      sections: [
        { properties: { page: pagina }, children: pre },
        { properties: { page: pagina },
          headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 20 })] })] }) },
          children: T },
      ],
    });
  }

  // ------------------------------------------------------- referências ABNT
  function formatarRef(r) {
    const L = s => String(s == null ? '' : s).trim();
    const fim = s => s.replace(/[.\s]+$/, '');
    const partes = [];
    const add = (t, b) => { if (t) partes.push({ t, b }); };
    const autor = () => {
      const sob = L(r.sobrenome), nome = L(r.nome);
      if (sob) return sob.toUpperCase() + (nome ? ', ' + nome : '') + '. ';
      return '';
    };
    let chave = (L(r.sobrenome) || L(r.titulo) || L(r.livre)).toUpperCase();
    const acesso = () => {
      const d = L(r.acesso);
      const m = d.match(/^(\d{4})-(\d{2})-(\d{2})$/);
      return m ? ` Acesso em: ${parseInt(m[3], 10)} ${MESES_ABNT[parseInt(m[2], 10) - 1]} ${m[1]}.` : (d ? ` Acesso em: ${d}.` : '');
    };
    if (r.tipo === 'livre') { add(fim(L(r.livre)) + '.'); return { chave, partes }; }
    if (r.tipo === 'livro') {
      add(autor());
      add(fim(L(r.titulo)) + (L(r.subtitulo) ? '' : '. '), true);
      if (L(r.subtitulo)) add(': ' + fim(L(r.subtitulo)) + '. ');
      if (L(r.edicao)) add(L(r.edicao).replace(/\D/g, '') + '. ed. ');
      add(`${L(r.cidade) || '[S. l.]'}: ${L(r.editora) || '[s. n.]'}, ${L(r.ano) || '[s. d.]'}.`);
    } else if (r.tipo === 'norma') {
      add(autor() || (L(r.entidade) ? L(r.entidade).toUpperCase() + '. ' : ''));
      if (!L(r.sobrenome) && L(r.entidade)) chave = L(r.entidade).toUpperCase();
      add(fim(L(r.titulo)) + '. ', true);
      add(`${L(r.cidade) || '[S. l.]'}: ${L(r.editora) || '[s. n.]'}, ${L(r.ano) || '[s. d.]'}.`);
      if (L(r.url)) add(` Disponível em: ${L(r.url)}.` + acesso());
    } else { // site
      add(autor() || (L(r.entidade) ? L(r.entidade).toUpperCase() + '. ' : ''));
      if (!L(r.sobrenome) && L(r.entidade)) chave = L(r.entidade).toUpperCase();
      add(fim(L(r.titulo)) + '. ', true);
      if (L(r.ano)) add(L(r.ano) + '. ');
      add(`Disponível em: ${L(r.url) || '[endereço]'}.` + acesso());
    }
    return { chave, partes };
  }

  const api = { construir, formatarRef };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.TCCDocx = api;
})(typeof self !== 'undefined' ? self : this);
