/* Assistente do TCC – Técnico em Mecânica (CEEP Erotides Ângelo Nichele)
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
    const num = v => { const x = parseFloat(String(v || '').replace(/\./g, '').replace(',', '.')); return isNaN(x) ? 0 : x; };
    const real = v => v.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

    const integrantes = (st.integrantes || []).map(limpa).filter(Boolean);
    const titulo = limpa(st.titulo).toUpperCase();
    const subtitulo = limpa(st.subtitulo);
    const objEsp = (st.objEsp || []).map(limpa).filter(Boolean);
    const equip = limpa(st.equipamento) || 'equipamento';

    // ================================================================ PRÉ-TEXTUAIS
    const pre = [];
    const tituloRuns = (o = {}) => {
      const r = [titulo ? tr(titulo, { bold: true }) : falta({ bold: true })('título do trabalho')];
      if (subtitulo) r.push(tr(': ' + subtitulo, { bold: false }));
      return r;
    };
    const nomes = (pb) => (integrantes.length ? integrantes : ['']).map((n, i) =>
      C([n ? tr(n.toUpperCase(), { bold: true }) : falta({ bold: true })('nomes dos integrantes')], { pageBreak: pb && i === 0 }));

    // capa
    pre.push(
      C([tr('CENTRO ESTADUAL DE EDUCAÇÃO PROFISSIONAL EROTIDES ÂNGELO NICHELE', { bold: true })]),
      C([tr((limpa(st.curso) || 'Curso Técnico em Mecânica').toUpperCase(), { bold: true })]),
      ...vazio(4), ...nomes(false), ...vazio(Math.max(2, 9 - integrantes.length)),
      C(tituloRuns()), ...vazio(9),
      C([tr('FAZENDA RIO GRANDE', { bold: true })]), C([tr(ano, { bold: true })]),
    );
    // folha de rosto
    const natureza = `Trabalho de Conclusão de Curso apresentado ao ${limpa(st.curso) || 'Curso Técnico em Mecânica'} do Centro Estadual de Educação Profissional Erotides Ângelo Nichele, como requisito parcial para a obtenção do título de Técnico em Mecânica.`;
    const orientador = limpa(st.orientador) || 'Prof. Egon Henrique Felício';
    const blocoNat = () => new Paragraph({ indent: { left: 4535 }, alignment: AlignmentType.JUSTIFIED, spacing: { line: 240, after: 240 }, children: [tr(natureza, { size: 22 })] });
    pre.push(
      ...nomes(true), ...vazio(Math.max(2, 8 - integrantes.length)), C(tituloRuns()), ...vazio(3), blocoNat(),
      new Paragraph({ indent: { left: 4535 }, spacing: { line: 240 }, children: [tr('Orientador: ' + orientador, { size: 22 })] }),
      ...vazio(7), C([tr('FAZENDA RIO GRANDE', { bold: true })]), C([tr(ano, { bold: true })]),
    );
    // folha de aprovação
    const assin = (nome, papel) => [
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 600, line: 240 }, children: [tr('_____________________________________________')] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { line: 240 }, children: [tr(nome)] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { line: 240 }, children: [tr(papel, { size: 22 })] }),
    ];
    const banca = (st.banca || []).filter(b => b && tem(b.nome));
    pre.push(
      ...nomes(true), ...vazio(1), C(tituloRuns()), ...vazio(1), blocoNat(),
      C([tr('BANCA EXAMINADORA', { bold: true })], { before: 240 }),
      ...assin(orientador, 'Orientador – CEEP Erotides Ângelo Nichele'),
      ...(banca.length ? banca : [{ nome: '', inst: '' }, { nome: '', inst: '' }]).flatMap(b =>
        assin(tem(b.nome) ? b.nome : ' ', tem(b.inst) ? b.inst : 'Membro da banca')),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 600 }, children: [tr(`Fazenda Rio Grande, ____ de ______________ de ${ano}.`)] }),
    );
    // dedicatória e agradecimentos
    if (tem(st.dedicatoria)) {
      pre.push(new Paragraph({ pageBreakBefore: true, children: [] }), ...vazio(20),
        ...paragrafos(st.dedicatoria).map(t => new Paragraph({ indent: { left: 4535 }, alignment: AlignmentType.JUSTIFIED, spacing: { line: L15 }, children: [tr(t, { italics: true })] })));
    }
    if (tem(st.agradecimentos)) pre.push(TIT('AGRADECIMENTOS'), ...textoCorrido(st.agradecimentos));
    // resumo
    pre.push(TIT('RESUMO'),
      new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { line: 240, after: 240 },
        children: [txt(paragrafos(st.resumo).join(' '), 'resumo')] }),
      new Paragraph({ spacing: { line: 240 }, children: [tr('Palavras-chave: ', { bold: true }), txt(st.palavras, 'palavras-chave')] }));

    // ===================================================== TEXTUAIS (montados antes das listas)
    const T = [];
    // 1 INTRODUÇÃO
    T.push(h1('1 INTRODUÇÃO', { first: true }), h2('1.1 CONTEXTUALIZAÇÃO'), ...textoCorrido(st.contexto, 'contextualização'),
      h2('1.2 PROBLEMA'), ...textoCorrido(st.problema, 'pergunta do problema'),
      h2('1.3 JUSTIFICATIVA'), ...textoCorrido(st.justificativa, 'justificativa'),
      h2('1.4 OBJETIVOS'), h3('1.4.1 Objetivo geral'), ...textoCorrido(st.objGeral, 'objetivo geral'),
      h3('1.4.2 Objetivos específicos'));
    if (objEsp.length) {
      T.push(P([tr('São objetivos específicos deste trabalho:')]));
      T.push(...lista(objEsp.map((o, i) => o.replace(/[.;]+$/, '') + (i === objEsp.length - 1 ? '.' : ';')).map(o => o.charAt(0).toLowerCase() + o.slice(1)), 'letras'));
    } else T.push(P([falta()('objetivos específicos')]));

    // 2 FUNDAMENTAÇÃO
    T.push(h1('2 FUNDAMENTAÇÃO TEÓRICA'));
    const fund = (st.fund || []).filter(f => f && (tem(f.titulo) || tem(f.texto)));
    if (!fund.length) T.push(P([falta()('tópicos da fundamentação teórica')]));
    fund.forEach((f, i) => { T.push(h2(`2.${i + 1} ${(limpa(f.titulo) || 'TÓPICO').toUpperCase()}`), ...textoCorrido(f.texto, 'texto do tópico')); });

    // 3 METODOLOGIA
    T.push(h1('3 METODOLOGIA'), ...textoCorrido(st.metodo, 'metodologia'));
    const etapas = rows(st.etapas, 3);
    if (etapas.length) {
      const q = quadro('Quadro', 'Etapas do desenvolvimento', ['Etapa', 'O que foi feito', 'Período'], etapas, [2, 5, 1.5]);
      T.splice(T.length, 0, P([tr(`As etapas do desenvolvimento estão apresentadas no Quadro ${q.n}.`)]), ...q.blocos);
    }
    if (tem(st.instrumentos)) T.push(P([tr('Para as medições, foram utilizados os seguintes instrumentos: ' + limpa(st.instrumentos).replace(/\.$/, '') + '.')]));
    T.push(P([tr('A divisão das tarefas entre os integrantes do grupo está descrita no Apêndice C.')]));

    // 4 PROJETO
    T.push(h1('4 PROJETO DO EQUIPAMENTO'), h2('4.1 PRODUTOS SEMELHANTES'));
    const prods = rows(st.produtos, 5);
    if (prods.length) {
      const q = quadro('Quadro', 'Produtos semelhantes', ['Modelo', 'Acionamento', 'Capacidade', 'Preço (R$)', 'Fonte'], prods, [2, 1.6, 2, 1.2, 2.2],
        `Fonte: os autores (${ano}), com base nas fontes indicadas.`);
      T.push(P([tr(`Para conhecer as soluções existentes, foram pesquisados modelos disponíveis no mercado. O Quadro ${q.n} apresenta a comparação.`)]), ...q.blocos);
    } else T.push(P([falta()('quadro de produtos semelhantes')]));
    T.push(...textoCorrido(st.diferencial, 'diferença do projeto em relação ao mercado'));
    T.push(h2('4.2 EVOLUÇÃO DA PROPOSTA'), ...textoCorrido(st.evolucao, 'evolução da proposta'));
    T.push(h2('4.3 PRINCÍPIO DE FUNCIONAMENTO'));
    const fF = figura(st.figFunc, `Partes principais do(a) ${equip}`, 'foto ou imagem com as partes numeradas');
    T.push(P([tr(`A Figura ${fF.n} apresenta as partes principais do equipamento.`)]), ...fF.blocos);
    const passos = paragrafos(st.funcionamento);
    if (passos.length) { T.push(P([tr('O funcionamento ocorre da seguinte forma:')])); T.push(...lista(passos.map(p => p.replace(/^\d+[.)-]\s*/, '')), 'numeros')); }
    else T.push(P([falta()('passo a passo do funcionamento')]));
    T.push(h2('4.4 DESENHO TÉCNICO E MODELAGEM'), ...textoCorrido(st.modelagem, 'texto sobre a modelagem'));
    const fC = figura(st.figCad, 'Vista geral do modelo em CAD', 'imagem do modelo em CAD');
    T.push(...fC.blocos);
    T.push(h2('4.5 LISTA DE MATERIAIS E COMPONENTES'));
    const mats = rows(st.materiais, 5);
    if (mats.length) {
      const q = quadro('Quadro', 'Lista de materiais e componentes', ['Item', 'Qtd.', 'Especificação (material e medidas)', 'Função', 'Origem'], mats, [1.6, 0.7, 2.8, 2, 1.8]);
      T.push(P([tr(`O Quadro ${q.n} apresenta os materiais e componentes utilizados no equipamento.`)]), ...q.blocos);
    } else T.push(P([falta()('lista de materiais')]));

    // 5 DIMENSIONAMENTO
    T.push(h1('5 DIMENSIONAMENTO'));
    if (tem(st.dimIntro)) T.push(...textoCorrido(st.dimIntro));
    const calcs = (st.calculos || []).filter(c => c && (tem(c.titulo) || tem(c.formula) || tem(c.dados)));
    if (!calcs.length) T.push(P([falta()('cálculos do projeto')]));
    calcs.forEach((c, i) => {
      T.push(h2(`5.${i + 1} ${(limpa(c.titulo) || 'CÁLCULO').toUpperCase()}`));
      T.push(...textoCorrido(c.dados, 'dados usados no cálculo'));
      if (tem(c.formula)) { const n = cont.eq + 1; T.push(P([tr(`O cálculo foi feito pela Equação ${n}:`)], { keepNext: true }), equacao(c.formula)); }
      else T.push(P([falta()('fórmula')]));
      if (tem(c.variaveis)) T.push(P([tr('em que ' + limpa(c.variaveis).replace(/^em que\s*/i, '').replace(/\.$/, '') + '. Substituindo os valores:')], { noIndent: true, keepNext: true }));
      if (tem(c.substituicao)) T.push(equacao(c.substituicao)); else T.push(P([falta()('substituição dos valores e resultado com unidade')]));
      T.push(...textoCorrido(c.significado, 'o que o resultado significa para o projeto'));
    });

    // 6 CONSTRUÇÃO
    T.push(h1('6 CONSTRUÇÃO DO PROTÓTIPO'), h2('6.1 ETAPAS DE FABRICAÇÃO E MONTAGEM'));
    const etC = (st.construcao || []).filter(e => e && (tem(e.texto) || (e.foto && e.foto.img)));
    if (!etC.length) T.push(P([falta()('etapas da construção com fotos')]));
    etC.forEach((e, i) => {
      const f = figura(e.foto, (e.foto && limpa(e.foto.legenda)) || `Etapa ${i + 1} da construção`, 'foto da etapa');
      const ps = paragrafos(e.texto);
      if (ps.length) { ps.forEach((t, j) => T.push(P([tr(t + (j === ps.length - 1 ? ` (Figura ${f.n}).` : '').replace(/\.\s*\(Figura/, ' (Figura'))]))); }
      else T.push(P([falta()('texto da etapa')]));
      T.push(...f.blocos);
    });
    T.push(h2('6.2 DIFICULDADES E ALTERAÇÕES'));
    const difs = rows(st.dificuldades, 4);
    if (difs.length) {
      const q = quadro('Quadro', 'Dificuldades e alterações', ['Problema', 'Causa', 'Solução adotada', 'Resultado'], difs, [2, 1.8, 2.4, 1.8]);
      T.push(P([tr(`Durante a construção, surgiram dificuldades que exigiram alterações no projeto. O Quadro ${q.n} apresenta os principais problemas e as soluções adotadas.`)]), ...q.blocos);
    } else T.push(P([falta()('quadro de dificuldades e alterações')]));
    if (tem(st.dificuldadesTexto)) T.push(...textoCorrido(st.dificuldadesTexto));

    // 7 CUSTOS
    T.push(h1('7 CUSTOS'));
    const custos = rows(st.custos, 4);
    if (custos.length) {
      let total = 0;
      const linhas = custos.map(r => { const t = num(r[1] || '1') * num(r[2]); total += t; return [r[0], r[1], real(num(r[2])), real(t), r[3]]; });
      linhas.push(['TOTAL', '', '', real(total), '']);
      const q = quadro('Tabela', 'Custos de fabricação', ['Item', 'Qtd.', 'Valor unitário (R$)', 'Total (R$)', 'Origem'], linhas, [2.6, 0.8, 1.6, 1.4, 2.4], null,
        { negritoUltima: true, direita: [1, 2, 3], podeVazio: true });
      T.push(P([tr(`A Tabela ${q.n} apresenta os custos de fabricação. Os materiais reaproveitados foram listados com valor zero.`)]), ...q.blocos);
    } else T.push(P([falta()('tabela de custos')]));
    T.push(...textoCorrido(st.custosTexto, 'comparação do custo com os produtos do mercado'));

    // 8 SEGURANÇA
    T.push(h1('8 SEGURANÇA'));
    const riscos = rows(st.riscos, 4);
    const nQ = cont.Quadro + 1;
    T.push(P([tr(`Os riscos do equipamento foram analisados tomando como referência a Norma Regulamentadora NR-12, que trata da segurança no trabalho em máquinas e equipamentos. ${riscos.length ? `O Quadro ${nQ} apresenta os riscos identificados e as medidas adotadas.` : ''}`)]));
    if (riscos.length) T.push(...quadro('Quadro', 'Riscos e medidas de segurança', ['Risco', 'Onde', 'O que pode acontecer', 'Medida adotada'], riscos, [1.8, 2, 2.2, 2.4]).blocos);
    else T.push(P([falta()('quadro de riscos')]));
    if (tem(st.segurancaTexto)) T.push(...textoCorrido(st.segurancaTexto));
    if (tem(st.epis)) T.push(P([tr('Durante a fabricação e os testes, foram utilizados os seguintes EPIs: ' + limpa(st.epis).replace(/\.$/, '') + '.')]));

    // 9 TESTES
    T.push(h1('9 TESTES E RESULTADOS'), h2('9.1 PROCEDIMENTO DE TESTE'), ...textoCorrido(st.testeProc, 'procedimento de teste'));
    T.push(h2('9.2 RESULTADOS'));
    const m1 = limpa(st.medida1) || 'Medida 1', m2 = limpa(st.medida2);
    const reps = rows(st.testes, 3).filter(r => tem(r[0]) || tem(r[1]) || tem(r[2]));
    if (reps.length) {
      const cab = m2 ? ['Repetição', m1, m2, 'Observação'] : ['Repetição', m1, 'Observação'];
      const linhas = reps.map((r, i) => m2 ? [String(i + 1), r[0], r[1], r[2]] : [String(i + 1), r[0], r[2]]);
      const media = k => { const v = reps.map(r => r[k]).filter(x => /\d/.test(x)).map(num); return v.length ? (v.reduce((a, b) => a + b, 0) / v.length).toLocaleString('pt-BR', { maximumFractionDigits: 2 }) : ''; };
      if (reps.length > 1) linhas.push(m2 ? ['Média', media(0), media(1), ''] : ['Média', media(0), '']);
      const q = quadro('Tabela', 'Resultados dos testes', cab, linhas, m2 ? [1.4, 2, 2, 3] : [1.4, 2.5, 4], null, { negritoUltima: reps.length > 1, podeVazio: true });
      T.push(P([tr(`Os valores medidos em cada repetição estão na Tabela ${q.n}.`)]), ...q.blocos);
    } else T.push(P([falta()('tabela com os resultados das repetições')]));
    if (st.figTeste && st.figTeste.img) T.push(...figura(st.figTeste, 'Teste do equipamento').blocos);
    T.push(...textoCorrido(st.testeAnalise, 'descrição dos resultados'));
    T.push(h2('9.3 DISCUSSÃO'));
    if (objEsp.length) {
      const disc = st.discussao || [];
      const linhas = objEsp.map((o, i) => [o, limpa((disc[i] || {}).resultado), limpa((disc[i] || {}).situacao)]);
      const q = quadro('Quadro', 'Comparação entre objetivos e resultados', ['Objetivo específico', 'Resultado obtido', 'Situação'], linhas, [3, 3.5, 2]);
      T.push(P([tr(`O Quadro ${q.n} compara os objetivos específicos com os resultados obtidos.`)]), ...q.blocos);
    }
    T.push(...textoCorrido(st.discussaoTexto, 'discussão dos resultados'));

    // 10 CONCLUSÃO
    T.push(h1('10 CONCLUSÃO'), ...textoCorrido(st.conclusao, 'conclusão'));

    // REFERÊNCIAS
    T.push(h1('REFERÊNCIAS', { center: true }));
    const refs = (st.referencias || []).filter(r => r && (tem(r.titulo) || tem(r.livre))).map(formatarRef).sort((a, b) => a.chave.localeCompare(b.chave, 'pt-BR'));
    if (!refs.length) T.push(P([falta()('referências')], { noIndent: true }));
    refs.forEach(r => T.push(new Paragraph({ alignment: AlignmentType.LEFT, spacing: { line: 240, after: 240 }, children: r.partes.map(p => tr(p.t, { bold: p.b })) })));

    // APÊNDICES
    let letra = 0;
    const apx = t => h1(`APÊNDICE ${'ABCDEFG'[letra++]} – ${t}`, { center: true });
    T.push(apx('DESENHOS TÉCNICOS'));
    const des = (st.desenhos || []).filter(d => d && d.img);
    if (!des.length) T.push(...figura(null, `Desenho de conjunto do(a) ${equip}`, 'desenho técnico cotado').blocos);
    des.forEach((d, i) => T.push(...figura(d, `Desenho técnico ${i + 1}`).blocos));
    T.push(apx('REGISTRO DO USO DE INTELIGÊNCIA ARTIFICIAL'));
    const ia = rows(st.ia, 4);
    if (st.naoUsouIA) T.push(P([tr('O grupo não utilizou ferramentas de inteligência artificial na elaboração deste trabalho.')]));
    else if (ia.length) T.push(...quadro('Quadro', 'Uso de inteligência artificial', ['Ferramenta', 'Para que foi usada', 'Em qual parte do TCC', 'Como o grupo conferiu'], ia, [1.6, 2.4, 1.8, 2.6]).blocos);
    else T.push(P([falta()('registro do uso de IA (ou a declaração de que não usaram)')]));
    T.push(apx('DIVISÃO DE TAREFAS'));
    const tarefas = integrantes.map((n, i) => [n, limpa((st.tarefas || [])[i])]);
    if (tarefas.length) T.push(...quadro('Quadro', 'Divisão de tarefas', ['Integrante', 'Tarefas principais'], tarefas, [2.5, 6]).blocos);
    if (tem(st.manual)) { T.push(apx('MANUAL DE OPERAÇÃO'), ...textoCorrido(st.manual)); }

    // ---------------------------------------------------- listas e sumário
    const linhaLista = (it) => new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: W, leader: 'dot' }], spacing: { line: L15 },
      children: [tr(it.legenda), new TextRun({ children: [new Tab()], font: FONT, size: 24 }), new PageReference(it.bm)] });
    const figs = ilustracoes.filter(i => i.tipo === 'Figura');
    const quads = ilustracoes.filter(i => i.tipo !== 'Figura').sort((a, b) => (a.tipo === b.tipo ? a.n - b.n : (a.tipo === 'Quadro' ? -1 : 1)));
    if (figs.length) pre.push(TIT('LISTA DE FIGURAS'), ...figs.map(linhaLista));
    if (quads.length) pre.push(TIT('LISTA DE QUADROS E TABELAS'), ...quads.map(linhaLista));
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
      title: titulo || 'TCC – Técnico em Mecânica',
      description: 'Gerado pelo Assistente do TCC – Fazenda de Projetos',
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
