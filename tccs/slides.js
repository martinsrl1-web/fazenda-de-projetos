/* Apresentação do TCC – Assistentes do TCC (CEEP FRG)
   1) Gera os slides sugeridos (PowerPoint) no padrão "Slide Padrão CEEP", a partir do que a equipe escreveu.
   2) Lê os slides que a equipe fez (.pptx) e sugere melhorias.
   Tudo no navegador, sem IA. Apresentação de 10 a 15 minutos. */
(function () {
  const BASE = new URL('.', document.currentScript.src).href;           // .../tccs/
  const VERMELHO = 'C00000', AZUL = '002060', NAVY = '312B50', FONTE = 'Calibri';
  const W = 13.333, H = 7.5;
  let o = null;

  /* ================= utilidades ================= */
  const L = s => String(s == null ? '' : s).trim();
  const tem = s => L(s) !== '';
  const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const palavras = s => (L(s).match(/[A-Za-zÀ-ÿ0-9]+/g) || []).length;
  const num = v => { const x = parseFloat(String(v || '').replace(/[^\d,.\-]/g, '').replace(/\.(?=\d{3}(\D|$))/g, '').replace(',', '.')); return isNaN(x) ? 0 : x; };
  const temN = v => /\d/.test(String(v || ''));
  const real = v => v.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const fmt = (v, d = 1) => v.toLocaleString('pt-BR', { maximumFractionDigits: d });
  const linhas = (t, k) => (t || []).filter(r => r && r.slice(0, k || r.length).some(tem));
  const semCit = s => s.replace(/\s*\((?:[A-ZÀ-Ý][^()]*?|[^()]*?et al\.?),?\s*\d{4}[a-z]?(?:,\s*p\.\s*[\d-]+)?\)/g, '').replace(/\s+([.,;:])/g, '$1');
  function frases(t) {
    return L(t).replace(/\s+/g, ' ').split(/(?<=[.!?])\s+(?=[A-ZÀ-Ý0-9“"(])/).map(s => semCit(s).trim()).filter(s => palavras(s) >= 3);
  }
  // encurta uma frase para virar tópico
  function curto(s, max = 20) {
    s = L(s).replace(/[.;:]+$/, '');
    const ws = s.split(/\s+/);
    if (ws.length <= max) return s;
    // tenta cortar numa vírgula entre a 8ª e a última palavra permitida
    let acc = '', melhor = '';
    for (let i = 0; i < max; i++) { acc += (i ? ' ' : '') + ws[i]; if (i >= 7 && /[,;]$/.test(ws[i])) melhor = acc.replace(/[,;]$/, ''); }
    return melhor || acc + '…';
  }
  const topicos = (t, n = 4, max = 20) => frases(t).slice(0, n).map(f => curto(f, max));
  const primeira = (t, max = 24) => { const f = frases(t)[0]; return f ? curto(f, max) : ''; };

  async function imagem(id) {
    if (!id) return null;
    const v = await o.lerImagem(id);
    if (!v) return null;
    const mime = v.type === 'png' ? 'image/png' : 'image/jpeg';
    return { data: mime + ';base64,' + v.b64, w: v.w || 4, h: v.h || 3 };
  }
  const figs = (...fs) => fs.filter(f => f && f.img);

  async function arquivoB64(url) {
    const r = await fetch(url); const b = await r.blob();
    return new Promise(res => { const fr = new FileReader(); fr.onload = () => res(String(fr.result).replace(/^data:/, '')); fr.readAsDataURL(b); });
  }
  function carregarPptx() {
    if (window.PptxGenJS) return Promise.resolve();
    return new Promise((res, rej) => {
      const s = document.createElement('script'); s.src = BASE + '../vendor/pptxgen.bundle.js';
      s.onload = res; s.onerror = () => rej(new Error('Não consegui carregar o gerador de PowerPoint. Verifiquem a internet.'));
      document.head.appendChild(s);
    });
  }

  /* ================= roteiros (o que vai em cada slide) =================
     Cada slide: { tipo, titulo, ... , notas, fonte (de qual parte do TCC veio) } */
  function capa(S, rotOrg) {
    return { tipo: 'capa', titulo: L(S.titulo) || 'Título do TCC', sub: L(S.subtitulo), integrantes: S.integrantes.filter(tem),
      orientador: L(S.orientador), curso: L(S.curso), ano: L(S.ano), extra: rotOrg,
      notas: 'Apresentem-se: cada integrante diz o nome. Digam o título e, em uma frase, o que o grupo fez.' };
  }
  function problemaSlide(S) {
    if (!tem(S.problema) && !tem(S.justificativa)) return null;
    return { tipo: 'destaque', titulo: 'Problema', destaque: L(S.problema), itens: topicos(S.justificativa, 2, 22), rotItens: 'Por que este trabalho é importante',
      notas: [S.problema, S.justificativa].filter(tem).join('\n\n'), parte: 'Problema e justificativa' };
  }
  function objetivosSlide(S) {
    const esp = (S.objEsp || []).filter(tem);
    if (!tem(S.objGeral) && !esp.length) return null;
    return { tipo: 'objetivos', titulo: 'Objetivos', geral: curto(L(S.objGeral), 30), itens: esp.slice(0, 6).map(e => curto(e, 16)),
      notas: ['Objetivo geral: ' + L(S.objGeral), ...esp.map(e => '• ' + e)].join('\n'), parte: 'Objetivos' };
  }
  function conclusaoSlide(S) {
    if (!tem(S.conclusao)) return null;
    return { tipo: 'topicos', titulo: 'Conclusão', itens: topicos(S.conclusao, 4, 20), notas: L(S.conclusao), parte: 'Conclusão' };
  }
  function discussaoSlide(S) {
    const esp = (S.objEsp || []).filter(tem), d = S.discussao || [];
    const rows = esp.map((e, i) => [curto(e, 10), curto(L((d[i] || {}).resultado), 12), L((d[i] || {}).situacao)]).filter(r => tem(r[1]) || tem(r[2]));
    if (!rows.length) return null;
    return { tipo: 'tabela', titulo: 'Objetivos alcançados', cab: ['Objetivo', 'Resultado', 'Situação'], linhas: rows, larg: [4.4, 5.2, 2.4],
      notas: L(S.discussaoTexto), parte: 'Discussão' };
  }
  function finalSlide(S) {
    return { tipo: 'final', titulo: 'Obrigado!', integrantes: S.integrantes.filter(tem), notas: 'Agradeçam e abram para as perguntas da banca.' };
  }

  async function roteiroMecanica(S) {
    const sl = [], fora = [];
    const add = (s, nome) => { if (s) sl.push(s); else fora.push(nome); };
    sl.push(capa(S, L(S.equipamento)));
    add(tem(S.contexto) ? { tipo: 'topicos', titulo: 'Introdução', itens: topicos(S.contexto, 4, 20), notas: L(S.contexto), parte: 'Contextualização' } : null, 'Introdução');
    add(problemaSlide(S), 'Problema');
    add(objetivosSlide(S), 'Objetivos');
    const et = linhas(S.etapas, 3).filter(r => tem(r[1]) || tem(r[2]));
    add(et.length ? { tipo: 'etapas', titulo: 'Como o trabalho foi feito', itens: et.map(r => [L(r[0]), curto(L(r[1]), 12)]),
      notas: [L(S.metodo), ...et.map(r => `${r[0]}: ${r[1]} ${r[2] ? '(' + r[2] + ')' : ''}`)].join('\n'), parte: 'Metodologia' } : null, 'Metodologia (etapas)');
    const prods = linhas(S.produtos, 1);
    add(prods.length || tem(S.diferencial) ? { tipo: 'tabela', titulo: 'Produtos semelhantes', cab: ['Modelo', 'Capacidade', 'Preço (R$)'],
      linhas: prods.slice(0, 4).map(r => [L(r[0]), L(r[2]), L(r[3])]), rodape: tem(S.diferencial) ? 'Diferencial do nosso projeto: ' + primeira(S.diferencial, 22) : '',
      larg: [5.6, 3.6, 2.8], notas: [L(S.diferencial), L(S.evolucao)].filter(tem).join('\n\n'), parte: 'Produtos semelhantes' } : null, 'Produtos semelhantes');
    const fF = figs(S.figFunc)[0];
    add(tem(S.funcionamento) || fF ? { tipo: 'topicos', titulo: 'Como funciona', itens: topicos(S.funcionamento, 4, 18), img: fF && await imagem(fF.img), legenda: fF && L(fF.legenda),
      notas: L(S.funcionamento), parte: 'Princípio de funcionamento' } : null, 'Como funciona');
    const fC = figs(S.figCad, ...(S.desenhos || []))[0];
    add(fC ? { tipo: 'imagem', titulo: 'Desenho e modelagem', img: await imagem(fC.img), legenda: L(fC.legenda), itens: topicos(S.modelagem, 2, 16),
      notas: L(S.modelagem), parte: 'Desenho e modelagem' } : null, 'Desenho / CAD (imagem)');
    const mats = linhas(S.materiais, 1);
    add(mats.length ? { tipo: 'tabela', titulo: 'Principais materiais', cab: ['Item', 'Qtd.', 'Material e medidas'], linhas: mats.slice(0, 6).map(r => [L(r[0]), L(r[1]), curto(L(r[2]), 10)]),
      larg: [4, 1.4, 6.6], notas: mats.map(r => r.filter(tem).join(' – ')).join('\n'), parte: 'Lista de materiais' } : null, 'Materiais');
    const cal = (S.calculos || []).filter(c => c && (tem(c.formula) || tem(c.substituicao)));
    cal.slice(0, 2).forEach(c => sl.push({ tipo: 'calculo', titulo: 'Dimensionamento' + (tem(c.titulo) ? ': ' + curto(c.titulo, 6) : ''),
      dados: primeira(c.dados, 24), formula: L(c.formula), subst: L(c.substituicao), conclusao: primeira(c.significado, 22),
      notas: [c.dados, c.formula, c.variaveis, c.substituicao, c.significado].filter(tem).join('\n'), parte: 'Dimensionamento' }));
    if (!cal.length) fora.push('Dimensionamento (cálculos)');
    const fotos = (S.construcao || []).filter(c => c && c.foto && c.foto.img);
    if (fotos.length) {
      const ims = []; for (const c of fotos.slice(0, 3)) ims.push({ ...(await imagem(c.foto.img)), legenda: curto(L(c.foto.legenda) || primeira(c.texto, 8), 8) });
      sl.push({ tipo: 'galeria', titulo: 'Construção', imgs: ims, notas: (S.construcao || []).map(c => L(c.texto)).filter(tem).join('\n\n'), parte: 'Etapas da construção' });
    } else if ((S.construcao || []).some(c => tem(c.texto))) {
      sl.push({ tipo: 'topicos', titulo: 'Construção', itens: (S.construcao || []).filter(c => tem(c.texto)).slice(0, 4).map(c => primeira(c.texto, 16)),
        notas: (S.construcao || []).map(c => L(c.texto)).filter(tem).join('\n\n'), parte: 'Etapas da construção' });
    } else fora.push('Construção (fotos)');
    const cs = linhas(S.custos, 4);
    if (cs.length) {
      let tot = 0; const rows = cs.map(r => { const t = num(r[1] || '1') * num(r[2]); tot += t; return [L(r[0]), L(r[1]), real(t)]; });
      const mostrar = rows.length > 6 ? rows.slice(0, 5).concat([['Outros itens', '', real(rows.slice(5).reduce((a, r) => a + num(r[2]), 0))]]) : rows;
      sl.push({ tipo: 'tabela', titulo: 'Custos', cab: ['Item', 'Qtd.', 'Total (R$)'], linhas: mostrar.concat([['TOTAL', '', 'R$ ' + real(tot)]]), negritoUltima: true,
        larg: [7, 1.6, 3.4], direita: [1, 2], rodape: primeira(S.custosTexto, 22), notas: L(S.custosTexto), parte: 'Custos' });
    } else fora.push('Custos');
    const reps = linhas(S.testes, 3);
    if (reps.length || tem(S.testeAnalise)) {
      const m1 = L(S.medida1) || 'Medida', m2 = L(S.medida2);
      const media = k => { const v = reps.map(r => r[k]).filter(temN).map(num); return v.length ? fmt(v.reduce((a, b) => a + b, 0) / v.length, 2) : ''; };
      const rows = reps.slice(0, 5).map((r, i) => m2 ? [String(i + 1), L(r[0]), L(r[1])] : [String(i + 1), L(r[0])]);
      if (reps.length > 1) rows.push(m2 ? ['Média', media(0), media(1)] : ['Média', media(0)]);
      const fT = figs(S.figTeste)[0];
      sl.push({ tipo: reps.length ? 'tabela' : 'topicos', titulo: 'Testes e resultados', cab: m2 ? ['Teste', m1, m2] : ['Teste', m1], linhas: rows, negritoUltima: reps.length > 1,
        itens: topicos(S.testeAnalise, 3, 18), rodape: primeira(S.testeAnalise, 22), img: fT && await imagem(fT.img),
        notas: [S.testeProc, S.testeAnalise].filter(tem).join('\n\n'), parte: 'Testes' });
    } else fora.push('Testes e resultados');
    add(discussaoSlide(S), 'Objetivos alcançados (discussão)');
    add(conclusaoSlide(S), 'Conclusão');
    sl.push(finalSlide(S));
    return caber15({ slides: sl, fora }, ['Principais materiais', 'Produtos semelhantes', 'Objetivos alcançados', 'Desenho e modelagem', 'Dimensionamento']);
  }

  // até 15 slides (10 a 15 minutos): corta primeiro os slides menos importantes
  function caber15(r, ordem) {
    for (const nome of ordem) {
      if (r.slides.length <= 15) break;
      let i = -1; r.slides.forEach((x, k) => { if (k > 0 && x.tipo !== 'final' && L(x.titulo).startsWith(nome)) i = k; });
      if (i > 0) { (r.cortados = r.cortados || []).push(r.slides[i].titulo); r.slides.splice(i, 1); }
    }
    return r;
  }

  async function roteiroPCP(S) {
    const sl = [], fora = [];
    const add = (s, nome) => { if (s) sl.push(s); else fora.push(nome); };
    sl.push(capa(S, L(S.organizacao)));
    const org = [S.historico, S.processo].filter(tem).join(' ');
    add(tem(org) ? { tipo: 'topicos', titulo: S.ficticia ? 'A organização estudada (fictícia)' : 'A organização estudada', itens: topicos(org, 4, 20),
      notas: [S.historico, S.produtosTexto, S.capacidade, S.processo].filter(tem).join('\n\n'), parte: 'Caracterização da organização' } : null, 'A organização');
    add(problemaSlide(S), 'Problema');
    add(objetivosSlide(S), 'Objetivos');
    const et = linhas(S.etapas, 3).filter(r => tem(r[1]) || tem(r[2]));
    add(et.length || tem(S.coleta) ? { tipo: 'etapas', titulo: 'Metodologia', itens: et.map(r => [L(r[0]), curto(L(r[1]), 12)]),
      notas: [S.tipoPesquisa, S.objeto, S.coleta, S.analise].filter(tem).join('\n\n'), parte: 'Metodologia' } : null, 'Metodologia (etapas)');
    const fF = figs(S.figFluxo)[0];
    add(tem(S.procAtual) || fF ? { tipo: fF ? 'imagem' : 'topicos', titulo: 'Como o processo funciona hoje', img: fF && await imagem(fF.img), legenda: fF && L(fF.legenda),
      itens: topicos(S.procAtual, fF ? 2 : 4, 18), notas: [S.procAtual, S.procAnalise].filter(tem).join('\n\n'), parte: 'Processo atual' } : null, 'Processo atual');
    const dados = linhas(S.dados, 2).filter(r => tem(r[0]) && temN(r[1])).sort((a, b) => num(b[1]) - num(a[1]));
    if (dados.length) {
      const tot = dados.reduce((a, r) => a + num(r[1]), 0); let ac = 0;
      const top = dados.slice(0, 7);
      sl.push({ tipo: 'pareto', titulo: 'Onde está o problema (Pareto)', rotulos: top.map(r => curto(L(r[0]), 4)),
        valores: top.map(r => num(r[1])), acumulado: top.map(r => { ac += tot ? num(r[1]) / tot * 100 : 0; return Math.min(100, Math.round(ac)); }),
        unidade: L(S.dadosUnid) || 'Ocorrências', rodape: primeira(S.dadosAnalise, 22), notas: [S.dadosTexto, S.dadosAnalise].filter(tem).join('\n\n'), parte: 'Levantamento de dados' });
    } else fora.push('Dados do problema (Pareto)');
    const fI = figs(S.figIshikawa)[0];
    const gut = linhas(S.gut, 4).map(r => ({ r, p: temN(r[1]) && temN(r[2]) && temN(r[3]) ? num(r[1]) * num(r[2]) * num(r[3]) : 0 })).sort((a, b) => b.p - a.p);
    if (fI) sl.push({ tipo: 'imagem', titulo: 'Causas do problema', img: await imagem(fI.img), legenda: L(fI.legenda), itens: [], notas: [S.causasTexto, S.causasAnalise].filter(tem).join('\n\n'), parte: 'Análise das causas' });
    if (gut.length) sl.push({ tipo: 'tabela', titulo: 'Causas prioritárias (GUT)', cab: ['Causa', 'G × U × T', 'Prioridade'],
      linhas: gut.slice(0, 5).map((g, i) => [curto(L(g.r[0]), 10), g.p ? String(g.p) : '', `${i + 1}ª`]), larg: [7.6, 2.2, 2.2], direita: [1],
      rodape: primeira(S.sintese, 22), notas: L(S.sintese), parte: 'Matriz GUT' });
    if (!fI && !gut.length) fora.push('Causas (Ishikawa ou GUT)');
    const acoes = linhas(S.acoes, 3);
    add(acoes.length || tem(S.planoTexto) ? { tipo: acoes.length ? 'tabela' : 'topicos', titulo: 'Proposta de melhoria', cab: ['Causa', 'Ação proposta', 'Ferramenta'],
      linhas: acoes.slice(0, 5).map(r => [curto(L(r[0]), 7), curto(L(r[1]), 12), curto(L(r[2]), 5)]), larg: [3.4, 6, 2.6], itens: topicos(S.planoTexto, 4, 18),
      rodape: temN(S.investimento) ? 'Investimento estimado: R$ ' + real(num(S.investimento)) : '', notas: L(S.planoTexto), parte: 'Proposta' } : null, 'Proposta de melhoria');
    const fs = (S.ferramentas || []).filter(f => f && ([f.titulo, f.formula, f.substituicao, f.significado].some(tem) || (f.fig && f.fig.img)));
    for (const f of fs.slice(0, 2)) {
      const fg = f.fig && f.fig.img ? await imagem(f.fig.img) : null;
      sl.push({ tipo: 'calculo', titulo: tem(f.titulo) ? curto(f.titulo, 7) : 'Ferramenta aplicada', dados: primeira(f.dados, 24), formula: L(f.formula), subst: L(f.substituicao),
        conclusao: primeira(f.significado, 22), img: fg, notas: [f.dados, f.formula, f.variaveis, f.substituicao, f.significado].filter(tem).join('\n'), parte: 'Aplicação das ferramentas' });
    }
    if (!fs.length) fora.push('Ferramentas aplicadas');
    const res = linhas(S.resultados, 3).filter(r => tem(r[0]));
    if (res.length) {
      const varia = (a, d) => { if (!temN(a) || !temN(d)) return '–'; const x = num(a), y = num(d); if (/%/.test(a + d)) { const pp = y - x; return (pp > 0 ? '+' : pp < 0 ? '−' : '') + fmt(Math.abs(pp)) + ' p.p.'; } if (!x) return '–'; const v = (y - x) / x * 100; return (v > 0 ? '+' : v < 0 ? '−' : '') + fmt(Math.abs(v), 0) + '%'; };
      sl.push({ tipo: 'tabela', titulo: 'Resultados: antes e depois', cab: ['Indicador', 'Antes', 'Depois', 'Variação', 'Meta'],
        linhas: res.slice(0, 6).map(r => [curto(L(r[0]), 7), L(r[1]) || '–', L(r[2]), varia(r[1], r[2]), L(r[3])]), larg: [4.4, 1.9, 1.9, 1.9, 1.9], direita: [1, 2, 3, 4],
        rodape: primeira(S.resultadosTexto, 22), notas: L(S.resultadosTexto), parte: 'Resultados' });
    } else fora.push('Resultados (antes e depois)');
    const inv = num(S.investimento), eco = num(S.economiaMensal);
    if (inv || eco) {
      const nums = [];
      if (inv) nums.push(['R$ ' + real(inv), 'investimento']);
      if (eco) nums.push(['R$ ' + real(eco), 'economia por mês']);
      if (inv && eco) { const pb = inv / eco; nums.push([fmt(pb, 1) + (pb < 2 ? ' mês' : ' meses'), 'tempo de retorno (payback)']); }
      sl.push({ tipo: 'numeros', titulo: 'Análise financeira', numeros: nums, rodape: primeira(S.financeiroTexto || S.economiaTexto, 22),
        notas: [S.economiaTexto, S.financeiroTexto].filter(tem).join('\n\n'), parte: 'Análise financeira' });
    } else fora.push('Análise financeira');
    add(discussaoSlide(S), 'Objetivos alcançados (discussão)');
    add(conclusaoSlide(S), 'Conclusão');
    sl.push(finalSlide(S));
    return caber15({ slides: sl, fora }, ['Objetivos alcançados', 'Causas prioritárias', 'Metodologia', 'A organização']);
  }

  /* ================= desenho dos slides (PptxGenJS) ================= */
  // calcula o maior tamanho de letra que cabe na caixa (sem depender do "encolher" do PowerPoint)
  function tamanho(conteudo, w, h, max, min, bullet) {
    const pars = []; let atual = '';
    if (typeof conteudo === 'string') conteudo.split('\n').forEach(x => pars.push(x));
    else { conteudo.forEach(r => { atual += r.text; if (r.options && r.options.breakLine) { pars.push(atual); atual = ''; } }); if (atual) pars.push(atual); }
    for (let t = max; t > min; t--) {
      const porLinha = Math.max(8, ((w - 0.2 - (bullet ? 0.4 : 0)) * 72) / (t * 0.5));
      let alt = 0; pars.forEach(x => { alt += Math.max(1, Math.ceil(x.length / porLinha)) * t * 1.18 / 72 + (bullet ? 10 / 72 : 0); });
      if (alt <= h - 0.08) return t;
    }
    return min;
  }
  function ajustarTexto(sl) {
    const orig = sl.addText.bind(sl);
    sl.addText = (c, op) => {
      if (op && op.fit === 'shrink') { op = Object.assign({}, op); op.fontSize = tamanho(c, op.w, op.h, op.fontSize, op.min || Math.min(op.fontSize, 18), op.bul); delete op.fit; delete op.min; delete op.bul; }
      return orig(c, op);
    };
    return sl;
  }
  function caixaImagem(sl, im, x, y, w, h, legenda) {
    if (!im) return;
    const hl = legenda ? 0.45 : 0;
    const k = Math.min(w / im.w, (h - hl) / im.h);
    const iw = im.w * k, ih = im.h * k;
    sl.addImage({ data: im.data, x: x + (w - iw) / 2, y: y + (h - hl - ih) / 2, w: iw, h: ih });
    if (legenda) sl.addText(legenda, { x, y: y + h - hl, w, h: hl, fontFace: FONTE, fontSize: 16, color: '595959', italic: true, align: 'center', valign: 'top', fit: 'shrink', min: 14 });
  }
  function titulo(sl, t) {
    sl.addText(t, { x: 0.42, y: 0.62, w: 10.6, h: 0.95, fontFace: FONTE, fontSize: 32, bold: true, color: VERMELHO, valign: 'middle', fit: 'shrink', min: 24, margin: 0 });
  }
  function bullets(sl, itens, x, y, w, h, tam = 24) {
    if (!itens || !itens.length) return;
    sl.addText(itens.map(t => ({ text: t, options: { bullet: { code: '25A0' }, paraSpaceAfter: 10, breakLine: true } })),
      { x, y, w, h, fontFace: FONTE, fontSize: tam, color: AZUL, valign: 'top', fit: 'shrink', bul: true, margin: 2 });
  }
  function rodape(sl, t) {
    if (!t) return;
    sl.addText(t, { x: 0.7, y: 5.95, w: 11.9, h: 0.65, fontFace: FONTE, fontSize: 18, color: AZUL, italic: true, valign: 'middle', fit: 'shrink', min: 16,
      fill: { color: 'F2F1F7' }, line: { color: 'F2F1F7' }, margin: 8 });
  }
  function tabela(sl, s, y = 1.85, largura = 11.9) {
    const larg = s.larg || s.cab.map(() => 12 / s.cab.length);
    const tot = larg.reduce((a, b) => a + b, 0), k = largura / tot;
    const n = s.linhas.length;
    const tam = n > 6 ? 14 : n > 4 ? 16 : 18;
    const rows = [s.cab.map(c => ({ text: c, options: { bold: true, color: 'FFFFFF', fill: { color: NAVY }, fontSize: tam } }))];
    s.linhas.forEach((r, i) => rows.push(r.map((c, j) => ({ text: String(c == null ? '' : c), options: {
      fontSize: tam, color: AZUL, bold: !!(s.negritoUltima && i === n - 1), fill: { color: i % 2 ? 'F2F1F7' : 'FFFFFF' },
      align: (s.direita || []).includes(j) ? 'right' : 'left' } }))));
    sl.addTable(rows, { x: 0.7, y, w: largura, colW: larg.map(v => v * k), fontFace: FONTE, border: { type: 'solid', pt: 0.75, color: 'BFBFD0' },
      margin: [4, 8, 4, 8], valign: 'middle', autoPage: false });
  }

  async function montar(roteiro, nomeArquivo) {
    await carregarPptx();
    const [capaB, fundoB] = await Promise.all([arquivoB64(BASE + 'slides/capa.jpg'), arquivoB64(BASE + 'slides/fundo.jpg')]);
    const p = new window.PptxGenJS();
    p.layout = 'LAYOUT_WIDE'; p.author = 'Assistente do TCC – CEEP FRG'; p.title = roteiro.slides[0].titulo;
    p.theme = { headFontFace: FONTE, bodyFontFace: FONTE };
    for (const s of roteiro.slides) {
      const sl = ajustarTexto(p.addSlide());
      sl.background = { data: s.tipo === 'capa' ? capaB : fundoB };
      if (s.notas) sl.addNotes('O QUE FALAR NESTE SLIDE (texto do TCC):\n\n' + s.notas);
      if (s.tipo === 'capa') {
        sl.addText(s.titulo, { x: 6.9, y: 0.75, w: 6.0, h: 2.6, fontFace: FONTE, fontSize: 34, bold: true, color: 'FFFFFF', align: 'center', valign: 'bottom', fit: 'shrink', min: 22 });
        if (s.sub) sl.addText(s.sub, { x: 6.9, y: 3.4, w: 6.0, h: 0.7, fontFace: FONTE, fontSize: 18, color: 'FFFFFF', align: 'center', valign: 'top', fit: 'shrink', min: 12 });
        sl.addText(s.integrantes.map(n => ({ text: n, options: { breakLine: true } })).concat(s.integrantes.length ? [] : [{ text: 'Nomes dos integrantes' }]),
          { x: 6.9, y: 4.2, w: 6.0, h: 1.5, fontFace: FONTE, fontSize: 16, color: 'FFFFFF', align: 'center', valign: 'top', fit: 'shrink', min: 12 });
        sl.addText([{ text: 'Orientador: ' + (s.orientador || ''), options: { breakLine: true } }, { text: `${s.curso} · ${s.ano}` }],
          { x: 6.9, y: 5.8, w: 6.0, h: 0.9, fontFace: FONTE, fontSize: 13, color: 'D9D7E8', align: 'center', valign: 'top' });
        continue;
      }
      if (s.tipo === 'final') {
        sl.addText('Obrigado!', { x: 0.7, y: 1.7, w: 11.9, h: 1.3, fontFace: FONTE, fontSize: 54, bold: true, color: VERMELHO, align: 'center' });
        sl.addText('Perguntas?', { x: 0.7, y: 2.9, w: 11.9, h: 0.8, fontFace: FONTE, fontSize: 30, color: AZUL, align: 'center' });
        if (s.integrantes.length) sl.addText(s.integrantes.join(' · '), { x: 0.7, y: 3.8, w: 11.9, h: 0.7, fontFace: FONTE, fontSize: 16, color: AZUL, align: 'center', fit: 'shrink' });
        sl.addText([
          { text: 'CEEP FAZENDA RIO GRANDE EROTIDES ÂNGELO NICHELE', options: { bold: true, breakLine: true } },
          { text: 'Rua Rio Pinhão esquina Rua Rio Xingu, 1172 – Iguaçu', options: { breakLine: true } },
          { text: 'Fazenda Rio Grande / PR · Fone: (41) 3627-6955' }],
          { x: 0.7, y: 4.9, w: 11.9, h: 1.2, fontFace: FONTE, fontSize: 14, color: AZUL, align: 'center' });
        continue;
      }
      titulo(sl, s.titulo);
      if (s.tipo === 'topicos') {
        if (s.img) { bullets(sl, s.itens, 0.7, 1.85, 6.2, 4.6, 22); caixaImagem(sl, s.img, 7.2, 1.75, 5.5, 4.8, s.legenda); }
        else bullets(sl, s.itens, 0.7, 1.85, 11.9, 4.7, 24);
      } else if (s.tipo === 'destaque') {
        if (s.destaque) {
          sl.addShape(p.ShapeType.rect, { x: 0.7, y: 1.85, w: 0.12, h: 1.6, fill: { color: VERMELHO }, line: { color: VERMELHO } });
          sl.addText(s.destaque, { x: 1.0, y: 1.8, w: 11.6, h: 1.7, fontFace: FONTE, fontSize: 26, bold: true, italic: true, color: AZUL, valign: 'middle', fit: 'shrink' });
        }
        if (s.itens.length) {
          sl.addText(s.rotItens, { x: 0.7, y: 3.75, w: 11.9, h: 0.5, fontFace: FONTE, fontSize: 18, bold: true, color: VERMELHO });
          bullets(sl, s.itens, 0.7, 4.25, 11.9, 2.3, 22);
        }
      } else if (s.tipo === 'objetivos') {
        if (s.geral) {
          sl.addText([{ text: 'Geral: ', options: { bold: true, color: VERMELHO } }, { text: s.geral }],
            { x: 0.7, y: 1.8, w: 11.9, h: 1.1, fontFace: FONTE, fontSize: 22, color: AZUL, valign: 'middle', fit: 'shrink', fill: { color: 'F2F1F7' }, margin: 8 });
        }
        if (s.itens.length) {
          sl.addText('Específicos', { x: 0.7, y: 3.05, w: 11.9, h: 0.45, fontFace: FONTE, fontSize: 18, bold: true, color: VERMELHO });
          bullets(sl, s.itens, 0.7, 3.5, 11.9, 3.0, 20);
        }
      } else if (s.tipo === 'etapas') {
        const n = Math.min(s.itens.length, 8), cols = n > 4 ? 2 : 1, porCol = Math.ceil(n / cols);
        s.itens.slice(0, 8).forEach((it, i) => {
          const c = Math.floor(i / porCol), r = i % porCol;
          const x = 0.7 + c * 6.05, y = 1.85 + r * (4.6 / porCol), h = Math.min(1.0, 4.6 / porCol - 0.12);
          sl.addShape(p.ShapeType.ellipse, { x, y: y + (h - 0.55) / 2, w: 0.55, h: 0.55, fill: { color: VERMELHO }, line: { color: VERMELHO } });
          sl.addText(String(i + 1), { x, y: y + (h - 0.55) / 2, w: 0.55, h: 0.55, fontFace: FONTE, fontSize: 16, bold: true, color: 'FFFFFF', align: 'center', valign: 'middle', margin: 0 });
          sl.addText([{ text: it[0], options: { bold: true } }].concat(it[1] ? [{ text: ' – ' + it[1] }] : []),
            { x: x + 0.7, y, w: 5.2, h, fontFace: FONTE, fontSize: 18, color: AZUL, valign: 'middle', fit: 'shrink', min: 16, margin: 2 });
        });
      } else if (s.tipo === 'tabela') {
        if (s.img) { tabela(sl, s, 1.85, 6.4); caixaImagem(sl, s.img, 7.4, 1.75, 5.3, s.rodape ? 4.05 : 4.85); }
        else tabela(sl, s, 1.85);
        rodape(sl, s.rodape);
      } else if (s.tipo === 'imagem') {
        if (s.itens && s.itens.length) { caixaImagem(sl, s.img, 0.7, 1.75, 7.4, 4.85, s.legenda); bullets(sl, s.itens, 8.4, 1.9, 4.3, 4.5, 20); }
        else caixaImagem(sl, s.img, 0.7, 1.75, 11.9, 4.85, s.legenda);
      } else if (s.tipo === 'galeria') {
        const n = s.imgs.length, w = (11.9 - (n - 1) * 0.3) / n;
        s.imgs.forEach((im, i) => caixaImagem(sl, im, 0.7 + i * (w + 0.3), 1.8, w, 4.8, im.legenda));
      } else if (s.tipo === 'calculo') {
        const largura = s.img ? 6.6 : 11.9;
        let y = 1.8;
        if (s.dados) { sl.addText(s.dados, { x: 0.7, y, w: largura, h: 0.95, fontFace: FONTE, fontSize: 18, color: AZUL, valign: 'top', fit: 'shrink' }); y += 1.05; }
        const eq = [s.formula, s.subst].filter(tem);
        if (eq.length) {
          sl.addText(eq.map(t => ({ text: t, options: { breakLine: true, paraSpaceAfter: 6 } })),
            { x: 0.7, y, w: largura, h: 1.7, fontFace: 'Cambria Math', fontSize: 22, color: AZUL, align: 'center', valign: 'middle', fill: { color: 'F2F1F7' }, fit: 'shrink', margin: 8 });
          y += 1.85;
        }
        if (s.conclusao) sl.addText([{ text: 'Resultado: ', options: { bold: true, color: VERMELHO } }, { text: s.conclusao }],
          { x: 0.7, y, w: largura, h: 6.55 - y, fontFace: FONTE, fontSize: 19, color: AZUL, valign: 'top', fit: 'shrink' });
        if (s.img) caixaImagem(sl, s.img, 7.6, 1.75, 5.1, 4.85);
      } else if (s.tipo === 'pareto') {
        sl.addChart(p.ChartType.bar, [{ name: s.unidade, labels: s.rotulos, values: s.valores }],
          { x: 0.7, y: 1.75, w: 11.9, h: s.rodape ? 4.1 : 4.85, barDir: 'col', chartColors: [NAVY], showValue: true, dataLabelFontSize: 12, dataLabelColor: AZUL,
            catAxisLabelFontSize: 12, catAxisLabelColor: AZUL, valAxisLabelFontSize: 11, valAxisLabelColor: '595959', valGridLine: { color: 'E5E5EE', size: 0.5 },
            showLegend: false, catAxisLabelFontFace: FONTE });
        rodape(sl, s.rodape || ('Acumulado: ' + s.rotulos.map((r, i) => `${r} ${s.acumulado[i]}%`).slice(0, 3).join(' · ')));
      } else if (s.tipo === 'numeros') {
        const n = s.numeros.length, w = (11.9 - (n - 1) * 0.4) / n;
        s.numeros.forEach(([v, r], i) => {
          const x = 0.7 + i * (w + 0.4);
          sl.addShape(p.ShapeType.roundRect, { x, y: 2.0, w, h: 2.9, rectRadius: 0.12, fill: { color: 'F2F1F7' }, line: { color: 'D9D7E8' } });
          sl.addText(v, { x, y: 2.3, w, h: 1.4, fontFace: FONTE, fontSize: 38, bold: true, color: i === n - 1 && n === 3 ? VERMELHO : AZUL, align: 'center', valign: 'middle', fit: 'shrink' });
          sl.addText(r, { x, y: 3.7, w, h: 0.8, fontFace: FONTE, fontSize: 18, color: AZUL, align: 'center', valign: 'top' });
        });
        rodape(sl, s.rodape);
      }
    }
    await p.writeFile({ fileName: nomeArquivo });
  }

  /* ================= leitura dos slides da equipe (.pptx) ================= */
  async function lerZip(buf) {
    const dv = new DataView(buf);
    let i = buf.byteLength - 22;
    while (i >= 0 && dv.getUint32(i, true) !== 0x06054b50) i--;
    if (i < 0) throw new Error('O arquivo não parece ser um PowerPoint (.pptx).');
    const n = dv.getUint16(i + 10, true); let p = dv.getUint32(i + 16, true);
    const files = {}; const td = new TextDecoder();
    for (let k = 0; k < n; k++) {
      const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true);
      const nlen = dv.getUint16(p + 28, true), xlen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true), loff = dv.getUint32(p + 42, true);
      files[td.decode(new Uint8Array(buf, p + 46, nlen))] = { method, csize, loff };
      p += 46 + nlen + xlen + clen;
    }
    return async name => {
      const f = files[name]; if (!f) return null;
      const ln = dv.getUint16(f.loff + 26, true), lx = dv.getUint16(f.loff + 28, true);
      const data = new Uint8Array(buf, f.loff + 30 + ln + lx, f.csize);
      let b = data.slice();
      if (f.method !== 0) b = new Uint8Array(await new Response(new Blob([data]).stream().pipeThrough(new DecompressionStream('deflate-raw'))).arrayBuffer());
      return new TextDecoder().decode(b);
    };
  }
  const xml = t => new DOMParser().parseFromString(t, 'application/xml');
  const tags = (el, nome) => Array.from(el.getElementsByTagName('*')).filter(e => e.localName === nome);
  const dir = p => p.slice(0, p.lastIndexOf('/') + 1);
  function resolver(base, alvo) {
    if (alvo.startsWith('/')) return alvo.slice(1);
    const partes = (dir(base) + alvo).split('/'); const r = [];
    partes.forEach(x => { if (x === '..') r.pop(); else if (x !== '.') r.push(x); });
    return r.join('/');
  }
  async function rels(get, arq) {
    const t = await get(dir(arq) + '_rels/' + arq.slice(arq.lastIndexOf('/') + 1) + '.rels');
    const m = {}; if (!t) return m;
    tags(xml(t), 'Relationship').forEach(r => { m[r.getAttribute('Id')] = { alvo: resolver(arq, r.getAttribute('Target')), tipo: r.getAttribute('Type') || '' }; });
    return m;
  }

  async function lerPptx(buf) {
    const get = await lerZip(buf);
    const pres = await get('ppt/presentation.xml');
    if (!pres) throw new Error('O arquivo não parece ser um PowerPoint (.pptx).');
    const pd = xml(pres);
    const sz = tags(pd, 'sldSz')[0];
    const cx = sz ? +sz.getAttribute('cx') : 12192000, cy = sz ? +sz.getAttribute('cy') : 6858000;
    const pr = await rels(get, 'ppt/presentation.xml');
    const ids = tags(pd, 'sldId').map(e => e.getAttribute('r:id') || e.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id'));
    const slides = [];
    for (const id of ids) {
      const arq = pr[id] && pr[id].alvo; if (!arq) continue;
      const d = xml(await get(arq) || '<x/>');
      const sr = await rels(get, arq);
      const s = { titulo: '', textos: [], tabTextos: [], miudas: 0, menor: 99, imagens: 0, tabelas: tags(d, 'tbl').length, graficos: 0, notas: '', fundoImagem: false };
      tags(d, 'sp').forEach(sp => {
        const ph = tags(sp, 'ph')[0], tipo = ph ? (ph.getAttribute('type') || 'body') : '';
        const af = tags(sp, 'normAutofit')[0], escala = af && af.getAttribute('fontScale') ? +af.getAttribute('fontScale') / 100000 : 1;
        const pars = tags(sp, 'p').map(pp => tags(pp, 't').map(t => t.textContent).join('')).filter(t => tem(t) && !/^\s*\d{1,2}\s*$/.test(t));
        if (!pars.length) return;
        tags(sp, 'r').forEach(r => { const rp = tags(r, 'rPr')[0], v = rp && rp.getAttribute('sz'); if (!v) return;
          const t = Math.round(+v / 100 * escala), pal = palavras(tags(r, 't').map(e => e.textContent).join(''));
          if (t < 16 && pal) { s.miudas += pal; s.menor = Math.min(s.menor, t); } });
        if ((tipo === 'title' || tipo === 'ctrTitle') && !s.titulo) s.titulo = pars.join(' ');
        else s.textos.push(...pars);
      });
      tags(d, 'tbl').forEach(t => tags(t, 'tc').forEach(c => { const x = tags(c, 't').map(e => e.textContent).join(''); if (tem(x)) s.tabTextos.push(x); }));
      s.linhasTab = tags(d, 'tr').length;
      tags(d, 'pic').forEach(pc => {
        const ext = tags(pc, 'ext')[0];
        const area = ext ? (+ext.getAttribute('cx') * +ext.getAttribute('cy')) / (cx * cy) : 0.2;
        if (area > 0.85) s.fundoImagem = true; else if (area > 0.01) s.imagens++;
      });
      if (tags(d, 'bg').some(b => tags(b, 'blip').length)) s.fundoImagem = true;
      tags(d, 'graphicData').forEach(g => { if (/chart/.test(g.getAttribute('uri') || '')) s.graficos++; });
      const nt = Object.values(sr).find(r => /notesSlide/.test(r.tipo));
      if (nt) { const nd = await get(nt.alvo); if (nd) s.notas = tags(xml(nd), 't').map(t => t.textContent).join(' ').replace(/\d+\s*$/, '').trim(); }
      if (!s.titulo && s.textos.length && palavras(s.textos[0]) <= 8) s.titulo = s.textos.shift();
      slides.push(s);
    }
    return slides;
  }

  /* ================= sugestões ================= */
  const PARTES = {
    mec: [['Problema', /problema|pergunta/i], ['Objetivos', /objetivo/i], ['Metodologia / etapas', /metodolog|etapa|como (o trabalho|foi feito)/i],
      ['Projeto / funcionamento', /funcion|projeto|desenho|cad|modelag|princ[íi]pio/i], ['Dimensionamento (cálculos)', /dimension|c[áa]lculo|for[çc]a|torque|resist/i],
      ['Construção', /constru|fabrica|montag/i], ['Custos', /custo|r\$|or[çc]amento/i], ['Testes e resultados', /teste|resultado|ensaio/i], ['Conclusão', /conclus|considera[çc][õo]es finais/i]],
    pcp: [['A organização', /organiza[çc][ãa]o|empresa|caracteriza/i], ['Problema', /problema|pergunta/i], ['Objetivos', /objetivo/i], ['Metodologia', /metodolog|etapa/i],
      ['Diagnóstico (processo e dados)', /diagn[óo]stico|processo atual|pareto|dados|fluxograma/i], ['Causas', /causa|ishikawa|gut|porqu[êe]s/i],
      ['Proposta / plano de ação', /proposta|plano de a[çc][ãa]o|a[çc][õo]es|5w2h|melhoria/i], ['Resultados / indicadores', /resultado|indicador|antes|depois/i],
      ['Análise financeira', /payback|financeir|economia|investimento|retorno/i], ['Conclusão', /conclus|considera[çc][õo]es finais/i]],
  };
  function textoTCC(S) {
    const out = [];
    const vai = v => { if (typeof v === 'string') { if (v.length > 80) out.push(v); } else if (Array.isArray(v)) v.forEach(vai); else if (v && typeof v === 'object') Object.values(v).forEach(vai); };
    vai(S);
    return out.join(' ').replace(/\s+/g, ' ').toLowerCase();
  }
  function sugerir(slides, S, curso) {
    const g = [], ps = [];
    const n = slides.length;
    const conteudo = slides.filter((s, i) => !(i === 0 || (i === n - 1 && /obrigad|pergunta|d[úu]vida|agradec/i.test([s.titulo, ...s.textos].join(' ')))));
    const minT = Math.round(conteudo.length * 0.8 + 1), maxT = Math.round(conteudo.length * 1.15 + 1.5);
    let resumo;
    if (conteudo.length < 8) resumo = { c: 'bad', t: `${n} slides: provavelmente a apresentação fica com menos de 10 minutos (cerca de ${minT} a ${maxT} min). Faltam partes do TCC ou os slides estão muito resumidos.` };
    else if (conteudo.length > 16) resumo = { c: 'bad', t: `${n} slides: é muito para 15 minutos (cerca de ${minT} a ${maxT} min). Juntem ou cortem slides até ficar com 12 a 15.` };
    else resumo = { c: 'ok', t: `${n} slides: dá para apresentar em 10 a 15 minutos (cerca de ${minT} a ${maxT} min, falando cerca de 1 minuto por slide).` };
    const tudo = slides.map(s => [s.titulo, ...s.textos, ...s.tabTextos].join(' ')).join(' \n ');
    const faltam = (PARTES[curso] || []).filter(([, rx]) => !rx.test(tudo)).map(([nome]) => nome);
    if (faltam.length) g.push(['bad', 'Não encontrei slides sobre: ' + faltam.join(', ') + '. A banca espera ver todas as partes do TCC.']);
    if (!slides.some(s => s.imagens > 0)) g.push(['bad', 'Não há nenhuma foto ou figura. Mostrem ' + (curso === 'mec' ? 'o equipamento, o desenho e a construção' : 'o fluxograma, o Ishikawa e os gráficos') + ': imagem prende mais a atenção do que texto.']);
    if (curso === 'pcp' && !slides.some(s => s.graficos > 0 || s.imagens > 0 && /pareto|gr[áa]fico|resultado/i.test(s.titulo))) g.push(['mid', 'Usem gráficos (Pareto, antes e depois). Número em gráfico é mais fácil de entender do que em tabela.']);
    const semFundo = slides.filter(s => !s.fundoImagem).length;
    if (semFundo > n / 2) g.push(['mid', 'Os slides parecem não usar o modelo do CEEP. Usem o “Slide Padrão CEEP” (ou gerem os slides sugeridos aqui e copiem o conteúdo para eles).']);
    if (!slides.some(s => tem(s.notas))) g.push(['mid', 'Dica: escrevam nas anotações (notas do orador) o que cada um vai falar em cada slide. Ajuda a ensaiar sem decorar.']);
    const integrantes = (S.integrantes || []).filter(tem);
    if (integrantes.length > 1) g.push(['ok', `Dividam: com ${integrantes.length} integrantes, cada um apresenta cerca de ${Math.round(conteudo.length / integrantes.length)} slides (uns ${Math.round(12.5 / integrantes.length)} minutos cada).`]);
    const tcc = textoTCC(S);
    slides.forEach((s, i) => {
      const nome = `Slide ${i + 1}${s.titulo ? ' – ' + curto(s.titulo, 7) : ''}`, r = [];
      const pal = s.textos.reduce((a, t) => a + palavras(t), 0);
      if (i > 0 && !tem(s.titulo)) r.push(['mid', 'Sem título. Todo slide precisa de um título curto.']);
      if (pal > 90) r.push(['bad', `${pal} palavras: texto demais. Deixem só tópicos curtos (até 6 linhas) e falem o resto.`]);
      else if (pal > 70) r.push(['mid', `${pal} palavras: está com bastante texto. Tentem resumir em tópicos.`]);
      if (s.linhasTab > 9) r.push(['mid', `Tabela com ${s.linhasTab - 1} linhas: fica difícil de ler na projeção. Mostrem só as principais (até 7).`]);
      const longos = s.textos.filter(t => palavras(t) > 25).length;
      if (longos && pal <= 90) r.push(['mid', `${longos === 1 ? 'Há um parágrafo longo' : 'Há ' + longos + ' parágrafos longos'}. Transformem em tópicos de uma linha.`]);
      if (s.textos.length > 8) r.push(['mid', `${s.textos.length} linhas de texto. O ideal é até 6.`]);
      const capaOuFim = i === 0 || i === n - 1;
      if (!capaOuFim && s.miudas >= 12) r.push(['bad', `Letra pequena (${s.menor} pt). Na projeção, usem no mínimo 18 pt no texto e 28 pt no título.`]);
      if (tcc) {
        const copiadas = s.textos.flatMap(t => frases(t)).filter(f => palavras(f) >= 25 && tcc.includes(f.toLowerCase().replace(/\s+/g, ' ').replace(/[.!?]$/, ''))).length;
        if (copiadas) r.push(['mid', `${copiadas === 1 ? 'Uma frase foi copiada' : copiadas + ' frases foram copiadas'} do TCC inteira. No slide vão só as palavras-chave; o texto completo vocês falam.`]);
      }
      if (/constru|funcion|projeto|processo|ishikawa|fluxo|pareto/i.test(s.titulo) && !s.imagens && !s.graficos) r.push(['mid', 'Este slide pediria uma imagem (foto, desenho, fluxograma ou gráfico).']);
      if (r.length) ps.push({ nome, itens: r });
    });
    return { resumo, gerais: g, porSlide: ps };
  }

  function htmlSugestoes(res, titulo) {
    const ic = { ok: '✓', mid: '!', bad: '✗' };
    const item = ([c, t]) => `<li class="sl-${c}"><span>${ic[c]}</span>${esc(t)}</li>`;
    return `<div class="sl-res"><h3>${esc(titulo)}</h3>
      <p class="sl-resumo sl-${res.resumo.c}">${esc(res.resumo.t)}</p>
      ${res.gerais.length ? `<ul class="sl-lista">${res.gerais.map(item).join('')}</ul>` : ''}
      ${res.porSlide.length ? `<h4>Slide por slide</h4>` + res.porSlide.map(p => `<div class="sl-slide"><b>${esc(p.nome)}</b><ul class="sl-lista">${p.itens.map(item).join('')}</ul></div>`).join('')
        : '<p class="sl-tudook">✓ Nenhum slide com excesso de texto ou letra pequena.</p>'}</div>`;
  }

  /* ================= tela ================= */
  const CSS = `
.sl-card{border:1px solid var(--line,#ddd); border-radius:14px; padding:16px 18px; margin:14px 0; background:#fff}
.sl-card h3{margin:0 0 6px; font-size:1.05rem}
.sl-card .sl-num{display:inline-flex; width:26px; height:26px; border-radius:50%; background:var(--marca,#913032); color:#fff; align-items:center; justify-content:center; font-weight:700; margin-right:8px; font-size:.9rem}
.sl-plano{columns:2; column-gap:24px; margin:10px 0; padding-left:20px; font-size:.9rem}
@media (max-width:700px){ .sl-plano{columns:1} }
.sl-plano li{margin:2px 0; break-inside:avoid}
.sl-fora{font-size:.85rem; color:#7a4b00; background:#fff6e0; border-radius:8px; padding:8px 10px; margin:8px 0}
.sl-res{margin-top:14px; border-top:1px dashed var(--line,#ddd); padding-top:12px}
.sl-res h3{font-size:1rem; margin:0 0 8px}.sl-res h4{font-size:.92rem; margin:14px 0 4px}
.sl-resumo{padding:10px 12px; border-radius:8px; font-weight:600; margin:0 0 8px}
.sl-resumo.sl-ok{background:#e6f4ec; color:#1d5e3a}.sl-resumo.sl-bad{background:#fde7e7; color:#8a1c1c}.sl-resumo.sl-mid{background:#fff1d6; color:#7a4b00}
.sl-lista{list-style:none; padding:0; margin:4px 0}
.sl-lista li{display:flex; gap:8px; padding:4px 0; font-size:.9rem}
.sl-lista li span{flex:0 0 18px; font-weight:700; text-align:center}
.sl-lista li.sl-bad span{color:#b3261e}.sl-lista li.sl-mid span{color:#a06800}.sl-lista li.sl-ok span{color:#1d7a4a}
.sl-slide{padding:6px 0; border-top:1px solid #f0f0f0}
.sl-tudook{color:#1d7a4a}
.sl-btnarq{display:inline-block; cursor:pointer; padding:9px 16px; border-radius:10px; border:1px solid var(--accent,#312B50); color:var(--accent,#312B50); font-weight:700; background:#fff}
.sl-btnarq:hover{background:#f2f1f7}
.sl-msg{font-size:.88rem; margin-left:8px}
`;
  function plano() {
    const S = o.S(), f = o.curso === 'pcp' ? roteiroPCP : roteiroMecanica;
    return f(S);
  }
  async function html() {
    const r = await plano();
    const nomes = r.slides.map(s => s.titulo === 'Obrigado!' ? 'Obrigado / perguntas' : s.tipo === 'capa' ? 'Capa' : s.titulo);
    return `<div class="panel">
      <div class="eyebrow">Final</div><h2 style="margin:4px 0 12px">Apresentação (slides)</h2>
      <div class="note"><p>A apresentação para a banca tem de <b>10 a 15 minutos</b>. Contem com cerca de <b>1 minuto por slide</b>: de 12 a 15 slides, com capa e encerramento.</p>
      <p>Aqui vocês podem <b>gerar os slides sugeridos</b> no padrão do CEEP e, depois de ajustar, <b>mandar os slides para receber sugestões</b>.</p></div>
      <div class="sl-card"><h3><span class="sl-num">1</span>Gerar os slides sugeridos</h3>
        <p class="small">O assistente monta um PowerPoint no padrão do CEEP com o que vocês escreveram no TCC: tópicos curtos, tabelas, fotos e gráficos.
        Nas <b>anotações</b> de cada slide fica o texto completo, para vocês saberem o que falar. É um <b>rascunho</b>: revisem, enxuguem e ajustem antes de apresentar.</p>
        <p class="small" style="margin-bottom:0"><b>${r.slides.length} slides</b> (cerca de ${Math.round((r.slides.length - 2) * 0.8 + 1)} a ${Math.round((r.slides.length - 2) * 1.15 + 1.5)} minutos):</p>
        <ol class="sl-plano">${nomes.map(n => `<li>${esc(n)}</li>`).join('')}</ol>
        ${r.fora.length ? `<div class="sl-fora"><b>Ficaram de fora</b> porque ainda não estão preenchidos no TCC: ${esc(r.fora.join(', '))}.</div>` : ''}
        ${(r.cortados || []).length ? `<div class="sl-fora" style="background:#eef1f8;color:#2b3a67"><b>Não entraram para caber em 15 minutos:</b> ${esc(r.cortados.join(', '))}. Se quiserem, coloquem depois no lugar de outro slide.</div>` : ''}
        <button class="primary" id="slGerar">Baixar os slides (PowerPoint)</button><span class="sl-msg" id="slMsg"></span>
        <p class="small" style="margin-top:10px">Para usar no Google Slides: no Drive, cliquem em <b>Novo → Upload de arquivo</b>, escolham o arquivo e abram com o Google Apresentações.</p>
        <div id="slResGerado"></div>
      </div>
      <div class="sl-card"><h3><span class="sl-num">2</span>Já fizeram os slides? Recebam sugestões</h3>
        <p class="small">Mandem o arquivo <b>PowerPoint (.pptx)</b>. No Google Slides: <b>Arquivo → Fazer download → Microsoft PowerPoint (.pptx)</b>.
        O assistente confere o tempo, o excesso de texto, o tamanho da letra, as imagens e se todas as partes do TCC aparecem.</p>
        <label class="sl-btnarq">Escolher o arquivo dos slides (.pptx)<input type="file" id="slArquivo" accept=".pptx,application/vnd.openxmlformats-officedocument.presentationml.presentation" hidden></label>
        <span class="sl-msg" id="slMsg2"></span>
        <div id="slResEquipe"></div>
      </div>
      <div class="sl-card"><h3><span class="sl-num">3</span>Dicas para o dia</h3>
        <ul class="sl-lista">
          <li class="sl-ok"><span>✓</span>Ensaiem com cronômetro pelo menos duas vezes. O tempo é de 10 a 15 minutos.</li>
          <li class="sl-ok"><span>✓</span>Não leiam o slide: olhem para a banca e expliquem com as próprias palavras.</li>
          <li class="sl-ok"><span>✓</span>Cada integrante apresenta uma parte. Combinem quem fala em cada slide.</li>
          <li class="sl-ok"><span>✓</span>Mostrem o equipamento, as fotos e os números. ${o.curso === 'pcp' ? 'Destaquem o antes e depois dos indicadores e o payback.' : 'Se possível, levem o protótipo e mostrem funcionando.'}</li>
          <li class="sl-ok"><span>✓</span>Preparem-se para as perguntas: por que escolheram essa solução, de onde vieram os números e o que fariam diferente.</li>
          <li class="sl-ok"><span>✓</span>Levem os slides no pendrive e no Drive, e testem no computador da sala antes.</li>
        </ul>
      </div>
    </div>`;
  }
  function nomeArq() {
    const S = o.S();
    const t = (L(S.titulo) || 'TCC').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^A-Za-z0-9]+/g, '_').slice(0, 40).replace(/_+$/, '');
    return `Slides_${t}.pptx`;
  }
  async function analisarArquivo(buf, alvo, titulo) {
    const slides = await lerPptx(buf);
    if (!slides.length) throw new Error('Não encontrei slides nesse arquivo.');
    const res = sugerir(slides, o.S(), o.curso);
    alvo.innerHTML = htmlSugestoes(res, titulo);
  }
  function ligar() {
    const g = document.getElementById('slGerar'); if (!g) return;
    g.onclick = async () => {
      const msg = document.getElementById('slMsg');
      g.disabled = true; msg.textContent = 'Montando os slides…';
      try {
        const r = await plano();
        await montar(r, nomeArq());
        msg.textContent = '✓ Slides baixados. Abram e revisem.';
      } catch (err) { console.error(err); msg.textContent = err.message || 'Não consegui gerar os slides.'; }
      finally { g.disabled = false; }
    };
    const a = document.getElementById('slArquivo');
    a.onchange = async () => {
      const f = a.files[0]; a.value = ''; if (!f) return;
      const msg = document.getElementById('slMsg2'), alvo = document.getElementById('slResEquipe');
      if (!/\.pptx$/i.test(f.name)) { msg.textContent = 'Mandem o arquivo em PowerPoint (.pptx).'; return; }
      msg.textContent = 'Lendo os slides…'; alvo.innerHTML = '';
      try { await analisarArquivo(await f.arrayBuffer(), alvo, 'Sugestões para “' + f.name + '”'); msg.textContent = ''; alvo.scrollIntoView({ behavior: 'smooth', block: 'start' }); }
      catch (err) { console.error(err); msg.textContent = err.message || 'Não consegui ler esse arquivo.'; }
    };
  }

  window.SlidesTCC = {
    iniciar(opcoes) { o = opcoes; const st = document.createElement('style'); st.textContent = CSS; document.head.appendChild(st); },
    async desenhar(alvo, rodapeNav) { alvo.innerHTML = await html() + rodapeNav; ligar(); },
    // para testes
    _plano: plano, _lerPptx: lerPptx, _sugerir: sugerir, _montar: montar,
  };
})();
