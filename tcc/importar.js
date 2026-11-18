/* Assistente do TCC – importação do que o grupo já escreveu.
   Lê Word (.docx), PDF ou texto (.txt), reconhece títulos, parágrafos, listas, tabelas e imagens
   e distribui cada parte na seção certa do TCC. Sem IA: só regras. */
(function (root) {
  'use strict';
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const norm = s => String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/\s+/g, ' ').trim();
  const limpa = s => String(s || '').replace(/ /g, ' ').replace(/[ \t]+/g, ' ').trim();

  /* ------------------------------------------------------------------ ZIP */
  async function lerZip(buf) {
    const dv = new DataView(buf);
    let i = buf.byteLength - 22;
    while (i >= 0 && dv.getUint32(i, true) !== 0x06054b50) i--;
    if (i < 0) throw new Error('O arquivo não parece ser um Word (.docx) válido.');
    const n = dv.getUint16(i + 10, true); let p = dv.getUint32(i + 16, true);
    const files = {}; const td = new TextDecoder();
    for (let k = 0; k < n; k++) {
      const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true);
      const nlen = dv.getUint16(p + 28, true), xlen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true), loff = dv.getUint32(p + 42, true);
      files[td.decode(new Uint8Array(buf, p + 46, nlen))] = { method, csize, loff };
      p += 46 + nlen + xlen + clen;
    }
    async function get(name) {
      const f = files[name]; if (!f) return null;
      const ln = dv.getUint16(f.loff + 26, true), lx = dv.getUint16(f.loff + 28, true);
      const data = new Uint8Array(buf, f.loff + 30 + ln + lx, f.csize);
      if (f.method === 0) return data.slice();
      const ds = new DecompressionStream('deflate-raw');
      return new Uint8Array(await new Response(new Blob([data]).stream().pipeThrough(ds)).arrayBuffer());
    }
    return { files, get, texto: async name => { const b = await get(name); return b ? new TextDecoder().decode(b) : null; } };
  }

  /* ----------------------------------------------------------------- DOCX */
  async function lerDocx(buf) {
    const z = await lerZip(buf);
    const xml = await z.texto('word/document.xml');
    if (!xml) throw new Error('Não encontrei o texto dentro do arquivo Word.');
    const dp = new DOMParser();
    const doc = dp.parseFromString(xml, 'application/xml');
    // estilos: id -> {nome, nivel}
    const estilos = {};
    const sx = await z.texto('word/styles.xml');
    if (sx) {
      const sd = dp.parseFromString(sx, 'application/xml');
      for (const s of sd.getElementsByTagNameNS(W, 'style')) {
        const id = s.getAttributeNS(W, 'styleId') || s.getAttribute('w:styleId');
        const nm = s.getElementsByTagNameNS(W, 'name')[0];
        const ol = s.getElementsByTagNameNS(W, 'outlineLvl')[0];
        estilos[id] = { nome: nm ? (nm.getAttributeNS(W, 'val') || nm.getAttribute('w:val')) : id, nivel: ol ? +(ol.getAttributeNS(W, 'val') || ol.getAttribute('w:val')) + 1 : 0 };
      }
    }
    // relações (imagens)
    const rels = {};
    const rx = await z.texto('word/_rels/document.xml.rels');
    if (rx) for (const r of dp.parseFromString(rx, 'application/xml').getElementsByTagName('Relationship')) rels[r.getAttribute('Id')] = r.getAttribute('Target');
    const val = el => el && (el.getAttributeNS(W, 'val') || el.getAttribute('w:val'));
    const filho = (el, nome) => { for (const c of el.children) if (c.localName === nome) return c; return null; };

    function textoPar(p) {
      let t = '';
      const walk = el => {
        for (const c of el.children) {
          const ln = c.localName;
          if (ln === 't') t += c.textContent;
          else if (ln === 'tab') t += ' ';
          else if (ln === 'br' || ln === 'cr') t += ' ';
          else if (ln === 'instrText' || ln === 'delText' || ln === 'pPr' || ln === 'rPr') continue;
          else walk(c);
        }
      };
      walk(p);
      return limpa(t);
    }
    const blocos = [];
    async function imagensDe(el) {
      const out = [];
      for (const b of el.getElementsByTagName('*')) {
        if (b.localName !== 'blip') continue;
        const id = b.getAttributeNS(R, 'embed') || b.getAttribute('r:embed');
        const alvo = rels[id]; if (!alvo) continue;
        const caminho = alvo.startsWith('/') ? alvo.slice(1) : 'word/' + alvo.replace(/^\.\//, '');
        const bytes = await z.get(caminho);
        const ext = (caminho.split('.').pop() || '').toLowerCase();
        if (bytes && /^(png|jpe?g|gif|bmp|webp)$/.test(ext)) out.push({ bytes, ext });
      }
      return out;
    }
    async function processa(el) {
      for (const c of el.children) {
        const ln = c.localName;
        if (ln === 'p') {
          const pPr = filho(c, 'pPr');
          const st = pPr && filho(pPr, 'pStyle'); const sid = val(st) || '';
          const info = estilos[sid] || { nome: sid, nivel: 0 };
          const ol = pPr && filho(pPr, 'outlineLvl');
          let nivelEstilo = ol ? +val(ol) + 1 : info.nivel;
          const nome = norm(info.nome);
          if (/^(toc|sumario|indice)/.test(nome) || /^(toc|sumrio)/i.test(sid)) continue; // linhas do sumário
          if (!nivelEstilo) { const m = nome.match(/^(heading|titulo|título)\s*(\d)/); if (m) nivelEstilo = +m[2]; }
          const lista = !!(pPr && filho(pPr, 'numPr')) || /list|lista|marcador/.test(nome);
          const imgs = await imagensDe(c);
          const texto = textoPar(c);
          if (texto) blocos.push({ tipo: 'p', texto, nivelEstilo: nome === 'title' ? 0 : nivelEstilo, lista });
          imgs.forEach(im => blocos.push({ tipo: 'img', img: im }));
        } else if (ln === 'tbl') {
          const linhas = [];
          for (const tr of c.children) {
            if (tr.localName !== 'tr') continue;
            const cel = [];
            for (const tc of tr.children) if (tc.localName === 'tc') cel.push(limpa(Array.from(tc.getElementsByTagNameNS(W, 'p')).map(textoPar).join(' ')));
            linhas.push(cel);
          }
          const imgs = await imagensDe(c);
          if (linhas.length) blocos.push({ tipo: 'tbl', linhas });
          imgs.forEach(im => blocos.push({ tipo: 'img', img: im }));
        } else if (ln === 'sdt') {
          const pr = filho(c, 'sdtPr');
          const gal = pr && pr.getElementsByTagNameNS(W, 'docPartGallery')[0];
          if (gal && /contents|sum/i.test(val(gal) || '')) continue;
          const ct = filho(c, 'sdtContent'); if (ct) await processa(ct);
        } else if (ln === 'body' || ln === 'sdtContent' || ln === 'customXml') await processa(c);
      }
    }
    await processa(doc.documentElement.getElementsByTagNameNS(W, 'body')[0]);
    return blocos;
  }

  /* ------------------------------------------------------------------ PDF */
  async function lerPdf(buf, base) {
    let pdfjs;
    try { pdfjs = await import(base + 'pdfjs/pdf.min.mjs'); }
    catch (e) { throw new Error('Não consegui abrir o leitor de PDF aqui. Enviem o arquivo em Word (.docx).'); }
    pdfjs.GlobalWorkerOptions.workerSrc = base + 'pdfjs/pdf.worker.min.mjs';
    const pdf = await pdfjs.getDocument({ data: new Uint8Array(buf) }).promise;
    const paginas = [];
    for (let i = 1; i <= pdf.numPages; i++) {
      const pg = await pdf.getPage(i);
      const tc = await pg.getTextContent();
      const linhas = [];
      let atual = null;
      for (const it of tc.items) {
        if (!('str' in it)) continue;
        const y = Math.round(it.transform[5]);
        if (!atual || Math.abs(atual.y - y) > 2) { atual = { y, x: it.transform[4], t: '' }; linhas.push(atual); }
        atual.t += it.str + (it.hasEOL ? ' ' : '');
      }
      paginas.push(linhas.map(l => ({ t: limpa(l.t), x: l.x })).filter(l => l.t));
    }
    // linhas repetidas em muitas páginas (cabeçalho/rodapé) saem
    const conta = {};
    paginas.forEach(ls => new Set(ls.map(l => l.t)).forEach(t => conta[t] = (conta[t] || 0) + 1));
    const blocos = [];
    let par = '';
    const fecha = () => { if (par) { blocos.push({ tipo: 'p', texto: limpa(par), nivelEstilo: 0, lista: /^[-•▪◦*]\s|^[a-z]\)\s/.test(par) }); par = ''; } };
    for (const ls of paginas) {
      for (const l of ls) {
        const t = l.t;
        if (/^\d{1,3}$/.test(t)) continue; // número de página
        if (conta[t] > 3 && paginas.length > 5 && t.length < 80) continue;
        if (/\.{4,}\s*\d+\s*$/.test(t)) continue; // sumário
        if (pareceTitulo(t) || /^[-•▪◦*]\s|^[a-z]\)\s/.test(t) || /^(figura|quadro|tabela|fonte)\b/i.test(t)) { fecha(); if (pareceTitulo(t) || /^(figura|quadro|tabela|fonte)\b/i.test(t)) { blocos.push({ tipo: 'p', texto: t, nivelEstilo: 0, lista: false }); continue; } }
        par = par ? (par.endsWith('-') ? par.slice(0, -1) + t : par + ' ' + t) : t;
        if (/[.!?:]$/.test(t) && t.length < 60) fecha();
      }
      // não fecha no fim da página: o parágrafo pode continuar
    }
    fecha();
    return blocos;
  }

  async function lerTxt(buf) {
    const t = new TextDecoder().decode(buf);
    return t.split(/\r?\n/).map(limpa).filter(Boolean).map(texto => ({ tipo: 'p', texto, nivelEstilo: 0, lista: /^[-•*]\s|^[a-z]\)\s/.test(texto) }));
  }

  /* ------------------------------------------------- reconhecer os títulos */
  const REGRAS = [
    [/^resumo$/, 'resumo'], [/^abstract/, 'ignorar'], [/^(sumario|lista de|indice|epigrafe|ficha catalog)/, 'ignorar'],
    [/^dedicatoria/, 'dedicatoria'], [/^agradecimento/, 'agradecimentos'],
    [/^introducao/, 'contexto'], [/contextualiza|^apresentacao|^tema\b/, 'contexto'],
    [/problema( de pesquisa)?$|^problematica|^pergunta/, 'problema'], [/^justificativa$|^justificativas?\b(?!.*escolha)/, 'justificativa'],
    [/^objetivo geral/, 'objGeral'], [/^objetivos? especific/, 'objEsp'], [/^objetivos?\b/, 'objetivos'],
    [/fundamentacao|referencial teorico|revisao (bibliografica|de literatura)|base teorica/, 'fund'],
    [/^metodologia|materiais e metodos|procedimentos metodolog/, 'metodo'],
    [/produtos? semelhantes|benchmark|concorrent|comparacao com|similares|no mercado|do mercado/, 'produtos'],
    [/evolucao|concepcao|ideia inicial|historico do projeto|escolha do projeto/, 'evolucao'],
    [/funcionamento|^principio/, 'funcionamento'],
    [/^desenho|modelagem|\bcad\b|solidworks|projeto tecnico|esbocos?\b/, 'modelagem'],
    [/^materiais$|lista de materiais|materiais utilizados|materiais e componentes|^componentes|principais componentes/, 'materiais'],
    [/dimensionamento|memorial de calculo|^calculos?\b/, 'calculos'],
    [/desenvolvimento do prototipo|construcao|fabricacao|montagem|^desenvolvimento$|^execucao/, 'construcao'],
    [/dificuldade|problemas encontrados|alteracoes/, 'dificuldades'],
    [/^custos?\b|orcamento|levantamento de custos/, 'custos'],
    [/seguranca|nr-?12|^riscos/, 'seguranca'],
    [/teste|ensaio|validacao|verificacao do prototipo/, 'testes'],
    [/^resultados?/, 'resultados'],
    [/discussao|analise dos resultados/, 'discussao'],
    [/conclus|consideracoes finais/, 'conclusao'],
    [/^referencias|^bibliografia/, 'referencias'],
    [/^manual\b/, 'manual'],
    [/^apendice|^anexo/, 'apendice'],
  ];
  const semNumero = t => norm(t).replace(/^(capitulo\s+)?(\d+(\.\d+)*|[ivx]+)\s*[.)\-–—:]?\s*/, '');
  function classifica(t) { const n = semNumero(t); for (const [rx, alvo] of REGRAS) if (rx.test(n)) return alvo; return null; }
  function numeracao(t) { const m = limpa(t).match(/^(\d+(?:\.\d+)*)\.?\s+\S/); return m ? m[1].split('.').length : 0; }
  function pareceTitulo(t) {
    t = limpa(t);
    if (!t || t.length > 100 || /[.;,]$/.test(t)) return false;
    const letras = t.replace(/[^A-Za-zÀ-ÿ]/g, '');
    const maiusc = letras.length && letras.replace(/[^A-ZÀ-Þ]/g, '').length / letras.length > 0.8;
    if (numeracao(t) && (maiusc || classifica(t))) return true;
    if (maiusc && letras.length >= 4 && t.length <= 80) return true;
    return false;
  }

  /* --------------------------------------------- distribuir nas seções */
  function mapear(blocos) {
    const D = {}; // dados importados (mesmos nomes do estado do assistente)
    const relat = {}; const sobras = []; const imagens = [];
    const marca = sec => relat[sec] = (relat[sec] || 0) + 1;
    const anexa = (k, t, sec) => { D[k] = D[k] ? D[k] + '\n\n' + t : t; marca(sec || k); };
    let alvo = 'capa', pai = 'capa', comecou = false;
    let ultimaLegenda = null; let ultimaImg = null; const capa = [];
    const cur = { fund: null, etapa: null, calc: null };
    const novoTopico = titulo => { D.fund = D.fund || []; cur.fund = { titulo: titulo || '', texto: '' }; D.fund.push(cur.fund); marca('fund'); };
    const novaEtapa = () => { D.construcao = D.construcao || []; cur.etapa = { texto: '', foto: { img: '', legenda: '', fonte: '' } }; D.construcao.push(cur.etapa); marca('construcao'); };
    const novoCalc = titulo => { D.calculos = D.calculos || []; cur.calc = { titulo: titulo || '', dados: '', formula: '', variaveis: '', substituicao: '', significado: '' }; D.calculos.push(cur.calc); marca('calculos'); };
    const grupoIntro = ['contexto', 'problema', 'justificativa', 'objetivos', 'objGeral', 'objEsp'];

    for (let i = 0; i < blocos.length; i++) {
      const b = blocos[i];
      if (b.tipo === 'p') {
        const t = b.texto;
        // palavras-chave em qualquer lugar
        const pk = t.match(/^palavras?[- ]chaves?\s*[:.]\s*(.+)$/i);
        if (pk) { D.palavras = limpa(pk[1]); marca('resumo'); continue; }
        if (/^keywords?\s*:/i.test(t)) continue;
        // legendas e fontes
        const lg = t.match(/^(figura|foto|imagem)\s*\d*\s*[-–—:.]\s*(.+)$/i);
        if (lg && t.length < 160) { ultimaLegenda = limpa(lg[2]); if (ultimaImg && !ultimaImg.legenda) { ultimaImg.legenda = ultimaLegenda; ultimaLegenda = null; } continue; }
        if (/^(quadro|tabela)\s*\d+\s*[-–—:.]/i.test(t) && t.length < 160) continue;
        if (/^fonte\s*:/i.test(t) && t.length < 200) { if (ultimaImg && !ultimaImg.fonte && !/os autores|o autor|pr[óo]prio/i.test(t)) ultimaImg.fonte = t; continue; }
        if (/\.{4,}\s*\d+\s*$/.test(t) || /\.{6,}/.test(t)) continue; // linhas de sumário
        if (numeracao(t) && /\s\d{1,3}$/.test(t) && t.length < 110 && pareceTitulo(t.replace(/\s\d{1,3}$/, ''))) continue; // sumário sem pontilhado

        const nivel = b.nivelEstilo || 0;
        const ehTitulo = (nivel > 0 && nivel <= 4 && t.length <= 140) || (pareceTitulo(t) && (!b.lista || !!classifica(t)));
        if (!comecou && alvo === 'capa' && (!ehTitulo || !classifica(t))) capa.push(t);
        if (ehTitulo) {
          const c = classifica(t);
          const sub = (numeracao(t) > 1) || nivel > 1;
          if (c === 'ignorar') { alvo = 'ignorar'; continue; }
          // subtítulos dentro de fundamentação, construção e cálculos viram tópicos/etapas/cálculos
          if (sub && pai === 'fund' && !['metodo', 'conclusao', 'referencias', 'calculos', 'construcao'].includes(c)) { novoTopico(semTitulo(t)); alvo = 'fund'; continue; }
          if (sub && pai === 'construcao' && c !== 'dificuldades' && c !== 'testes' && c !== 'custos') { novaEtapa(); alvo = 'construcao'; continue; }
          if (sub && pai === 'calculos' && !c) { novoCalc(semTitulo(t)); alvo = 'calculos'; continue; }
          if (c) {
            comecou = true;
            alvo = c;
            if (!sub || ['fund', 'metodo', 'construcao', 'calculos', 'conclusao', 'referencias', 'apendice', 'resumo', 'custos', 'seguranca', 'testes'].includes(c)) pai = c;
            if (c === 'fund') { pai = 'fund'; cur.fund = null; }
            if (c === 'construcao') { pai = 'construcao'; cur.etapa = null; }
            if (c === 'calculos') { pai = 'calculos'; cur.calc = null; }
            continue;
          }
          if (!comecou) continue; // títulos da capa
          // título desconhecido
          if (grupoIntro.includes(alvo) || alvo === 'fund') { pai = 'fund'; novoTopico(semTitulo(t)); alvo = 'fund'; continue; }
          if (sub && !['capa', 'ignorar', 'sobra'].includes(alvo)) continue; // subtítulo dentro de uma parte conhecida: o texto continua nela
          sobras.push(t); alvo = 'sobra'; continue;
        }
        if (!comecou && alvo === 'capa') continue;
        // texto comum
        switch (alvo) {
          case 'ignorar': case 'capa': break;
          case 'resumo': anexa('resumo', t); break;
          case 'dedicatoria': anexa('dedicatoria', t); break;
          case 'agradecimentos': anexa('agradecimentos', t); break;
          case 'contexto': anexa('contexto', t); break;
          case 'problema': anexa('problema', t); break;
          case 'justificativa': anexa('justificativa', t); break;
          case 'objGeral': if (!D.objGeral) { D.objGeral = t; marca('objetivos'); } else sobras.push(t); break;
          case 'objEsp': (D.objEsp = D.objEsp || []).push(t.replace(/^([-•▪◦*]|[a-z]\)|\d+[.)])\s*/, '').replace(/[;.]$/, '')); marca('objetivos'); break;
          case 'objetivos':
            if (b.lista || /^([-•▪◦*]|[a-z]\))\s/.test(t)) { (D.objEsp = D.objEsp || []).push(t.replace(/^([-•▪◦*]|[a-z]\)|\d+[.)])\s*/, '').replace(/[;.]$/, '')); marca('objetivos'); }
            else if (!D.objGeral && /^[A-ZÀ-Ý][a-zà-ÿ]+r\b/.test(t.replace(/^(o\s+)?objetivo geral( deste trabalho)?\s*(é|e|:)\s*/i, ''))) { D.objGeral = t.replace(/^(o\s+)?objetivo geral( deste trabalho)?\s*(é|e|:)\s*/i, ''); marca('objetivos'); }
            else if (!D.objGeral) { D.objGeral = t; marca('objetivos'); }
            else if (/espec[íi]fic/i.test(t)) { /* frase de introdução da lista */ }
            else sobras.push(t);
            break;
          case 'fund': if (!cur.fund) novoTopico(''); cur.fund.texto = cur.fund.texto ? cur.fund.texto + '\n\n' + t : t; break;
          case 'metodo': anexa('metodo', t); break;
          case 'produtos': anexa('diferencial', t, 'produtos'); break;
          case 'evolucao': anexa('evolucao', t); break;
          case 'funcionamento': anexa('funcionamento', t); break;
          case 'modelagem': anexa('modelagem', t); break;
          case 'materiais':
            if (b.lista || /^([-•▪◦*]|\d+[.)])\s/.test(t)) { (D.materiais = D.materiais || []).push([t.replace(/^([-•▪◦*]|\d+[.)])\s*/, ''), '', '', '', '']); marca('materiais'); }
            else sobras.push(t);
            break;
          case 'calculos': if (cur.calc) { cur.calc.dados = cur.calc.dados ? cur.calc.dados + '\n\n' + t : t; } else anexa('dimIntro', t, 'calculos'); break;
          case 'construcao': if (!cur.etapa) novaEtapa(); cur.etapa.texto = cur.etapa.texto ? cur.etapa.texto + '\n\n' + t : t; break;
          case 'dificuldades': anexa('dificuldadesTexto', t, 'dificuldades'); break;
          case 'custos': anexa('custosTexto', t, 'custos'); break;
          case 'seguranca': anexa('segurancaTexto', t, 'seguranca'); break;
          case 'testes': anexa('testeProc', t, 'testes'); break;
          case 'resultados': anexa('testeAnalise', t, 'testes'); break;
          case 'discussao': anexa('discussaoTexto', t, 'discussao'); break;
          case 'conclusao': anexa('conclusao', t); break;
          case 'referencias': if (t.length > 15) { (D.referencias = D.referencias || []).push({ tipo: 'livre', livre: t }); marca('referencias'); } break;
          case 'manual': anexa('manual', t, 'apendices'); break;
          default: sobras.push(t);
        }
      } else if (b.tipo === 'tbl') {
        const res = tabela(b.linhas, alvo);
        if (res) { D[res.k] = (D[res.k] || []).concat(res.linhas); marca(res.sec); }
        else if (alvo !== 'ignorar' && comecou) sobras.push(b.linhas.map(l => l.join(' | ')).join('\n'));
      } else if (b.tipo === 'img') {
        if (!comecou || alvo === 'ignorar') continue;
        const im = { bytes: b.img.bytes, ext: b.img.ext, legenda: '', fonte: '', alvo: null };
        if (ultimaLegenda) { im.legenda = ultimaLegenda; ultimaLegenda = null; }
        if (alvo === 'construcao' || pai === 'construcao') {
          if (!cur.etapa || cur.etapa._temFoto) novaEtapa();
          cur.etapa._temFoto = true; im.alvo = { tipo: 'etapa', obj: cur.etapa };
        } else if (alvo === 'funcionamento') im.alvo = { tipo: 'campo', k: 'figFunc' };
        else if (alvo === 'modelagem' || alvo === 'evolucao') im.alvo = { tipo: 'campo', k: 'figCad' };
        else if (alvo === 'testes' || alvo === 'resultados') im.alvo = { tipo: 'campo', k: 'figTeste' };
        else if (alvo === 'apendice' || pai === 'apendice') im.alvo = { tipo: 'desenho' };
        else im.alvo = { tipo: 'sobra' };
        imagens.push(im); ultimaImg = im;
      }
    }
    (D.construcao || []).forEach(e => delete e._temFoto);
    lerCapa(capa, D);
    // limpa itens vazios
    if (D.fund) D.fund = D.fund.filter(f => limpa(f.texto));
    if (D.construcao) D.construcao = D.construcao.filter(e => e.texto || imagens.some(im => im.alvo && im.alvo.obj === e));
    return { dados: D, imagens, relatorio: relat, sobras: sobras.filter(s => limpa(s).length > 2) };
  }
  // nomes dos integrantes e título a partir da capa
  function lerCapa(linhas, D) {
    const NAO = /centro|estadual|educa|profission|curso|tecnico|mecanica|fazenda|grande|trabalho|conclus|orientador|orientacao|professor|prof\b|banca|monografia|universidade|faculdade|colegio|ceep|requisito|apresentad|projeto|desenvolvimento|sistema|maquina|manual|prensa|serra|moedor|dobradeira|ponte|macaco|calculo|apoio|disciplina|\d/;
    const nomes = [], vistos = new Set();
    let titulo = '';
    for (const l of linhas) {
      const n = norm(l);
      const pal = l.split(/\s+/);
      const conector = pal.some(w => /^(em|para|com|no|na|nos|nas|a|o|as|os|por|sobre|ao|aos|um|uma)$/i.test(w));
      const ehNome = !conector && pal.length >= 2 && pal.length <= 7 && pal.every(w => /^[A-Za-zÀ-ÿ'.-]+$/.test(w)) && !NAO.test(n) && pal.filter(w => w.length >= 3).length >= 2
        && (l === l.toUpperCase() || pal.every(w => /^[A-ZÀ-Ý]/.test(w) || /^(de|da|do|dos|das|e)$/i.test(w)));
      if (ehNome) { if (!vistos.has(n)) { vistos.add(n); nomes.push(l.toLowerCase().replace(/(^|\s)(\S)/g, (m, a, c) => a + c.toUpperCase()).replace(/\s(De|Da|Do|Dos|Das|E)\s/g, m => m.toLowerCase())); } continue; }
      if (!titulo && l.length >= 15 && l.length <= 220 && !/^(trabalho|monografia|orientador|orienta|banca|prof|apoio|fazenda rio grande|centro estadual|curso)/.test(n) && !/^\d{4}$/.test(l) && !/^ceep\b/.test(n) && !/universidade|faculdade|instituto|colegio|escola/.test(n) && !/\d{1,2}\/\d{1,2}\/\d{2,4}|\d{1,2}:\d{2}|\d{1,2}h\d{0,2}\b/.test(l)) titulo = l;
    }
    if (nomes.length) D.integrantes = nomes.slice(0, 6);
    if (titulo) D.titulo = titulo.charAt(0) + titulo.slice(1).toLowerCase();
  }
  function semTitulo(t) { return limpa(t).replace(/^(\d+(\.\d+)*)\.?\s+/, ''); }

  // reconhece tabelas pelo cabeçalho
  function tabela(linhas, alvo) {
    if (linhas.length < 2) return null;
    const cab = linhas[0].map(norm);
    const acha = (...rx) => cab.findIndex(c => rx.some(r => r.test(c)));
    const corpo = linhas.slice(1).filter(l => l.some(c => limpa(c)));
    const monta = idx => corpo.map(l => idx.map(i => (i >= 0 ? limpa(l[i]) : '')));
    const iItem = acha(/^item|componente|material|peca|descricao|^nome|produto|modelo/);
    const iQtd = acha(/qtd|quant/);
    const iValor = acha(/valor|preco|custo|r\$/);
    if (acha(/risco|perigo/) >= 0) return { k: 'riscos', sec: 'seguranca', linhas: monta([acha(/risco|perigo/), acha(/onde|local|parte/), acha(/acontec|consequ|dano/), acha(/medida|protec|solucao|controle/)]) };
    if (acha(/problema|dificuldade/) >= 0) return { k: 'dificuldades', sec: 'dificuldades', linhas: monta([acha(/problema|dificuldade/), acha(/causa/), acha(/solucao|acao/), acha(/resultado/)]) };
    if (acha(/^etapa/) >= 0 && acha(/periodo|mes|data|prazo/) >= 0) return { k: 'etapas', sec: 'metodo', linhas: monta([acha(/^etapa/), acha(/feito|descri|atividade/), acha(/periodo|mes|data|prazo/)]) };
    if (acha(/modelo|marca|fabricante/) >= 0 && iValor >= 0 && alvo !== 'custos') return { k: 'produtos', sec: 'produtos', linhas: monta([acha(/modelo|marca|fabricante/), acha(/acionamento|tipo/), acha(/capacidade|caracter/), iValor, acha(/fonte|site|link/)]) };
    if (iValor >= 0 && (iItem >= 0 || alvo === 'custos')) return { k: 'custos', sec: 'custos', linhas: monta([iItem >= 0 ? iItem : 0, iQtd, acha(/unit/) >= 0 ? acha(/unit/) : iValor, acha(/origem|fornecedor|onde/)]) };
    if (iItem >= 0 && (iQtd >= 0 || alvo === 'materiais' || acha(/especifica|dimens|medida|funcao/) >= 0)) return { k: 'materiais', sec: 'materiais', linhas: monta([iItem, iQtd, acha(/especifica|dimens|medida|descri/), acha(/funcao|aplicacao/), acha(/origem|obtenc|fornecedor/)]) };
    if (acha(/ferramenta/) >= 0 && acha(/conferi|usada|uso/) >= 0) return { k: 'ia', sec: 'apendices', linhas: monta([acha(/ferramenta/), acha(/para que|uso|usada/), acha(/parte/), acha(/conferi/)]) };
    return null;
  }

  async function ler(file, base) {
    const buf = await file.arrayBuffer();
    const nome = (file.name || '').toLowerCase();
    if (nome.endsWith('.docx')) return lerDocx(buf);
    if (nome.endsWith('.pdf')) return lerPdf(buf, base);
    if (nome.endsWith('.txt') || nome.endsWith('.md')) return lerTxt(buf);
    if (nome.endsWith('.doc')) throw new Error('Arquivo .doc antigo: abra no Word ou Google Docs e salve como .docx.');
    throw new Error('Use um arquivo Word (.docx), PDF ou texto (.txt).');
  }

  root.TCCImport = { ler, mapear, classifica, pareceTitulo };
})(typeof self !== 'undefined' ? self : this);
