/* Assistente do TCC do Técnico em PCP – importação do que a equipe já escreveu.
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
  // títulos do Modelo de TCC do Técnico em PCP (9 capítulos)
  const REGRAS = [
    [/^resumo$/, 'resumo'], [/^abstract/, 'ignorar'], [/^(sumario|lista de|indice|epigrafe|ficha catalog)/, 'ignorar'],
    [/^dedicatoria/, 'dedicatoria'], [/^agradecimento/, 'agradecimentos'],
    [/^introducao/, 'tema'], [/contextualizacao|^apresentacao do problema/, 'contexto'],
    [/^problema( de pesquisa)?$|^problema de pesquisa|^problematica|^pergunta/, 'problema'], [/^justificativa/, 'justificativa'],
    [/^objetivo geral/, 'objGeral'], [/^objetivos? especific/, 'objEsp'], [/^objetivos?\b/, 'objetivos'], [/^estrutura do trabalho/, 'ignorar'],
    [/caracterizacao|^a (empresa|organizacao)$|apresentacao da (empresa|organizacao)/, 'carac'],
    [/historico|segmento de atuacao/, 'historico'], [/^produtos|^servicos|mix de produtos/, 'produtosTexto'],
    [/estrutura organizacional|organograma/, 'estrutura'], [/estrutura fisica|capacidade produtiva|^capacidade/, 'capacidade'],
    [/^processo produtivo|fluxo produtivo/, 'processo'],
    [/fundamentacao|referencial teorico|revisao (bibliografica|de literatura)|base teorica/, 'fund'],
    [/^metodologia|procedimentos metodolog/, 'metodo'], [/tipo de pesquisa|classificacao da pesquisa/, 'tipoPesquisa'],
    [/objeto de estudo/, 'objeto'], [/coleta de dados/, 'coleta'], [/analise dos dados/, 'analise'], [/etapas da pesquisa/, 'etapas'],
    [/^diagnostico/, 'diag'], [/(descricao|mapeamento) do processo atual|^processo atual/, 'procAtual'],
    [/levantamento de dados|dados do problema|^levantamento/, 'dadosTexto'], [/analise das causas|^causas/, 'causasTexto'],
    [/priorizacao|matriz gut/, 'gut'], [/sintese do diagnostico/, 'sintese'],
    [/^proposta|proposta de melhoria/, 'proposta'], [/relacao entre causas|causas e acoes/, 'acoes'], [/plano de acao|5w2h/, 'planoTexto'],
    [/indicadores de desempenho|^indicadores/, 'indicadores'],
    [/aplicacao das ferramentas|^aplicacao/, 'aplic'],
    [/^resultados e discussao|^resultados/, 'resultadosTexto'], [/apresentacao dos resultados/, 'resultadosTexto'],
    [/analise financeira|analise economica|viabilidade/, 'economiaTexto'], [/^discussao/, 'discussaoTexto'],
    [/conclus|consideracoes finais/, 'conclusao'],
    [/^referencias|^bibliografia/, 'referencias'],
    [/roteiro de entrevista/, 'roteiro'], [/uso de inteligencia|inteligencia artificial|registro de uso/, 'ia'],
    [/^apendice|^anexo/, 'apendice'],
  ];
  // trechos do exemplo do modelo (Confeitaria Doce Sabor): não são do grupo
  const EXEMPLO = new Set(["1 levantamento bibliografico livros e artigos ma", "2 caracterizacao da empresa entrevista abr. 2026", "3 mapeamento do processo atual fluxograma, obser", "4 analise das causas pareto, ishikawa, 5 porques", "5 elaboracao da proposta 5w2h jun. 2026", "5 porques: tecnica de perguntas sucessivas para ", "5s: programa de organizacao baseado nos sensos d", "5w2h what, why, where, when, who, how, how much", "5w2h: plano de acao que define o que sera feito,", "6 aplicacao das ferramentas curva abc, ponto de ", "7 avaliacao dos resultados indicadores, simulaca", "[cole aqui o trecho copiado exatamente como apar", "a classificacao abc partiu do valor consumido de", "a confeitaria conta com 14 colaboradores, distri", "a confeitaria doce sabor foi fundada em 2015 pel", "a confeitaria doce sabor, localizada em fazenda ", "a curva abc e um metodo de classificacao dos ite", "a empresa atua no segmento alimenticio, com prod", "a empresa ocupa 180 m², divididos em area de ate", "a empresa opera em um turno, com jornada de 44 h", "a empresa trabalha com tres familias de produtos", "a figura 3 destaca os dois indicadores ligados d", "a gestao de estoques busca responder a duas perg", "a gestao de estoques ocupa posicao central nesse", "a reducao das rupturas confirma o que tubino (20", "a tabela 8 compara os indicadores antes da melho", "abc classificacao de itens por importancia (curv", "abnt associacao brasileira de normas tecnicas", "acompanhar os indicadores medir os resultados es", "acuracia do estoque itens com saldo correto ÷ it", "acuracia do estoque – 91% – 90% atingida", "administrativo compras, financeiro e estoque 1", "almoxarifado sem identificacao 3 3 3 27 3ª", "almoxarifado sem identificacao organizar, identi", "apenas 3 dos 28 insumos (11%) concentram 72% do ", "aplicar o 5s no almoxarifado organizar e identif", "as causas das categorias metodo e medida foram a", "as causas identificadas foram priorizadas pela m", "as ferramentas da qualidade sao tecnicas que aux", "as possiveis causas das rupturas foram levantada", "as rupturas de estoque afetam tres grupos: os cl", "associacao brasileira de normas tecnicas. nbr 14", "atendimento balcao e pedidos online 2", "ausencia de ponto de pedido 5 5 4 100 1ª", "ausencia de ponto de pedido calcular o ponto de ", "avaliar os resultados por meio de indicadores de", "bolos chocolate, cenoura, bolos de festa 180 45", "cada compra emergencial custava, em media, r$ 60", "cada insumo da classe a foi dividido em duas cai", "calcular o ponto de pedido evitar compras tardia", "campo conteudo", "campo descricao", "causa (capitulo 5) acao proposta ferramenta", "causa g u t g × u × t prioridade", "chocolate 300 kg 12 kg 3 12 kg 48 kg", "chocolate 48 kg 70 kg 82 kg", "chocolate em barra 18 33 33", "classe nº de itens % dos itens valor mensal (r$)", "classificar insumos (curva abc) priorizar itens ", "classificar os insumos por meio da curva abc;", "codigo ins-001", "com que frequencia falta algum insumo? quais fal", "como e feito hoje o controle dos insumos?", "como limitacao, destaca-se que os resultados for", "como limitacao, os resultados foram obtidos por ", "como o resultado foi utilizado as correcoes fora", "como o resultado foi utilizado as formulas suger", "como o resultado foi utilizado serviu como ponto", "compras centralizadas em uma pessoa 3 2 3 18 4ª", "compras emergenciais nº de compras fora da progr", "compras emergenciais por mes 9 2 eœ¢?78% ate 2 a", "compras emergenciais por mes 9 2 −78% ate 2 atin", "creme de leite 600 caixas 24 caixas 3 24 caixas ", "creme de leite 9 17 76", "creme de leite 96 caixas 144 caixas 168 caixas", "custo extra das compras emergenciais (%) 15 nota", "das 54 rupturas registradas no periodo, 41 (76%)", "data 10/07/2026", "data 15/04/2026", "data 20/09/2026", "declara-se que as ferramentas de ia foram utiliz", "diagrama de ishikawa (causa e efeito): organizac", "diagrama de pareto: ordenacao das ocorrencias po", "diante desse cenario, este trabalho busca respon", "dimensionar o ponto de pedido e o sistema kanban", "disciplina (shitsuke) auditoria 5s semanal com n", "doces finos brigadeiro, beijinho, bombons 2.400 ", "embalagem e expedicao auxiliares de embalagem e ", "embalagens 4 7 93", "entre as vantagens do sistema estao a reducao de", "es estoque de seguranca", "este apendice registra os prompts (comandos) env", "este trabalho esta organizado em nove secoes. a ", "este trabalho partiu da seguinte questao: como a", "estoques sao acumulos de recursos materiais entr", "etapa atividade ferramenta periodo", "existe algum registro de consumo ou de compras?", "falta de registro de consumo 5 4 4 80 2ª", "falta de registro de consumo criar planilha de c", "falta percebida so quando o item acaba implantar", "familia de produtos exemplos producao semanal (u", "ferramenta [nome da ferramenta de ia e versao]", "figura 1 – fluxograma do processo atual de contr", "figura 2 – diagrama de ishikawa das rupturas de ", "figura 3 – rupturas e compras emergenciais por m", "finalidade apoio na montagem da planilha", "finalidade pesquisa inicial sobre o tema", "finalidade revisao ortografica", "fluxograma: representacao grafica da sequencia d", "fonte: os autores (2026), com dados simulados.", "fonte: os autores (2026).", "fornecedor distribuidora x (ficticia)", "gerencia proprietaria e gerente geral 1", "gut gravidade, urgencia e tendencia", "identificar as causas das rupturas de estoque;", "implantar o kanban de duas caixas sinalizar a re", "indicador antes depois variacao meta situacao", "indicador formula valor atual meta", "indicador valor atual como foi obtido", "insumo caixa 2 (reserva = pp) lote de compra est", "insumo consumo mensal d (por dia) lt (dias) es p", "insumo rupturas % % acumulado", "item chocolate em barra 1 kg", "keywords: production planning and control. inven", "leite condensado 14 26 59", "leite condensado 600 latas 24 latas 3 24 latas 9", "leite condensado 96 latas 144 latas 168 latas", "limpeza (seiso) limpeza diaria de 10 minutos ao ", "local almoxarifado seco, prateleira a1", "lt lead time (prazo de entrega do fornecedor)", "mapear o processo atual de recebimento, armazena", "matriz gut: priorizacao das causas pelas notas d", "morango 5 9 85", "nao se sabe quais itens priorizar classificar os", "nota: o sinal “–” indica que o indicador nao era", "o controle de insumos da confeitaria doce sabor ", "o custo total estimado da proposta e de r$ 470,0", "o diagnostico demonstrou que as rupturas se conc", "o investimento, portanto, e recuperado em pouco ", "o kanban e um sistema de controle da producao e ", "o objetivo geral deste trabalho e analisar o pro", "o objetivo geral, de analisar o processo de cont", "o objeto de estudo e o processo de controle de i", "o planejamento e controle da producao (pcp) cons", "o planejamento e controle da producao (pcp) e a ", "o planejamento e controle da producao contribui ", "o plano de acao foi estruturado pela ferramenta ", "o ponto de pedido corresponde a quantidade em es", "o ponto de pedido foi calculado pela equacao 1, ", "o processo produtivo segue o fluxo geral: fornec", "o programa 5s foi aplicado no almoxarifado para ", "o quadro 2 apresenta a sequencia de etapas do tr", "o quadro 4 relaciona cada causa identificada no ", "o que acontece com a producao quando falta um in", "o que por que onde quando quem como quanto", "o trabalho se justifica por propor uma solucao d", "observa-se que a reposicao depende exclusivament", "onde pp e o ponto de pedido, d e a demanda media", "ordenacao (seiton) prateleiras por classe abc, i", "orientador(a): prof.(º) rodrigo leandro martins", "os dados foram coletados por tres meios: entrevi", "os dados foram organizados em planilha eletronic", "os itens da classe a, poucos e de alto valor, re", "outros (23 itens) 4 7 100", "padronizacao (seiketsu) etiquetas com nome, codi", "palavras-chave: planejamento e controle da produ", "para a empresa, a principal contribuicao e um si", "para alcanca-lo, foram definidos os seguintes ob", "para avaliar a melhoria, foram definidos os indi", "para dimensionar o problema, foram levantadas as", "para o chocolate, o consumo medio e de 300 kg po", "para slack, brandon-jones e johnston (2018), o p", "payback = investimento ÷ economia mensal = 470 ÷", "pcp planejamento e controle da producao", "pdca plan, do, check, act", "pedidos entregues com atraso (%) 18 registro de ", "pedidos no prazo 82% 94% +12 p.p. 95% parcial", "pedidos no prazo pedidos no prazo ÷ total de ped", "pergunta resposta", "por que faltou chocolate? porque o pedido ao for", "por que nao ha quantidade minima? porque o consu", "por que o consumo nao e registrado? porque a emp", "por que o pedido foi feito tarde? porque a falta", "por que so foi percebida nesse momento? porque n", "prazo de entrega 3 dias", "producao confeiteira-chefe 1", "producao confeiteiros e auxiliares 7", "producao parada por falta de insumo (h/mes) 6 re", "production planning and control helps companies ", "prompt utilizado explique, de forma simples, a d", "prompt utilizado revise a ortografia, a pontuaca", "prompt utilizado tenho uma planilha com 28 insum", "quadro 1 – estrutura organizacional da confeitar", "quadro 2 – etapas da pesquisa", "quadro 2 – etapas da pesquisa 18", "quadro 3 – aplicacao dos 5 porques", "quadro 3 – aplicacao dos 5 porques 21", "quadro 4 – relacao entre causas, acoes e ferrame", "quadro 5 – plano de acao 5w2h", "quadro 5 – plano de acao 5w2h 23", "quadro 6 – indicadores de desempenho", "quadro 6 – indicadores de desempenho 24", "quadro 7 – modelo de cartao kanban", "quadro 7 – modelo de cartao kanban 26", "quadro 8 – acoes do programa 5s", "quadro 8 – acoes do programa 5s 26", "qual e o prazo de entrega de cada fornecedor?", "quantidade de reposicao 70 kg", "quanto a natureza, trata-se de uma pesquisa apli", "que melhorias a senhora considera importantes?", "quem decide quando e quanto comprar?", "rupturas por mes 9 2 eœ¢?78% ate 3 atingida", "rupturas por mes 9 2 −78% ate 3 atingida", "rupturas por mes 9 media de 6 meses", "rupturas por mes nº de vezes em que um insumo fa", "senso acao realizada evidencia", "setor funcao quantidade", "tabela 1 – familias de produtos e producao media", "tabela 2 – rupturas de estoque por insumo em sei", "tabela 3 – indicadores do processo atual", "tabela 3 – indicadores do processo atual 20", "tabela 4 – priorizacao das causas pela matriz gu", "tabela 5 – classificacao abc dos insumos", "tabela 5 – classificacao abc dos insumos 25", "tabela 6 – calculo do ponto de pedido dos itens ", "tabela 7 – dimensionamento do kanban de duas cai", "tabela 8 – indicadores antes e depois da melhori", "tortas holandesa, limao, morango 60 20", "total 14", "total 28 100 18.000 100 –", "total 54 100 –", "total – 100", "trabalho de conclusao de curso apresentado ao cu", "treinar os colaboradores garantir o uso do siste", "tubino (2017) explica que o kanban trabalha com ", "utilizacao (seiri) descarte de embalagens vencid"]);
  const LINHAS_EX = new Set(["1 | levantamento bibliografico | livros ", "2 | caracterizacao da empresa | entrevis", "3 | mapeamento do processo atual | fluxo", "4 | analise das causas | pareto, ishikaw", "5 | elaboracao da proposta | 5w2h | jun.", "6 | aplicacao das ferramentas | curva ab", "7 | avaliacao dos resultados | indicador", "a | 3 | 11 | 12.960 | 72 | 72", "acompanhar os indicadores | medir os res", "acuracia do estoque | itens com saldo co", "acuracia do estoque | – | 91% | – | 90% ", "administrativo | compras, financeiro e e", "almoxarifado sem identificacao | 3 | 3 |", "almoxarifado sem identificacao | organiz", "aplicar o 5s no almoxarifado | organizar", "atendimento | balcao e pedidos online | ", "ausencia de ponto de pedido | 5 | 5 | 4 ", "ausencia de ponto de pedido | calcular o", "b | 4 | 14 | 3.600 | 20 | 92", "bolos | chocolate, cenoura, bolos de fes", "c | 21 | 75 | 1.440 | 8 | 100", "calcular o ponto de pedido | evitar comp", "chocolate em barra | 18 | 33 | 33", "chocolate | 300 kg | 12 kg | 3 | 12 kg |", "chocolate | 48 kg | 70 kg | 82 kg", "classificar insumos (curva abc) | priori", "codigo | ins-001", "como o resultado foi utilizado | as corr", "como o resultado foi utilizado | as form", "como o resultado foi utilizado | serviu ", "compras centralizadas em uma pessoa | 3 ", "compras emergenciais por mes | 9 | 2 | e", "compras emergenciais por mes | 9 | 2 | −", "compras emergenciais | nº de compras for", "creme de leite | 600 caixas | 24 caixas ", "creme de leite | 9 | 17 | 76", "creme de leite | 96 caixas | 144 caixas ", "custo extra das compras emergenciais (%)", "data | 10/07/2026", "data | 15/04/2026", "data | 20/09/2026", "disciplina (shitsuke) | auditoria 5s sem", "doces finos | brigadeiro, beijinho, bomb", "embalagem e expedicao | auxiliares de em", "embalagens | 4 | 7 | 93", "falta de registro de consumo | 5 | 4 | 4", "falta de registro de consumo | criar pla", "falta percebida so quando o item acaba |", "ferramenta | [nome da ferramenta de ia e", "finalidade | apoio na montagem da planil", "finalidade | pesquisa inicial sobre o te", "finalidade | revisao ortografica", "fornecedor | distribuidora x (ficticia)", "gerencia | proprietaria e gerente geral ", "implantar o kanban de duas caixas | sina", "item | chocolate em barra 1 kg", "leite condensado | 14 | 26 | 59", "leite condensado | 600 latas | 24 latas ", "leite condensado | 96 latas | 144 latas ", "limpeza (seiso) | limpeza diaria de 10 m", "local | almoxarifado seco, prateleira a1", "morango | 5 | 9 | 85", "nao se sabe quais itens priorizar | clas", "ordenacao (seiton) | prateleiras por cla", "outros (23 itens) | 4 | 7 | 100", "padronizacao (seiketsu) | etiquetas com ", "pedidos entregues com atraso (%) | 18 | ", "pedidos no prazo | 82% | 94% | +12 p.p. ", "pedidos no prazo | pedidos no prazo ÷ to", "por que faltou chocolate? | porque o ped", "por que nao ha quantidade minima? | porq", "por que o consumo nao e registrado? | po", "por que o pedido foi feito tarde? | porq", "por que so foi percebida nesse momento? ", "prazo de entrega | 3 dias", "producao parada por falta de insumo (h/m", "producao | confeiteira-chefe | 1", "producao | confeiteiros e auxiliares | 7", "prompt utilizado | explique, de forma si", "prompt utilizado | revise a ortografia, ", "prompt utilizado | tenho uma planilha co", "quantidade de reposicao | 70 kg", "rupturas por mes | 9 | 2 | eœ¢?78% | ate", "rupturas por mes | 9 | 2 | −78% | ate 3 ", "rupturas por mes | 9 | media de 6 meses", "rupturas por mes | nº de vezes em que um", "tortas | holandesa, limao, morango | 60 ", "total | 28 | 100 | 18.000 | 100 | –", "total | 54 | 100 | –", "total | | 14", "total | | – | 100", "treinar os colaboradores | garantir o us", "utilizacao (seiri) | descarte de embalag"]);
  const LEGENDAS_EX = new Set(["diagrama de ishikawa das rupturas de estoque", "fluxograma do processo atual de controle de insumos", "rupturas e compras emergenciais por mes, antes e depois da melhoria"]);
  const doExemplo = t => { const n = norm(t).replace(/^(\d+[.)]|[a-z]\))\s+/, ''); return n.length >= 22 && EXEMPLO.has(n.slice(0, 48)); };
  const RX_EX = /doce sabor|marcia ribeiro|confeiteira-chefe|chocolate em barra|leite condensado|creme de leite|distribuidora x/;
  // frases que o próprio assistente escreve no Word: não precisam voltar para o texto
  const AUTO = [/^(a|o) (tabela|quadro|figura) \d+ (apresenta|relaciona|compara|reune|organiza|ilustra)/, /^as possiveis causas foram organizadas no diagrama/,
    /^para aprofundar a causa principal/, /^as causas identificadas foram priorizadas pela matriz gut/, /^o plano de acao foi estruturado pela ferramenta 5w2h/, /^o plano de acao esta apresentado/,
    /^o custo total estimado da proposta e de/, /^para avaliar a melhoria, foram definidos os indicadores/, /^o calculo (foi feito pela equacao|resultou em)/,
    /^substituindo os valores/, /^como o investimento total foi de/, /^a organizacao conta com \d+ colaboradores/, /^para alcanca-lo, foram definidos/,
    /^este trabalho esta organizado em nove secoes/, /^o resultado esta apresentado na figura/];
  const auto = t => AUTO.some(r => r.test(norm(t)));
  const SIGLAS = /^(PCP|ABC|GUT|SMED|OEE|5S|5W2H|PDCA|MRP|ERP|TPM|JIT|VSM|SIPOC|CEP|PP|ES|LT|NR|ABNT|KPI|OTIF|FIFO|PEPS|TRF|TI|5W1H|POP|CLP|IOT|IA)$/;
  const capitaliza = t => { t = limpa(t); if (!t || t !== t.toUpperCase()) return t; const w = t.toLowerCase().split(' ').map(x => SIGLAS.test(x.toUpperCase()) ? x.toUpperCase() : x).join(' '); return w.charAt(0).toUpperCase() + w.slice(1); };
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
  // capítulo atual -> para onde vai o texto de um subtítulo desconhecido
  const PADRAO_CAP = { tema: 'tema', carac: 'historico', fund: 'fund', metodo: 'tipoPesquisa', diag: 'procAtual', proposta: 'planoTexto', aplic: 'aplic', resultadosTexto: 'resultadosTexto', apendice: 'sobra' };
  function mapear(blocos) {
    const D = {}; const sobras = []; const imagens = []; let modelo = false;
    const anexa = (k, t) => { D[k] = D[k] ? D[k] + '\n\n' + t : t; };
    let alvo = 'capa', cap = 'capa', comecou = false;
    let ultimaLegenda = null; let ultimaImg = null; const capa = [];
    const cur = { fund: null, ferr: null };
    const novoTopico = titulo => { D.fund = D.fund || []; cur.fund = { titulo: titulo || '', texto: '' }; D.fund.push(cur.fund); };
    const novaFerr = titulo => { D.ferramentas = D.ferramentas || []; cur.ferr = { titulo: titulo || '', dados: '', formula: '', variaveis: '', substituicao: '', significado: '' }; D.ferramentas.push(cur.ferr); };
    const tiraMarca = t => t.replace(/^([-•▪◦*]|[a-z]\)|\d+[.)])\s*/, '').replace(/[;.]$/, '');

    for (let i = 0; i < blocos.length; i++) {
      const b = blocos[i];
      if (b.tipo === 'p') {
        const t = b.texto;
        const pk = t.match(/^palavras?[- ]chaves?\s*[:.]\s*(.+)$/i);
        if (pk) { if (!doExemplo(t)) D.palavras = limpa(pk[1]); continue; }
        if (/^keywords?\s*:/i.test(t)) continue;
        const lg = t.match(/^(figura|foto|imagem|grafico|gráfico)\s*\d*\s*[-–—:.]\s*(.+)$/i);
        if (lg && t.length < 160) { ultimaLegenda = limpa(lg[2]); if (ultimaImg && !ultimaImg.legenda) { ultimaImg.legenda = ultimaLegenda; ultimaLegenda = null; } continue; }
        if (/^(quadro|tabela)\s*\d+\s*[-–—:.]/i.test(t) && t.length < 160) continue;
        if (/^fonte\s*:/i.test(t) && t.length < 200) { if (ultimaImg && !ultimaImg.fonte && !/os autores|o autor|pr[óo]prio/i.test(t)) ultimaImg.fonte = t; continue; }
        if (/^nota\s*:/i.test(t) && t.length < 200) continue;
        if (/\.{4,}\s*\d+\s*$/.test(t) || /\.{6,}/.test(t)) continue;
        if (numeracao(t) && /\s\d{1,3}$/.test(t) && t.length < 110 && pareceTitulo(t.replace(/\s\d{1,3}$/, ''))) continue;

        const nivel = b.nivelEstilo || 0;
        const ehTitulo = (nivel > 0 && nivel <= 4 && t.length <= 140) || (pareceTitulo(t) && (!b.lista || !!classifica(t)));
        if (!comecou && alvo === 'capa' && (!ehTitulo || !classifica(t))) capa.push(t);
        if (ehTitulo) {
          const c = classifica(t);
          const sub = (numeracao(t) > 1) || nivel > 1;
          if (c === 'ignorar') { alvo = 'ignorar'; continue; }
          if (sub && cap === 'fund' && !['metodo', 'tipoPesquisa', 'conclusao', 'referencias'].includes(c)) { novoTopico(t.replace(/^(\d+(\.\d+)*)\.?\s+/, '')); alvo = 'fund'; continue; }
          if (sub && cap === 'aplic' && !['resultadosTexto', 'economiaTexto', 'discussaoTexto', 'conclusao'].includes(c)) { novaFerr(t.replace(/^(\d+(\.\d+)*)\.?\s+/, '')); alvo = 'aplic'; continue; }
          if (c) {
            comecou = true;
            if (!sub) cap = c;
            if (['fund', 'aplic', 'metodo', 'diag', 'proposta', 'carac', 'tema', 'resultadosTexto', 'conclusao', 'referencias', 'apendice'].includes(c) && !sub) cap = c;
            if (c === 'fund') cur.fund = null;
            if (c === 'aplic') cur.ferr = null;
            alvo = { carac: 'historico', diag: 'procAtual', proposta: 'planoTexto', metodo: 'tipoPesquisa' }[c] || c;
            continue;
          }
          if (!comecou) continue;
          if (cap === 'fund' || ['tema', 'contexto'].includes(alvo) && !sub) { cap = 'fund'; novoTopico(t.replace(/^(\d+(\.\d+)*)\.?\s+/, '')); alvo = 'fund'; continue; }
          if (sub && !['capa', 'ignorar', 'sobra'].includes(alvo)) continue;
          sobras.push(t); alvo = 'sobra'; continue;
        }
        if (!comecou && alvo === 'capa') continue;
        if (alvo !== 'referencias' && RX_EX.test(norm(t))) { modelo = true; continue; }
        { const m = t.match(/economia estimada [ée] de R\$\s?([\d.]+,\d{2})/i); if (m) D.economiaMensal = m[1]; const c = t.match(/custo total estimado da proposta [ée] de R\$\s?([\d.]+,\d{2})/i); if (c) D.investimento = c[1]; }
        if (alvo !== 'referencias' && doExemplo(t)) continue; // frase igual à do modelo (o Word gerado escreve de novo)
        if (auto(t)) { const m = t.match(/economia estimada [ée] de R\$\s?([\d.]+,\d{2})/i); if (m) D.economiaMensal = m[1]; const c = t.match(/custo total estimado da proposta [ée] de R\$\s?([\d.]+,\d{2})/i); if (c) D.investimento = c[1]; continue; }
        // equações numeradas: na aplicação viram fórmula/cálculo
        const eq = t.match(/^(.*=.*?)\s*\(\d+\)$/);
        if (eq && alvo === 'aplic') { if (!cur.ferr) novaFerr(''); if (!cur.ferr.formula) cur.ferr.formula = limpa(eq[1]); else cur.ferr.substituicao = limpa(eq[1]); continue; }
        if (eq && alvo === 'economiaTexto' && /payback/i.test(t)) continue;
        if (alvo === 'aplic' && /^(onde|em que)\s/i.test(t) && cur.ferr && cur.ferr.formula && !cur.ferr.variaveis) { cur.ferr.variaveis = t.replace(/^(onde|em que)\s*/i, ''); continue; }
        if (/^\[.*\]$/.test(t)) continue; // marcações do modelo
        if (/^registro \d+$/i.test(t)) continue;
        switch (alvo) {
          case 'ignorar': case 'capa': break;
          case 'objGeral': if (!D.objGeral) D.objGeral = t.replace(/^(o\s+)?objetivo geral( deste trabalho)?\s*(é|e|:)\s*/i, ''); else sobras.push(t); break;
          case 'objEsp': if (/espec[íi]fic/i.test(t) && /:$/.test(t)) break; (D.objEsp = D.objEsp || []).push(tiraMarca(t)); break;
          case 'objetivos':
            if (b.lista || /^([-•▪◦*]|[a-z]\))\s/.test(t)) (D.objEsp = D.objEsp || []).push(tiraMarca(t));
            else if (/espec[íi]fic/i.test(t) && t.length < 140) { /* frase de introdução da lista */ }
            else if (!D.objGeral) D.objGeral = t.replace(/^(o\s+)?objetivo geral( deste trabalho)?\s*(é|e|:)\s*/i, '');
            else sobras.push(t);
            break;
          case 'fund': if (!cur.fund) novoTopico(''); cur.fund.texto = cur.fund.texto ? cur.fund.texto + '\n\n' + t : t; break;
          case 'aplic': if (!cur.ferr) novaFerr(''); cur.ferr.dados = cur.ferr.dados ? cur.ferr.dados + '\n\n' + t : t; break;
          case 'estrutura': anexa('historico', t); break;
          case 'etapas': case 'acoes': case 'indicadores': case 'gut': anexa({ etapas: 'analise', acoes: 'planoTexto', indicadores: 'planoTexto', gut: 'sintese' }[alvo], t); break;
          case 'gutTexto': anexa('sintese', t); break;
          case 'roteiro': anexa('roteiro', tiraMarca(t)); break;
          case 'ia': sobras.push(t); break;
          case 'referencias': if (t.length > 15) (D.referencias = D.referencias || []).push({ tipo: 'livre', livre: t }); break;
          case 'apendice': case 'sobra': sobras.push(t); break;
          default: anexa(alvo, t);
        }
      } else if (b.tipo === 'tbl') {
        const plano = norm(b.linhas.map(l => l.join(' ')).join(' '));
        if (/^orientacao|apague esta caixa/.test(plano) || /apague esta caixa/.test(plano)) continue;
        if (RX_EX.test(plano)) { modelo = true; continue; }
        const corpoT = b.linhas.slice(1).filter(l => l.some(c => limpa(c)));
        if (corpoT.length && corpoT.filter(l => LINHAS_EX.has(norm(l.map(limpa).join(' | ')).slice(0, 40))).length >= Math.ceil(corpoT.length / 2)) { modelo = true; continue; }
        const res = tabela(b.linhas, alvo);
        if (res && res.k === 'ignorar') { modelo = true; continue; }
        if (res && res.k === 'w2h') {
          const cab = ['O quê', 'Por quê', 'Onde', 'Quando', 'Quem', 'Como', 'Quanto'];
          res.linhas.forEach(l => anexa('planoTexto', cab.map((c, i) => limpa(l[i]) ? c + ': ' + limpa(l[i]) : '').filter(Boolean).join('; ') + '.'));
          const tot = res.linhas.reduce((a, l) => { const x = parseFloat(String(l[6] || '').replace(/[^\d,.]/g, '').replace(/\.(?=\d{3})/g, '').replace(',', '.')); return a + (isNaN(x) ? 0 : x); }, 0);
          if (tot && !D.investimento) D.investimento = tot.toFixed(2).replace('.', ',');
          continue;
        }
        if (res) D[res.k] = (D[res.k] || []).concat(res.linhas);
        else if (alvo !== 'ignorar' && comecou) sobras.push(b.linhas.map(l => l.join(' | ')).join('\n'));
      } else if (b.tipo === 'img') {
        if (!comecou || alvo === 'ignorar') continue;
        const im = { bytes: b.img.bytes, ext: b.img.ext, legenda: '', fonte: '', alvo: null };
        if (ultimaLegenda) { im.legenda = ultimaLegenda; ultimaLegenda = null; }
        const leg = norm(im.legenda);
        if (leg && LEGENDAS_EX.has(leg)) { modelo = true; continue; }
        if (/ishikawa|causa e efeito|espinha/.test(leg) || alvo === 'causasTexto') im.alvo = { tipo: 'campo', k: 'figIshikawa' };
        else if (/fluxograma|mapeamento/.test(leg) || alvo === 'procAtual') im.alvo = { tipo: 'campo', k: 'figFluxo' };
        else if (cap === 'aplic' && D.ferramentas && D.ferramentas.length) im.alvo = { tipo: 'ferramenta', i: D.ferramentas.length - 1 };
        else if (/antes e depois|resultado/.test(leg) || alvo === 'resultadosTexto') im.alvo = { tipo: 'campo', k: 'figResultado' };
        else im.alvo = { tipo: 'sobra' };
        imagens.push(im); ultimaImg = im;
      }
    }
    lerCapa(capa, D);
    if (D.fund) { D.fund = D.fund.filter(f => limpa(f.texto)); D.fund.forEach(f => f.titulo = capitaliza(f.titulo)); }
    (D.ferramentas || []).forEach(f => f.titulo = capitaliza(f.titulo));
    if (D.objGeral) D.objGeral = D.objGeral.charAt(0).toUpperCase() + D.objGeral.slice(1);
    if (D.economiaTexto) { const m = D.economiaTexto.match(/R\$\s?([\d.]+,\d{2})\s*(?:por|ao|\/)\s*m[eê]s/i); if (m) D.economiaMensal = m[1]; }
    if (D.ferramentas) D.ferramentas = D.ferramentas.filter(f => limpa(f.dados));
    // separa interpretação: último parágrafo do bloco da ferramenta
    (D.ferramentas || []).forEach(f => { const ps = f.dados.split('\n\n'); if (ps.length > 1) { f.significado = ps.pop(); f.dados = ps.join('\n\n'); } });
    if (D.objEsp && !D.objEsp.length) delete D.objEsp;
    return { dados: D, imagens, sobras: sobras.filter(s => limpa(s).length > 2), modelo };
  }
  // nomes dos integrantes e título a partir da capa
  function lerCapa(linhas, D) {
    const NAO = /centro|estadual|educa|profission|curso|tecnico|planejamento|controle|producao|fazenda|grande|trabalho|conclus|orientador|orientacao|professor|prof\b|banca|monografia|universidade|faculdade|colegio|ceep|requisito|apresentad|integrante|nome completo|erotides|nichele|aplicacao|empresa|sistema/;
    const nomes = [], vistos = new Set();
    let titulo = '';
    for (const l of linhas) {
      const n = norm(l);
      if (/^\[/.test(l)) continue;
      const pal = l.split(/\s+/);
      const conector = pal.some(w => /^(em|para|com|no|na|nos|nas|a|o|as|os|por|sobre|ao|aos|um|uma)$/i.test(w));
      const ehNome = !conector && pal.length >= 2 && pal.length <= 7 && pal.every(w => /^[A-Za-zÀ-ÿ'.-]+$/.test(w)) && !NAO.test(n) && pal.filter(w => w.length >= 3).length >= 2
        && (l === l.toUpperCase() || pal.every(w => /^[A-ZÀ-Ý]/.test(w) || /^(de|da|do|dos|das|e)$/i.test(w)));
      if (ehNome) { if (!vistos.has(n)) { vistos.add(n); nomes.push(l.toLowerCase().replace(/(^|\s)(\S)/g, (m, a, c) => a + c.toUpperCase()).replace(/\s(De|Da|Do|Dos|Das|E)\s/g, m => m.toLowerCase())); } continue; }
      if (!titulo && l.length >= 15 && l.length <= 220 && !/^(trabalho|monografia|orientador|orienta|banca|prof|apoio|fazenda rio grande|centro estadual|curso)/.test(n) && !/^\d{4}$/.test(l) && !/^ceep\b/.test(n) && !/universidade|faculdade|instituto|colegio|erotides/.test(n) && !RX_EX.test(n) && !/kanban para reducao de rupturas/.test(n)) titulo = l;
    }
    if (nomes.length) D.integrantes = nomes.slice(0, 4);
    if (titulo) D.titulo = titulo.charAt(0) + titulo.slice(1).toLowerCase();
  }

  // reconhece tabelas pelo cabeçalho
  function tabela(linhas, alvo) {
    if (linhas.length < 2) return null;
    const cab = linhas[0].map(norm);
    const acha = (...rx) => cab.findIndex(c => rx.some(r => r.test(c)));
    const corpo = linhas.slice(1).filter(l => l.some(c => limpa(c)) && !/^total\b/i.test(limpa(l[0])));
    const monta = idx => corpo.map(l => idx.map(i => (i >= 0 ? limpa(l[i]).replace(/^[—–-]$/, '') : '')));
    if (acha(/^g$/) >= 0 && acha(/^u$/) >= 0 && acha(/^t$/) >= 0) return { k: 'gut', linhas: monta([acha(/causa|problema/) >= 0 ? acha(/causa|problema/) : 0, acha(/^g$/), acha(/^u$/), acha(/^t$/)]) };
    if (acha(/^o que|^o que\b|what/) >= 0 && acha(/por que|why/) >= 0) return { k: 'w2h', linhas: monta([acha(/^o que|what/), acha(/por que|why/), acha(/onde|where/), acha(/quando|when/), acha(/quem|who/), acha(/^como|^how$/), acha(/quanto|how much|custo/)]) };
    if (acha(/^pergunta/) >= 0 && acha(/^resposta/) >= 0) return { k: 'porques', linhas: monta([acha(/^pergunta/), acha(/^resposta/)]) };
    if (acha(/objetivo/) >= 0 && acha(/resultado/) >= 0) return { k: 'discussao', linhas: monta([acha(/objetivo/), acha(/resultado/), acha(/situacao/)]) };
    if (acha(/acumulad/) >= 0 || (acha(/ocorrencia|quantidade|frequencia/) >= 0 && acha(/^%$/) >= 0)) { const iq = acha(/ocorrencia|quantidade|frequencia/) >= 0 ? acha(/ocorrencia|quantidade|frequencia/) : 1; return { k: 'dados', linhas: monta([0, iq]) }; }
    if (acha(/^antes/) >= 0 && acha(/^depois/) >= 0) return { k: 'resultados', linhas: monta([acha(/indicador/) >= 0 ? acha(/indicador/) : 0, acha(/^antes/), acha(/^depois/), acha(/^meta/), acha(/situacao/)]) };
    if (acha(/indicador/) >= 0 && acha(/formula/) >= 0) return { k: 'indicadores', linhas: monta([acha(/indicador/), acha(/formula/), acha(/valor atual|atual/), acha(/^meta/)]) };
    if (acha(/indicador/) >= 0 && acha(/valor/) >= 0) return { k: 'indAtuais', linhas: monta([acha(/indicador/), acha(/valor/), acha(/como|obtido|fonte/)]) };
    if (acha(/^causa/) >= 0 && acha(/acao/) >= 0) return { k: 'acoes', linhas: monta([acha(/^causa/), acha(/acao/), acha(/ferramenta/)]) };
    if (acha(/^setor/) >= 0 && acha(/quant|qtd|n[º°o]/) >= 0) return { k: 'estrutura', linhas: monta([acha(/^setor/), acha(/funcao|cargo/), acha(/quant|qtd|n[º°o]/)]) };
    if (acha(/familia|produto|servico/) >= 0 && acha(/producao|semanal|mensal|quantidade|volume|faturamento/) >= 0) return { k: 'produtos', linhas: monta([acha(/familia|produto|servico/), acha(/exemplo/), acha(/producao|semanal|mensal|quantidade|volume/), acha(/faturamento|%/)]) };
    if (acha(/^etapa|atividade/) >= 0 && acha(/periodo|mes|data|prazo/) >= 0) return { k: 'etapas', linhas: monta([acha(/atividade/) >= 0 ? acha(/atividade/) : acha(/^etapa/), acha(/ferramenta|instrumento/), acha(/periodo|mes|data|prazo/)]) };
    if (acha(/ferramenta/) >= 0 && acha(/prompt|pedido|conferi|verific/) >= 0) return { k: 'ia', linhas: monta([acha(/ferramenta/), acha(/data/), acha(/finalidade|para que|uso/), acha(/prompt|pedido/), acha(/conferi|verific|como/)]) };
    if (alvo === 'dadosTexto' && cab.length >= 2) { const iq = cab.findIndex((c, i) => i > 0 && corpo.some(l => /\d/.test(l[i] || ''))); if (iq > 0) return { k: 'dados', linhas: monta([0, iq]) }; }
    // registro de IA no formato do modelo (uma tabela Campo | Descrição por uso)
    if (acha(/^campo/) >= 0 && corpo.some(l => /^ferramenta/i.test(limpa(l[0])))) {
      const v = rx => { const l = corpo.find(x => rx.test(norm(x[0]))); return l ? limpa(l[1]) : ''; };
      const linha = [v(/^ferramenta/), v(/^data/), v(/^finalidade|para que/), v(/^prompt|pedido/), v(/^como|resultado/)];
      if (/^\[/.test(linha[0]) || doExemplo(linha[3])) return { k: 'ignorar', linhas: [] };
      return { k: 'ia', linhas: [linha] };
    }
    if (acha(/integrante/) >= 0 && acha(/tarefa/) >= 0) return { k: 'tarefasTab', linhas: monta([acha(/integrante/), acha(/tarefa/)]) };
    if (acha(/^sigla/) >= 0) return { k: 'siglas', linhas: monta([acha(/^sigla/), acha(/significado/)]) };
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
