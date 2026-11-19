/* Revisor de português – Assistentes do TCC (CEEP FRG)
   Coloca um botão "Revisar português" embaixo de cada caixa de texto.
   Usa o LanguageTool (api.languagetool.org): ortografia, concordância, crase, pontuação…
   Mostra cada erro com a correção e, embaixo, o texto inteiro já corrigido. */
(function () {
  const API = 'https://api.languagetool.org/v2/check';
  const LIMITE = 15000; // caracteres por envio
  const IGN_CHAVE = 'tcc-revisor-ignorar';
  let o = null;
  const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const ignoradas = () => { try { return JSON.parse(localStorage.getItem(IGN_CHAVE) || '[]'); } catch (e) { return []; } };
  const ignorar = p => { const l = ignoradas(); if (!l.includes(p)) l.push(p); try { localStorage.setItem(IGN_CHAVE, JSON.stringify(l.slice(-300))); } catch (e) {} };

  const css = `
.rev-bar{margin:6px 0 2px; display:flex; gap:8px; align-items:center; flex-wrap:wrap}
.rev-btn{font-size:.82rem; padding:5px 10px; border-radius:8px; border:1px solid var(--line,#ccc); background:#fff; color:var(--accent,#312B50); font-weight:600; cursor:pointer}
.rev-btn:hover{border-color:var(--accent,#312B50)}
.rev-msg{font-size:.82rem; color:var(--muted,#666)}
.rev-painel{margin:8px 0 14px; border:1px solid var(--line,#ddd); border-left:4px solid var(--accent,#312B50); border-radius:10px; background:#fff; padding:12px 14px; font-size:.92rem}
.rev-painel h4{margin:0 0 8px; font-size:.95rem}
.rev-item{display:flex; gap:10px; align-items:flex-start; justify-content:space-between; padding:8px 0; border-top:1px dashed var(--line,#ddd)}
.rev-item:first-of-type{border-top:0}
.rev-item .rev-o{text-decoration:line-through; text-decoration-color:#b3261e; color:#b3261e}
.rev-item .rev-n{color:#1d7a4a; font-weight:600}
.rev-item .rev-exp{display:block; font-size:.8rem; color:var(--muted,#666); margin-top:2px}
.rev-item .rev-acoes{display:flex; gap:6px; flex-shrink:0}
.rev-item button, .rev-rodape button{font-size:.78rem; padding:4px 9px; border-radius:7px; border:1px solid var(--line,#ccc); background:#fff; cursor:pointer}
.rev-item button.ok, .rev-rodape button.ok{background:#1d7a4a; border-color:#1d7a4a; color:#fff; font-weight:600}
.rev-item button.rev-alt{padding:1px 7px; font-size:.8rem; color:#1d7a4a; border-color:#9fd4b4; margin:2px 0}
.rev-sug{margin-top:12px; padding:10px 12px; background:#f3f7f4; border-radius:8px; white-space:pre-wrap; line-height:1.55}
.rev-sug mark{background:#c9ecd6; color:inherit; padding:0 2px; border-radius:3px}
.rev-rodape{display:flex; gap:8px; flex-wrap:wrap; margin-top:10px; align-items:center}
.rev-tipo{display:inline-block; font-size:.68rem; text-transform:uppercase; letter-spacing:.04em; padding:1px 6px; border-radius:5px; background:#eee; color:#555; margin-right:4px; vertical-align:1px}
.rev-tipo.ort{background:#fde7e7; color:#9b1c1c}.rev-tipo.gra{background:#fff1d6; color:#7a4b00}.rev-tipo.pon{background:#e6eefc; color:#1f4fa3}
`;

  function tipo(m) {
    const c = m.rule.category.id, t = m.rule.issueType;
    if (t === 'misspelling' || c === 'TYPOS') return ['ort', 'ortografia'];
    if (c === 'PUNCTUATION' || c === 'TYPOGRAPHY') return ['pon', 'pontuação'];
    if (c === 'CASING') return ['ort', 'maiúscula'];
    if (c === 'COLLOQUIALISMS') return ['', 'informal'];
    return ['gra', c === 'GRAMMAR' ? 'gramática' : 'escrita'];
  }
  // tira o que mais atrapalha do que ajuda: dicas de estilo (ex.: trocar "fazer" por "delinear") e nomes próprios
  function serve(m, texto) {
    const c = m.rule.category.id;
    if (m.rule.issueType === 'style' && c !== 'COLLOQUIALISMS') return false;
    if (['STYLE', 'REDUNDANCY', 'FORMAL', 'CLICHES', 'REPETITIONS_STYLE'].includes(c)) return false;
    const trecho = texto.substr(m.offset, m.length);
    if (ignoradas().includes(trecho.toLowerCase())) return false;
    if (m.rule.issueType === 'misspelling' && /^[A-ZÀ-Ý]/.test(trecho) && m.offset > 0 && !/[.!?]\s*$/.test(texto.slice(0, m.offset))) return false; // nome próprio no meio da frase
    if (m.rule.issueType === 'misspelling' && /\d|[A-Z]{2,}/.test(trecho)) return false; // siglas, códigos, medidas
    return true;
  }

  function pedacos(texto) {
    if (texto.length <= LIMITE) return [[0, texto]];
    const r = []; let i = 0;
    while (i < texto.length) {
      let f = Math.min(texto.length, i + LIMITE);
      if (f < texto.length) { const q = texto.lastIndexOf('\n', f); const p = texto.lastIndexOf('. ', f); const corte = Math.max(q, p + 1); if (corte > i + 1000) f = corte + 1; }
      r.push([i, texto.slice(i, f)]); i = f;
    }
    return r;
  }

  async function verificar(texto) {
    const todos = [];
    for (const [ini, parte] of pedacos(texto)) {
      let r;
      try {
        r = await fetch(API, { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
          body: new URLSearchParams({ text: parte, language: 'pt-BR', level: 'default' }) });
      } catch (e) { throw new Error('Não consegui falar com o revisor. Verifiquem a internet (ou ela pode estar bloqueada na escola).'); }
      if (r.status === 429) throw new Error('Muitas revisões ao mesmo tempo no laboratório. Esperem um minuto e cliquem de novo.');
      if (!r.ok) throw new Error('O revisor não respondeu agora. Tentem de novo em instantes.');
      const j = await r.json();
      (j.matches || []).forEach(m => { m.offset += ini; todos.push(m); });
    }
    return todos.filter(m => serve(m, texto) && m.replacements && m.replacements.length)
      .map(m => ({ off: m.offset, len: m.length, troca: m.replacements[0].value, alt: m.replacements.slice(1, 3).map(x => x.value),
        msg: m.message, tipo: tipo(m) }));
  }

  // aplica as trocas (do fim para o começo) e devolve o texto e o HTML com as mudanças marcadas
  function corrigido(texto, itens) {
    const ord = itens.slice().sort((a, b) => a.off - b.off);
    let html = '', txt = '', pos = 0;
    ord.forEach(it => {
      if (it.off < pos) return; // trocas sobrepostas: fica a primeira
      html += esc(texto.slice(pos, it.off)) + '<mark>' + esc(it.troca) + '</mark>';
      txt += texto.slice(pos, it.off) + it.troca;
      pos = it.off + it.len;
    });
    html += esc(texto.slice(pos)); txt += texto.slice(pos);
    return { txt, html };
  }

  function escrever(ta, valor) {
    ta.value = valor;
    ta.dispatchEvent(new Event('input', { bubbles: true }));
  }

  function painel(ta, bar, estado) {
    let p = bar.nextElementSibling;
    if (!p || !p.classList.contains('rev-painel')) { p = document.createElement('div'); p.className = 'rev-painel'; bar.after(p); }
    const { texto, itens } = estado;
    if (!itens.length) {
      p.innerHTML = '<h4 style="color:#1d7a4a">✓ Não encontrei erros de português neste texto.</h4><p class="rev-msg" style="margin:0">O revisor automático ajuda muito, mas não pega tudo: leiam o texto em voz alta antes de entregar.</p><div class="rev-rodape"><button data-rev="fechar">Fechar</button></div>';
    } else {
      const c = corrigido(texto, itens);
      p.innerHTML = `<h4>${itens.length === 1 ? '1 ponto para corrigir' : itens.length + ' pontos para corrigir'}</h4>` +
        itens.map((it, i) => `<div class="rev-item"><div><span class="rev-tipo ${it.tipo[0]}">${it.tipo[1]}</span>
          <span class="rev-o">${esc(texto.substr(it.off, it.len)) || '(espaço)'}</span> → <span class="rev-n">${esc(it.troca) || '(apagar)'}</span>
          ${it.alt.length ? `<span class="rev-msg"> ou </span>${it.alt.map((a, k) => `<button class="rev-alt" data-rev="alt" data-i="${i}" data-a="${k}" title="Usar esta forma">${esc(a)}</button>`).join(' ')}` : ''}
          <span class="rev-exp">${esc(it.msg)}</span></div>
          <div class="rev-acoes"><button class="ok" data-rev="um" data-i="${i}">Corrigir</button><button data-rev="ign" data-i="${i}" title="Não é erro (nome, termo técnico…)">Não é erro</button></div></div>`).join('') +
        `<div class="rev-sug"><b style="display:block;margin-bottom:4px;white-space:normal">Sugestão: o texto com as correções (em verde)</b>${c.html}</div>
        <div class="rev-rodape"><button class="ok" data-rev="todos">Usar o texto corrigido</button><button data-rev="fechar">Fechar</button>
        <span class="rev-msg">Confiram se a sugestão não mudou o sentido do que vocês queriam dizer.</span></div>`;
    }
    p.onclick = e => {
      const b = e.target.closest('button[data-rev]'); if (!b) return;
      const acao = b.dataset.rev;
      if (acao === 'fechar') { p.remove(); return; }
      if (ta.value !== estado.texto) { p.innerHTML = '<p class="rev-msg" style="margin:0">O texto mudou depois da revisão. Cliquem em <b>Revisar português</b> de novo.</p>'; return; }
      if (acao === 'todos') {
        escrever(ta, corrigido(texto, itens).txt);
        p.innerHTML = '<h4 style="color:#1d7a4a;margin:0">✓ Texto corrigido. Leiam de novo para conferir.</h4>';
        return;
      }
      const i = +b.dataset.i, it = itens[i];
      if (acao === 'ign') ignorar(texto.substr(it.off, it.len).toLowerCase());
      let novo = texto;
      if (acao === 'alt') { it.troca = it.alt[+b.dataset.a]; }
      if (acao === 'um' || acao === 'alt') {
        novo = texto.slice(0, it.off) + it.troca + texto.slice(it.off + it.len);
        const delta = it.troca.length - it.len;
        itens.forEach(x => { if (x.off > it.off) x.off += delta; });
        escrever(ta, novo);
      }
      itens.splice(i, 1);
      estado.texto = novo;
      painel(ta, bar, estado);
    };
  }

  async function revisar(ta, bar) {
    const btn = bar.querySelector('.rev-btn'), msg = bar.querySelector('.rev-msg');
    const texto = ta.value;
    if (texto.trim().split(/\s+/).length < 3) { msg.textContent = 'Escrevam o texto primeiro.'; return; }
    btn.disabled = true; btn.textContent = 'Revisando…'; msg.textContent = '';
    try {
      const itens = await verificar(texto);
      if (ta.value !== texto) { msg.textContent = 'O texto mudou durante a revisão. Cliquem de novo.'; return; }
      painel(ta, bar, { texto, itens });
    } catch (err) {
      msg.textContent = err.message;
    } finally { btn.disabled = false; btn.textContent = '✎ Revisar português'; }
  }

  function decorar(raiz) {
    raiz.querySelectorAll('textarea').forEach(ta => {
      if (ta.readOnly || ta.dataset.revOk || ta.dataset.obj === 'referencias') return;
      ta.dataset.revOk = '1';
      ta.spellcheck = true; ta.lang = 'pt-BR';
      const bar = document.createElement('div'); bar.className = 'rev-bar';
      bar.innerHTML = '<button type="button" class="rev-btn">✎ Revisar português</button><span class="rev-msg"></span>';
      ta.after(bar);
      bar.querySelector('.rev-btn').onclick = () => revisar(ta, bar);
    });
  }

  window.RevisorTCC = {
    iniciar(opcoes) {
      o = opcoes || {};
      const st = document.createElement('style'); st.textContent = css; document.head.appendChild(st);
      const main = o.main || document.getElementById('main');
      decorar(main);
      new MutationObserver(() => decorar(main)).observe(main, { childList: true, subtree: true });
    },
  };
})();
