/* Nuvem das equipes – Assistentes do TCC (CEEP FRG)
   Guarda o trabalho em andamento no Drive do professor (pelo Apps Script do professor),
   com código da equipe + senha. Assim a equipe abre o trabalho em qualquer computador.
   Não precisa de login do Google. */
(function () {
  const AUTO_ESPERA = 8000;      // espera depois da última mudança
  const AUTO_MINIMO = 40000;     // intervalo mínimo entre salvamentos automáticos
  const LOTE_IMG = 8 * 1024 * 1024;
  let o = null, st = null, timer = null, ocupado = false, pendente = false, ultimo = 0, pausado = false;

  const chave = () => o.app + '-nuvem';
  function ler() { try { return JSON.parse(localStorage.getItem(chave()) || 'null'); } catch (e) { return null; } }
  function gravar() { try { if (st) localStorage.setItem(chave(), JSON.stringify(st)); else localStorage.removeItem(chave()); } catch (e) {} }
  const hora = t => { if (!t) return ''; const d = new Date(t); const hoje = new Date().toDateString() === d.toDateString();
    return hoje ? 'hoje, ' + d.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' }) : d.toLocaleString('pt-BR', { dateStyle: 'short', timeStyle: 'short' }); };

  async function post(corpo) {
    const r = await fetch(o.url(), { method: 'POST', body: JSON.stringify(Object.assign({ app: o.app }, corpo)), redirect: 'follow' });
    return r.json();
  }
  function status(txt, cor) {
    const el = o.elStatus(); if (!el) return;
    el.textContent = txt; el.style.color = cor || '';
  }
  function botao() {
    const b = o.botao(); if (!b) return;
    b.textContent = st ? '☁ ' + st.codigo : '☁ Nuvem';
    b.title = st ? 'Trabalho guardado na nuvem com o código ' + st.codigo : 'Guardar o trabalho na nuvem para abrir em qualquer computador';
  }
  const ok = t => status('Salvo na nuvem ✓ ' + hora(t), 'var(--ok)');

  /* ---------- salvar ---------- */
  function mudou() {
    if (!st || pausado) return;
    st.uso = Date.now(); gravar();
    status('Alterações ainda não enviadas à nuvem…');
    clearTimeout(timer);
    const falta = Math.max(AUTO_ESPERA, AUTO_MINIMO - (Date.now() - ultimo));
    timer = setTimeout(() => salvarAgora(false), falta);
  }

  async function salvarAgora(forcar, avisar) {
    if (!st) return false;
    if (ocupado) { pendente = true; return false; }
    clearTimeout(timer); timer = null; ocupado = true; pendente = false; ultimo = Date.now();
    status('Salvando na nuvem…');
    try {
      const ids = await o.idsUsados();
      const r = await post({ acao: 'salvar', codigo: st.codigo, senha: st.senha, base: st.v, forcar: !!forcar,
        trabalho: JSON.stringify(o.trabalho()), ids, nome: o.nome() });
      if (r.conflito && !forcar && Number(r.v) < Number(st.v)) { ocupado = false; return salvarAgora(true, avisar); } // envio ao fechar a página que não chegou
      if (r.conflito) { ocupado = false; return conflito(r); }
      if (!r.ok) throw Object.assign(new Error(r.erro || 'Erro na nuvem.'), { servidor: true });
      st.v = r.v; st.t = r.t; gravar();
      if (r.faltam && r.faltam.length) await enviarImagens(r.faltam, ids);
      ok(st.t);
      if (avisar) o.toast('Salvo na nuvem ✓');
      return true;
    } catch (err) {
      console.error(err);
      if (err.servidor && /Senha errada|Código não encontrado/.test(err.message)) { await perdeuAcesso(err.message); return false; }
      status('Sem conexão com a nuvem: salvo só neste computador', 'var(--bad)');
      if (avisar) o.toast('Não consegui salvar na nuvem. Verifiquem a internet.');
      clearTimeout(timer); timer = setTimeout(() => salvarAgora(false), 60000);
      return false;
    } finally {
      ocupado = false;
      if (pendente) mudou();
    }
  }

  async function enviarImagens(faltam, usados) {
    let lote = {}, tam = 0, n = 0;
    const manda = async () => {
      if (!Object.keys(lote).length) return;
      status('Enviando fotos para a nuvem… ' + n + '/' + faltam.length);
      const r = await post({ acao: 'imagens', codigo: st.codigo, senha: st.senha, imagens: lote, usados });
      if (!r.ok) throw Object.assign(new Error(r.erro || 'Erro ao enviar as fotos.'), { servidor: true });
      lote = {}; tam = 0;
    };
    for (const id of faltam) {
      const v = await o.lerImagem(id); if (!v) continue;
      lote[id] = v; tam += v.b64.length; n++;
      if (tam > LOTE_IMG) await manda();
    }
    await manda();
  }

  async function conflito(r) {
    clearTimeout(timer); pausado = true;
    const esc = o.esc;
    const v = await o.dialogo(`<p><b>Outro computador salvou o trabalho da equipe na nuvem</b> (${esc(hora(r.t))}).</p>
      <p class="small">Se continuarem aqui, podem apagar o que foi feito lá. Qual versão vale?</p>`,
      [['Decidir depois', false], ['Ficar com a deste computador', 'aqui'], ['Abrir a versão da nuvem', 'nuvem', 'primary']]);
    pausado = false;
    if (v === 'nuvem') return baixarDaNuvem(false);
    if (v === 'aqui') return salvarAgora(true, true);
    pausado = true;
    status('Nuvem em pausa: há duas versões. Clique no botão da nuvem para decidir.', 'var(--bad)');
    return false;
  }

  async function perdeuAcesso(msg) {
    st = null; gravar(); botao();
    status('Salvo neste navegador');
    await o.dialogo(`<p><b>Este computador saiu da nuvem da equipe.</b></p><p class="small">${o.esc(msg)} Entrem de novo pelo botão ☁ Nuvem.</p>`, [['Ok', true, 'primary']]);
  }

  /* ---------- abrir ---------- */
  async function baixarDaNuvem(perguntar) {
    if (perguntar && !o.vazio()) {
      const s = await o.dialogo('<p><b>Abrir o trabalho da nuvem?</b></p><p class="small">O que está agora neste computador será substituído pela versão da nuvem.</p>',
        [['Cancelar', false], ['Abrir', true, 'primary']]);
      if (!s) return false;
    }
    status('Abrindo da nuvem…');
    try {
      const r = await post({ acao: 'abrir', codigo: st.codigo, senha: st.senha });
      if (!r.ok) throw Object.assign(new Error(r.erro), { servidor: true });
      if (r.trabalho) { await o.aplicar(JSON.parse(r.trabalho), r.imagens || {}); clearTimeout(timer); timer = null; }
      st.v = r.v; st.t = r.t; st.uso = Date.now(); gravar(); pausado = false;
      if (!r.trabalho) await salvarAgora(true);
      ok(st.t); o.toast('Trabalho aberto da nuvem ✓');
      return true;
    } catch (err) {
      console.error(err);
      if (err.servidor && /Senha errada|Código não encontrado/.test(err.message)) { await perdeuAcesso(err.message); return false; }
      status('Não consegui abrir da nuvem', 'var(--bad)');
      await o.dialogo('<p><b>Não consegui abrir o trabalho da nuvem.</b></p><p class="small">Verifiquem a internet e tentem de novo.</p>', [['Ok', true, 'primary']]);
      return false;
    }
  }

  /* ---------- janela da nuvem ---------- */
  async function janela() {
    const esc = o.esc;
    if (st) {
      const v = await o.dialogo(`<p><b>☁ Trabalho guardado na nuvem</b></p>
        <p style="font-size:1.6rem;font-weight:700;letter-spacing:.08em;margin:6px 0">${esc(st.codigo)}</p>
        <p class="small">${st.t ? 'Último salvamento: ' + esc(hora(st.t)) + '.' : 'Ainda não salvo.'} O trabalho é salvo sozinho enquanto vocês escrevem.</p>
        <p class="small">Em outro computador: <b>☁ Nuvem → Já temos código</b>, com esse código e a senha da equipe.</p>
        <p class="small">No fim da aula, num computador da escola, cliquem em <b>Sair deste computador</b>.</p>`,
        [['Fechar', false], ['Sair deste computador', 'sair', 'danger'], ['Abrir a versão da nuvem', 'abrir'], ['Salvar agora', 'salvar', 'primary']]);
      if (v === 'salvar') { pausado = false; return salvarAgora(false, true); }
      if (v === 'abrir') return baixarDaNuvem(true);
      if (v === 'sair') return sair();
      return;
    }
    const v = await o.dialogo(`<p><b>☁ Guardar o trabalho na nuvem</b></p>
      <p class="small">Com a nuvem, a equipe abre o trabalho em <b>qualquer computador</b>, só com o código da equipe e uma senha. Não precisa de e-mail.</p>
      <p class="small">Uma pessoa da equipe cria o código <b>uma vez só</b>. Os outros usam “Já temos código”.</p>`,
      [['Cancelar', false], ['Já temos código', 'entrar'], ['Criar código da equipe', 'criar', 'primary']]);
    if (v === 'criar') return criar();
    if (v === 'entrar') return entrar();
  }

  async function criar() {
    const v = await o.dialogo(`<p><b>Criar o código da equipe</b></p>
      <p class="small">Escolham uma senha que todos da equipe vão lembrar (pelo menos 4 letras ou números). Não usem a senha do e-mail.</p>
      <label class="small">Senha da equipe<br><input id="nvSenha" type="text" autocomplete="off" style="width:100%;margin-top:4px"></label>`,
      [['Cancelar', false], ['Criar', true, 'primary']]);
    if (!v) return;
    const senha = (document.getElementById('nvSenha').value || '').trim();
    if (senha.length < 4) { await o.dialogo('<p>A senha precisa ter pelo menos 4 letras ou números.</p>', [['Ok', true, 'primary']]); return criar(); }
    status('Criando o código…');
    let r;
    try { r = await post({ acao: 'criar', senha, nome: o.nome() }); } catch (e) { r = null; }
    if (!r || !r.ok) {
      status('Salvo neste navegador');
      await o.dialogo(`<p><b>Não consegui falar com a nuvem.</b></p><p class="small">${o.esc((r && r.erro) || 'Verifiquem a internet. Se a internet da escola bloquear, usem “Salvar cópia”.')}</p>`, [['Ok', true, 'primary']]);
      return;
    }
    st = { codigo: r.codigo, senha, v: r.v, t: '', uso: Date.now() }; gravar(); botao();
    await salvarAgora(true);
    await o.dialogo(`<p><b>Pronto! Anotem no caderno:</b></p>
      <p style="font-size:2rem;font-weight:700;letter-spacing:.08em;margin:6px 0;text-align:center">${o.esc(r.codigo)}</p>
      <p style="text-align:center">senha: <b>${o.esc(senha)}</b></p>
      <p class="small">Em outro computador, abram o assistente, cliquem em <b>☁ Nuvem → Já temos código</b> e digitem o código e a senha.</p>`,
      [['Anotei', true, 'primary']]);
  }

  async function entrar() {
    const v = await o.dialogo(`<p><b>Abrir o trabalho da equipe</b></p>
      <label class="small">Código da equipe<br><input id="nvCod" type="text" autocomplete="off" placeholder="${o.prefixo}-XXXX" style="width:100%;margin-top:4px;text-transform:uppercase"></label>
      <label class="small" style="display:block;margin-top:10px">Senha da equipe<br><input id="nvSenha" type="text" autocomplete="off" style="width:100%;margin-top:4px"></label>`,
      [['Cancelar', false], ['Abrir', true, 'primary']]);
    if (!v) return;
    const codigo = (document.getElementById('nvCod').value || '').trim();
    const senha = (document.getElementById('nvSenha').value || '').trim();
    if (!codigo || !senha) return entrar();
    st = { codigo: codigo.toUpperCase(), senha, v: 0, t: '', uso: Date.now() };
    status('Conferindo o código…');
    let r;
    try { r = await post({ acao: 'info', codigo: st.codigo, senha }); } catch (e) { r = null; }
    if (!r || !r.ok) {
      st = null; status('Salvo neste navegador');
      await o.dialogo(`<p><b>Não deu certo.</b></p><p class="small">${o.esc((r && r.erro) || 'Não consegui falar com a nuvem. Verifiquem a internet.')}</p>`, [['Ok', true, 'primary']]);
      return;
    }
    const cod = st.codigo.replace(/[^A-Z0-9]/g, '');
    st.codigo = cod.indexOf(o.prefixo) === 0 ? o.prefixo + '-' + cod.slice(o.prefixo.length) : o.prefixo + '-' + cod;
    const abriu = await baixarDaNuvem(true);
    if (!abriu) { if (st) { st = null; gravar(); } status('Salvo neste navegador'); }
    botao();
  }

  async function sair() {
    const v = await o.dialogo(`<p><b>Sair da nuvem neste computador?</b></p>
      <p class="small">O trabalho continua guardado na nuvem com o código ${o.esc(st.codigo)}.</p>
      <p class="small">Em computador da escola, é melhor <b>apagar o trabalho deste computador</b>, para outra turma não mexer nele.</p>`,
      [['Cancelar', false], ['Só sair', 'sair'], ['Sair e apagar deste computador', 'apagar', 'danger']]);
    if (!v) return;
    status('Salvando na nuvem…');
    const salvou = pausado ? true : await salvarAgora(false);
    if (!salvou) {
      const s = await o.dialogo('<p><b>A última versão não foi para a nuvem.</b></p><p class="small">Se saírem agora, as mudanças mais recentes ficam só neste computador.</p>',
        [['Cancelar', false], ['Sair mesmo assim', true, 'danger']]);
      if (!s) return;
    }
    st = null; gravar(); botao(); clearTimeout(timer); pausado = false;
    if (v === 'apagar') await o.limpar();
    status('Salvo neste navegador');
    o.toast(v === 'apagar' ? 'Saiu da nuvem e apagou deste computador ✓' : 'Saiu da nuvem ✓');
  }

  /* ---------- ao abrir a página ---------- */
  async function aoAbrir() {
    if (!st) return;
    const parado = Date.now() - (st.uso || 0) > 30 * 60 * 1000;
    if (parado) {
      const v = await o.dialogo(`<p><b>Este computador está ligado à nuvem da equipe ${o.esc(st.codigo)}.</b></p>
        <p class="small">${o.esc(o.nome() || '')}</p><p class="small">É a sua equipe?</p>`,
        [['Não, sair', 'nao', 'danger'], ['Sim, continuar', 'sim', 'primary']]);
      if (v === 'nao') {
        const a = await o.dialogo('<p><b>Apagar deste computador o trabalho dessa equipe?</b></p><p class="small">Ele continua guardado na nuvem.</p>',
          [['Não apagar', false], ['Apagar', true, 'danger']]);
        st = null; gravar(); botao(); status('Salvo neste navegador');
        if (a) await o.limpar();
        return;
      }
    }
    st.uso = Date.now(); gravar();
    status('Conferindo a nuvem…');
    let r;
    try { r = await post({ acao: 'info', codigo: st.codigo, senha: st.senha }); } catch (e) { r = null; }
    if (!r) { status('Sem conexão com a nuvem: salvo só neste computador', 'var(--bad)'); return; }
    if (!r.ok) { if (/Senha errada|Código não encontrado/.test(r.erro || '')) return perdeuAcesso(r.erro); return; }
    if (Number(r.v) > Number(st.v)) {
      const v = await o.dialogo(`<p><b>Há uma versão mais nova do trabalho na nuvem</b> (${o.esc(hora(r.t))}), salva em outro computador.</p>`,
        [['Ficar com a deste computador', 'aqui'], ['Abrir a versão da nuvem', 'nuvem', 'primary']]);
      if (v === 'nuvem') return baixarDaNuvem(false);
      st.v = r.v; gravar(); return salvarAgora(true);
    }
    ok(r.t || st.t);
  }

  function aoEsconder() { if (st && !pausado && timer && !ocupado) salvarAgora(false); }
  function aoSair() {
    // tenta mandar o texto (sem as fotos) se ficou alguma mudança sem enviar
    if (!st || pausado || !timer || ocupado) return;
    try {
      const corpo = JSON.stringify({ app: o.app, acao: 'salvar', codigo: st.codigo, senha: st.senha, base: st.v, trabalho: JSON.stringify(o.trabalho()), ids: [], nome: o.nome() });
      if (corpo.length < 60000 && navigator.sendBeacon) { navigator.sendBeacon(o.url(), new Blob([corpo], { type: 'text/plain' })); st.v = Number(st.v) + 1; gravar(); clearTimeout(timer); timer = null; }
    } catch (e) {}
  }

  window.NuvemTCC = {
    iniciar(opcoes) {
      o = opcoes; st = ler(); botao();
      const b = o.botao(); if (b) b.onclick = janela;
      if (!o.url()) { if (b) b.hidden = true; return; }
      document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'hidden') aoEsconder(); });
      window.addEventListener('pagehide', aoSair);
      aoAbrir();
    },
    mudou() { if (o) mudou(); },
    ligado: () => !!st,
  };
})();
