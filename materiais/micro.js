/* Fazenda de Projetos · Tecnologia dos Materiais
   Gerador de microestruturas (desenhadas pelo computador, não são fotos).
   Usado pelo Microscópio Virtual e pelo Forno Virtual.
   MICRO.gerar(tipo, {N, seed, atacado, ...}) devolve um canvas N×N.
   Escala: 1 px do canvas = UM_PX micrômetros. */
const MICRO = (() => {
  const UM_PX = 0.55;

  function rng(seed) {
    let s = (seed >>> 0) || 123456789;
    return () => { s ^= s << 13; s >>>= 0; s ^= s >>> 17; s ^= s << 5; s >>>= 0; return s / 4294967296; };
  }

  /* Grãos: diagrama de Voronoi com sementes numa grade sorteada */
  function voronoi(N, cel, R) {
    const g = Math.ceil(N / cel), sx = new Float32Array(g * g), sy = new Float32Array(g * g);
    for (let j = 0; j < g; j++) for (let i = 0; i < g; i++) {
      const k = j * g + i; sx[k] = (i + 0.05 + 0.9 * R()) * cel; sy[k] = (j + 0.05 + 0.9 * R()) * cel;
    }
    const id = new Int32Array(N * N);
    for (let y = 0; y < N; y++) {
      const cy = Math.floor(y / cel);
      for (let x = 0; x < N; x++) {
        const cx = Math.floor(x / cel);
        let best = 1e18, bi = 0;
        for (let dj = -1; dj <= 1; dj++) {
          const jj = cy + dj; if (jj < 0 || jj >= g) continue;
          for (let di = -1; di <= 1; di++) {
            const ii = cx + di; if (ii < 0 || ii >= g) continue;
            const k = jj * g + ii, dx = sx[k] - x, dy = sy[k] - y, d = dx * dx + dy * dy;
            if (d < best) { best = d; bi = k; }
          }
        }
        id[y * N + x] = bi;
      }
    }
    return { id, n: g * g };
  }

  function contornos(id, N) {
    const b = new Uint8Array(N * N);
    for (let y = 0; y < N - 1; y++) for (let x = 0; x < N - 1; x++) {
      const i = y * N + x;
      if (id[i] !== id[i + 1]) { b[i] = 1; b[i + 1] = 1; }
      if (id[i] !== id[i + N]) { b[i] = 1; b[i + N] = 1; }
    }
    return b;
  }

  /* Engorda uma máscara em r pixels (para a rede de cementita) */
  function dilatar(m, N, r) {
    const t = new Uint8Array(N * N), o = new Uint8Array(N * N);
    for (let y = 0; y < N; y++) { let c = 0; const row = y * N;
      for (let x = -r; x < N; x++) { if (x + r < N && m[row + x + r]) c++; if (x - r - 1 >= 0 && m[row + x - r - 1]) c--; if (x >= 0) t[row + x] = c > 0 ? 1 : 0; } }
    for (let x = 0; x < N; x++) { let c = 0;
      for (let y = -r; y < N; y++) { if (y + r < N && t[(y + r) * N + x]) c++; if (y - r - 1 >= 0 && t[(y - r - 1) * N + x]) c--; if (y >= 0) o[y * N + x] = c > 0 ? 1 : 0; } }
    return o;
  }

  const clamp = v => v < 0 ? 0 : v > 255 ? 255 : v;

  function novoCanvas(N) {
    const c = document.createElement('canvas'); c.width = c.height = N; return c;
  }

  /* Superfície polida sem ataque: lisa, com riscos e inclusões */
  function semAtaque(ctx, N, R, cor) {
    const im = ctx.createImageData(N, N), d = im.data;
    for (let i = 0; i < N * N; i++) { const n = (R() - 0.5) * 10;
      d[i * 4] = clamp(cor[0] + n); d[i * 4 + 1] = clamp(cor[1] + n); d[i * 4 + 2] = clamp(cor[2] + n); d[i * 4 + 3] = 255; }
    ctx.putImageData(im, 0, 0);
    riscos(ctx, N, R, 0.5); inclusoes(ctx, N, R);
  }
  function riscos(ctx, N, R, alfa) {
    ctx.save(); ctx.globalAlpha = 0.18 * alfa; ctx.strokeStyle = '#6b6b66'; ctx.lineWidth = 0.8;
    const ang = R() * Math.PI;
    for (let k = 0; k < 22; k++) { const x = R() * N, y = R() * N, a = ang + (R() - 0.5) * 0.25, L = 200 + R() * 700;
      ctx.beginPath(); ctx.moveTo(x - Math.cos(a) * L / 2, y - Math.sin(a) * L / 2); ctx.lineTo(x + Math.cos(a) * L / 2, y + Math.sin(a) * L / 2); ctx.stroke(); }
    ctx.restore();
  }
  function inclusoes(ctx, N, R) {
    ctx.save(); ctx.fillStyle = '#4a4744';
    const n = Math.round(N * N / 90000);
    for (let k = 0; k < n; k++) { const x = R() * N, y = R() * N, w = 2 + R() * 7, h = 1 + R() * 1.6;
      ctx.beginPath(); ctx.ellipse(x, y, w, h, 0.1, 0, Math.PI * 2); ctx.fill(); }
    ctx.restore();
  }

  /* ---------- Aços: ferrita + perlita, perlita, perlita + cementita ---------- */
  function aco(ctx, N, R, o) {
    const { id, n } = voronoi(N, o.grao, R), b = contornos(id, N);
    const tipo = new Uint8Array(n), ang = new Float32Array(n), lam = new Float32Array(n), ph = new Float32Array(n), tom = new Float32Array(n);
    for (let k = 0; k < n; k++) { tipo[k] = R() < o.pf ? 1 : 0; ang[k] = R() * Math.PI; lam[k] = o.lam * (0.75 + 0.5 * R()); ph[k] = R(); tom[k] = (R() - 0.5) * 14; }
    const cem = o.cem ? dilatar(b, N, o.cem) : null;
    const im = ctx.createImageData(N, N), d = im.data;
    for (let y = 0; y < N; y++) for (let x = 0; x < N; x++) {
      const i = y * N + x, g = id[i]; let r, gg, bb;
      if (cem && cem[i]) { r = 244; gg = 241; bb = 232; }
      else if (tipo[g]) {
        const c = Math.cos(ang[g]), s = Math.sin(ang[g]);
        let p = ((x * c + y * s) / lam[g] + ph[g]) % 1; if (p < 0) p += 1;
        if (p < 0.34) { r = 52; gg = 45; bb = 40; } else { r = 158; gg = 148; bb = 132; }
        r += tom[g]; gg += tom[g]; bb += tom[g];
      } else { r = 234 + tom[g] * 0.5; gg = 231 + tom[g] * 0.5; bb = 219 + tom[g] * 0.5; }
      if (b[i] && !(cem && cem[i] && !b[i])) { if (!(cem && cem[i])) { r *= 0.42; gg *= 0.42; bb *= 0.42; } }
      const nn = (R() - 0.5) * 12;
      d[i * 4] = clamp(r + nn); d[i * 4 + 1] = clamp(gg + nn); d[i * 4 + 2] = clamp(bb + nn); d[i * 4 + 3] = 255;
    }
    if (cem) { // contorno escuro fininho nas bordas da cementita
      for (let y = 1; y < N - 1; y++) for (let x = 1; x < N - 1; x++) { const i = y * N + x;
        if (cem[i] && (!cem[i - 1] || !cem[i + 1] || !cem[i - N] || !cem[i + N])) { d[i * 4] *= 0.45; d[i * 4 + 1] *= 0.45; d[i * 4 + 2] *= 0.45; } }
    }
    ctx.putImageData(im, 0, 0);
    riscos(ctx, N, R, 0.15); inclusoes(ctx, N, R);
  }

  /* ---------- Martensita (com ferrita, bainita e revenimento opcionais) ---------- */
  function martensita(ctx, N, R, o) {
    const { id, n } = voronoi(N, o.grao || 140, R), b = contornos(id, N);
    const tipo = new Uint8Array(n), angs = [];
    for (let k = 0; k < n; k++) {
      const u = R(); tipo[k] = u < (o.ferr || 0) ? 1 : u < (o.ferr || 0) + (o.bain || 0) ? 2 : u < (o.ferr || 0) + (o.bain || 0) + (o.perl || 0) ? 3 : 0;
      const a0 = R() * Math.PI; angs.push([a0, a0 + Math.PI / 3, a0 + 2 * Math.PI / 3]);
    }
    const rev = o.rev || 0;
    const im = ctx.createImageData(N, N), d = im.data;
    for (let i = 0; i < N * N; i++) { const t = tipo[id[i]], nn = (R() - 0.5) * 14;
      let c = t === 1 ? [232, 229, 216] : t === 2 ? [128, 116, 100] : t === 3 ? [96, 86, 76] : [168 - 50 * rev, 158 - 50 * rev, 144 - 46 * rev];
      if (b[i] && t === 1) c = c.map(v => v * 0.45);
      d[i * 4] = clamp(c[0] + nn); d[i * 4 + 1] = clamp(c[1] + nn); d[i * 4 + 2] = clamp(c[2] + nn); d[i * 4 + 3] = 255; }
    ctx.putImageData(im, 0, 0);
    // agulhas: crescem até bater no contorno do grão de austenita
    const qtd = Math.round(N * N / 170);
    ctx.lineCap = 'round';
    for (let k = 0; k < qtd; k++) {
      const x0 = R() * N, y0 = R() * N, g = id[(y0 | 0) * N + (x0 | 0)];
      if (tipo[g] === 1) continue;
      const bain = tipo[g] === 2 || tipo[g] === 3;
      const a = angs[g][(R() * 3) | 0] + (R() - 0.5) * 0.35, c = Math.cos(a), s = Math.sin(a);
      const Lmax = bain ? 8 + R() * 22 : 4 + Math.pow(R(), 2) * 55;
      const fim = dir => { let L = 0; while (L < Lmax) { const xx = x0 + dir * c * L, yy = y0 + dir * s * L;
        if (xx < 0 || yy < 0 || xx >= N || yy >= N || id[(yy | 0) * N + (xx | 0)] !== g) break; L += 1.5; } return L; };
      const L1 = fim(1), L2 = fim(-1);
      ctx.strokeStyle = bain ? `rgba(40,32,26,${0.5 + R() * 0.4})` : (R() < 0.5 ? `rgba(70,62,54,${0.55 + R() * 0.4 - rev * 0.2})` : `rgba(220,214,200,${0.45 + R() * 0.4 - rev * 0.3})`);
      ctx.lineWidth = bain ? 1 + R() : 0.8 + R() * 2.2;
      ctx.beginPath(); ctx.moveTo(x0 - c * L2, y0 - s * L2); ctx.lineTo(x0 + c * L1, y0 + s * L1); ctx.stroke();
    }
    if (rev > 0) { // carbonetos finos do revenimento: pontinhos escuros
      ctx.fillStyle = 'rgba(30,24,20,0.55)';
      const m = Math.round(N * N / 120 * rev);
      for (let k = 0; k < m; k++) { const x = R() * N, y = R() * N; if (tipo[id[(y | 0) * N + (x | 0)]] === 1) continue; ctx.fillRect(x, y, 1.2, 1.2); }
    }
    riscos(ctx, N, R, 0.15); inclusoes(ctx, N, R);
  }

  /* ---------- Esferoidita: bolinhas de cementita na ferrita ---------- */
  function esferoidita(ctx, N, R, o) {
    aco(ctx, N, R, { grao: o.grao || 95, pf: 0, lam: 4 });
    const m = Math.round(N * N / 420 * (o.c || 0.8) / 0.8);
    for (let k = 0; k < m; k++) { const x = R() * N, y = R() * N, r = 1.2 + R() * 2.6;
      ctx.beginPath(); ctx.arc(x, y, r + 0.9, 0, 7); ctx.fillStyle = 'rgba(60,52,46,0.85)'; ctx.fill();
      ctx.beginPath(); ctx.arc(x, y, r, 0, 7); ctx.fillStyle = '#f3f0e6'; ctx.fill(); }
  }

  /* ---------- Ferros fundidos ---------- */
  function veios(ctx, N, R) {
    ctx.strokeStyle = '#15120f'; ctx.lineCap = 'round'; ctx.lineJoin = 'round';
    const qtd = Math.round(N * N / 9000);
    for (let k = 0; k < qtd; k++) {
      let x = R() * N, y = R() * N, a = R() * Math.PI * 2; const seg = 4 + (R() * 6 | 0), passo = 12 + R() * 30;
      ctx.lineWidth = 2 + R() * 4; ctx.beginPath(); ctx.moveTo(x, y);
      for (let s = 0; s < seg; s++) { a += (R() - 0.5) * 0.9; x += Math.cos(a) * passo; y += Math.sin(a) * passo; ctx.lineTo(x, y); }
      ctx.stroke();
    }
  }
  function nodulos(ctx, N, R, halo) {
    const qtd = Math.round(N * N / 11000), L = [];
    for (let k = 0, t = 0; k < qtd && t < qtd * 30; t++) {
      const x = R() * N, y = R() * N, r = 7 + R() * 13;
      if (L.some(([a, b, q]) => (a - x) ** 2 + (b - y) ** 2 < ((r + q) * 2.3) ** 2)) continue;
      L.push([x, y, r]); k++;
    }
    if (halo) L.forEach(([x, y, r]) => { ctx.beginPath(); ctx.arc(x, y, r * 2.0, 0, 7); ctx.fillStyle = '#ece9de'; ctx.fill();
      ctx.lineWidth = 1.2; ctx.strokeStyle = 'rgba(60,52,46,0.6)'; ctx.stroke(); });
    L.forEach(([x, y, r]) => { const gr = ctx.createRadialGradient(x - r * 0.3, y - r * 0.3, r * 0.1, x, y, r);
      gr.addColorStop(0, '#3a3532'); gr.addColorStop(1, '#0d0b0a'); ctx.beginPath(); ctx.arc(x, y, r, 0, 7); ctx.fillStyle = gr; ctx.fill(); });
  }
  function branco(ctx, N, R, atacado) {
    ctx.fillStyle = '#efebe0'; ctx.fillRect(0, 0, N, N);
    if (!atacado) { semAtaque(ctx, N, R, [236, 233, 224]); return; }
    // ledeburita: pontinhos escuros em áreas brancas
    ctx.fillStyle = 'rgba(70,60,52,0.75)';
    for (let k = 0; k < N * N / 55; k++) { const x = R() * N, y = R() * N; if (((x / 140 | 0) + (y / 170 | 0)) % 3 === 0 && R() < 0.6) continue; ctx.fillRect(x, y, 1.6, 1.6); }
    // dendritas de perlita (escuras)
    ctx.fillStyle = '#4c4238';
    const qtd = Math.round(N * N / 26000);
    for (let k = 0; k < qtd; k++) {
      let x = R() * N, y = R() * N; const a = R() * Math.PI, c = Math.cos(a), s = Math.sin(a), L = 60 + R() * 160;
      for (let t = -L / 2; t < L / 2; t += 9) {
        const xx = x + c * t, yy = y + s * t, w = 6 + R() * 7;
        ctx.beginPath(); ctx.ellipse(xx, yy, w, w * 0.7, a, 0, 7); ctx.fill();
        if (R() < 0.6) { const l2 = 12 + R() * 26, sg = R() < 0.5 ? 1 : -1;
          ctx.beginPath(); ctx.ellipse(xx - s * sg * l2 / 2, yy + c * sg * l2 / 2, l2 / 2, 3.5 + R() * 3, a + Math.PI / 2, 0, 7); ctx.fill(); }
      }
    }
    riscos(ctx, N, R, 0.1);
  }

  /* ---------- Cobre e latão: grãos coloridos com maclas ---------- */
  function naoFerroso(ctx, N, R, pal) {
    const { id, n } = voronoi(N, 115, R), b = contornos(id, N);
    const c1 = [], c2 = [], tw = [];
    for (let k = 0; k < n; k++) {
      c1.push(pal[(R() * pal.length) | 0]); c2.push(pal[(R() * pal.length) | 0]);
      tw.push(R() < 0.6 ? [R() * Math.PI, 28 + R() * 60, 5 + R() * 16, R()] : null);
    }
    const im = ctx.createImageData(N, N), d = im.data;
    for (let y = 0; y < N; y++) for (let x = 0; x < N; x++) {
      const i = y * N + x, g = id[i]; let c = c1[g];
      const t = tw[g];
      if (t) { let p = ((x * Math.cos(t[0]) + y * Math.sin(t[0])) / t[1] + t[3]) % 1; if (p < 0) p += 1; if (p < t[2] / t[1]) c = c2[g] === c1[g] ? c.map(v => v * 0.82) : c2[g]; }
      let f = b[i] ? 0.4 : 1; const nn = (R() - 0.5) * 9;
      d[i * 4] = clamp(c[0] * f + nn); d[i * 4 + 1] = clamp(c[1] * f + nn); d[i * 4 + 2] = clamp(c[2] * f + nn); d[i * 4 + 3] = 255;
    }
    ctx.putImageData(im, 0, 0); riscos(ctx, N, R, 0.12);
  }
  const PAL_COBRE = [[214, 140, 102], [190, 112, 80], [232, 168, 128], [160, 92, 66], [205, 128, 96], [240, 186, 150]];
  const PAL_LATAO = [[222, 190, 110], [196, 160, 84], [236, 210, 140], [170, 136, 70], [210, 176, 98], [244, 224, 164]];

  /* ---------- catálogo ---------- */
  function gerar(tipo, o = {}) {
    const N = o.N || 1600, R = rng(o.seed || 7), cv = novoCanvas(N), ctx = cv.getContext('2d');
    const at = o.atacado !== false;
    const aco_ = p => at ? aco(ctx, N, R, p) : semAtaque(ctx, N, R, [226, 225, 219]);
    switch (tipo) {
      case 'ferrita': aco_({ grao: 115, pf: 0, lam: 4 }); break;
      case 'fp': aco_({ grao: o.grao || 100, pf: o.pf, lam: o.lam || 4 }); break;
      case 'perlita': aco_({ grao: o.grao || 110, pf: 1, lam: o.lam || 4 }); break;
      case 'perlita-cementita': aco_({ grao: o.grao || 115, pf: 1, lam: o.lam || 4, cem: o.cem || 3 }); break;
      case 'martensita': at ? martensita(ctx, N, R, o) : semAtaque(ctx, N, R, [226, 225, 219]); break;
      case 'esferoidita': at ? esferoidita(ctx, N, R, o) : semAtaque(ctx, N, R, [226, 225, 219]); break;
      case 'cinzento':
        if (at) aco(ctx, N, R, { grao: 90, pf: 0.93, lam: 3.5 }); else semAtaque(ctx, N, R, [228, 227, 221]);
        veios(ctx, N, R); break;
      case 'nodular':
        if (at) aco(ctx, N, R, { grao: 85, pf: 0.95, lam: 3.5 }); else semAtaque(ctx, N, R, [228, 227, 221]);
        nodulos(ctx, N, R, at); break;
      case 'branco': branco(ctx, N, R, at); break;
      case 'cobre': at ? naoFerroso(ctx, N, R, PAL_COBRE) : semAtaque(ctx, N, R, [226, 156, 118]); break;
      case 'latao': at ? naoFerroso(ctx, N, R, PAL_LATAO) : semAtaque(ctx, N, R, [232, 204, 130]); break;
    }
    return cv;
  }

  /* Pirâmide de imagens (1/1, 1/2, 1/4) para reduzir sem formar listras falsas */
  function piramide(cv) {
    const L = [cv]; let src = cv;
    for (let k = 0; k < 2; k++) {
      const N = src.width, M = N >> 1, a = src.getContext('2d').getImageData(0, 0, N, N).data;
      const c = novoCanvas(M), cx = c.getContext('2d'), im = cx.createImageData(M, M), d = im.data;
      for (let y = 0; y < M; y++) for (let x = 0; x < M; x++) {
        const i0 = (2 * y * N + 2 * x) * 4, i1 = i0 + 4, i2 = i0 + N * 4, i3 = i2 + 4, o = (y * M + x) * 4;
        for (let c2 = 0; c2 < 3; c2++) d[o + c2] = (a[i0 + c2] + a[i1 + c2] + a[i2 + c2] + a[i3 + c2]) >> 2;
        d[o + 3] = 255;
      }
      cx.putImageData(im, 0, 0); L.push(c); src = c;
    }
    return L;
  }
  /* Desenha a região (sx, sy, lado fw em px da imagem original) num retângulo de lado D */
  function desenharRegiao(ctx, pir, sx, sy, fw, x, y, D) {
    let k = 0; while (k < pir.length - 1 && fw / (1 << (k + 1)) >= D * 0.9) k++;
    const f = 1 << k;
    ctx.imageSmoothingEnabled = true; ctx.imageSmoothingQuality = 'high';
    ctx.drawImage(pir[k], sx / f, sy / f, fw / f, fw / f, x, y, D, D);
  }

  return { gerar, piramide, desenharRegiao, UM_PX };
})();
