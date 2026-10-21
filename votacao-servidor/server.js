// Servidor da votação do juramentista — sem dependências, roda com "node server.js"
const http = require('http');
const fs = require('fs');
const path = require('path');

const PORT = process.env.PORT || 3000;
const SENHA = process.env.SENHA || 'jur-54f039';
const CANDIDATAS = ['Sibele', 'Ana Roik', 'Mariane', 'Sophia', 'Nicolly', 'Yasmim'];
const ARQ = path.join(__dirname, 'votos.json');

let estado = { aberta: false, votos: [] };
try { estado = JSON.parse(fs.readFileSync(ARQ, 'utf8')); } catch (e) {}
function salvar() { try { fs.writeFileSync(ARQ, JSON.stringify(estado)); } catch (e) {} }

function json(res, obj) {
  res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store', 'Access-Control-Allow-Origin': '*' });
  res.end(JSON.stringify(obj));
}
function arquivo(res, nome) {
  fs.readFile(path.join(__dirname, 'public', nome), (err, dados) => {
    if (err) { res.writeHead(404); res.end('Página não encontrada'); return; }
    res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8', 'Cache-Control': 'no-store' });
    res.end(dados);
  });
}

http.createServer((req, res) => {
  const url = new URL(req.url, 'http://x');
  if (url.pathname === '/api') {
    if (req.method === 'GET') {
      const p = Object.fromEntries(url.searchParams);
      if (p.acao === 'status') return json(res, { aberta: estado.aberta, candidatas: CANDIDATAS });
      if (p.senha !== SENHA) return json(res, { erro: 'senha' });
      if (p.acao === 'abrir') estado.aberta = true;
      if (p.acao === 'fechar') estado.aberta = false;
      if (p.acao === 'zerar') estado.votos = [];
      if (p.acao) salvar();
      return json(res, { aberta: estado.aberta, candidatas: CANDIDATAS, votos: estado.votos });
    }
    if (req.method === 'POST') {
      let corpo = '';
      req.on('data', c => { corpo += c; if (corpo.length > 5000) req.destroy(); });
      req.on('end', () => {
        if (!estado.aberta) return json(res, { ok: false, motivo: 'fechada' });
        let d = {};
        try { d = JSON.parse(corpo); } catch (e) { return json(res, { ok: false, motivo: 'invalido' }); }
        const aparelho = String(d.aparelho || '').slice(0, 40);
        if (aparelho && estado.votos.some(v => v.aparelho === aparelho)) return json(res, { ok: false, motivo: 'repetido' });
        const notas = CANDIDATAS.map((_, i) => {
          const n = Math.round(Number(d.notas && d.notas[i]));
          return n >= 1 && n <= 5 ? n : '';
        });
        estado.votos.push({ data: new Date().toISOString(), aparelho, notas });
        salvar();
        json(res, { ok: true });
      });
      return;
    }
  }
  if (url.pathname === '/' || url.pathname === '/index.html') return arquivo(res, 'index.html');
  if (url.pathname === '/resultados' || url.pathname === '/resultados.html') return arquivo(res, 'resultados.html');
  res.writeHead(404); res.end('Página não encontrada');
}).listen(PORT, () => console.log('Votação no ar na porta ' + PORT));
