# Fazenda de Projetos

Site de ensino técnico do Prof. Rodrigo Leandro Martins: simuladores, jogos, atividades e avaliações, com acesso livre e sem cadastro.

O site começa por Metrologia. Desenho Mecânico entra em seguida.

## Como o site está organizado

| Arquivo | O que é |
|---|---|
| `index.html` | Página principal: início, disciplinas, recursos, professores, sobre e contato |
| `metrologia/paquimetro-partes.html` | Jogo "Ache a peça": as 14 partes do paquímetro |
| `metrologia/paquimetro-nonio-10.html` | Simulador de leitura, nônio de 10 divisões (0,1 mm) |
| `metrologia/paquimetro-nonio-20.html` | Simulador de leitura, nônio de 20 divisões (0,05 mm) |
| `metrologia/paquimetro-nonio-50.html` | Simulador de leitura, nônio de 50 divisões (0,02 mm) |
| `metrologia/paquimetro-jogo-fases.html` | Jogo "Domine o Paquímetro", em 5 fases |
| `metrologia/aula-pratica-medicao.html` | Aula prática: folha de medição com régua e paquímetro, sem nota |
| `metrologia/paquimetro-avaliacao.html` | Avaliação de paquímetro: 60 leituras e 10 questões, 50 pontos |
| `vendor/jspdf.umd.min.js` | Biblioteca jsPDF 2.5.1 (licença MIT), usada para gerar os PDFs |

São arquivos HTML comuns, sem etapa de montagem: o que está aqui é o que o navegador abre.

## Como alterar

- **Lista de disciplinas e de recursos:** fica dentro do `index.html`, nas constantes `AREAS`, `VISIVEIS` e `ASSUNTOS`. Para mostrar uma disciplina nova, inclua o id dela em `VISIVEIS`.
- **Recurso novo:** crie o arquivo na pasta da disciplina e aponte para ele em `ASSUNTOS`, trocando a situação de `breve` para `ok`.
- **Visual:** todas as páginas usam as mesmas cores e letras (Barlow e Big Shoulders Display) e o mesmo cabeçalho, com o nome do site e a faixa de régua.

## Como publicar

Qualquer hospedagem de site estático serve. No GitHub Pages: Settings, Pages, "Deploy from a branch", ramo `main`, pasta `/ (root)`.
