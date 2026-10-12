# Roda de Conversa ao vivo

Os alunos escaneiam o QR Code no telão, respondem pelo celular e os resultados aparecem na hora.

## Arquivos

| Arquivo | Para que serve |
|---|---|
| `index.html` | Telão: QR Code, cartões, resultados ao vivo e mural |
| `aluno.html` | Página que abre no celular do aluno |
| `roda.js` | Textos do kit, carinhas e conexão |
| `estilo.css` | Cores e letras |
| `qrcode.js` | Gera o QR Code (biblioteca livre, licença MIT) |
| `config.js` | **Onde você cola a configuração do Firebase** |
| `regras-firebase.json` | Regras de segurança para colar no Firebase |

Mantenha todos os arquivos juntos na mesma pasta.

## 1. Testar antes (sem Firebase)

Com o `config.js` como veio (`null`), tudo funciona em **modo teste**. Abra o `index.html` em uma aba e o link do aluno (aparece embaixo do código da sala) em outra aba do **mesmo navegador**. Vote na aba do aluno e veja o resultado aparecer no telão.

Para o QR Code funcionar nos celulares, as páginas precisam estar publicadas no site (passo 3) e o Firebase configurado (passo 2).

## 2. Criar o Firebase (gratuito, uma vez só)

1. Entre em **console.firebase.google.com** com sua conta Google.
2. Clique em **Criar um projeto**, dê um nome (ex.: `roda-de-conversa`) e pode desativar o Google Analytics.
3. No menu da esquerda, abra **Criação › Realtime Database** e clique em **Criar banco de dados**. Escolha o local sugerido e **Iniciar no modo bloqueado**.
4. Na aba **Regras** do banco, apague o que estiver lá, cole o conteúdo de `regras-firebase.json` e clique em **Publicar**.
5. Clique na engrenagem › **Configurações do projeto**. Em **Seus apps**, clique no ícone **`</>`** (Web), dê um apelido e clique em **Registrar app**.
6. O Firebase mostra um bloco `const firebaseConfig = { ... }`. Copie só o que está entre as chaves e cole no `config.js`, no lugar do `null`:

```js
window.FIREBASE_CONFIG = {
  apiKey: "...",
  authDomain: "...",
  databaseURL: "https://...firebaseio.com",
  projectId: "...",
  storageBucket: "...",
  messagingSenderId: "...",
  appId: "..."
};
```

Confira se aparece a linha `databaseURL`. Se não aparecer, copie o endereço que está no topo da página do Realtime Database e acrescente essa linha.

A chave `apiKey` não é senha. Ela fica visível em qualquer site que usa Firebase, e quem protege os dados são as regras do passo 4.

## 3. Publicar no site

Coloque a pasta inteira no repositório do site (por exemplo, em `roda-de-conversa/`). O telão fica em `.../roda-de-conversa/` e o QR Code já aponta para o `aluno.html` certo.

## Na aula

1. Abra o telão. Um código de 4 números é criado e o QR Code aparece à esquerda.
2. Os alunos escaneiam o QR Code ou entram pelo link e digitam o código. O contador mostra quantos aparelhos entraram.
3. **Situações, Perguntas e Discussão:** cada aluno toca na emoção que sente, e as barras sobem no telão. "Esconder resultados" evita que a turma se influencie antes de todos responderem.
4. **Constrói ou machuca?:** os alunos votam e você toca em **Revelar resposta**. A resposta aparece no telão e também no celular de cada aluno.
5. **Fechamento:** os alunos completam a frase no celular. As respostas **só vão para o mural depois que você aprova**. Use "Aprovar pelo meu celular" para revisar no seu aparelho, sem projetar as respostas antes. Feche o QR de moderação logo depois de escanear.
6. **Nova turma** cria outro código e apaga os dados da turma anterior.

## Observações

- Os alunos não precisam de conta nem de cadastro. Nenhum nome é pedido.
- Cada aparelho tem um voto por cartão e pode trocar a resposta.
- O plano gratuito do Firebase (Spark) aguenta com folga uma turma inteira.
- As regras deixam qualquer pessoa com o código escrever na sala. Para uso em sala de aula isso basta, mas não divulgue o código fora da turma.
