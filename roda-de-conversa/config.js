/*
  CONFIGURAÇÃO DO FIREBASE
  ------------------------
  Enquanto estiver "null", as páginas funcionam em MODO TESTE:
  o telão e o aluno conversam só entre abas do MESMO navegador
  (bom para conferir tudo antes da aula).

  Para usar com os celulares dos alunos, troque o null pelo bloco
  que o Firebase mostra em "Configurações do projeto > Seus apps".
  Exemplo do formato:

  window.FIREBASE_CONFIG = {
    apiKey: "AIza...",
    authDomain: "roda-de-conversa.firebaseapp.com",
    databaseURL: "https://roda-de-conversa-default-rtdb.firebaseio.com",
    projectId: "roda-de-conversa",
    storageBucket: "roda-de-conversa.appspot.com",
    messagingSenderId: "123456789",
    appId: "1:123456789:web:abc123"
  };

  Confira se a linha "databaseURL" está presente.
*/
window.FIREBASE_CONFIG = null;
