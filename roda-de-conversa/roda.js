/* Roda de Conversa ao vivo — conteúdo do kit, carinhas e conexão (Firebase ou modo teste) */
(function(){
const EMO=[
 {id:"alegre",l:"Alegre",c:"#F7C531"},{id:"triste",l:"Triste",c:"#4DA6E8"},{id:"irritado",l:"Irritado(a)",c:"#E8473F"},{id:"medo",l:"Com medo",c:"#A57BD6"},
 {id:"calmo",l:"Calmo(a)",c:"#58C27A"},{id:"surpreso",l:"Surpreso(a)",c:"#F59A3B"},{id:"timido",l:"Tímido(a)",c:"#F6A9A0"},{id:"orgulhoso",l:"Orgulhoso(a)",c:"#2BB3B1"},
 {id:"preocupado",l:"Preocupado(a)",c:"#A3A9B5"},{id:"amoroso",l:"Amoroso(a)",c:"#EE6F9A"},{id:"envergonhado",l:"Envergonhado(a)",c:"#6CC6A0"},{id:"pensativo",l:"Pensativo(a)",c:"#8F86E8"}];

const RULES=[
 ["Respeitar quem estiver falando.","Cada opinião importa!"],
 ["Não interromper quem estiver falando.","Ouvir também é uma forma de respeito!"],
 ["Não rir ou debochar da opinião de ninguém.","Respeito também é atitude!"],
 ["Não citar nomes de colegas durante a conversa.","Vamos falar sobre ideias, não sobre pessoas!"],
 ["Quem preferir não falar, pode passar a vez.","Aqui, todas as vozes são importantes!"]];

const DECKS={
 sit:{name:"Situações",eyebrow:"Situação para reflexão",ask:"Você já se sentiu assim? O que poderia ajudar?",items:[
  ["Um colega erra a resposta e alguns alunos começam a rir."],
  ["Um aluno fala de modo diferente dos outros porque veio de outro país."],
  ["Uma aluna tem dificuldade para prestar atenção e alguns colegas dizem que ela é preguiçosa."],
  ["Um aluno prefere ficar mais quieto e precisa de mais tempo para responder."],
  ["Um colega faz algo diferente e a turma começa a imitá-lo."],
  ["Um colega é chamado por apelidos que ele não gosta."],
  ["Um aluno demonstra sinais de ansiedade e se sente desconfortável em se expor."],
  ["Uma colega não participa das atividades porque não tem o mesmo acesso aos recursos que os outros."],
  ["Um aluno com TEA (Transtorno do Espectro Autista) tem um jeito diferente de se comunicar."],
  ["Um aluno com TDAH (Transtorno do Déficit de Atenção e Hiperatividade) se distrai com facilidade e tem dificuldade para finalizar as tarefas."]]},
 perg:{name:"Perguntas da roda",eyebrow:"Pergunta para a roda",items:[
  ["Você acha que existe diferença entre brincadeira e desrespeito?","Como podemos identificar quando uma situação passa do limite?"],
  ["Como você se sente quando é chamado por apelidos que não gosta?","O que você faria nessa situação?"],
  ["Quais são os desafios de quem vem de outro país e ainda não conhece bem a nossa cultura?","Como podemos ajudar?"],
  ["Como podemos incluir colegas que têm formas diferentes de aprender, de se comunicar ou de se expressar?"],
  ["Você já presenciou alguma situação de falta de respeito?","O que você fez ou poderia ter feito?"],
  ["O que você pode fazer para tornar a nossa turma um lugar mais acolhedor, respeitoso e inclusivo?"],
  ["Você consegue falar sobre o que sente?","Em caso de dificuldade, a quem você pode recorrer?"],
  ["Como podemos respeitar as diferenças sem fazer piadas ou julgamentos?"],
  ["Qual é a importância da empatia no dia a dia?","Como ela pode mudar a nossa convivência?"],
  ["Você acha que a maneira como falamos e escrevemos pode influenciar na forma como as pessoas nos veem?","Por quê?"],
  ["O que podemos fazer quando alguém está sozinho ou parece estar triste?"],
  ["Qual mensagem você quer deixar para os seus colegas sobre respeito, união e convivência?"]]},
 disc:{name:"Cartões de discussão",eyebrow:"Cartão de discussão",items:[
  ["Nós falamos do mesmo jeito com todas as pessoas?","Você fala com seus amigos da mesma forma que fala com seus professores?"],
  ["As redes sociais mudaram a maneira como escrevemos?","Isso é Português errado ou é uma forma diferente de comunicação?"],
  ["É possível machucar alguém com palavras, sem encostar nela?","Por que alguns apelidos podem deixar de ser brincadeira?"],
  ["Como podemos incluir colegas que são de outro país e ainda não conhecem bem a nossa cultura?","O que podemos fazer para que se sintam bem e participem da turma?"],
  ["Como podemos ajudar colegas que aprendem de maneiras diferentes?","Como podemos comunicar melhor e mostrar respeito?"],
  ["O que podemos fazer quando presenciamos uma situação de agressão, bullying ou exclusão?","Ficar em silêncio ajuda ou atrapalha?"],
  ["Qual é a importância do respeito e da empatia no dia a dia?","Como eles podem transformar o ambiente da nossa turma?"]]}
};

const PAIRS=[
 ["Você é fedido.","Você parece estar se sentindo mal. Posso te ajudar?"],
 ["Você é burro(a).","Você não conseguiu, mas pode tentar de outra forma. Eu acredito em você!"],
 ["Sai daqui!","Você pode ficar um pouco mais longe, por favor?"],
 ["Você fala estranho.","Que legal! Você fala de um jeito diferente. Pode me ensinar alguma palavra?"],
 ["Você não consegue nada.","Você já conseguiu outras coisas. Vai dar certo, você tem capacidade!"],
 ["Você não é normal.","Você é único(a)! E isso é especial!"]];
const FRASES=[];PAIRS.forEach((p,i)=>{FRASES.push({id:"m"+i,t:p[0],ok:false,swap:p[1]});FRASES.push({id:"c"+i,t:p[1],ok:true})});
const FRASE_ORDEM=["m0","c3","m2","c1","m4","c5","m1","c0","m5","c2","m3","c4"];

function face(id,size){
  const e=EMO.find(x=>x.id===id),k="#1C2140";
  const eyes=(r)=>`<circle cx="35" cy="42" r="${r||5}" fill="${k}"/><circle cx="65" cy="42" r="${r||5}" fill="${k}"/>`;
  const closed=`<path d="M27 43 Q35 50 43 43M57 43 Q65 50 73 43" stroke="${k}" stroke-width="4" fill="none" stroke-linecap="round"/>`;
  const blush=`<circle cx="24" cy="58" r="7" fill="#FF6F91" opacity=".55"/><circle cx="76" cy="58" r="7" fill="#FF6F91" opacity=".55"/>`;
  const worry=`<path d="M26 32 L42 27M74 32 L58 27" stroke="${k}" stroke-width="4" stroke-linecap="round"/>`;
  const m=(d)=>`<path d="${d}" stroke="${k}" stroke-width="4.5" fill="none" stroke-linecap="round"/>`;
  const heart=(x)=>`<path d="M${x} 48 C${x-10} 40 ${x-9} 30 ${x} 37 C${x+9} 30 ${x+10} 40 ${x} 48Z" fill="#D2364A"/>`;
  let inner="";
  switch(id){
    case"alegre":inner=eyes()+m("M30 60 Q50 82 70 60");break;
    case"triste":inner=eyes()+m("M32 74 Q50 58 68 74")+`<path d="M30 50 q-4 8 0 11 q4 -3 0 -11z" fill="#8EC9F5"/>`;break;
    case"irritado":inner=eyes()+`<path d="M25 30 L43 38M75 30 L57 38" stroke="${k}" stroke-width="5" stroke-linecap="round"/>`+m("M33 72 Q50 60 67 72");break;
    case"medo":inner=eyes(6)+worry+`<ellipse cx="50" cy="69" rx="9" ry="11" fill="${k}"/>`;break;
    case"calmo":inner=closed+m("M36 63 Q50 73 64 63");break;
    case"surpreso":inner=`<circle cx="35" cy="42" r="8" fill="#fff" stroke="${k}" stroke-width="3"/><circle cx="65" cy="42" r="8" fill="#fff" stroke="${k}" stroke-width="3"/><circle cx="35" cy="43" r="3.5" fill="${k}"/><circle cx="65" cy="43" r="3.5" fill="${k}"/><circle cx="50" cy="70" r="8" fill="${k}"/>`;break;
    case"timido":inner=`<circle cx="36" cy="46" r="4" fill="${k}"/><circle cx="64" cy="46" r="4" fill="${k}"/>`+blush+m("M42 66 Q50 71 58 66");break;
    case"orgulhoso":inner=`<path d="M27 45 Q35 37 43 45M57 45 Q65 37 73 45" stroke="${k}" stroke-width="4" fill="none" stroke-linecap="round"/>`+m("M32 61 Q52 78 70 58");break;
    case"preocupado":inner=eyes()+worry+m("M33 70 q5.6 -6 11.3 0 t11.3 0 t11.3 0");break;
    case"amoroso":inner=heart(35)+heart(65)+m("M32 61 Q50 78 68 61");break;
    case"envergonhado":inner=`<circle cx="30" cy="43" r="4.5" fill="${k}"/><circle cx="58" cy="43" r="4.5" fill="${k}"/>`+blush+m("M38 69 q4 -4 8 0 t8 0 t8 0");break;
    case"pensativo":inner=`<path d="M27 30 L42 33M58 26 L73 30" stroke="${k}" stroke-width="4" stroke-linecap="round"/>`+eyes()+m("M40 70 L62 66")+`<circle cx="74" cy="76" r="6" fill="${e.c}" stroke="${k}" stroke-width="3"/>`;break;
  }
  return `<svg viewBox="0 0 100 100" width="${size||56}" height="${size||56}" aria-hidden="true"><circle cx="50" cy="50" r="45" fill="${e.c}" stroke="${k}" stroke-width="4"/>${inner}</svg>`;
}

const esc=s=>String(s==null?"":s).replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));

/* ---------- Conexão ---------- */
function firebaseDB(cfg){
  firebase.initializeApp(cfg);
  const d=firebase.database();
  return {
    mode:"firebase",
    watch(path,cb){const r=d.ref(path);const h=s=>cb(s.val());r.on("value",h);return()=>r.off("value",h)},
    set(path,v){return d.ref(path).set(v)},
    update(path,v){return d.ref(path).update(v)},
    push(path,v){return d.ref(path).push(v)},
    remove(path){return d.ref(path).remove()},
    presence(path){const r=d.ref(path);r.onDisconnect().remove();return r.set(Date.now())},
    onConn(cb){d.ref(".info/connected").on("value",s=>cb(!!s.val()))}
  };
}
function localDB(){
  const KEY="roda-ao-vivo-teste";let ch=null;try{ch=new BroadcastChannel("roda-ao-vivo")}catch(e){}
  const watchers=new Set();
  const load=()=>{try{return JSON.parse(localStorage.getItem(KEY))||{}}catch(e){return{}}};
  const store=t=>{try{localStorage.setItem(KEY,JSON.stringify(t))}catch(e){}};
  const parts=p=>p.split("/").filter(Boolean);
  const get=(t,p)=>parts(p).reduce((o,k)=>o&&typeof o==="object"?o[k]:undefined,t);
  const prune=(o)=>{if(o&&typeof o==="object"){for(const k of Object.keys(o)){o[k]=prune(o[k]);if(o[k]==null)delete o[k]}return Object.keys(o).length?o:null}return o};
  const put=(p,v)=>{const t=load(),ks=parts(p);let o=t;ks.slice(0,-1).forEach(k=>{if(!o[k]||typeof o[k]!=="object")o[k]={};o=o[k]});o[ks[ks.length-1]]=v==null?null:JSON.parse(JSON.stringify(v));store(prune(t)||{});notify();if(ch)ch.postMessage(1)};
  const notify=()=>{const t=load();watchers.forEach(w=>{const v=get(t,w.p);const s=JSON.stringify(v??null);if(s!==w.last){w.last=s;w.cb(v??null)}})};
  if(ch)ch.onmessage=notify;
  window.addEventListener("storage",e=>{if(e.key===KEY)notify()});
  return {
    mode:"teste",
    watch(p,cb){const w={p,cb,last:undefined};watchers.add(w);setTimeout(()=>{const v=get(load(),p);w.last=JSON.stringify(v??null);cb(v??null)},0);return()=>watchers.delete(w)},
    set(p,v){put(p,v);return Promise.resolve()},
    update(p,v){Object.entries(v).forEach(([k,x])=>put(p+"/"+k,x));return Promise.resolve()},
    push(p,v){const id=Date.now().toString(36)+Math.random().toString(36).slice(2,6);put(p+"/"+id,v);return Promise.resolve()},
    remove(p){put(p,null);return Promise.resolve()},
    presence(p){put(p,Date.now());window.addEventListener("pagehide",()=>put(p,null));return Promise.resolve()},
    onConn(cb){cb(true)}
  };
}
function connect(){
  const cfg=window.FIREBASE_CONFIG;
  if(cfg&&window.firebase){try{return firebaseDB(cfg)}catch(e){console.error(e)}}
  return localDB();
}

window.RODA={EMO,RULES,DECKS,FRASES,FRASE_ORDEM,face,esc,connect};
})();
