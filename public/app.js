const coordenadores = ['Johnny Dias Carvalho','Jose Nailson Goncalves Ferreira','Vania Novaes Reboucas','Rosa Aparecida Acacia de Oliveira'];
const base = {
  'Johnny Dias Carvalho': { dia:'Quinta-feira', inicio:'16:40', fim:'', professores:[
    ['Andrea Fernandes','166.999.278-09'],['Claudio de Jesus','129.832.548-01'],['Celso Eduardo Firmino de Oliveira','184.744.378-81'],['Katia Nunes Caldeira','218.994.968-82'],['Marinalva Eva da Silva Leite','265.153.818-32'],['Michele Almeida Santana','014.454.355-94'],['Roberto Valdomiro de Brito','022.910.668-46'],['Tatiana Gonçalves Silva Soares','418.702.428-07']
  ]},
  'Jose Nailson Goncalves Ferreira':{ dia:'Quinta-feira', inicio:'19:00', fim:'20:30', professores:[
    ['Adriana Morais dos Santos','152.118.918-80'],['Alexandra Karine Silva da Rocha','390.067.968-16'],['Alexandre Brito de Oliveira','110.891.158-73'],['Brigite Belluco Inoue','074.683.998-76'],['Cicero Luis da Silva Santos','083.960.778-43'],['Erico Marques da Costa','360.899.108-56'],['Edson de Souza e Silva','220.072.978-25'],['Fabio Augusto de Moraes','001.308.908-04'],['Laura Ribeiro Goncalves de Alvarenga','196.739.128-90'],['Nilma Aparecida de Sousa Lares','311.270.308-17'],['Sonia Borges Soares','165.442.468-40'],['Soraia Vicente de Sousa Dias','375.800.498-56'],['Vinicius Nonato dos Santos','372.618.088-57'],['Alexandre Marques','225.603.718-55'],['Aline de Souza Campos Santos Andrade','']
  ]},
  'Vania Novaes Reboucas':{ dia:'Quarta-feira', inicio:'13:00', fim:'', professores:[
    ['Anderson Borges','219.110.938-19'],['Ariel Fischer Gloria','068.286.168-55'],['Renata Ribeiro dos Santos Gandra','278.072.738-10'],['Lucas Oliveira Araujo','352.875.268-81'],['Daiane Silva Assunção','400.418.298-08'],['Hugo dos Santos Lima','358.985.098-10'],['Robson de Jesus','254.861.258-95'],['Nilma Aparecida de Sousa Lares','311.270.308-17'],['Gilvanir Arcanjo Soares Rodrigues','042.171.228-70'],['Alex da Silva Moreira','427.868.118-64'],['Marcia Cristina S. A. Taha','327.925.638-00'],['Anderson Pereira de Souza','4368089-5'],['Evelyn Pereira de Souza','48.715.984-6'],['Luiza Martins Araújo Rego','079116388-10']
  ]},
  'Rosa Aparecida Acacia de Oliveira':{ dia:'Quarta-feira', inicio:'10:40', fim:'', professores:[
    ['Soraia Vicente de Sousa Dias','375.800.498-56'],['Priscila Costa de Santana','277.039.068-67'],['Marinalva Eva da Silva Leite','265.153.818-32'],['Lya de Jesus Oliveira','766.023.545-15'],['Cristiane Frotta dos Santos','152.061.798-41'],['Carla Maria de Castro Martins','931.402.747-34'],['Thiago Alessandro Xavier','375.619.388-88'],['Gabriela Vitoria Lourenço Pinto','485.115.118-73'],['Marlucia Rosa da Silva Batista','082.984.308-67'],['Sirleide Maria Alves Gonzalez','415.729.708-39'],['Josiane Santos do Carmo','220.700.308-65'],['Renata Ribeiro dos Santos Gandra','278.072.738-10']
  ]}
};

const $ = id => document.getElementById(id);
let professoresAtuais = [];

function hoje(){ return new Date().toISOString().slice(0,10); }
function brData(v){ if(!v)return ''; const [a,m,d]=v.split('-'); return `${d}/${m}/${a}`; }
function diaSemana(v){ if(!v)return ''; return new Intl.DateTimeFormat('pt-BR',{weekday:'long',timeZone:'UTC'}).format(new Date(v+'T12:00:00Z')); }
function dataExtenso(v){ if(!v)return ''; return new Intl.DateTimeFormat('pt-BR',{day:'numeric',month:'long',year:'numeric',timeZone:'UTC'}).format(new Date(v+'T12:00:00Z')); }

function init(){
  coordenadores.forEach(c => $('coordenador').add(new Option(c,c)));
  $('data').value = hoje();
  $('coordenador').value = 'Johnny Dias Carvalho';
  loadCoord();
  $('coordenador').addEventListener('change',loadCoord);
  ['data','inicio','fim','pauta','informacoes','encaminhamentos','tamanhoAta','ata','redator','redatorOutro'].forEach(id => $(id).addEventListener('input',renderPreview));
  $('addProfessor').onclick=()=>{ professoresAtuais.push({nome:'Novo professor',documento:'',situacao:'Presente'}); renderProfessores(); };
  $('gerarAta').onclick=gerarAta;
  $('docx').onclick=()=>downloadDoc('/api/docx','.docx');
  $('pdf').onclick=()=>downloadDoc('/api/pdf','.pdf');
  $('salvar').onclick=salvarHistorico;
  $('btnHistorico').onclick=abrirHistorico; $('fecharHistorico').onclick=()=>$('historicoDialog').close();
}

function loadCoord(){
  const c=$('coordenador').value, cfg=base[c];
  $('inicio').value=cfg.inicio; $('fim').value=cfg.fim;
  professoresAtuais=cfg.professores.map(([nome,documento])=>({nome,documento,situacao:'Presente'}));
  $('redatorOutro').value='';
  $('coordsExtras').innerHTML='';
  coordenadores.filter(x=>x!==c).forEach(x=>{
    const lab=document.createElement('label');
    lab.innerHTML=`<input type="checkbox" value="${x}"> ${x}`;
    lab.querySelector('input').addEventListener('change',renderPreview);
    $('coordsExtras').appendChild(lab);
  });
  renderProfessores(); updateRedatorOptions(); renderPreview();
}

function updateRedatorOptions(preferido=''){
  const sel=$('redator');
  const atual=preferido || sel.value;
  sel.innerHTML='';
  professoresAtuais.forEach(p=>{ if(p.nome.trim()) sel.add(new Option(p.nome,p.nome)); });
  if([...sel.options].some(o=>o.value===atual)) sel.value=atual;
  else if(sel.options.length) sel.selectedIndex=0;
}

function renderProfessores(){
  const body=$('professores'); body.innerHTML='';
  professoresAtuais.forEach((p,i)=>{
    const tr=document.createElement('tr');
    tr.className=`status-${p.situacao}`;
    tr.innerHTML=`<td><input class="teacher-name" value="${escapeHtml(p.nome)}"></td><td><input class="teacher-doc" value="${escapeHtml(p.documento||'')}" placeholder="CPF ou RG"></td><td><select class="teacher-status"><option>Presente</option><option>Ausente</option><option>Justificado</option></select></td><td><button class="remove" title="Remover desta ATPC">✕</button></td>`;
    const name=tr.querySelector('.teacher-name'), doc=tr.querySelector('.teacher-doc'), status=tr.querySelector('.teacher-status');
    status.value=p.situacao;
    name.oninput=e=>{const anterior=p.nome; p.nome=e.target.value; const eraSelecionado=$('redator').value===anterior; updateRedatorOptions(eraSelecionado?p.nome:$('redator').value); renderPreview()}; doc.oninput=e=>{p.documento=e.target.value;renderPreview()};
    status.onchange=e=>{p.situacao=e.target.value; renderProfessores();};
    tr.querySelector('.remove').onclick=()=>{professoresAtuais.splice(i,1);renderProfessores();updateRedatorOptions()};
    body.appendChild(tr);
  });
  updateRedatorOptions(); renderPreview();
}
function escapeHtml(s=''){return String(s).replace(/[&<>\"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]))}
function extras(){return [...$('coordsExtras').querySelectorAll('input:checked')].map(x=>x.value)}
function payload(){
  return {
    coordenadorPrincipal:$('coordenador').value,
    coordenadoresParticipantes:extras(),
    data:brData($('data').value), dataISO:$('data').value,
    dataExtenso:dataExtenso($('data').value), diaSemana:diaSemana($('data').value),
    horarioInicio:$('inicio').value, horarioFim:$('fim').value,
    pauta:$('pauta').value.trim(), informacoes:$('informacoes').value.trim(), encaminhamentos:$('encaminhamentos').value.trim(),
    redator:$('redatorOutro').value.trim() || $('redator').value,
    tamanhoAta:$('tamanhoAta').value || 'muito_longa',
    ata:$('ata').value.trim(), professores:professoresAtuais
  }
}

async function gerarAta(){
  if(!$('informacoes').value.trim() && !$('pauta').value.trim()){ alert('Preencha pelo menos a pauta ou as informações discutidas.'); return; }
  const b=$('gerarAta'); b.disabled=true; b.textContent='Gerando...';
  try{ const r=await fetch('/api/gerar-ata',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload())}); const j=await r.json(); $('ata').value=j.ata||''; renderPreview(); if(j.modo==='modelo-local') console.info('Ata gerada pelo modelo local. Configure OPENAI_API_KEY para usar IA.'); }
  catch(e){alert('Não foi possível gerar a ata.');}
  finally{b.disabled=false;b.textContent='✨ Gerar ata com IA'}
}

function renderPreview(){
  const d=payload(); const rows=d.professores.map((p,i)=>`<tr><td>${i+1}</td><td>${escapeHtml(p.nome)}</td><td>${escapeHtml(p.situacao)}</td><td class="signature-preview">____________________________</td></tr>`).join('');
  $('preview').innerHTML=`<div class="meta"><b>Data:</b> ${escapeHtml(d.data)}<b>Horário:</b> ${escapeHtml(d.horarioInicio)}${d.horarioFim?' às '+escapeHtml(d.horarioFim):''}<b>Coordenador:</b> ${escapeHtml(d.coordenadorPrincipal)}</div><h3>ATA DO ATPC</h3><p>${escapeHtml(d.ata||'A ata gerada aparecerá aqui.').replace(/\n/g,'<br>')}</p>${d.encaminhamentos?`<h3>ENCAMINHAMENTOS / COMBINADOS</h3><p>${escapeHtml(d.encaminhamentos).replace(/\n/g,'<br>')}</p>`:''}<h3>LISTA DE PRESENÇA</h3><table><thead><tr><th>Nº</th><th>Professor</th><th>Situação</th><th>Assinatura</th></tr></thead><tbody>${rows}</tbody></table>`;
}

async function downloadDoc(url,ext){
  if(!$('ata').value.trim()){alert('Gere ou escreva a ata antes de criar o documento.');return;}
  const r=await fetch(url,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload())});
  if(!r.ok){const j=await r.json().catch(()=>({}));alert(j.error||'Erro ao gerar documento.');return;}
  const blob=await r.blob(), a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download=`ATPC_${$('coordenador').value}_${$('data').value}${ext}`; a.click(); setTimeout(()=>URL.revokeObjectURL(a.href),1000);
}
function salvarHistorico(){const h=JSON.parse(localStorage.getItem('atpc-historico')||'[]');h.unshift({...payload(),salvoEm:new Date().toISOString()});localStorage.setItem('atpc-historico',JSON.stringify(h.slice(0,100)));alert('ATPC salva no histórico deste navegador.');}
function abrirHistorico(){const h=JSON.parse(localStorage.getItem('atpc-historico')||'[]');$('historicoLista').innerHTML=h.length?h.map((x,i)=>`<div class="history-item"><div><b>${escapeHtml(x.coordenadorPrincipal)} • ${escapeHtml(x.data)}</b><br><small>${escapeHtml(x.horarioInicio)}${x.horarioFim?'–'+escapeHtml(x.horarioFim):''} • ${x.professores.filter(p=>p.situacao==='Presente').length} presentes</small></div><button class="secondary" onclick="carregarHistorico(${i})">Carregar</button></div>`).join(''):'<p>Nenhuma ATPC salva ainda.</p>';$('historicoDialog').showModal();}
window.carregarHistorico=i=>{const h=JSON.parse(localStorage.getItem('atpc-historico')||'[]'),x=h[i];if(!x)return;$('coordenador').value=x.coordenadorPrincipal;loadCoord();$('data').value=x.dataISO||'';$('inicio').value=x.horarioInicio||'';$('fim').value=x.horarioFim||'';$('pauta').value=x.pauta||'';$('informacoes').value=x.informacoes||'';$('encaminhamentos').value=x.encaminhamentos||'';$('tamanhoAta').value=x.tamanhoAta||'muito_longa';$('ata').value=x.ata||'';professoresAtuais=x.professores||[];renderProfessores();updateRedatorOptions(x.redator||'');$('redatorOutro').value=[...$('redator').options].some(o=>o.value===x.redator)?'':(x.redator||'');(x.coordenadoresParticipantes||[]).forEach(c=>{const cb=[...$('coordsExtras').querySelectorAll('input')].find(y=>y.value===c);if(cb)cb.checked=true});renderPreview();$('historicoDialog').close()};
init();
