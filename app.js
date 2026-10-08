const C={cy:'#22d3ee',mg:'#ff3d8b',ye:'#fbbf24',vi:'#8b5cf6',ok:'#34e8a5'};
const $=s=>document.querySelector(s),nf=n=>Math.round(n).toLocaleString('es-PE');
const norm=s=>String(s).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f\s_]/g,'');
const A={op:['odtcod','op'],desc:['odtdescrip','descripcion'],fecha:['fecha'],oper:['desoperador','operador'],maq:['maqdes','maquina'],proc:['proceso'],min:['minutos','minutostotalesdeproceso'],b:['pliegosparc','cantidadbuena'],m:['pliegosparcmal','cantidadmala']};
let rows=[];
const iso=v=>{if(typeof v=='number')return new Date(Math.round((Math.floor(v)-25569)*864e5)).toISOString().slice(0,10);const d=new Date(v);return isNaN(d)?'':d.toISOString().slice(0,10)};
function setRows(r,name){rows=r;$('#info').textContent=name+' · '+nf(r.length)+' registros';
 const u=k=>[...new Set(rows.map(x=>x[k]))].sort().map(v=>[v,v]);
 const dm={};rows.forEach(x=>dm[x.op]=dm[x.op]||x.desc);
 msBuild('maq',u('maq'));msBuild('oper',u('oper'));msBuild('op',u('op').map(([v])=>[v,v+' · '+(dm[v]||'').slice(0,40)]));
 const ds=rows.map(x=>x.d).filter(Boolean).sort();$('#f-d1').value=ds[0]||'';$('#f-d2').value=ds[ds.length-1]||'';render()}
function load(buf,name){const wb=XLSX.read(buf,{type:'array'});const js=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:null});
 const keys=Object.keys(js[0]||{}),m={};for(const k in A)m[k]=keys.find(h=>A[k].includes(norm(h)));
 const miss=Object.keys(m).filter(k=>!m[k]);if(miss.length){alert('No encuentro estas columnas: '+miss.join(', ')+'\nEncabezados leídos: '+keys.join(', '));return}
 const t=v=>String(v??'').trim(),n=v=>+v||0;
 setRows(js.map(r=>({op:t(r[m.op]),desc:t(r[m.desc]),d:iso(r[m.fecha]),oper:t(r[m.oper]),maq:t(r[m.maq]),proc:t(r[m.proc]),min:n(r[m.min]),b:n(r[m.b]),m:n(r[m.m])})),name)}
const sel={maq:new Set(),oper:new Set(),op:new Set()},LB={maq:'Todas',oper:'Todos',op:'Todas'};
const esc=v=>String(v).replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/</g,'&lt;');
function msSum(k){const s=sel[k];$('#f-'+k+' summary').textContent=!s.size?LB[k]:s.size==1?[...s][0]:s.size+' seleccionados'}
function msBuild(k,vals){sel[k].clear();$('#f-'+k+' .opts').innerHTML='<input class="q" placeholder="Buscar..." autocomplete="off"><a class="clr" href="#">Limpiar selección</a>'+vals.map(([v,t])=>`<label><input type="checkbox" value="${esc(v)}"> ${esc(t)}</label>`).join('');msSum(k)}
document.addEventListener('change',e=>{const d=e.target.closest('.ms');if(!d)return;const k=d.id.slice(2);e.target.checked?sel[k].add(e.target.value):sel[k].delete(e.target.value);msSum(k);render()});
document.addEventListener('click',e=>{const c=e.target.closest('.clr');if(c){e.preventDefault();const d=c.closest('.ms'),k=d.id.slice(2);sel[k].clear();d.querySelectorAll('input').forEach(i=>i.checked=false);msSum(k);render();return}
 document.querySelectorAll('.ms[open]').forEach(d=>{if(!d.contains(e.target))d.open=false})});
$('#file').onchange=e=>{const f=e.target.files[0];if(f)f.arrayBuffer().then(b=>load(b,f.name))};
const isProd=p=>norm(p).startsWith('produccion');
function agg(list,key){const M=new Map();for(const r of list){const k=r[key];let a=M.get(k);if(!a)M.set(k,a={k,b:0,m:0,min:0,pmin:0,rows:[]});a.b+=r.b;a.m+=r.m;a.min+=r.min;if(isProd(r.proc))a.pmin+=r.min;a.rows.push(r)}return [...M.values()].sort((x,y)=>y.b-x.b)}
const cal=a=>a.b+a.m>0?a.b/(a.b+a.m)*100:null,bph=a=>a.pmin>0?a.b/(a.pmin/60):null,ut=a=>a.min>0?a.pmin/a.min*100:null;
const cls=c=>c==null?'':c>=98?'g':c>=95?'w':'r',pc=c=>c==null?'–':c.toFixed(1)+'%';
function filtered(){const d1=$('#f-d1').value,d2=$('#f-d2').value,tx=$('#f-desc').value.toLowerCase(),ok=(k,v)=>!sel[k].size||sel[k].has(v);
 return rows.filter(r=>(!d1||r.d>=d1)&&(!d2||r.d<=d2)&&ok('maq',r.maq)&&ok('oper',r.oper)&&ok('op',r.op)&&(!tx||r.desc.toLowerCase().includes(tx)))}
const tr=(a,label,c)=>`<tr class="${c}"><td>${label}</td><td>${nf(a.b)}</td><td>${nf(a.m)}</td><td class="${cls(cal(a))}">${pc(cal(a))}<span class="bar"><i style="width:${cal(a)??0}%"></i></span></td><td>${nf(a.pmin)}</td><td>${bph(a)==null?'–':nf(bph(a))}</td><td>${pc(ut(a))}</td></tr>`;
function render(){const L=filtered(),T=agg(L,'x')[0]||{b:0,m:0,min:0,pmin:0,rows:[]};
 const all=agg(L.map(r=>({...r,x:1})),'x')[0]||{b:0,m:0,min:0,pmin:0};
 $('#kpis').innerHTML=[['','Cantidad buena',nf(all.b)],['m','Cantidad mala',nf(all.m)],['','% calidad (buena / total)',pc(cal(all))],['y','Minutos de producción',nf(all.pmin)],['y','Buenas por hora productiva',bph(all)==null?'–':nf(bph(all))],['t','% tiempo en producción',pc(ut(all))]].map(([c,l,v])=>`<div class="k ${c}"><b>${v}</b><span>${l}</span></div>`).join('');
 const v=$('#vista').value,P={om:['oper','maq'],mo:['maq','oper'],o:['oper'],m:['maq']}[v];
 $('#hint').textContent={om:'Cada operador y, debajo, las máquinas donde trabajó',mo:'Cada máquina y, debajo, los operadores que la usaron',o:'Compara operadores; marca una o varias máquinas en el filtro para acotar',m:'Compara máquinas; marca uno o varios operadores en el filtro para acotar'}[v];
 let h='';for(const a of agg(L,P[0])){h+=tr(a,a.k,'p');if(P[1])for(const c of agg(a.rows,P[1]))h+=tr(c,c.k,'c')}
 $('#tabla').innerHTML=L.length?`<table><thead><tr><th>${P[1]?'Grupo / detalle':P[0]=='oper'?'Operador':'Máquina'}</th><th>Buena</th><th>Mala</th><th>% calidad</th><th>Min. prod.</th><th>Buenas/h</th><th>% t. prod.</th></tr></thead><tbody>${h}</tbody></table>`:'<div class="empty">Sin registros con estos filtros. Amplía las fechas o limpia algún filtro.</div>';
 const ps=agg(L,'proc'),mx=Math.max(1,...ps.map(a=>a.min)),sorted=ps.sort((a,b)=>b.min-a.min);
 charts(L,P);$('#procs').innerHTML=sorted.map(a=>`<div class="row"><div title="${a.k}">${a.k}</div><div><i style="width:${a.min/mx*100}%"></i></div><div>${nf(a.min)} min</div></div>`).join('')}
['#f-d1','#f-d2','#f-desc','#vista'].forEach(s=>$(s).addEventListener('input',render));
if(typeof DATA!=='undefined')setRows(DATA.map(r=>({op:r[0],desc:r[1],d:r[2],oper:r[3],maq:r[4],proc:r[5],min:r[6],b:r[7],m:r[8]})),'EJEMPLO.xlsx (datos de muestra)');
else setRows([],'Sin datos · carga un Excel');

function charts(L,P){
 const D=agg(L,'d').filter(a=>a.k).sort((a,b)=>a.k<b.k?-1:1),n=D.length;
 if(!n){['#ch-trend','#ch-rank','#ch-donut'].forEach(s=>$(s).innerHTML='<div class="empty">Sin datos con estos filtros</div>');return}
 const W=640,H=210,p=26,mx=Math.max(1,...D.map(a=>a.b)),bw=(W-2*p)/n,lo=Math.max(0,Math.min(...D.map(cal).filter(c=>c!=null),99)-1);
 const bars=D.map((a,i)=>{const h=a.b/mx*(H-2*p);return `<rect x="${p+i*bw+(bw-Math.min(Math.max(1,bw*.76),28))/2}" y="${H-p-h}" width="${Math.min(Math.max(1,bw*.76),28)}" height="${h}" rx="2" fill="url(#g1)"><title>${a.k}: ${nf(a.b)} buenas, ${nf(a.m)} malas</title></rect>`}).join('');
 const pts=D.map((a,i)=>cal(a)==null?null:[(p+i*bw+bw/2).toFixed(1),(H-p-(cal(a)-lo)/(100-lo)*(H-2*p)).toFixed(1)]).filter(Boolean);
 const line=`<polyline points="${pts.join(' ')}" fill="none" stroke="${C.mg}" stroke-width="2" style="filter:drop-shadow(0 0 4px ${C.mg})"/>`+pts.map(q=>`<circle cx="${q[0]}" cy="${q[1]}" r="2.5" fill="${C.mg}"/>`).join('');
 $('#ch-trend').innerHTML=`<svg viewBox="0 0 ${W} ${H}" width="100%"><defs><linearGradient id="g1" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="${C.cy}"/><stop offset="1" stop-color="${C.vi}" stop-opacity=".5"/></linearGradient></defs>${bars}${line}<text x="${p}" y="${H-7}" fill="#7f93b5" font-size="11">${D[0].k}</text><text x="${W-p}" y="${H-7}" fill="#7f93b5" font-size="11" text-anchor="end">${D[n-1].k}</text></svg><div class="lg"><i style="background:${C.cy}"></i>Cantidad buena por día <i style="background:${C.mg}"></i>% calidad (escala ${lo.toFixed(0)}–100%)</div>`;
 const R=agg(L,P[0]).slice(0,10),rm=Math.max(1,...R.map(a=>a.b+a.m));
 $('#ch-rank-t').textContent='Ranking por '+(P[0]=='oper'?'operador':'máquina')+' (top 10)';
 $('#ch-rank').innerHTML=R.map(a=>`<div class="rk"><div title="${esc(a.k)}">${esc(a.k)}</div><div class="rt"><i style="width:${a.b/rm*100}%"></i><u style="width:${a.m/rm*100}%"></u></div><div class="${cls(cal(a))}">${pc(cal(a))}</div></div>`).join('');
 const S=agg(L,'proc').sort((a,b)=>b.min-a.min),tot=S.reduce((s,a)=>s+a.min,0)||1,col=[C.cy,C.mg,C.ye,C.vi,C.ok,'#64748b'];
 const top=S.slice(0,5),rest=tot-top.reduce((s,a)=>s+a.min,0),sl=rest>0?[...top,{k:'Otros',min:rest}]:top,r=52,cc=2*Math.PI*r;let off=0;
 const arcs=sl.map((a,i)=>{const f=a.min/tot,e=`<circle r="${r}" cx="70" cy="70" fill="none" stroke="${col[i]}" stroke-width="16" stroke-dasharray="${f*cc} ${cc}" stroke-dashoffset="${-off*cc}" transform="rotate(-90 70 70)"/>`;off+=f;return e}).join('');
 const pr=L.filter(x=>isProd(x.proc)).reduce((s,x)=>s+x.min,0)/tot*100;
 $('#ch-donut').innerHTML=`<svg viewBox="0 0 140 140">${arcs}<text x="70" y="68" text-anchor="middle" fill="#e8f0ff" font-size="22" font-weight="700">${pr.toFixed(0)}%</text><text x="70" y="86" text-anchor="middle" fill="#7f93b5" font-size="9">en producción</text></svg><div class="dl">${sl.map((a,i)=>`<div><i style="background:${col[i]}"></i>${esc(a.k)}<b>${(a.min/tot*100).toFixed(0)}%</b></div>`).join('')}</div>`}

document.addEventListener('input',e=>{if(!e.target.classList.contains('q'))return;const t=norm(e.target.value);e.target.closest('.opts').querySelectorAll('label').forEach(l=>l.style.display=norm(l.textContent).includes(t)?'':'none')});
document.addEventListener('toggle',e=>{if(e.target.open){const q=e.target.querySelector('.q');if(q)q.focus()}},true);
