/* Abell drawing engine. Labels, selections and export use the same geometry. */
'use strict';
const W=1500,H=1060;
const F_SANS='Arial, sans-serif',F_SERIF='Georgia, serif';
const PAL=['#2b5963','#af7950','#638b71','#7d6b97','#567fab','#b86663'];
const AXES={clients:{label:'Clienti',singular:'cliente',color:'#407c68',help:'Chi servi? Inserisci i gruppi di clienti.',placeholder:'Es. Pendolari'},functions:{label:'Funzioni',singular:'funzione',color:'#b7804c',help:'Quali bisogni soddisfi? Inserisci le funzioni d’uso.',placeholder:'Es. Spostarsi in città'},modes:{label:'Modalità',singular:'modalità',color:'#847494',help:'Come soddisfi quei bisogni? Inserisci le modalità.',placeholder:'Es. Noleggio a stazioni fisse'}};
const esc=s=>String(s??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
const clone=o=>JSON.parse(JSON.stringify(o));
const uid=()=> 'a'+Date.now().toString(36)+Math.random().toString(36).slice(2,8);
const num=v=>{let s=String(v??'').trim().replace(/[\s%€]/g,'');if(s.includes(','))s=s.replace(/\./g,'').replace(',','.');return Number(s)||0;};
const sum=a=>a.reduce((s,v)=>s+v,0);
const fmtN=n=>Number(n).toLocaleString('it-IT',{maximumFractionDigits:2});
function wrap(s,width=100){const out=[];String(s??'').split('\n').forEach(par=>{let line='';for(let word of par.split(/\s+/).filter(Boolean)){while(word.length>width){if(line){out.push(line);line='';}out.push(word.slice(0,width));word=word.slice(width);}if(!word)continue;if(line&&line.length+word.length+1>width){out.push(line);line=word;}else line+=(line?' ':'')+word;}out.push(line);});return out;}
function item(label='',weight=''){return{id:uid(),label,weight};}
function blank(){return{id:uid(),title:'Nuovo business',subtitle:'',clients:[item('Cliente 1'),item('Cliente 2'),item('Cliente 3')],functions:[item('Funzione 1'),item('Funzione 2'),item('Funzione 3')],modes:[item('Modalità 1')],weightsOn:false,weightsBasis:'',revenueOn:false,revenue:'',revenueUnit:'milioni di €',revenuePeriod:'',revenueMethod:'',criteriaOn:false,criteria:'',legend:'',closure:'outer',cubes:[],offsets:{}};}
function example(){const p=blank();p.title='Noleggio di biciclette elettriche';p.subtitle='Italia · Esempio modificabile';p.clients=[item('Pendolari','55'),item('Turisti','30'),item('Corrieri e rider','15')];p.functions=[item('Spostamento casa-lavoro'),item('Visita turistica della città'),item('Consegne dell’ultimo miglio')];p.modes=[item('Noleggio a stazioni fisse')];return p;}
function normalize(raw){
 if(!raw||typeof raw!=='object'||Array.isArray(raw))throw Error('Formato non valido');
 const p=blank();const text=(v,n=1000)=>String(v??'').slice(0,n);const idMap={};
 p.title=text(raw.title??raw.nome??p.title,120);p.subtitle=text(raw.subtitle??raw.merc,140);p.id=uid();
 for(const axis of Object.keys(AXES)){const old={clients:'cli',functions:'fun',modes:'mod'}[axis];const a=raw[axis]??raw[old];if(Array.isArray(a))p[axis]=a.slice(0,8).map(v=>{const n=item(text(v&&typeof v==='object'?(v.label??v.l):v,160),text(v?.weight??v?.v,20));if(v?.id)idMap[v.id]=n.id;return n;});if(!p[axis].length)p[axis]=[item('')];}
 for(const k of ['weightsOn','revenueOn','criteriaOn'])p[k]=raw[k]===true;
 for(const k of ['weightsBasis','revenue','revenueUnit','revenuePeriod','revenueMethod','criteria','legend'])if(raw[k]!=null)p[k]=text(raw[k]);
 if(raw.crit!=null)p.criteria=text(raw.crit);if(raw.pesi!=null)p.weightsBasis=text(raw.pesi);
 if(raw.vol){p.revenue=String(sum((raw.vol.righe||[]).map(r=>num(r.v))));p.revenueUnit=text(raw.vol.unita||'milioni di €');p.revenuePeriod=text(raw.vol.intest);p.revenueMethod=text(raw.vol.met);}
 const cubes=raw.cubes??raw.cubi;p.closure=raw.closure==='selected'||raw.agg?'selected':'outer';
 if(Array.isArray(cubes))p.cubes=cubes.slice(0,8).map((c,i)=>({id:uid(),name:text(c.name??c.n??('Cubo '+(i+1)),80),c:Math.max(1,Math.min(p.clients.length,Math.round(num(c.c)||1))),f:Math.max(1,Math.min(p.functions.length,Math.round(num(c.f)||1))),m:Math.max(1,Math.min(p.modes.length,Math.round(num(c.m??c.t)||1))),closed:c.closed!==false,color:/^#[0-9a-f]{6}$/i.test(c.color)?c.color:PAL[i%PAL.length]}));
 for(const [key,value] of Object.entries(raw.offsets||{})){const newKey=idMap[key]||(['title','subtitle','axis-functions','axis-clients','axis-modes','criteria','weights','revenue','method','legend'].includes(key)?key:null);if(newKey&&Array.isArray(value)&&value.length===2&&value.every(n=>Number.isFinite(n)&&Math.abs(n)<3000))p.offsets[newKey]=value;}
 return p;
}
function shares(p){const ws=p.clients.map(v=>Math.max(0,num(v.weight)));const t=sum(ws);return p.clients.map((_,i)=>p.weightsOn&&t>0?ws[i]/t:1/p.clients.length);}
function newCube(p){return{id:uid(),name:'Cubo '+(p.cubes.length+1),c:p.clients.length,f:p.functions.length,m:p.modes.length,closed:true,color:PAL[p.cubes.length%PAL.length]};}
function buildModel(p,interactive=false){
 const gfx=[],labels=[],anchors={};const add=s=>gfx.push(s);
 const T=(key,x,y,s,size=20,o={})=>{const off=p.offsets[key]||[0,0];labels.push({key,x:x+off[0],y:y+off[1],s,lines:wrap(s,o.width||100),size,anchor:o.anchor||'start',fill:o.fill||'#334449',family:o.family||F_SANS,weight:o.weight||'normal',style:'normal',halo:!!o.halo,axis:o.axis,itemId:o.itemId});};
 const line=(a,b,color='#274d58',width=2,dash='')=>`<line x1="${a[0]}" y1="${a[1]}" x2="${b[0]}" y2="${b[1]}" stroke="${color}" stroke-width="${width}"${dash?` stroke-dasharray="${dash}"`:''}/>`;
 const nc=p.clients.length,nf=p.functions.length,nm=p.modes.length,cum=[0];shares(p).forEach(s=>cum.push(cum.at(-1)+s));
 const ox=570,oy=635,lx=730,fh=390/nf,dx=215/nm,dy=170/nm;
 const P=(c,f,m)=>[ox+(cum[c]??1)*lx-m*dx,oy-f*fh+m*dy];
 T('title',70,70,p.title||'Nuovo business',32,{family:F_SERIF,width:80});
 if(p.subtitle.trim())T('subtitle',70,112,p.subtitle,17,{fill:'#83918e',width:130});
 const edges=[['000','100'],['100','110'],['110','010'],['010','000'],['001','101'],['101','111'],['111','011'],['011','001'],['000','001'],['100','101'],['110','111'],['010','011']];
 const box=(c,f,m,color,closed,id,index)=>{
  const v={};for(const a of [0,1])for(const b of [0,1])for(const d of [0,1])v[`${a}${b}${d}`]=P(a*c,b*f,d*m);
  let s='';if(closed){s+=`<polygon points="${['010','110','111','011'].map(k=>v[k].join(',')).join(' ')}" fill="${color}" fill-opacity=".045"/>`;s+=`<polygon points="${['101','111','110','100'].map(k=>v[k].join(',')).join(' ')}" fill="${color}" fill-opacity=".025"/>`;for(const [a,b]of edges)s+=line(v[a],v[b],color,2.2);}
  if(interactive&&p.closure==='selected'){
   if(!closed)for(const [a,b]of edges)s+=line(v[a],v[b],color,1.2,'5 7');
   const hit=edges.map(([a,b])=>line(v[a],v[b],'transparent',18)).join('');
   const q=v['111'];s+=`<g class="cube-hit" data-cube="${id}" role="checkbox" tabindex="0" aria-checked="${closed}" aria-label="Chiudi ${esc(p.cubes[index].name)}">${hit}<circle cx="${q[0]}" cy="${q[1]}" r="15" fill="${closed?color:'#fff'}" stroke="${color}" stroke-width="2"/><text x="${q[0]}" y="${q[1]+5}" fill="${closed?'#fff':color}" text-anchor="middle" font-size="14" font-family="Arial">${index+1}</text></g>`;
  }add(s);
 };
 if(p.closure==='outer')box(nc,nf,nm,PAL[0],true);
 else p.cubes.map((c,i)=>({c,i})).sort((a,b)=>(b.c.c*b.c.f*b.c.m)-(a.c.c*a.c.f*a.c.m)).forEach(({c,i})=>{if(c.closed||interactive)box(c.c,c.f,c.m,c.color,c.closed,c.id,i);});
 // Fine reference marks remain attached to the dimensions, independent of label placement.
 for(let i=1;i<nc;i++)add(line(P(i,0,0),P(i,0,nm),'#dce6e5',1,'4 6'));
 const axisEnd=[[ox,oy-470],[ox+lx+65,oy],[ox-265,oy+209]];
 axisEnd.forEach(a=>add(`<path d="M${ox} ${oy} L${a[0]} ${a[1]}" stroke="#536862" stroke-width="1.8" fill="none" marker-end="url(#arrow)"/>`));
 T('axis-functions',ox,oy-500,'FUNZIONI',18,{fill:AXES.functions.color,weight:'bold',anchor:'middle'});
 T('axis-clients',ox+lx+35,oy+36,'CLIENTI',18,{fill:AXES.clients.color,weight:'bold',anchor:'middle'});
 T('axis-modes',ox-265,oy+250,'MODALITÀ',18,{fill:AXES.modes.color,weight:'bold',anchor:'middle'});
 p.functions.forEach((v,i)=>{const q=P(0,i+1,0);anchors[v.id]={axis:'functions',i,point:q};add(line([q[0]-6,q[1]],[q[0]+6,q[1]],AXES.functions.color,2.5));T(v.id,q[0]-20,q[1],v.label||`Funzione ${i+1}`,19,{anchor:'end',width:26,axis:'functions',itemId:v.id,halo:true});});
 p.modes.forEach((v,i)=>{const q=P(0,0,i+1);anchors[v.id]={axis:'modes',i,point:q};add(line([q[0]-5,q[1]-6],[q[0]+5,q[1]+6],AXES.modes.color,2.5));T(v.id,q[0]-20,q[1]+2,v.label||`Modalità ${i+1}`,19,{anchor:'end',width:24,axis:'modes',itemId:v.id,halo:true});});
 let graphBottom=880;
 const actualShares=shares(p);
 p.clients.forEach((v,i)=>{const a=P(i,0,nm),b=P(i+1,0,nm),mid=(a[0]+b[0])/2,span=b[0]-a[0];anchors[v.id]={axis:'clients',i,point:[mid,a[1]+35]};add(`<path d="M${a[0]+3} ${a[1]+10} v9 H${Math.max(a[0]+3,b[0]-3)} v-9" stroke="#c9d5cc" stroke-width="1.5" fill="none"/>`);
  const label=(v.label||`Cliente ${i+1}`)+(p.weightsOn&&sum(p.clients.map(v=>Math.max(0,num(v.weight))))>0?' ('+fmtN(actualShares[i]*100)+'%)':'');
  let x=mid,y=a[1]+44,width=Math.max(8,Math.floor(span/10));
  if(p.weightsOn&&span<90){y+=75+i*21;x=mid;add(line([mid,a[1]+20],[mid,y-14],'#c6d3ce',1));width=28;}
  const count=wrap(label,width).length;T(v.id,x,y+(count-1)*10.8,label,19,{anchor:'middle',width,axis:'clients',itemId:v.id,halo:true});graphBottom=Math.max(graphBottom,y+count*23+18);
 });
 let fy=graphBottom+24;
 const footer=(key,title,value)=>{if(!String(value).trim())return;const lines=wrap(`${title}${value}`,135);T(key,70,fy+(lines.length-1)*10,lines.join('\n'),17,{fill:'#64786f',width:135});fy+=lines.length*21+13;};
 if(p.criteriaOn&&p.criteria.trim())footer('criteria','Criterio di segmentazione: ',p.criteria);
 if(p.weightsOn&&p.weightsBasis.trim())footer('weights','Base dei pesi: ',p.weightsBasis);
 if(p.revenueOn&&String(p.revenue).trim())footer('revenue','Fatturato stimato: ',`${fmtN(num(p.revenue))} ${p.revenueUnit}${p.revenuePeriod?' · '+p.revenuePeriod:''}`);
 if(p.revenueOn&&p.revenueMethod.trim())footer('method','Metodo di stima: ',p.revenueMethod);
 if(p.legend.trim())footer('legend','Legenda: ',p.legend);
 const contentBottom=Math.max(fy,...labels.map(l=>l.y+l.lines.length*l.size*.65));const scale=Math.min(1,1010/contentBottom),tx=(W-W*scale)/2;
 if(scale<1){for(const l of labels){l.x=l.x*scale+tx;l.y*=scale;l.size*=scale;}for(const a of Object.values(anchors))a.point=[a.point[0]*scale+tx,a.point[1]*scale];}
 return{gfx:scale<1?[`<g transform="translate(${tx} 0) scale(${scale})">${gfx.join('')}</g>`]:gfx,labels,anchors,scale,tx};
}
const DEFS='<defs><marker id="arrow" viewBox="0 0 10 10" refX="8" refY="5" markerWidth="7" markerHeight="7" orient="auto"><path d="M0 0 10 5 0 10Z" fill="#536862"/></marker></defs>';
function textSvg(l,interactive=false){const y=l.y-(l.lines.length-1)*l.size*1.15/2+l.size*.35;return `<g data-label="${esc(l.key)}"${interactive?` class="draggable" tabindex="0" role="graphics-symbol" aria-label="${esc(l.s)}"${l.axis?` data-axis="${l.axis}" data-item="${l.itemId}"`:''}`:''}><text x="${l.x}" y="${y}" font-family="${l.family}" font-size="${l.size}" fill="${l.fill}" font-weight="${l.weight}" text-anchor="${l.anchor}"${l.halo?' stroke="white" stroke-width="5" stroke-linejoin="round" paint-order="stroke"':''}>${l.lines.map((v,i)=>`<tspan x="${l.x}" dy="${i?l.size*1.15:0}">${esc(v)}</tspan>`).join('')}</text></g>`;}
function fullSvg(p,plates,options={}){const m=buildModel(p,options.interactive);const svg=`<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}">${DEFS}${options.bg===false?'':`<rect width="${W}" height="${H}" fill="white"/>`}${m.gfx.join('')}${options.labels===false?'':m.labels.map(l=>textSvg(l,options.interactive)).join('')}</svg>`;return{svg,m};}
