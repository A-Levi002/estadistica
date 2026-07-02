/**
 * ════════════════════════════════════════════════════════════════
 *  app.js — Estadística · Hábitos de Sueño UPDS
 * ════════════════════════════════════════════════════════════════
 *
 *  Estructura:
 *    1.  Estado global (data, charts)
 *    2.  Navegación entre secciones
 *    3.  Excel / Tally uploader (SheetJS)
 *    4.  Cálculos estadísticos (stats, recalc, calcMuestra)
 *    5.  Tablas de frecuencia estilo SPSS (updateFreqTable)
 *    6.  Gráficos Chart.js (updateAllCharts)
 *    7.  Probabilidades (updateProbabilities, calcProbInteractive)
 *    8.  Intervalos de confianza (updateIC, positionCI)
 *    9.  Prueba de hipótesis (updateHypSteps, animateHypothesis)
 *   10.  Resumen general (updateResumen)
 *   11.  Pizarra — motor universal (drawOnBoard, tokenize)
 *   12.  Pizarra — Estadísticas descriptivas
 *   13.  Pizarra — Probabilidades
 *   14.  Pizarra — Intervalos de confianza
 *   15.  Pizarra — Hipótesis
 *   16.  Notificaciones (notify)
 *   17.  Inicialización
 *
 *  Dependencias externas:
 *    - Chart.js 4.4.1
 *    - SheetJS (xlsx) 0.18.5
 *
 * ════════════════════════════════════════════════════════════════
 */

// ════════════════════════════════════════════
// DATOS GLOBALES
// ════════════════════════════════════════════
let data=[];let charts={};

// ════════════════════════════════════════════
// NAV
// ════════════════════════════════════════════
function showTab(tab,btn){
  document.querySelectorAll('.section').forEach(s=>s.classList.remove('active'));
  document.querySelectorAll('nav button').forEach(b=>b.classList.remove('active'));
  document.getElementById('sec-'+tab).classList.add('active');
  if(btn) btn.classList.add('active');
  updateAllCharts();updateIC();updateHypSteps();updateResumen();
}

// ════════════════════════════════════════════
// TEMAS GLOBALES — cambian TODO el sitio
// ════════════════════════════════════════════
const THEMES=[
  {id:'default', label:'Bosque oscuro',   icon:'🌿'},
  {id:'theme-ocean',  label:'Océano',     icon:'🌊'},
  {id:'theme-violet', label:'Violeta',    icon:'🔮'},
  {id:'theme-amber',  label:'Ámbar',      icon:'🌙'},
  {id:'theme-rose',   label:'Rosa',       icon:'🌸'},
  {id:'theme-emerald',label:'Esmeralda',  icon:'💎'},
];
let themeIndex=0;

function cycleBackground(){
  // Remove all theme classes from body
  THEMES.forEach(t=>document.body.classList.remove(t.id));
  themeIndex=(themeIndex+1)%THEMES.length;
  const t=THEMES[themeIndex];
  if(t.id!=='default') document.body.classList.add(t.id);
  // Update button tooltip
  const btn=document.getElementById('bg-toggle-btn');
  if(btn) btn.title=t.label;
  notify(t.icon+' Tema: '+t.label);
}
function toggleTheme(){ cycleBackground(); }

// ════════════════════════════════════════════
// EXCEL / TALLY UPLOADER
// ════════════════════════════════════════════
let rawRows=[];     // filas crudas del Excel
let rawHeaders=[];  // cabeceras detectadas

// Drag & drop en la zona de carga
document.addEventListener('DOMContentLoaded',()=>{
  const zone=document.getElementById('upload-zone');
  if(!zone) return;
  zone.addEventListener('dragover',e=>{e.preventDefault();zone.style.borderColor='var(--accent)';zone.style.background='rgba(255,255,255,.03)';});
  zone.addEventListener('dragleave',()=>{zone.style.borderColor='var(--border2)';zone.style.background='var(--surface2)';});
  zone.addEventListener('drop',e=>{e.preventDefault();zone.style.borderColor='var(--border2)';zone.style.background='var(--surface2)';const f=e.dataTransfer.files[0];if(f) processFile(f);});
});

function handleExcelUpload(evt){
  const f=evt.target.files[0];
  if(f) processFile(f);
  evt.target.value=''; // reset input
}

function processFile(file){
  showStatus('Leyendo archivo...','info');
  const reader=new FileReader();
  reader.onload=e=>{
    try{
      const wb=XLSX.read(e.target.result,{type:'array'});
      const ws=wb.Sheets[wb.SheetNames[0]];
      const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});
      if(rows.length<2){showStatus('El archivo parece vacío o no tiene datos.','error');return;}
      rawHeaders=rows[0].map(h=>String(h).trim());
      rawRows=rows.slice(1).filter(r=>r.some(c=>c!==''));
      showStatus(`Archivo leído: ${rawRows.length} respuestas, ${rawHeaders.length} columnas.`,'ok');
      showColMapper();
    }catch(err){
      showStatus('Error al leer el archivo: '+err.message,'error');
    }
  };
  reader.readAsArrayBuffer(file);
}

function showStatus(msg,type){
  const el=document.getElementById('upload-status');
  el.style.display='block';
  const colors={info:'rgba(96,165,250,.12)',ok:'rgba(74,222,128,.1)',error:'rgba(248,113,113,.1)'};
  const textColors={info:'var(--blue)',ok:'var(--green)',error:'var(--red)'};
  el.style.background=colors[type]||colors.info;
  el.style.color=textColors[type]||textColors.info;
  el.style.border=`1px solid ${textColors[type]||textColors.info}`;
  el.textContent=msg;
}

// ════════════════════════════════════════════
// CAMPOS COMPLETOS — 15 preguntas de Tally
// ════════════════════════════════════════════
const CAMPOS=[
  {key:'edad',      label:'1. Edad',                         hint:'¿Cuántos años tienes?'},
  {key:'genero',    label:'2. Sexo',                         hint:'¿Cuál es tu sexo?'},
  {key:'carrera',   label:'3. Carrera',                      hint:'¿Qué carrera estudias?'},
  {key:'semestre',  label:'4. Semestre',                     hint:'¿En qué semestre estás?'},
  {key:'horario',   label:'5. Horario de clases',            hint:'¿En qué horario estudias?'},
  {key:'levanta_clases', label:'6. Hora de levantarse (clases)', hint:'¿A qué hora te levantas en días de clases?'},
  {key:'duerme_clases',  label:'7. Hora de dormirse (clases)',   hint:'¿A qué hora te duermes noche anterior a clases?'},
  {key:'duerme_finde',   label:'7b. Hora dormirse (fin de semana)', hint:'¿A qué hora te duermes los fines de semana?'},
  {key:'levanta_finde',  label:'8. Hora levantarse (fin de semana)', hint:'¿A qué hora te levantas los fines de semana?'},
  {key:'horas',     label:'9. Horas de sueño (clases)',      hint:'¿Cuántas horas duermes en día de clases?'},
  {key:'bien',      label:'10. ¿Horas suficientes?',         hint:'¿Sientes que las horas que duermes son suficientes?'},
  {key:'desvela',   label:'11. ¿Te quedas despierto hasta tarde?', hint:'¿Con qué frecuencia te quedas despierto hasta tarde?'},
  {key:'somnolencia',label:'12. Somnolencia en clases',      hint:'¿Has sentido somnolencia durante clases esta semana?'},
  {key:'cafe',      label:'13. Café / bebidas energéticas',  hint:'¿Cuántas tazas consumes al día?'},
  {key:'pantalla',  label:'14. Uso de pantallas antes de dormir', hint:'¿Utilizas pantallas justo antes de dormir?'},
  {key:'actividad', label:'15. Actividad física regular',    hint:'¿Realizas actividad física de forma regular?'},
];

function showColMapper(){
  document.getElementById('upload-zone').style.display='none';
  const grid=document.getElementById('mapper-grid');
  grid.innerHTML=CAMPOS.map(c=>{
    const auto=autoDetect(c.key,rawHeaders);
    const opts=`<option value="-1">— no usar —</option>`+
      rawHeaders.map((h,i)=>`<option value="${i}" ${i===auto?'selected':''}>${h}</option>`).join('');
    return`<div style="background:var(--surface2);border:1px solid var(--border);border-radius:5px;padding:10px">
      <div style="font-size:11px;font-weight:600;color:var(--accent);margin-bottom:2px">${c.label}</div>
      <div style="font-size:10px;color:var(--muted);margin-bottom:6px">${c.hint}</div>
      <select id="map-${c.key}" style="width:100%;font-size:11px">${opts}</select>
    </div>`;
  }).join('');
  document.getElementById('col-mapper').style.display='block';
}

function autoDetect(key,headers){
  const P={
    edad:          [/^\d+\.\s*¿cuántos años/i, /años tienes/i, /edad/i],
    genero:        [/cuál es tu sexo/i, /sexo/i, /género/i],
    carrera:       [/qué carrera estudias/i, /carrera/i, /facultad/i],
    semestre:      [/qué semestre estás/i, /semestre/i],
    horario:       [/qué horario estudias/i, /horario/i],
    levanta_clases:[/levantas normalmente en d.*as de clases/i, /levantarte.*día de clases/i],
    duerme_clases: [/dormirte la noche anterior/i, /dormirte.*noche anterior/i],
    duerme_finde:  [/dormirte habitualmente los fines/i, /dormirte.*fines de semana/i],
    levanta_finde: [/levantarte los fines de semana/i, /levantarte.*fines/i],
    horas:         [/cuántas horas duermes normalmente/i, /horas duermes/i],
    bien:          [/horas que duermes son suficientes/i, /suficientes/i],
    desvela:       [/frecuencia te quedas despierto/i, /quedas despierto/i],
    somnolencia:   [/somnolencia o cansancio durante clases/i, /somnolencia/i],
    cafe:          [/tazas de café o bebidas/i, /café.*energéticas/i, /bebidas energéticas/i],
    pantalla:      [/utilizas pantallas.*justo antes/i, /pantallas.*dormir/i, /pantalla/i],
    // Muy específico: "actividad física de forma regular" — evita matchear col 18 "actividades"
    actividad:     [/actividad física de forma regular/i, /realizas actividad física/i],
  };
  const pats=P[key]||[];
  for(let i=0;i<headers.length;i++){
    if(pats.some(p=>p.test(headers[i]))) return i;
  }
  return -1;
}

// ── Normalizadores por pregunta ──
function normEdad(v){
  if(/menos.*18/i.test(v)) return 17;
  if(/18.*20/i.test(v))    return 19;
  if(/20.*22/i.test(v))    return 21;
  if(/22.*m[aá]s/i.test(v))return 23;
  const n=parseFloat(v); return isNaN(n)?20:n;
}
function normSemestre(v){
  if(/1r|2d/i.test(v))  return 1;
  if(/3r|4t/i.test(v))  return 3;
  if(/5t|6t/i.test(v))  return 5;
  if(/7m|8v/i.test(v))  return 7;
  if(/9n|más|mas/i.test(v)) return 9;
  const n=parseInt(v); return isNaN(n)?1:n;
}
function normHoras(v){
  if(/m[aá]s.*8|9|10|11|12/i.test(v)) return 9;
  if(/4.*5/i.test(v))  return 4.5;
  if(/\b6\b/i.test(v)) return 6;
  if(/\b7\b/i.test(v)) return 7;
  if(/\b8\b/i.test(v)) return 8;
  const n=parseFloat(v); return isNaN(n)?null:n;
}
function normGenero(v){
  if(/^m(asc|asculino)?$/i.test(v)||v==='1') return 'Masculino';
  if(/^f(em|emenino)?$/i.test(v)||v==='2')   return 'Femenino';
  if(/prefiero|otro|no.*decir/i.test(v))      return 'Prefiero no decir';
  return v||'Otro';
}
function normBien(v){
  if(/^s[ií]/i.test(v)||/descansad/i.test(v)) return 'Sí, descansado/a';
  if(/a veces|depende/i.test(v))               return 'A veces';
  if(/^no/i.test(v)||/cansad/i.test(v))        return 'No, cansado/a';
  return v||'A veces';
}
function normDesvela(v){
  if(/nunca/i.test(v))          return 'Nunca';
  if(/1.*2/i.test(v))           return '1-2 veces/semana';
  if(/3.*4/i.test(v))           return '3-4 veces/semana';
  if(/casi.*todos|todos/i.test(v)) return 'Casi todos los días';
  return v||'Nunca';
}
function normSomnolencia(v){
  if(/^no$/i.test(v))           return 'No';
  if(/1.*2/i.test(v))           return '1-2 días';
  if(/3.*4/i.test(v))           return '3-4 días';
  if(/todos/i.test(v))          return 'Todos los días';
  return v||'No';
}
function normCafe(v){
  if(/ninguna|0/i.test(v))      return 'Ninguna';
  if(/1\s*taza|^1$/i.test(v))   return '1 taza';
  if(/2.*3/i.test(v))           return '2-3 tazas';
  if(/m[aá]s.*3|4|5/i.test(v)) return 'Más de 3';
  return v||'Ninguna';
}
function normHoraLev(v){
  // Normaliza rangos de hora a etiqueta corta
  if(!v||v.trim()==='') return null;
  const m=v.match(/(\d+:\d+)\s*(am|pm)/gi);
  if(m&&m.length>=1) return m[0];
  return v.trim().substring(0,20);
}
function normHoraDuerm(v){
  if(!v||v.trim()==='') return null;
  const m=v.match(/(\d+:\d+)\s*(am|pm)/gi);
  if(m&&m.length>=1) return m[0];
  return v.trim().substring(0,20);
}
function normSiNo(v){
  if(/^s[ií]/i.test(v)||v==='1') return 'Sí';
  if(/^no/i.test(v)||v==='0')    return 'No';
  return v||'No';
}

function normalize(val,key){
  const v=String(val).trim();
  if(!v) return null;
  switch(key){
    case 'edad':           return normEdad(v);
    case 'semestre':       return normSemestre(v);
    case 'horas':          return normHoras(v);
    case 'genero':         return normGenero(v);
    case 'bien':           return normBien(v);
    case 'desvela':        return normDesvela(v);
    case 'somnolencia':    return normSomnolencia(v);
    case 'cafe':           return normCafe(v);
    case 'pantalla':
    case 'actividad':      return normSiNo(v);
    case 'levanta_clases':
    case 'levanta_finde':  return normHoraLev(v);
    case 'duerme_clases':
    case 'duerme_finde':   return normHoraDuerm(v);
    default:               return v;
  }
}

function applyMapping(){
  const mapping={};
  CAMPOS.forEach(c=>{
    const sel=document.getElementById('map-'+c.key);
    mapping[c.key]=sel?parseInt(sel.value):-1;
  });
  if(mapping.horas===-1){showStatus('Debes mapear la columna de horas de sueño','error');return;}

  // Detectar columnas extra de levanta_clases (cols 2 y 3) y duerme_clases (cols 2 y 3)
  // Tally genera columnas duplicadas por lógica condicional — las fusionamos
  const levantaExtra = [];
  const duermeExtra  = [];
  rawHeaders.forEach((h,i)=>{
    if(i===mapping.levanta_clases) return;
    if(/levantas normalmente en d.*as de clases/i.test(h)) levantaExtra.push(i);
    if(/dormirte la noche anterior/i.test(h)) duermeExtra.push(i);
  });

  data=[];
  rawRows.forEach((row)=>{
    const get=(key)=>{
      const idx=mapping[key];
      return idx>=0&&idx<row.length?String(row[idx]).trim():'';
    };
    // Fusionar columnas duplicadas: tomar el primer valor no vacío
    const getLevanta=()=>{
      let v=get('levanta_clases');
      if(v) return v;
      for(const i of levantaExtra){ const c=String(row[i]||'').trim(); if(c) return c; }
      return '';
    };
    const getDuerme=()=>{
      let v=get('duerme_clases');
      if(v) return v;
      for(const i of duermeExtra){ const c=String(row[i]||'').trim(); if(c) return c; }
      return '';
    };

    const horas=normalize(get('horas'),'horas');
    if(horas===null||horas<1||horas>16) return;

    // Carrera: quitar "Facultad de " para que sea más corto
    const carreraRaw=get('carrera')||'—';
    const carrera=carreraRaw.replace(/^Facultad de /i,'');

    // Horario: tomar solo la parte antes del paréntesis
    const horarioRaw=get('horario')||'—';
    const horario=horarioRaw.split('(')[0].trim();

    data.push({
      id:            data.length+1,
      edad:          normalize(get('edad'),'edad')||20,
      edadLabel:     get('edad')||'—',
      genero:        normalize(get('genero'),'genero')||'—',
      carrera,
      carreraFull:   carreraRaw,
      semestre:      normalize(get('semestre'),'semestre')||1,
      semestreLabel: get('semestre')||'—',
      horario,
      horarioFull:   horarioRaw,
      levanta_clases:getLevanta()||'—',
      duerme_clases: getDuerme()||'—',
      duerme_finde:  normalize(get('duerme_finde'),'duerme_finde')||'—',
      levanta_finde: normalize(get('levanta_finde'),'levanta_finde')||'—',
      horas,
      bien:          normalize(get('bien'),'bien')||'A veces',
      desvela:       normalize(get('desvela'),'desvela')||'Nunca',
      somnolencia:   normalize(get('somnolencia'),'somnolencia')||'No',
      cafe:          normalize(get('cafe'),'cafe')||'Ninguna',
      pantalla:      normSiNo(get('pantalla'))||'No',
      actividad:     normSiNo(get('actividad'))||'No',
    });
  });

  if(data.length===0){showStatus('No se encontraron filas válidas. Revisa el mapeo.','error');return;}
  document.getElementById('col-mapper').style.display='none';
  showStatus('\u2713 '+data.length+' respuestas cargadas correctamente desde Tally.','ok');
  renderPreview();
  recalc();
  notify(data.length+' respuestas cargadas');
}



function renderPreview(){
  document.getElementById('preview-wrap').style.display='block';
  document.getElementById('row-count').textContent=data.length+' respuestas cargadas';
  document.getElementById('m-nreal').textContent=data.length;
  document.getElementById('hdr-n').textContent=data.length;
  const bC=v=>v.includes('Sí')||v==='Sí'?'var(--green)':v.includes('No')||v==='No'?'var(--red)':'var(--yellow)';
  // Update table headers
  const thead=document.querySelector('#preview-table thead tr');
  if(thead) thead.innerHTML=`<th>#</th><th>Edad</th><th>Sexo</th><th>Carrera</th><th>Sem.</th><th>Horario</th><th>Horas sueño</th><th>¿Suficiente?</th><th>Desvela</th><th>Somnolencia</th><th>Café</th><th>Pantallas</th><th>Act. física</th>`;
  document.getElementById('data-body').innerHTML=data.map(r=>`<tr>
    <td><span class="tag-n">${r.id}</span></td>
    <td>${r.edad}a</td>
    <td>${r.genero}</td>
    <td style="max-width:120px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:10px">${r.carrera.replace('Facultad de ','')}</td>
    <td>${r.semestre}°</td>
    <td style="font-size:10px">${r.horario.split(' ')[0]}</td>
    <td style="color:var(--blue);font-weight:700">${r.horas}h</td>
    <td style="color:${bC(r.bien)};font-size:10px">${r.bien}</td>
    <td style="font-size:10px;color:var(--muted2)">${r.desvela}</td>
    <td style="font-size:10px;color:var(--muted2)">${r.somnolencia}</td>
    <td style="font-size:10px">${r.cafe}</td>
    <td style="color:${r.pantalla==='Sí'?'var(--purple)':'var(--muted2)'}">${r.pantalla}</td>
    <td style="color:${r.actividad==='Sí'?'var(--yellow)':'var(--muted2)'}">${r.actividad}</td>
  </tr>`).join('');
}

function resetUpload(){
  data=[];rawRows=[];rawHeaders=[];
  document.getElementById('upload-zone').style.display='block';
  document.getElementById('upload-status').style.display='none';
  document.getElementById('col-mapper').style.display='none';
  document.getElementById('preview-wrap').style.display='none';
  document.getElementById('data-body').innerHTML='';
  document.getElementById('row-count').textContent='0 respuestas';
  document.getElementById('m-nreal').textContent='0';
  document.getElementById('hdr-n').textContent='0';
  recalc();
}

// renderTable kept as alias for compatibility
function renderTable(){ if(data.length) renderPreview(); }

// ════════════════════════════════════════════
// CÁLCULOS
// ════════════════════════════════════════════
function recalc(){
  calcMuestra();
  if(data.length<2) return;
  calcDescriptivo();updateFreqTable();updateAllCharts();
  updateProbabilities();updateIC();updateHypSteps();updateResumen();
  syncBoardInputsFromData();
}

function calcMuestra(){
  const N=+document.getElementById('pop-N').value||3200;
  const z=+document.getElementById('conf-level').value||1.96;
  const e=+document.getElementById('margin-e').value||0.05;
  const p=+document.getElementById('prop-p').value||0.5;
  const q=1-p,n0=(z*z*p*q)/(e*e),n=n0/(1+(n0-1)/N);
  const nc=Math.ceil(n0),na=Math.ceil(n);
  document.getElementById('m-ncalc').textContent=nc;
  document.getElementById('m-nadj').textContent=na;
  document.getElementById('formula-muestra').innerHTML=
    `<span class="cm">// Fórmula muestra infinita:</span>
n₀ = z² × p × q / e²
n₀ = <span class="vl">${z}²</span> × <span class="vl">${p}</span> × <span class="vl">${q}</span> / <span class="vl">${e}²</span>
n₀ = <span class="vl">${(z*z).toFixed(4)}</span> × <span class="vl">${(p*q).toFixed(4)}</span> / <span class="vl">${(e*e).toFixed(4)}</span>
n₀ = <span class="rs">${nc}</span>

<span class="cm">// Ajuste por población finita (N = ${N}):</span>
n = n₀ / (1 + (n₀−1)/N) = ${nc} / (1 + ${nc-1}/${N})
n = <span class="rs">${na}</span>  ← muestra final`;
}

function esSí(v){ return v && (v === 'Sí' || v.indexOf('Sí,')===0 || v.indexOf('Sí ')===0); }

function stats(){
  if(!data.length) return {};
  const horas=data.map(d=>d.horas).sort((a,b)=>a-b);
  const n=horas.length,mean=horas.reduce((a,b)=>a+b,0)/n;
  const median=n%2===0?(horas[n/2-1]+horas[n/2])/2:horas[Math.floor(n/2)];
  const variance=horas.reduce((a,b)=>a+(b-mean)**2,0)/(n-1);
  const std=Math.sqrt(variance);
  const bienSi   =data.filter(d=>esSí(d.bien)).length;
  const pantallaSi=data.filter(d=>d.pantalla==='Sí').length;
  const actividadSi=data.filter(d=>d.actividad==='Sí').length;
  const desvelaSi=data.filter(d=>d.desvela&&d.desvela!=='Nunca'&&d.desvela!=='—').length;
  const somnolenciaSi=data.filter(d=>d.somnolencia&&d.somnolencia!=='No'&&d.somnolencia!=='—').length;
  return{n,mean,median,variance,std,horas,
    bienSi,pantallaSi,actividadSi,desvelaSi,somnolenciaSi,
    pBien:bienSi/n,pPantalla:pantallaSi/n,
    pActividad:actividadSi/n,pDesvela:desvelaSi/n};
}

function calcDescriptivo(){
  const s=stats();
  document.getElementById('d-mean').textContent=s.mean.toFixed(2);
  document.getElementById('d-median').textContent=s.median.toFixed(1);
  document.getElementById('d-std').textContent=s.std.toFixed(2);
  document.getElementById('d-var').textContent=s.variance.toFixed(2);
}

function destroyChart(id){if(charts[id]){charts[id].destroy();delete charts[id];}}

function freqTable(horas){
  const classes=[[2,4],[4,6],[6,8],[8,10],[10,12],[12,14]];
  const labels=classes.map(c=>`[${c[0]}–${c[1]})`);
  const freqs=classes.map(c=>horas.filter(h=>h>=c[0]&&h<c[1]).length);
  const n=horas.length,relFreqs=freqs.map(f=>f/n);
  let cum=0;const cumFreqs=freqs.map(f=>{cum+=f;return cum;});
  return{labels,freqs,relFreqs,cumFreqs,n};
}

function updateFreqTable(){
  const s=stats(); if(!s.n) return;

  // ── Tabla resumen SPSS ──
  const skew=calcSkewness(s.horas, s.mean, s.std);
  const modeHoras=calcMode(s.horas);
  document.getElementById('spss-summary-body').innerHTML=`
    <tr><td>Horas de sueño</td><td>${s.n}</td><td>${Math.min(...s.horas).toFixed(1)}</td><td>${Math.max(...s.horas).toFixed(1)}</td>
        <td>${s.mean.toFixed(4)}</td><td>${s.median.toFixed(4)}</td><td>${modeHoras}</td>
        <td>${s.std.toFixed(4)}</td><td>${s.variance.toFixed(4)}</td><td>${skew.toFixed(4)}</td></tr>
    <tr><td>Edad</td>
        ${(()=>{const v=data.map(d=>d.edad);const mn=v.reduce((a,b)=>a+b,0)/v.length;const md=calcMedian(v);const vr=v.reduce((a,b)=>a+(b-mn)**2,0)/(v.length-1);const sd=Math.sqrt(vr);const sk=calcSkewness([...v].sort((a,b)=>a-b),mn,sd);return`<td>${v.length}</td><td>${Math.min(...v)}</td><td>${Math.max(...v)}</td><td>${mn.toFixed(4)}</td><td>${md.toFixed(4)}</td><td>${calcMode(v)}</td><td>${sd.toFixed(4)}</td><td>${vr.toFixed(4)}</td><td>${sk.toFixed(4)}</td>`})()}</tr>
    <tr><td>Semestre</td>
        ${(()=>{const v=data.map(d=>d.semestre);const mn=v.reduce((a,b)=>a+b,0)/v.length;const md=calcMedian(v);const vr=v.reduce((a,b)=>a+(b-mn)**2,0)/(v.length-1);const sd=Math.sqrt(vr);const sk=calcSkewness([...v].sort((a,b)=>a-b),mn,sd);return`<td>${v.length}</td><td>${Math.min(...v)}</td><td>${Math.max(...v)}</td><td>${mn.toFixed(4)}</td><td>${md.toFixed(4)}</td><td>${calcMode(v)}</td><td>${sd.toFixed(4)}</td><td>${vr.toFixed(4)}</td><td>${sk.toFixed(4)}</td>`})()}</tr>`;

  // ── Horas de sueño — por valor individual ──
  const horasUniq=[...new Set(s.horas)].sort((a,b)=>a-b);
  let cumH=0;
  const horasRows=horasUniq.map(v=>{
    const fi=s.horas.filter(h=>h===v).length;
    const fri=fi/s.n; cumH+=fi;
    const isMode=v==modeHoras;
    return`<tr class="${isMode?'tr-hl':''}"><td>${v} h${isMode?' ← moda':''}</td><td>${fi}</td><td>${(fri*100).toFixed(2)}%</td><td>${(cumH/s.n*100).toFixed(2)}%</td><td>${fri.toFixed(4)}</td><td>${(fri*100).toFixed(2)}</td></tr>`;
  });
  document.getElementById('freq-horas-body').innerHTML=horasRows.join('')+
    `<tr class="tr-total"><td>Total</td><td>${s.n}</td><td>100.00%</td><td>100.00%</td><td>1.0000</td><td>100.00</td></tr>`;
  document.getElementById('spss-horas-note').textContent=
    `N válido = ${s.n} · Media = ${s.mean.toFixed(4)} · Mediana = ${s.median.toFixed(4)} · Moda = ${modeHoras} h · Desv. estándar = ${s.std.toFixed(4)} · Rango = ${(Math.max(...s.horas)-Math.min(...s.horas)).toFixed(1)}`;

  // ── Función genérica para variables dicotómicas ──
  function freqDicot(vals, bodyId, noteId){
    const cats=[...new Set(vals)].sort();
    let cum=0;
    const rows=cats.map(c=>{
      const fi=vals.filter(v=>v===c).length;
      const fri=fi/vals.length; cum+=fi;
      return`<tr><td>${c}</td><td>${fi}</td><td>${(fri*100).toFixed(2)}%</td><td>${(cum/vals.length*100).toFixed(2)}%</td><td>${fri.toFixed(4)}</td><td>${(fri*100).toFixed(2)}</td></tr>`;
    });
    document.getElementById(bodyId).innerHTML=rows.join('')+
      `<tr class="tr-total"><td>Total</td><td>${vals.length}</td><td>100.00%</td><td>100.00%</td><td>1.0000</td><td>100.00</td></tr>`;
    if(noteId) document.getElementById(noteId).textContent=
      `N = ${vals.length} · Categorías: ${cats.join(', ')}`;
    return cats.map(c=>({cat:c,fi:vals.filter(v=>v===c).length}));
  }

  const bienCats  =freqDicot(data.map(d=>d.bien),    'freq-bien-body',    'spss-bien-note');
  const pantCats  =freqDicot(data.map(d=>d.pantalla), 'freq-pantalla-body','spss-pantalla-note');
  const siestaCats=freqDicot(data.map(d=>d.actividad),'freq-siesta-body',  'spss-siesta-note');

  // ── Género ──
  const genVals=data.map(d=>d.genero);
  const genCats=[...new Set(genVals)].sort();
  let cumG=0;
  document.getElementById('freq-genero-body').innerHTML=genCats.map(c=>{
    const fi=genVals.filter(v=>v===c).length; const fri=fi/genVals.length; cumG+=fi;
    return`<tr><td>${c}</td><td>${fi}</td><td>${(fri*100).toFixed(2)}%</td><td>${(cumG/genVals.length*100).toFixed(2)}%</td><td>${fri.toFixed(4)}</td><td>${(fri*100).toFixed(2)}</td></tr>`;
  }).join('')+`<tr class="tr-total"><td>Total</td><td>${genVals.length}</td><td>100.00%</td><td>100.00%</td><td>1.0000</td><td>100.00</td></tr>`;

  // ── Semestre ──
  const semVals=data.map(d=>d.semestre);
  const semCats=[...new Set(semVals)].sort((a,b)=>a-b);
  let cumS=0;
  document.getElementById('freq-semestre-body').innerHTML=semCats.map(c=>{
    const fi=semVals.filter(v=>v===c).length; const fri=fi/semVals.length; cumS+=fi;
    return`<tr><td>${c}°</td><td>${fi}</td><td>${(fri*100).toFixed(2)}%</td><td>${(cumS/semVals.length*100).toFixed(2)}%</td><td>${fri.toFixed(4)}</td><td>${(fri*100).toFixed(2)}</td></tr>`;
  }).join('')+`<tr class="tr-total"><td>Total</td><td>${semVals.length}</td><td>100.00%</td><td>100.00%</td><td>1.0000</td><td>100.00</td></tr>`;

  // ── Carrera ──
  const carCats=freqDicot(data.map(d=>d.carrera.replace('Facultad de ','')),'freq-carrera-body','');
  // ── Horario ──
  const horCats=freqDicot(data.map(d=>d.horario.split('(')[0].trim()),'freq-horario-body','');
  // ── Desvela ──
  const devCats=freqDicot(data.map(d=>d.desvela),'freq-desvela-body','');
  // ── Somnolencia ──
  const somCats=freqDicot(data.map(d=>d.somnolencia),'freq-somnolencia-body','');
  // ── Café ──
  const cafCats=freqDicot(data.map(d=>d.cafe),'freq-cafe-body','');

  // Pasar datos a charts
  window._chartData={horasUniq,s,bienCats,pantCats,siestaCats,genCats,genVals,semCats,semVals,
    carCats,horCats,devCats,somCats,cafCats};
}

function calcSkewness(sorted, mean, std){
  if(!sorted.length||std===0) return 0;
  const n=sorted.length;
  return sorted.reduce((a,x)=>a+((x-mean)/std)**3,0)*n/((n-1)*(n-2)||1);
}
function calcMedian(arr){
  const s=[...arr].sort((a,b)=>a-b); const n=s.length;
  return n%2===0?(s[n/2-1]+s[n/2])/2:s[Math.floor(n/2)];
}
function calcMode(arr){
  const freq={};arr.forEach(v=>{freq[v]=(freq[v]||0)+1;});
  let mx=0,mo=arr[0];
  Object.entries(freq).forEach(([k,v])=>{if(v>mx){mx=v;mo=+k||k;}});
  return mo;
}

function updateAllCharts(){
  if(!data.length) return;
  const cd=window._chartData||{};
  const s=cd.s||stats();
  const horasUniq=cd.horasUniq||[...new Set(s.horas)].sort((a,b)=>a-b);

  const ids=['hist','ojiva','bien','bien-pie','pantalla-bar','pantalla',
             'siesta-bar','siesta-pie','genero-bar','genero-pie','semestre-bar','semestre-line'];
  ids.forEach(id=>destroyChart(id));

  const CO={grid:'rgba(255,255,255,.04)',ticks:'#666'};
  const base=()=>({responsive:true,maintainAspectRatio:false,
    plugins:{legend:{display:false},tooltip:{backgroundColor:'#111',borderColor:'#333',borderWidth:1,titleColor:'#eee',bodyColor:'#999'}},
    scales:{x:{grid:{color:CO.grid},ticks:{color:CO.ticks}},y:{grid:{color:CO.grid},ticks:{color:CO.ticks},beginAtZero:true}}});
  const pie_opts=(pos='bottom')=>({responsive:true,maintainAspectRatio:false,
    plugins:{legend:{position:pos,labels:{color:'#888',boxWidth:10,font:{size:11}}},
             tooltip:{backgroundColor:'#111',borderColor:'#333',borderWidth:1,titleColor:'#eee',bodyColor:'#999'}}});

  const PALETTE=['rgba(96,165,250,.75)','rgba(74,222,128,.75)','rgba(251,191,36,.75)','rgba(167,139,250,.75)','rgba(248,113,113,.75)','rgba(34,211,238,.75)'];
  const PAL_B=['rgba(96,165,250,.3)','rgba(74,222,128,.3)','rgba(251,191,36,.3)','rgba(167,139,250,.3)','rgba(248,113,113,.3)','rgba(34,211,238,.3)'];

  // Histograma horas
  const horasFi=horasUniq.map(v=>s.horas.filter(h=>h===v).length);
  const c1=document.getElementById('chart-hist');
  if(c1) charts.hist=new Chart(c1,{type:'bar',data:{labels:horasUniq.map(v=>v+'h'),datasets:[{label:'Frecuencia',data:horasFi,backgroundColor:'rgba(96,165,250,.3)',borderColor:'rgba(96,165,250,.8)',borderWidth:1.5,borderRadius:4}]},options:{...base(),plugins:{...base().plugins,title:{display:true,text:'Distribución de horas de sueño',color:'#888',font:{size:11}}}}});

  // Ojiva
  let cumO=0; const cumArr=[0,...horasUniq.map(v=>{cumO+=s.horas.filter(h=>h===v).length;return cumO;})];
  const c2=document.getElementById('chart-ojiva');
  if(c2) charts.ojiva=new Chart(c2,{type:'line',data:{labels:['',...horasUniq.map(v=>v+'h')],datasets:[{label:'Fi acum.',data:cumArr,borderColor:'rgba(251,191,36,.9)',backgroundColor:'rgba(251,191,36,.08)',tension:.35,fill:true,pointRadius:4,pointBackgroundColor:'rgba(251,191,36,.9)'}]},options:{...base(),plugins:{...base().plugins,title:{display:true,text:'Ojiva — frecuencia acumulada',color:'#888',font:{size:11}}}}});

  // Duerme bien — barras
  if(cd.bienCats){
    const bl=cd.bienCats.map(c=>c.cat),bf=cd.bienCats.map(c=>c.fi);
    const bColors=bl.map(l=>l==='Sí'||l==='Si'?'rgba(74,222,128,.35)':'rgba(248,113,113,.35)');
    const bBorders=bl.map(l=>l==='Sí'||l==='Si'?'rgba(74,222,128,.9)':'rgba(248,113,113,.9)');
    const c3=document.getElementById('chart-bien');
    if(c3) charts.bien=new Chart(c3,{type:'bar',data:{labels:bl,datasets:[{data:bf,backgroundColor:bColors,borderColor:bBorders,borderWidth:1.5,borderRadius:4}]},options:{...base(),plugins:{...base().plugins,title:{display:true,text:'¿Duerme bien?',color:'#888',font:{size:11}}}}});
    const c3p=document.getElementById('chart-bien-pie');
    if(c3p) charts['bien-pie']=new Chart(c3p,{type:'doughnut',data:{labels:bl,datasets:[{data:bf,backgroundColor:bColors,borderColor:bBorders,borderWidth:1.5}]},options:pie_opts()});
  }

  // Pantallas
  if(cd.pantCats){
    const pl=cd.pantCats.map(c=>c.cat),pf=cd.pantCats.map(c=>c.fi);
    const pc=pl.map((_,i)=>PALETTE[i]||PALETTE[0]),pcb=pl.map((_,i)=>PAL_B[i]||PAL_B[0]);
    const c4b=document.getElementById('chart-pantalla-bar');
    if(c4b) charts['pantalla-bar']=new Chart(c4b,{type:'bar',data:{labels:pl,datasets:[{data:pf,backgroundColor:pcb,borderColor:pc,borderWidth:1.5,borderRadius:4}]},options:{...base(),plugins:{...base().plugins,title:{display:true,text:'Uso de pantallas',color:'#888',font:{size:11}}}}});
    const c4=document.getElementById('chart-pantalla');
    if(c4) charts.pantalla=new Chart(c4,{type:'doughnut',data:{labels:pl,datasets:[{data:pf,backgroundColor:pcb,borderColor:pc,borderWidth:1.5}]},options:pie_opts()});
  }

  // Siestas
  if(cd.siestaCats){
    const sl=cd.siestaCats.map(c=>c.cat),sf=cd.siestaCats.map(c=>c.fi);
    const sc=sl.map((_,i)=>PALETTE[i+2]||PALETTE[0]),scb=sl.map((_,i)=>PAL_B[i+2]||PAL_B[0]);
    const c5b=document.getElementById('chart-siesta-bar');
    if(c5b) charts['siesta-bar']=new Chart(c5b,{type:'bar',data:{labels:sl,datasets:[{data:sf,backgroundColor:scb,borderColor:sc,borderWidth:1.5,borderRadius:4}]},options:{...base(),plugins:{...base().plugins,title:{display:true,text:'Toma de siestas',color:'#888',font:{size:11}}}}});
    const c5p=document.getElementById('chart-siesta-pie');
    if(c5p) charts['siesta-pie']=new Chart(c5p,{type:'doughnut',data:{labels:sl,datasets:[{data:sf,backgroundColor:scb,borderColor:sc,borderWidth:1.5}]},options:pie_opts()});
  }

  // Género
  if(cd.genCats&&cd.genVals){
    const gl=cd.genCats,gf=gl.map(c=>cd.genVals.filter(v=>v===c).length);
    const gc=gl.map((_,i)=>PALETTE[i]||PALETTE[0]),gcb=gl.map((_,i)=>PAL_B[i]||PAL_B[0]);
    const c6b=document.getElementById('chart-genero-bar');
    if(c6b) charts['genero-bar']=new Chart(c6b,{type:'bar',data:{labels:gl,datasets:[{data:gf,backgroundColor:gcb,borderColor:gc,borderWidth:1.5,borderRadius:4}]},options:{...base(),plugins:{...base().plugins,title:{display:true,text:'Distribución por género',color:'#888',font:{size:11}}}}});
    const c6p=document.getElementById('chart-genero-pie');
    if(c6p) charts['genero-pie']=new Chart(c6p,{type:'doughnut',data:{labels:gl,datasets:[{data:gf,backgroundColor:gcb,borderColor:gc,borderWidth:1.5}]},options:pie_opts()});
  }

  // Semestre
  if(cd.semCats&&cd.semVals){
    const sl2=cd.semCats.map(c=>c+'°'),sf2=cd.semCats.map(c=>cd.semVals.filter(v=>v===c).length);
    const c7b=document.getElementById('chart-semestre-bar');
    if(c7b) charts['semestre-bar']=new Chart(c7b,{type:'bar',data:{labels:sl2,datasets:[{data:sf2,backgroundColor:'rgba(167,139,250,.3)',borderColor:'rgba(167,139,250,.8)',borderWidth:1.5,borderRadius:4}]},options:{...base(),plugins:{...base().plugins,title:{display:true,text:'Frecuencia por semestre',color:'#888',font:{size:11}}}}});
    const c7l=document.getElementById('chart-semestre-line');
    if(c7l) charts['semestre-line']=new Chart(c7l,{type:'line',data:{labels:sl2,datasets:[{data:sf2,borderColor:'rgba(167,139,250,.9)',backgroundColor:'rgba(167,139,250,.08)',tension:.3,fill:true,pointRadius:4,pointBackgroundColor:'rgba(167,139,250,.9)'}]},options:{...base(),plugins:{...base().plugins,title:{display:true,text:'Polígono de frecuencias — semestre',color:'#888',font:{size:11}}}}});
  }

  // Helper para crear par de gráficos (bar + pie) desde cats
  function makePair(cats,barId,pieId,color,title){
    if(!cats||!cats.length) return;
    destroyChart(barId); destroyChart(pieId);
    const labels=cats.map(c=>c.cat), freqs=cats.map(c=>c.fi);
    const colors=labels.map((_,i)=>{const h=PALETTE[i%PALETTE.length];return h;});
    const colorsB=labels.map((_,i)=>PAL_B[i%PAL_B.length]);
    const cb=document.getElementById(barId);
    if(cb) charts[barId]=new Chart(cb,{type:'bar',data:{labels,datasets:[{data:freqs,backgroundColor:colorsB,borderColor:colors,borderWidth:1.5,borderRadius:4}]},options:{...base(),plugins:{...base().plugins,title:{display:true,text:title,color:'#888',font:{size:11}}}}});
    const cp=document.getElementById(pieId);
    if(cp) charts[pieId]=new Chart(cp,{type:'doughnut',data:{labels,datasets:[{data:freqs,backgroundColor:colorsB,borderColor:colors,borderWidth:1.5}]},options:pie_opts()});
  }

  makePair(cd.carCats,'chart-carrera-bar','chart-carrera-pie','#38bdf8','Carrera / Facultad');
  makePair(cd.horCats,'chart-horario-bar','chart-horario-pie','#2dd4bf','Horario de clases');
  makePair(cd.devCats,'chart-desvela-bar','chart-desvela-pie','#f87171','Frecuencia de desvelo');
  makePair(cd.somCats,'chart-somnolencia-bar','chart-somnolencia-pie','#fbbf24','Somnolencia en clases');
  makePair(cd.cafCats,'chart-cafe-bar','chart-cafe-pie','#c084fc','Café / bebidas energéticas');
}

// ════════════════════════════════════════════
// PROBABILIDADES
// ════════════════════════════════════════════
function probCard(label,p,color){
  return`<div class="prob-card"><div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:3px"><span style="font-size:12px;color:var(--muted2)">${label}</span><span style="font-size:18px;font-weight:700;color:${color||'var(--accent)'}">${(p*100).toFixed(1)}%</span></div><div class="prob-bar-wrap"><div class="prob-bar" style="width:${Math.min(p*100,100).toFixed(1)}%;background:${color||'var(--accent)'}"></div></div><div style="font-size:10px;color:var(--muted)">p = ${p.toFixed(4)} · ${Math.round(p*(stats().n||1))} de ${stats().n||'?'}</div></div>`;
}
function updateProbabilities(){
  const s=stats();if(!s.n) return;
  const n=s.n;
  const hb=data.filter(d=>d.horas<=6).length;
  const ha=data.filter(d=>d.horas>=8).length;
  const desv=data.filter(d=>d.desvela&&d.desvela!=='Nunca'&&d.desvela!=='—').length;
  const somn=data.filter(d=>d.somnolencia&&d.somnolencia!=='No'&&d.somnolencia!=='—').length;
  const cafe=data.filter(d=>d.cafe&&d.cafe!=='Ninguna'&&d.cafe!=='—').length;

  document.getElementById('prob-simple-list').innerHTML=
    probCard('P( horas de sueño suficientes )',s.pBien,'var(--green)')+
    probCard('P( horas insuficientes )',1-s.pBien,'var(--red)')+
    probCard('P( usa pantallas antes de dormir )',s.pPantalla,'var(--purple)')+
    probCard('P( realiza actividad física )',s.pActividad,'var(--yellow)')+
    probCard('P( se desvela frecuentemente )',desv/n,'var(--orange, #f97316)')+
    probCard('P( somnolencia en clases )',somn/n,'var(--blue)')+
    probCard('P( consume café/energéticas )',cafe/n,'#c084fc')+
    probCard('P( horas ≤ 6 )',hb/n,'var(--red)')+
    probCard('P( horas ≥ 8 )',ha/n,'var(--blue)');

  const pantallaSi=data.filter(d=>d.pantalla==='Sí').length;
  const actividadSi=data.filter(d=>d.actividad==='Sí').length;
  const bYp=data.filter(d=>esSí(d.bien)&&d.pantalla==='Sí').length;
  const bYa=data.filter(d=>esSí(d.bien)&&d.actividad==='Sí').length;
  const nYp=data.filter(d=>!esSí(d.bien)&&d.pantalla==='Sí').length;
  const pYd=data.filter(d=>d.pantalla==='Sí'&&d.desvela&&d.desvela!=='Nunca').length;
  const dYs=data.filter(d=>d.desvela&&d.desvela!=='Nunca'&&d.somnolencia&&d.somnolencia!=='No').length;

  document.getElementById('prob-conjunta-list').innerHTML=
    probCard('P( horas suficientes ∩ usa pantallas )',bYp/n,'var(--purple)')+
    probCard('P( horas suficientes | usa pantallas )',pantallaSi>0?bYp/pantallaSi:0,'var(--blue)')+
    probCard('P( insuficiente | usa pantallas )',pantallaSi>0?nYp/pantallaSi:0,'var(--red)')+
    probCard('P( suficiente ∩ act. física )',bYa/n,'var(--green)')+
    probCard('P( pantallas ∩ se desvela )',pYd/n,'var(--purple)')+
    probCard('P( se desvela ∩ somnolencia )',dYs/n,'#f97316');
  calcProbInteractive();
}

function getFilter(k){
  if(k==='bien_si')     return function(d){ return esSí(d.bien); };
  if(k==='bien_no')     return function(d){ return d.bien && !esSí(d.bien); };
  if(k==='pantalla_si') return function(d){ return d.pantalla === 'Sí'; };
  if(k==='pantalla_no') return function(d){ return d.pantalla === 'No'; };
  if(k==='actividad_si')return function(d){ return d.actividad === 'Sí'; };
  if(k==='desvela_si')  return function(d){ return d.desvela && d.desvela !== 'Nunca' && d.desvela !== '—'; };
  if(k==='horas_bajo')  return function(d){ return d.horas <= 6; };
  if(k==='horas_alto')  return function(d){ return d.horas >= 8; };
  return function(){ return true; };
}
function getLabel(k){const m={bien_si:'Horas suficientes',bien_no:'No suficientes',pantalla_si:'Usa pantallas',pantalla_no:'No usa pantallas',actividad_si:'Actividad física',desvela_si:'Se desvela',horas_bajo:'Horas ≤ 6',horas_alto:'Horas ≥ 8'};return m[k]||k;}
function calcProbInteractive(){
  const s=stats();if(!s.n) return;
  const kA=document.getElementById('prob-varA').value,kB=document.getElementById('prob-varB').value;
  const fA=getFilter(kA),fB=getFilter(kB);
  const cA=data.filter(fA).length,cB=data.filter(fB).length,cAB=data.filter(d=>fA(d)&&fB(d)).length;
  const pA=cA/s.n,pB=cB/s.n,pAB=cAB/s.n,pAgB=cB>0?cAB/cB:0;
  const lA=getLabel(kA),lB=getLabel(kB);
  document.getElementById('prob-result-box').innerHTML=`<div class="fbox">
<span class="cm">// n = ${s.n} encuestados</span>
P(A) = P(${lA}) = ${cA}/${s.n} = <span class="rs">${pA.toFixed(4)}</span>  [${(pA*100).toFixed(1)}%]
P(B) = P(${lB}) = ${cB}/${s.n} = <span class="rs">${pB.toFixed(4)}</span>  [${(pB*100).toFixed(1)}%]

<span class="cm">// Conjunta — ¿cuántos cumplen AMBAS?</span>
P(A∩B) = ${cAB}/${s.n} = <span class="rs">${pAB.toFixed(4)}</span>  [${(pAB*100).toFixed(1)}%]

<span class="cm">// ¿Independientes? P(A)×P(B) = ${(pA*pB).toFixed(4)}  vs  P(A∩B) = ${pAB.toFixed(4)}</span>
<span class="op">${Math.abs(pA*pB-pAB)<0.05?'≈ Aproximadamente independientes':'≠ Dependientes (uno influye en el otro)'}</span>

<span class="cm">// Condicional — dado que B ocurrió, ¿qué tan probable es A?</span>
P(A|B) = P(A∩B)/P(B) = ${pAB.toFixed(4)}/${pB.toFixed(4)} = <span class="rs">${pAgB.toFixed(4)}</span>  [${(pAgB*100).toFixed(1)}%]</div>`;
}

// ════════════════════════════════════════════
// INTERVALOS — campana normal + tabla unificada
// ════════════════════════════════════════════
function normalPDF(x,mu,sigma){return Math.exp(-0.5*((x-mu)/sigma)**2)/(sigma*Math.sqrt(2*Math.PI));}

function drawBellChart(canvasId,mu,sigma,lo,hi,label,color){
  destroyChart(canvasId);
  const c=document.getElementById(canvasId); if(!c) return;
  const pts=120;
  const xMin=mu-4*sigma, xMax=mu+4*sigma;
  const xs=Array.from({length:pts},(_,i)=>xMin+(xMax-xMin)*i/(pts-1));
  const ys=xs.map(x=>normalPDF(x,mu,sigma));
  const fillColors=xs.map(x=>x>=lo&&x<=hi?color.fill:color.out);
  charts[canvasId]=new Chart(c,{
    type:'line',
    data:{
      labels:xs.map(x=>x.toFixed(3)),
      datasets:[
        // Área sombreada IC
        {type:'bar',label:'IC 95%',data:xs.map((x,i)=>x>=lo&&x<=hi?ys[i]:null),
         backgroundColor:color.fill,borderWidth:0,barPercentage:1,categoryPercentage:1},
        // Curva normal
        {type:'line',label:'Distribución normal',data:ys,
         borderColor:color.line,borderWidth:2,pointRadius:0,tension:.4,fill:false},
      ]
    },
    options:{
      responsive:true,maintainAspectRatio:false,
      plugins:{
        legend:{display:false},
        tooltip:{enabled:false},
        annotation:{},
      },
      scales:{
        x:{
          ticks:{
            maxTicksLimit:7,
            callback:(v,i)=>{
              const x=xs[i];
              if(Math.abs(x-mu)<(xMax-xMin)/pts*2) return label;
              if(Math.abs(x-lo)<(xMax-xMin)/pts*2) return lo.toFixed?lo.toFixed(3):lo;
              if(Math.abs(x-hi)<(xMax-xMin)/pts*2) return hi.toFixed?hi.toFixed(3):hi;
              return '';
            },
            color:'#666',font:{size:10}
          },
          grid:{color:'rgba(255,255,255,.04)'}
        },
        y:{display:false,grid:{display:false}}
      }
    }
  });
}

function updateIC(){
  const s=stats();if(!s.n) return;
  const z=+document.getElementById('conf-level').value||1.96,n=s.n;

  // IC Proporción
  const ph=s.pBien,se_p=Math.sqrt(ph*(1-ph)/n),me_p=z*se_p;
  const lo_p=Math.max(0,ph-me_p),hi_p=Math.min(1,ph+me_p);
  document.getElementById('ic-prop-steps').innerHTML=`<div class="fbox">
<span class="cm">// IC 95% para proporción p (¿duerme bien?)</span>
p̂ = ${s.bienSi}/${n} = <span class="vl">${ph.toFixed(4)}</span>
SE = √(p̂(1−p̂)/n) = √(${ph.toFixed(3)}×${(1-ph).toFixed(3)}/${n}) = <span class="vl">${se_p.toFixed(4)}</span>
ME = z×SE = ${z}×${se_p.toFixed(4)} = <span class="op">${me_p.toFixed(4)}</span>
IC₉₅% = [ <span class="rs">${lo_p.toFixed(4)}</span> , <span class="rs">${hi_p.toFixed(4)}</span> ]  =  [ <span class="rs">${(lo_p*100).toFixed(1)}%</span> , <span class="rs">${(hi_p*100).toFixed(1)}%</span> ]
Interpretación: Con 95% de confianza, entre <span class="rs">${(lo_p*100).toFixed(1)}%</span> y <span class="rs">${(hi_p*100).toFixed(1)}%</span> de los estudiantes duerme bien.</div>`;
  drawBellChart('chart-ic-prop',ph,se_p,lo_p,hi_p,'p̂='+ph.toFixed(3),
    {fill:'rgba(96,165,250,.25)',out:'rgba(96,165,250,.04)',line:'rgba(96,165,250,.9)'});

  // IC Media
  const xbar=s.mean,se_m=s.std/Math.sqrt(n),me_m=z*se_m,lo_m=xbar-me_m,hi_m=xbar+me_m;
  document.getElementById('ic-media-steps').innerHTML=`<div class="fbox">
<span class="cm">// IC 95% para media μ (horas de sueño)</span>
x̄ = ${xbar.toFixed(4)}  |  s = ${s.std.toFixed(4)}  |  n = ${n}
SE = s/√n = ${s.std.toFixed(4)}/√${n} = <span class="vl">${se_m.toFixed(4)}</span>
ME = z×SE = ${z}×${se_m.toFixed(4)} = <span class="op">${me_m.toFixed(4)}</span>
IC₉₅% = [ <span class="rs">${lo_m.toFixed(4)}</span> , <span class="rs">${hi_m.toFixed(4)}</span> ] horas
Interpretación: Con 95% de confianza, el promedio real de horas de sueño está entre <span class="rs">${lo_m.toFixed(2)}</span> y <span class="rs">${hi_m.toFixed(2)}</span> horas.</div>`;
  drawBellChart('chart-ic-media',xbar,se_m,lo_m,hi_m,'x̄='+xbar.toFixed(2)+'h',
    {fill:'rgba(74,222,128,.25)',out:'rgba(74,222,128,.04)',line:'rgba(74,222,128,.9)'});

  // Tabla unificada de frecuencias
  updateUnifiedFreqTable(s);
}

function updateUnifiedFreqTable(s){
  if(!s||!s.n) return;
  const tbody=document.getElementById('unified-freq-body');
  if(!tbody) return;
  const n=s.n;

  function varBlock(varName,color,vals){
    const cats=[...new Set(vals.filter(v=>v&&v!=='—'&&v!=='null'))].sort((a,b)=>String(a).localeCompare(String(b)));
    if(!cats.length) return '';
    let cum=0;
    const hdr=`<tr style="background:#0e0e0e">
      <td colspan="2" style="text-align:left;font-weight:700;color:${color};font-size:11px;border-top:2px solid ${color}30;padding:8px 12px">${varName}</td>
      <td style="border-top:2px solid ${color}30"></td><td style="border-top:2px solid ${color}30"></td>
      <td style="border-top:2px solid ${color}30"></td><td style="border-top:2px solid ${color}30"></td>
      <td style="border-top:2px solid ${color}30"></td></tr>`;
    const rows=cats.map(c=>{
      const fi=vals.filter(v=>String(v)===String(c)).length;
      const fri=fi/n; cum+=fi;
      return`<tr>
        <td style="width:16px;padding-left:20px"><span style="display:inline-block;width:7px;height:7px;border-radius:50%;background:${color};opacity:.5"></span></td>
        <td style="text-align:left;font-size:11px">${c}</td>
        <td>${fi}</td><td>${fri.toFixed(4)}</td>
        <td>${(fri*100).toFixed(2)}</td>
        <td>${cum}</td><td>${(cum/n*100).toFixed(2)}%</td>
      </tr>`;
    });
    const tot_fi=vals.filter(v=>v&&v!=='—').length;
    const total=`<tr class="tr-total"><td colspan="2" style="text-align:left;padding-left:20px">Total</td>
      <td>${tot_fi}</td><td>1.0000</td><td>100.00</td><td>${tot_fi}</td><td>100.00%</td></tr>`;
    return hdr+rows.join('')+total;
  }

  tbody.innerHTML=[
    varBlock('P9. Horas de sueño (días de clases)',     '#60a5fa', s.horas),
    varBlock('P10. ¿Las horas son suficientes?',        '#4ade80', data.map(d=>d.bien)),
    varBlock('P11. ¿Con qué frecuencia se desvela?',    '#f87171', data.map(d=>d.desvela)),
    varBlock('P12. Somnolencia en clases',              '#fbbf24', data.map(d=>d.somnolencia)),
    varBlock('P13. Café / bebidas energéticas',         '#34d399', data.map(d=>d.cafe)),
    varBlock('P14. Usa pantallas antes de dormir',      '#a78bfa', data.map(d=>d.pantalla)),
    varBlock('P15. Realiza actividad física',           '#f97316', data.map(d=>d.actividad)),
    varBlock('P1. Edad',                                '#e2e8f0', data.map(d=>d.edad+'a')),
    varBlock('P2. Sexo',                                '#fb7185', data.map(d=>d.genero)),
    varBlock('P3. Carrera',                             '#38bdf8', data.map(d=>d.carrera)),
    varBlock('P4. Semestre',                            '#818cf8', data.map(d=>d.semestre+'°')),
    varBlock('P5. Horario de clases',                   '#2dd4bf', data.map(d=>d.horario)),
    varBlock('P6. Hora de levantarse (clases)',         '#facc15', data.map(d=>d.levanta_clases)),
    varBlock('P7. Hora de dormirse (clases)',           '#c084fc', data.map(d=>d.duerme_clases)),
    varBlock('P7b. Hora dormirse (fin de semana)',      '#94a3b8', data.map(d=>d.duerme_finde)),
    varBlock('P8. Hora levantarse (fin de semana)',     '#6ee7b7', data.map(d=>d.levanta_finde)),
  ].join('');
}

function positionCI(prefix,lo,mid,hi,minV,maxV,loL,hiL){
  const r=maxV-minV;
  const loP=(lo-minV)/r*80+10,hiP=(hi-minV)/r*80+10,midP=(mid-minV)/r*80+10;
  const fill=document.getElementById(prefix+'-fill');
  const dot=document.getElementById(prefix+'-dot');
  if(fill){fill.style.left=loP+'%';fill.style.width=(hiP-loP)+'%';}
  if(dot) dot.style.left=midP+'%';
  const lo_el=document.getElementById(prefix+'-lo');
  const hi_el=document.getElementById(prefix+'-hi');
  const pm_el=document.getElementById(prefix+'-pm');
  if(lo_el){lo_el.style.left=loP+'%';lo_el.textContent=loL;}
  if(hi_el){hi_el.style.left=hiP+'%';hi_el.textContent=hiL;}
  if(pm_el) pm_el.style.left=midP+'%';
}

// ════════════════════════════════════════════
// HIPÓTESIS (panel inferior)
// ════════════════════════════════════════════
function updateHypSteps(){
  const s=stats();if(!s.n) return;
  const n=s.n,ph=s.pBien,p0=0.50;
  const se=Math.sqrt(p0*(1-p0)/n);
  const z=(ph-p0)/se,reject=Math.abs(z)>1.96;
  document.getElementById('hyp-steps').innerHTML=`
    <div class="step-block" id="step-h1">
      <div class="sb-label">Paso 1 — Datos del problema</div>
      <div class="formula-step-visual">
        <span class="fv-term fv-n">n = ${n}</span>
        <span class="fv-term fv-p">p̂ = ${s.bienSi}/${n} = ${ph.toFixed(4)}</span>
        <span class="fv-term fv-p0">p₀ = 0.50</span>
        <span class="fv-term fv-z">z_crit = ±1.96</span>
      </div>
      <div class="fv-note">Reunimos lo que sabemos: cuántos encuestamos (n), qué proporción observamos (p̂), el valor que queremos comprobar (p₀) y el límite de la zona de rechazo (z_crit).</div>
    </div>
    <div class="step-block" id="step-h2">
      <div class="sb-label">Paso 2 — Fórmula del estadístico Z</div>
      <div class="formula-step-visual formula-big">
        <span class="fv-letter">z</span> <span class="fv-op">=</span>
        <span class="fv-frac">
          <span class="fv-frac-top">(<span class="fv-letter">p̂</span> − <span class="fv-letter">p₀</span>)</span>
          <span class="fv-frac-bot">√(<span class="fv-letter">p₀</span>(1−<span class="fv-letter">p₀</span>)/<span class="fv-letter">n</span>)</span>
        </span>
      </div>
      <div class="formula-step-visual formula-big fv-substituted">
        <span class="fv-letter">z</span> <span class="fv-op">=</span>
        <span class="fv-frac">
          <span class="fv-frac-top">(<span class="fv-num">${ph.toFixed(4)}</span> − <span class="fv-num">${p0}</span>)</span>
          <span class="fv-frac-bot">√(<span class="fv-num">${p0}</span>×<span class="fv-num">${(1-p0).toFixed(2)}</span>/<span class="fv-num">${n}</span>)</span>
        </span>
      </div>
      <div class="fv-note">Sustituimos cada letra por su número real: la diferencia entre lo observado y lo esperado, dividida entre el error estándar.</div>
    </div>
    <div class="step-block" id="step-h3">
      <div class="sb-label">Paso 3 — Resultado del estadístico</div>
      <div class="formula-step-visual formula-big">
        <span class="fv-letter">z</span> <span class="fv-op">=</span>
        <span class="fv-num">${(ph-p0).toFixed(4)}</span> <span class="fv-op">/</span> <span class="fv-num">${se.toFixed(4)}</span>
        <span class="fv-op">=</span>
        <span class="fv-result">${z.toFixed(4)}</span>
      </div>
      <div class="fv-note">Hacemos la resta y la división. Este número nos dice qué tan lejos está nuestro resultado de lo esperado, medido en desviaciones estándar.</div>
    </div>
    <div class="step-block" id="step-h4">
      <div class="sb-label">Paso 4 — ¿Cae en la zona de rechazo?</div>
      <div class="formula-step-visual formula-big">
        <span class="fv-letter">|z|</span> <span class="fv-op">=</span> <span class="fv-num">${Math.abs(z).toFixed(4)}</span>
        <span class="fv-op" style="color:${reject?'var(--red)':'var(--green)'}">${Math.abs(z)>1.96?'>':'<'}</span>
        <span class="fv-num">1.96</span>
      </div>
      <div class="hyp-result-banner ${reject?'reject':'no-reject'}">
        ${reject?'RECHAZAR H₀':'NO RECHAZAR H₀'}
      </div>
      <div class="fv-note">${reject?'El valor calculado supera el límite crítico: cae en la zona de rechazo.':'El valor calculado no supera el límite crítico: no hay evidencia suficiente para rechazar.'}</div>
    </div>`;
  document.getElementById('hyp-conclusion-card').style.display='block';
  document.getElementById('hyp-conclusion').innerHTML=`<div class="${reject?'hyp-fail':'hyp-pass'}">
    <strong style="color:${reject?'var(--red)':'var(--green)'}">${reject?'Se rechaza H₀':'No se rechaza H₀'}</strong><br>
    <span style="font-size:12px;color:var(--muted2)">Con z=${z.toFixed(4)}, α=0.05: ${reject?`La proporción ${(ph*100).toFixed(1)}% difiere significativamente del 50%.`:`Con ${(ph*100).toFixed(1)}% no hay evidencia suficiente para rechazar p=50%.`}</span></div>`;
}
let hypAnim=false;
function animateHypothesis(){
  if(hypAnim) return;hypAnim=true;
  ['step-h1','step-h2','step-h3','step-h4'].forEach(id=>document.getElementById(id)?.classList.remove('show','done'));
  document.getElementById('prog-hyp').style.width='0%';
  ['step-h1','step-h2','step-h3','step-h4'].forEach((id,i)=>setTimeout(()=>{
    const el=document.getElementById(id);if(el){el.classList.add('show');setTimeout(()=>el.classList.add('done'),400);}
    document.getElementById('prog-hyp').style.width=((i+1)/4*100)+'%';
    if(i===3) hypAnim=false;
  },i*900));
}
function loadHypFromData(){
  const s=stats();if(!s.n){notify('Primero carga datos en la pestaña Datos');return;}
  document.getElementById('hyp-ph').value=s.pBien.toFixed(4);
  document.getElementById('hyp-n').value=s.n;
}

// ════════════════════════════════════════════
// RESUMEN
// ════════════════════════════════════════════
function updateResumen(){
  const s=stats();
  if(!s.n){document.getElementById('resumen-content').innerHTML='<p style="color:var(--muted)">Ingresa datos para ver el resumen.</p>';return;}
  const z=+document.getElementById('conf-level').value||1.96;
  const ph=s.pBien,se_p=Math.sqrt(ph*(1-ph)/s.n),me_p=z*se_p;
  const lo_p=Math.max(0,ph-me_p),hi_p=Math.min(1,ph+me_p);
  const se_m=s.std/Math.sqrt(s.n),me_m=z*se_m;
  const zStat=(ph-0.5)/Math.sqrt(0.25/s.n),reject=Math.abs(zStat)>1.96;
  document.getElementById('resumen-content').innerHTML=`
    <div class="metrics-row">
      <div class="metric"><div class="metric-label">Encuestados</div><div class="metric-value">${s.n}</div></div>
      <div class="metric"><div class="metric-label">Promedio sueño</div><div class="metric-value">${s.mean.toFixed(2)}</div><div class="metric-sub">horas</div></div>
      <div class="metric"><div class="metric-label">Mediana</div><div class="metric-value">${s.median.toFixed(1)}</div><div class="metric-sub">horas</div></div>
      <div class="metric"><div class="metric-label">Desv. estándar</div><div class="metric-value">${s.std.toFixed(2)}</div></div>
    </div>
    <div class="sep"></div>
    <div class="grid2" style="margin-bottom:12px">
      <div class="metric" style="text-align:left;padding:12px 14px"><div class="metric-label">IC 95% proporción duerme bien</div><div style="margin-top:7px;font-size:13px">p̂ = <b>${(ph*100).toFixed(1)}%</b></div><div style="font-size:13px">IC = [<b>${(lo_p*100).toFixed(1)}%</b> , <b>${(hi_p*100).toFixed(1)}%</b>]</div><div style="font-size:10px;color:var(--muted);margin-top:3px">±${(me_p*100).toFixed(2)}%</div></div>
      <div class="metric" style="text-align:left;padding:12px 14px"><div class="metric-label">IC 95% horas de sueño</div><div style="margin-top:7px;font-size:13px">x̄ = <b>${s.mean.toFixed(2)} h</b></div><div style="font-size:13px">IC = [<b>${(s.mean-me_m).toFixed(2)}</b> , <b>${(s.mean+me_m).toFixed(2)}</b>] h</div><div style="font-size:10px;color:var(--muted);margin-top:3px">±${me_m.toFixed(3)} h</div></div>
    </div>
    <div class="sep"></div>
    <div style="font-size:10px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;margin-bottom:7px">Prueba de hipótesis — H₀: p = 0.50</div>
    <div class="${reject?'hyp-fail':'hyp-pass'}">
      <strong style="color:${reject?'var(--red)':'var(--green)'}">${reject?'Se rechaza H₀':'No se rechaza H₀'}</strong>  ·  z = ${zStat.toFixed(4)} · z_crit = ±1.96<br>
      <span style="font-size:12px;color:var(--muted2)">${reject?`Con p̂=${(ph*100).toFixed(1)}%, difiere significativamente del 50%.`:`Con p̂=${(ph*100).toFixed(1)}%, no hay evidencia para rechazar p=50%.`}</span>
    </div>
    <div class="sep"></div>
    <div class="grid3">
      <div class="metric"><div class="metric-label">P(duerme bien)</div><div class="metric-value" style="font-size:20px">${(ph*100).toFixed(1)}%</div></div>
      <div class="metric"><div class="metric-label">P(usa pantallas)</div><div class="metric-value" style="font-size:20px">${(s.pPantalla*100).toFixed(1)}%</div></div>
      <div class="metric"><div class="metric-label">P(horas ≤ 6)</div><div class="metric-value" style="font-size:20px">${(data.filter(d=>d.horas<=6).length/s.n*100).toFixed(1)}%</div></div>
    </div>`;
}

// ════════════════════════════════════════════
// PIZARRA — MOTOR UNIVERSAL
// ════════════════════════════════════════════
let boardSolving={};

function esc(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}

function tokenize(text){
  const tokens=[];let i=0;
  while(i<text.length){
    const nm=text.slice(i).match(/^-?\d+\.?\d*/);
    if(nm&&(i===0||!/[A-Za-z(₀₁₂]/.test(text[i-1]))){tokens.push({type:'num',val:nm[0]});i+=nm[0].length;continue;}
    const pm=text.slice(i).match(/^P\([^)]+\)/);
    if(pm){tokens.push({type:'label',val:pm[0]});i+=pm[0].length;continue;}
    tokens.push({type:'other',val:text[i]});i++;
  }
  return tokens;
}





function drawOnBoard(boardId, idleId, sdotId, stxtId, sstepId, formulaNombre, formulaGeneral, result){
  const bid=boardId;
  if(boardSolving[bid]) return;
  boardSolving[bid]=true;

  document.getElementById(idleId).style.opacity='0';
  document.getElementById(sdotId).classList.add('active');

  const bc=document.getElementById(bid);
  // Limpiar todo excepto idle
  Array.from(bc.children).forEach(c=>{if(!c.classList.contains('board-idle-state')) c.remove();});

  // Título
  const title=document.createElement('div');
  title.className='board-title';title.textContent=formulaNombre+' — paso a paso';
  bc.appendChild(title);setTimeout(()=>title.classList.add('show'),50);

  // Fórmula general (letras)
  const fDiv=document.createElement('div');
  fDiv.className='formula-display';
  fDiv.innerHTML=`<span style="color:var(--chalk);font-size:.82em">${esc(formulaGeneral)}</span>`;
  bc.appendChild(fDiv);setTimeout(()=>fDiv.classList.add('show'),300);

  // Sustitución animada
  const sDiv=document.createElement('div');
  sDiv.className='formula-display';sDiv.style.opacity='0';bc.appendChild(sDiv);

  // Divisor
  const div=document.createElement('div');
  div.style.cssText='width:80%;height:1px;background:rgba(240,244,240,.08);margin:5px 0 12px;opacity:0;transition:opacity .5s;position:relative;z-index:1';
  bc.appendChild(div);

  // Pasos
  const stepsDiv=document.createElement('div');stepsDiv.className='steps-area';bc.appendChild(stepsDiv);

  // Resultado
  const rBox=document.createElement('div');rBox.className='result-box';
  rBox.innerHTML=`<div><div class="result-label">${esc(result.label)} =</div></div>
    <div class="result-value">${typeof result.result==='number'?result.result.toFixed(4):result.result}</div>
    ${result.pct!==false?`<div class="result-pct">${typeof result.result==='number'?(result.result*100).toFixed(2)+'%':''}</div>`:''}`;
  bc.appendChild(rBox);

  let delay=600;
  setBS(stxtId,sstepId,'Escribiendo la fórmula...',0,result.steps.length);

  // Animación sustitución
  setTimeout(()=>{
    fDiv.style.transition='opacity .4s';fDiv.style.opacity='.25';
    sDiv.innerHTML='';sDiv.style.opacity='1';sDiv.style.transition='opacity .5s';
    const tokens=tokenize(result.steps[0].tex);
    tokens.forEach((tok,i)=>{
      const sp=document.createElement('span');
      if(tok.type==='num'){
        sp.style.cssText='color:var(--chalk-y);font-weight:700;text-shadow:0 0 12px rgba(245,230,66,.5);opacity:0;font-family:Kalam,cursive;font-size:32px';
        sp.textContent=tok.val;
        setTimeout(()=>{sp.style.transition='opacity .3s,transform .4s cubic-bezier(.34,1.56,.64,1)';sp.style.opacity='1';sp.classList.add('tok-anim');},i*65+200);
      } else if(tok.type==='label'){
        sp.style.cssText='color:var(--chalk-c);font-weight:700;font-family:Kalam,cursive;font-size:32px';sp.textContent=tok.val;
      } else {
        sp.style.cssText='color:var(--chalk-dim);font-family:Kalam,cursive;font-size:32px';sp.textContent=tok.val;
      }
      sDiv.appendChild(sp);
    });
  },delay);
  delay+=800;

  setTimeout(()=>{div.style.opacity='1';},delay);delay+=200;

  result.steps.forEach((step,i)=>{
    setTimeout(()=>{
      setBS(stxtId,sstepId,'Resolviendo...',i+1,result.steps.length);
      const line=document.createElement('div');line.className='step-line';
      const prefix=i===0?'→  ':i===result.steps.length-1?'∴  ':'   ';
      const cls=step.underline?'s-underline':'';
      const color=step.final?'color:var(--chalk-o);':'';
      const sz=step.final?'font-size:29px':'font-size:21px';
      line.innerHTML=`<span class="${cls}" id="sul-${bid}-${i}" style="${color}${sz}">${prefix}${esc(step.tex)}</span>`;
      if(step.ann) line.innerHTML+=`<span class="step-ann" id="ann-${bid}-${i}">← ${step.ann}</span>`;
      stepsDiv.appendChild(line);
      setTimeout(()=>{
        line.classList.add('show');
        if(step.ann) setTimeout(()=>document.getElementById(`ann-${bid}-${i}`)?.classList.add('show'),280);
        if(step.underline) setTimeout(()=>document.getElementById(`sul-${bid}-${i}`)?.classList.add('anim'),400);
      },50);
    },delay+i*860);
  });
  delay+=result.steps.length*860+200;

  setTimeout(()=>{
    rBox.classList.add('show');
    setBS(stxtId,sstepId,'Resolución completa',result.steps.length,result.steps.length);
    boardSolving[bid]=false;
  },delay);
}

function setBS(stxtId,sstepId,msg,cur,tot){
  document.getElementById(stxtId).textContent=msg;
  if(cur!==undefined) document.getElementById(sstepId).textContent=`Paso ${cur} / ${tot}`;
}

function clearBoardById(boardId){
  boardSolving[boardId]=false;
  const bc=document.getElementById(boardId);
  Array.from(bc.children).forEach(c=>{if(!c.classList.contains('board-idle-state')) c.remove();});
  const idle=bc.querySelector('.board-idle-state');if(idle) idle.style.opacity='1';
  const prefix=boardId.replace('board-','');
  const sdot=document.getElementById('sdot-'+prefix);if(sdot) sdot.classList.remove('active');
  const stxt=document.getElementById('stxt-'+prefix);if(stxt) stxt.textContent='Listo';
  const sstep=document.getElementById('sstep-'+prefix);if(sstep) sstep.textContent='';
}

// ════════════════════════════════════════════
// PIZARRA — DESCRIPTIVO
// ════════════════════════════════════════════
let currentDescFormula='media';
const DESC_FORMULAS={
  media:{
    nombre:'Media aritmética (x̄)',general:'x̄ = Σxᵢ / n',
    desc:'Suma todos los valores y divide entre la cantidad de datos. El "promedio" clásico.',
    fields:[{key:'suma',label:'Suma de todos los valores (Σxᵢ)',color:'a',ph:'ej: 175'},{key:'n',label:'Cantidad de datos (n)',color:'b',ph:'ej: 25'}],
    solve:(v)=>{const res=v.suma/v.n;return{steps:[
      {tex:`x̄ = ${v.suma} / ${v.n}`,ann:`Sustituyo: Σxᵢ=${v.suma}, n=${v.n}`},
      {tex:`x̄ = ${res.toFixed(4)}`,ann:'Divido suma entre n',underline:true},
      {tex:`x̄ = ${res.toFixed(4)} horas`,ann:'Resultado: promedio de horas de sueño',final:true},
    ],result:res,label:'x̄',pct:false};}
  },
  mediana:{
    nombre:'Mediana',general:'Me = valor central de datos ordenados',
    desc:'El valor que queda en el centro cuando ordenas todos los datos de menor a mayor.',
    fields:[{key:'pos1',label:'Valor en posición n/2',color:'a',ph:'ej: 7'},{key:'pos2',label:'Valor en posición n/2+1 (si n es par, 0 si impar)',color:'b',ph:'ej: 7'}],
    solve:(v)=>{const res=v.pos2===0?v.pos1:(v.pos1+v.pos2)/2;
      const steps=v.pos2===0?[
        {tex:`n es impar → Me = valor central`,ann:'Solo hay un valor central'},
        {tex:`Me = ${v.pos1}`,ann:'Resultado directo',final:true,underline:true},
      ]:[
        {tex:`n es par → Me = (${v.pos1} + ${v.pos2}) / 2`,ann:'Promedio de los dos centrales'},
        {tex:`Me = ${(v.pos1+v.pos2).toFixed(1)} / 2`,ann:'Sumo los dos valores'},
        {tex:`Me = ${res.toFixed(1)}`,ann:'Divido entre 2',final:true,underline:true},
      ];
      return{steps,result:res,label:'Me',pct:false};}
  },
  std:{
    nombre:'Desviación estándar (s)',general:'s = √[ Σ(xᵢ − x̄)² / (n−1) ]',
    desc:'Mide cuánto se alejan los datos del promedio. Mayor s = más dispersión.',
    fields:[{key:'sv',label:'Σ(xᵢ−x̄)² — suma de diferencias al cuadrado',color:'a',ph:'ej: 24.5'},{key:'n',label:'n — cantidad de datos',color:'b',ph:'ej: 25'}],
    solve:(v)=>{const var_=v.sv/(v.n-1),std=Math.sqrt(var_);return{steps:[
      {tex:`s² = ${v.sv} / (${v.n}−1)`,ann:`Varianza muestral: divido entre n−1`},
      {tex:`s² = ${v.sv} / ${v.n-1} = ${var_.toFixed(4)}`,ann:'Calculo la varianza'},
      {tex:`s = √(${var_.toFixed(4)})`,ann:'Raíz cuadrada de la varianza'},
      {tex:`s = ${std.toFixed(4)}`,ann:'Desviación estándar final',final:true,underline:true},
    ],result:std,label:'s',pct:false};}
  },
  varianza:{
    nombre:'Varianza muestral (s²)',general:'s² = Σ(xᵢ − x̄)² / (n−1)',
    desc:'Promedio de las diferencias al cuadrado respecto a la media. Base para calcular la desviación.',
    fields:[{key:'sv',label:'Σ(xᵢ−x̄)² — suma de cuadrados',color:'a',ph:'ej: 24.5'},{key:'n',label:'n — cantidad de datos',color:'b',ph:'ej: 25'}],
    solve:(v)=>{const res=v.sv/(v.n-1);return{steps:[
      {tex:`s² = ${v.sv} / (${v.n}−1)`,ann:'Sustituyo los valores'},
      {tex:`s² = ${v.sv} / ${v.n-1}`,ann:`Resto 1 al denominador (muestral)`},
      {tex:`s² = ${res.toFixed(4)}`,ann:'Divido',final:true,underline:true},
    ],result:res,label:'s²',pct:false};}
  },
};

function selectBoardFormula(section,formula,btn){
  if(section==='desc'){
    currentDescFormula=formula;
    document.querySelectorAll('#bfs-desc .bfs-btn').forEach(b=>b.classList.remove('active'));
    btn.classList.add('active');
    renderDescInputs();clearBoardById('board-desc');
  }
}

function renderDescInputs(){
  const f=DESC_FORMULAS[currentDescFormula];
  document.getElementById('bi-desc-desc').textContent=f.desc;
  document.getElementById('bi-desc-fields').innerHTML=f.fields.map(fi=>`
    <div class="bi-field"><label><span class="bi-badge ${fi.color==='b'?'b':''}">${fi.color==='a'?'A':'B'}</span> ${fi.label}</label>
    <input type="number" id="df-${fi.key}" step="0.001" placeholder="${fi.ph}" value=""></div>`).join('');
  // Ejemplos rápidos desde datos
  document.getElementById('bi-desc-ex').innerHTML=`<span class="bi-ex" onclick="loadDescFromData()">Cargar de mis datos</span>`;
}

function loadDescFromData(){
  const s=stats();if(!s.n){notify('Primero carga datos en la pestaña Datos');return;}
  const horas=s.horas;
  const suma=horas.reduce((a,b)=>a+b,0);
  const sv=horas.reduce((a,b)=>a+(b-s.mean)**2,0);
  const f=currentDescFormula;
  if(f==='media'){
    const e=document.getElementById('df-suma');if(e) e.value=suma.toFixed(2);
    const e2=document.getElementById('df-n');if(e2) e2.value=s.n;
  } else if(f==='mediana'){
    const sorted=[...horas].sort((a,b)=>a-b);
    const n=sorted.length;
    e=document.getElementById('df-pos1');if(e) e.value=n%2===0?sorted[n/2-1]:sorted[Math.floor(n/2)];
    e2=document.getElementById('df-pos2');if(e2) e2.value=n%2===0?sorted[n/2]:0;
  } else if(f==='std'||f==='varianza'){
    e=document.getElementById('df-sv');if(e) e.value=sv.toFixed(4);
    e2=document.getElementById('df-n');if(e2) e2.value=s.n;
  }
  notify('Datos cargados de la encuesta');
}

function solveBoardDesc(){
  const f=DESC_FORMULAS[currentDescFormula];
  const vals={};
  for(const fi of f.fields){
    const el=document.getElementById('df-'+fi.key);
    if(!el||el.value===''){notify('Completa todos los campos');return;}
    vals[fi.key]=parseFloat(el.value);
  }
  const result=f.solve(vals);
  drawOnBoard('board-desc','idle-desc','sdot-desc','stxt-desc','sstep-desc',f.nombre,f.general,result);
}

function syncBoardInputsFromData(){
  // Auto-sync desc si hay datos
}

// ════════════════════════════════════════════
// PIZARRA — PROBABILIDADES
// ════════════════════════════════════════════
const PROB_FORMULAS={
  laplace:{nombre:'Regla de Laplace',general:'P(A) = casos favorables / casos posibles',
    desc:'Probabilidad clásica: divide los casos que quieres entre todos los posibles.',
    fields:[{key:'fav',label:'Casos favorables',color:'a',ph:'ej: 15'},{key:'pos',label:'Casos posibles (total)',color:'b',ph:'ej: 25'}],
    ejemplos:[{label:'Duermen bien',vals:{fav:15,pos:25}},{label:'Usan pantallas',vals:{fav:18,pos:25}}],
    solve:(v)=>{const res=v.fav/v.pos;return{steps:[
      {tex:`P(A) = ${v.fav} / ${v.pos}`,ann:`favorables=${v.fav}, posibles=${v.pos}`},
      {tex:`P(A) = ${res.toFixed(4)}`,ann:'Divido',underline:true},
      {tex:`P(A) = ${(res*100).toFixed(2)}%`,ann:'En porcentaje',final:true},
    ],result:res,label:'P(A)'};}},
  complemento:{nombre:'Evento complementario',general:'P(Ā) = 1 − P(A)',
    desc:'Si algo tiene 60% de pasar, su contrario tiene 40%.',
    fields:[{key:'pA',label:'P(A) — probabilidad del evento',color:'a',ph:'ej: 0.60'}],
    ejemplos:[{label:'No duerme bien',vals:{pA:0.60}},{label:'No usa pantallas',vals:{pA:0.72}}],
    solve:(v)=>{const res=1-v.pA;return{steps:[
      {tex:`P(Ā) = 1 − ${v.pA}`,ann:`P(A)=${v.pA}`},
      {tex:`P(Ā) = ${res.toFixed(4)}`,ann:'Resto de 1',final:true,underline:true},
    ],result:res,label:'P(Ā)'};}},
  union_excl:{nombre:'Unión — eventos excluyentes',general:'P(A∪B) = P(A) + P(B)',
    desc:'Cuando A y B NO pueden ocurrir al mismo tiempo, solo se suman.',
    fields:[{key:'pA',label:'P(A)',color:'a',ph:'ej: 0.40'},{key:'pB',label:'P(B)',color:'b',ph:'ej: 0.35'}],
    ejemplos:[{label:'Sueño corto o muy corto',vals:{pA:0.24,pB:0.08}}],
    solve:(v)=>{const res=v.pA+v.pB;return{steps:[
      {tex:`P(A∪B) = ${v.pA} + ${v.pB}`,ann:'Sustituyo P(A) y P(B)'},
      {tex:`P(A∪B) = ${res.toFixed(4)}`,ann:'Sumo directamente',final:true,underline:true},
    ],result:res,label:'P(A∪B)'};}},
  union_noexcl:{nombre:'Unión — eventos no excluyentes',general:'P(A∪B) = P(A) + P(B) − P(A∩B)',
    desc:'Cuando A y B SÍ pueden ocurrir juntos, hay que restar lo que se conta doble.',
    fields:[{key:'pA',label:'P(A)',color:'a',ph:'ej: 0.60'},{key:'pB',label:'P(B)',color:'b',ph:'ej: 0.72'},{key:'pAB',label:'P(A∩B)',color:'r',ph:'ej: 0.44'}],
    ejemplos:[{label:'Duerme bien O usa pantallas',vals:{pA:0.60,pB:0.72,pAB:0.44}}],
    solve:(v)=>{const s2=v.pA+v.pB,res=s2-v.pAB;return{steps:[
      {tex:`P(A∪B) = ${v.pA} + ${v.pB} − ${v.pAB}`,ann:'Sustituyo los tres valores'},
      {tex:`P(A∪B) = ${s2.toFixed(4)} − ${v.pAB}`,ann:`Sumo P(A)+P(B)=${s2.toFixed(4)}`},
      {tex:`P(A∪B) = ${res.toFixed(4)}`,ann:'Resto la intersección',final:true,underline:true},
    ],result:res,label:'P(A∪B)'};}},
  conjunta_indep:{nombre:'Conjunta — independientes',general:'P(A∩B) = P(A) · P(B)',
    desc:'Cuando A no influye en B, la probabilidad de ambos es el producto.',
    fields:[{key:'pA',label:'P(A)',color:'a',ph:'ej: 0.60'},{key:'pB',label:'P(B)',color:'b',ph:'ej: 0.72'}],
    ejemplos:[{label:'Duerme bien Y usa pantallas',vals:{pA:0.60,pB:0.72}}],
    solve:(v)=>{const res=v.pA*v.pB;return{steps:[
      {tex:`P(A∩B) = ${v.pA} × ${v.pB}`,ann:'Sustituyo P(A) y P(B)'},
      {tex:`P(A∩B) = ${res.toFixed(4)}`,ann:'Multiplico',final:true,underline:true},
    ],result:res,label:'P(A∩B)'};}},
  conjunta_dep:{nombre:'Conjunta — dependientes',general:'P(A∩B) = P(A) · P(B|A)',
    desc:'Cuando B depende de A. Usamos la probabilidad condicional de B dado A.',
    fields:[{key:'pA',label:'P(A)',color:'a',ph:'ej: 0.60'},{key:'pBdA',label:'P(B|A) — B dado que A ocurrió',color:'b',ph:'ej: 0.73'}],
    ejemplos:[{label:'Pantallas → duerme bien',vals:{pA:0.72,pBdA:0.61}}],
    solve:(v)=>{const res=v.pA*v.pBdA;return{steps:[
      {tex:`P(A∩B) = ${v.pA} × ${v.pBdA}`,ann:'Sustituyo P(A) y P(B|A)'},
      {tex:`P(A∩B) = ${res.toFixed(4)}`,ann:'Multiplico',final:true,underline:true},
    ],result:res,label:'P(A∩B)'};}},
  condicional:{nombre:'Probabilidad condicional',general:'P(A|B) = P(A∩B) / P(B)',
    desc:'"Dado que B ocurrió, ¿qué tan probable es A?" Divide la intersección entre P(B).',
    fields:[{key:'pAB',label:'P(A∩B) — intersección',color:'r',ph:'ej: 0.44'},{key:'pB',label:'P(B)',color:'b',ph:'ej: 0.72'}],
    ejemplos:[{label:'Duerme bien | usa pantallas',vals:{pAB:0.44,pB:0.72}}],
    solve:(v)=>{const res=v.pAB/v.pB;return{steps:[
      {tex:`P(A|B) = ${v.pAB} / ${v.pB}`,ann:'Sustituyo P(A∩B) y P(B)'},
      {tex:`P(A|B) = ${res.toFixed(4)}`,ann:'Divido',final:true,underline:true},
    ],result:res,label:'P(A|B)'};}},
  diferencia:{nombre:'Diferencia de eventos',general:'P(A−B) = P(A) − P(A∩B)',
    desc:'A ocurre pero B NO. "Duerme bien pero no usa pantallas."',
    fields:[{key:'pA',label:'P(A)',color:'a',ph:'ej: 0.60'},{key:'pAB',label:'P(A∩B)',color:'r',ph:'ej: 0.44'}],
    ejemplos:[{label:'Duerme bien sin pantallas',vals:{pA:0.60,pAB:0.44}}],
    solve:(v)=>{const res=v.pA-v.pAB;return{steps:[
      {tex:`P(A−B) = ${v.pA} − ${v.pAB}`,ann:'Sustituyo P(A) y P(A∩B)'},
      {tex:`P(A−B) = ${res.toFixed(4)}`,ann:'Resto',final:true,underline:true},
    ],result:res,label:'P(A−B)'};}},
  prob_total:{nombre:'Probabilidad total',general:'P(B) = P(B|A₁)·P(A₁) + P(B|A₂)·P(A₂)',
    desc:'Calcula la probabilidad de B sumando todos los caminos posibles.',
    fields:[{key:'pA1',label:'P(A₁)',color:'a',ph:'ej: 0.60'},{key:'pA2',label:'P(A₂)',color:'b',ph:'ej: 0.40'},{key:'pBdA1',label:'P(B|A₁)',color:'a',ph:'ej: 0.75'},{key:'pBdA2',label:'P(B|A₂)',color:'b',ph:'ej: 0.30'}],
    ejemplos:[{label:'Duerme bien total',vals:{pA1:0.72,pA2:0.28,pBdA1:0.50,pBdA2:0.86}}],
    solve:(v)=>{const t1=v.pBdA1*v.pA1,t2=v.pBdA2*v.pA2,res=t1+t2;return{steps:[
      {tex:`P(B) = P(B|A₁)·P(A₁) + P(B|A₂)·P(A₂)`,ann:'Expando la sumatoria'},
      {tex:`P(B) = ${v.pBdA1}·${v.pA1} + ${v.pBdA2}·${v.pA2}`,ann:'Sustituyo todos los valores'},
      {tex:`P(B) = ${t1.toFixed(4)} + ${t2.toFixed(4)}`,ann:'Multiplico cada término'},
      {tex:`P(B) = ${res.toFixed(4)}`,ann:'Sumo',final:true,underline:true},
    ],result:res,label:'P(B)'};}},
  bayes:{nombre:'Teorema de Bayes',general:'P(A₁|B) = P(B|A₁)·P(A₁) / P(B)',
    desc:'Actualiza la probabilidad de una causa con nueva evidencia.',
    fields:[{key:'pA1',label:'P(A₁)',color:'a',ph:'ej: 0.72'},{key:'pA2',label:'P(A₂)',color:'b',ph:'ej: 0.28'},{key:'pBdA1',label:'P(B|A₁)',color:'a',ph:'ej: 0.50'},{key:'pBdA2',label:'P(B|A₂)',color:'b',ph:'ej: 0.86'}],
    ejemplos:[{label:'Bayes encuesta',vals:{pA1:0.72,pA2:0.28,pBdA1:0.50,pBdA2:0.86}}],
    solve:(v)=>{const num=v.pBdA1*v.pA1,den=v.pBdA1*v.pA1+v.pBdA2*v.pA2,res=num/den;return{steps:[
      {tex:`Numerador = ${v.pBdA1} × ${v.pA1} = ${num.toFixed(4)}`,ann:'Calculo el numerador'},
      {tex:`Denominador = ${num.toFixed(4)} + ${(v.pBdA2*v.pA2).toFixed(4)} = ${den.toFixed(4)}`,ann:'Probabilidad total'},
      {tex:`P(A₁|B) = ${num.toFixed(4)} / ${den.toFixed(4)}`,ann:'Divido numerador entre denominador'},
      {tex:`P(A₁|B) = ${res.toFixed(4)}`,ann:'Resultado final',final:true,underline:true},
    ],result:res,label:'P(A₁|B)'};}},
};

let currentProbFormula='laplace';
function selectProbFormula(id,btn){
  currentProbFormula=id;
  document.querySelectorAll('#bfs-prob .bfs-btn').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  renderProbInputs();clearBoardById('board-prob');
}
function renderProbInputs(){
  const f=PROB_FORMULAS[currentProbFormula];
  document.getElementById('prob-board-desc').textContent=f.desc;
  document.getElementById('prob-board-fields').innerHTML=f.fields.map(fi=>`
    <div class="bi-field"><label><span class="bi-badge ${fi.color==='b'?'b':fi.color==='r'?'r':''}">${fi.color==='a'?'A':fi.color==='b'?'B':'∩'}</span> ${fi.label}</label>
    <input type="number" id="pf-${fi.key}" step="0.001" min="0" max="${fi.key==='fav'||fi.key==='pos'?99999:1}" placeholder="${fi.ph}" value=""></div>`).join('');
  document.getElementById('prob-board-ex').innerHTML=f.ejemplos.map(ej=>`<span class="bi-ex" onclick='loadProbEx(${JSON.stringify(ej.vals)})'>${ej.label}</span>`).join('');
}
function buildRealDataBtns(formulaId){
  const s=stats(); if(!s.n) return '';
  const n=s.n;
  const pB=s.pBien, pP=s.pPantalla, pSi=s.pActividad;
  const bienSi=s.bienSi, pantallaSi=s.pantallaSi;
  const bYp=data.filter(d=>esSí(d.bien)&&d.pantalla==='Sí').length/n;
  const bYs=data.filter(d=>esSí(d.bien)&&d.actividad==='Sí').length/n;
  const r=(v)=>+v.toFixed(4);
  const presets={
    laplace:[
      {label:`Bien (${bienSi}/${n})`,vals:{fav:bienSi,pos:n}},
      {label:`Pantallas (${pantallaSi}/${n})`,vals:{fav:pantallaSi,pos:n}},
      {label:`Horas≤6 (${data.filter(d=>d.horas<=6).length}/${n})`,vals:{fav:data.filter(d=>d.horas<=6).length,pos:n}},
    ],
    complemento:[
      {label:`P(no suficiente)`,vals:{pA:r(pB)}},
      {label:`P(no pantallas)`,vals:{pA:r(pP)}},
    ],
    union_excl:[{label:`Bien excl. Pantallas`,vals:{pA:r(pB),pB:r(pP)}}],
    union_noexcl:[
      {label:`Bien ∪ Pantallas`,vals:{pA:r(pB),pB:r(pP),pAB:r(bYp)}},
      {label:`Bien ∪ Act.Física`,vals:{pA:r(pB),pB:r(pSi),pAB:r(bYs)}},
    ],
    conjunta_indep:[{label:`Bien ∩ Pantallas`,vals:{pA:r(pB),pB:r(pP)}}],
    conjunta_dep:[{label:`Pantallas→Bien`,vals:{pA:r(pP),pBdA:pantallaSi>0?r(data.filter(d=>esSí(d.bien)&&d.pantalla==='Sí').length/pantallaSi):0}}],
    condicional:[
      {label:`Bien | pantallas`,vals:{pAB:r(bYp),pB:r(pP)}},
      {label:`Bien | act. física`,vals:{pAB:r(bYs),pB:r(pSi)}},
    ],
    diferencia:[{label:`Bien sin pantallas`,vals:{pA:r(pB),pAB:r(bYp)}}],
    prob_total:[{label:`P(bien) total`,vals:{pA1:r(pP),pA2:r(1-pP),pBdA1:pantallaSi>0?r(data.filter(d=>esSí(d.bien)&&d.pantalla==='Sí').length/pantallaSi):0,pBdA2:(n-pantallaSi)>0?r(data.filter(d=>esSí(d.bien)&&d.pantalla==='No').length/(n-pantallaSi)):0}}],
    bayes:[{label:`Bayes pantallas/bien`,vals:{pA1:r(pP),pA2:r(1-pP),pBdA1:pantallaSi>0?r(data.filter(d=>esSí(d.bien)&&d.pantalla==='Sí').length/pantallaSi):0,pBdA2:(n-pantallaSi)>0?r(data.filter(d=>esSí(d.bien)&&d.pantalla==='No').length/(n-pantallaSi)):0}}],
  };
  const list=presets[formulaId]||[];
  if(!list.length) return '';
  return list.map(ej=>`<span class="bi-ex" style="border-color:rgba(245,230,66,.25);color:rgba(245,230,66,.7)" onclick='loadProbEx(${JSON.stringify(ej.vals)})'>${ej.label}</span>`).join('');
}

function loadProbEx(vals){Object.entries(vals).forEach(([k,v])=>{const el=document.getElementById('pf-'+k);if(el) el.value=v;});}
function solveProb(){
  const f=PROB_FORMULAS[currentProbFormula];const vals={};
  for(const fi of f.fields){const el=document.getElementById('pf-'+fi.key);if(!el||el.value===''){notify('Completa todos los campos');return;}vals[fi.key]=parseFloat(el.value);}
  const result=f.solve(vals);
  drawOnBoard('board-prob','idle-prob','sdot-prob','stxt-prob','sstep-prob',f.nombre,f.general,result);
}

// ════════════════════════════════════════════
// PIZARRA — INTERVALOS
// ════════════════════════════════════════════
const IC_FORMULAS={
  ic_prop:{nombre:'IC para proporción',general:'IC₉₅% = p̂ ± z · √(p̂(1−p̂)/n)',
    desc:'Intervalo donde está la verdadera proporción poblacional con 95% de confianza.',
    fields:[{key:'ph',label:'p̂ — proporción observada',color:'a',ph:'ej: 0.60'},{key:'n',label:'n — número de encuestados',color:'b',ph:'ej: 25'},{key:'z',label:'z — valor crítico (95%=1.96)',color:'b',ph:'1.96'}],
    ejemplos:[{label:'De mis datos',auto:true}],
    solve:(v)=>{const se=Math.sqrt(v.ph*(1-v.ph)/v.n),me=v.z*se,lo=Math.max(0,v.ph-me),hi=Math.min(1,v.ph+me);return{steps:[
      {tex:`p̂ = ${v.ph}  |  n = ${v.n}  |  z = ${v.z}`,ann:'Datos del problema'},
      {tex:`SE = √(${v.ph}×${(1-v.ph).toFixed(4)}/${v.n})`,ann:'Calculo el error estándar'},
      {tex:`SE = ${se.toFixed(4)}`,ann:'Resultado del error estándar'},
      {tex:`ME = ${v.z} × ${se.toFixed(4)} = ${me.toFixed(4)}`,ann:'Margen de error = z × SE'},
      {tex:`IC = ${v.ph} ± ${me.toFixed(4)}`,ann:'Aplicar el margen al centro'},
      {tex:`IC = [ ${lo.toFixed(4)} , ${hi.toFixed(4)} ]`,ann:`[${(lo*100).toFixed(1)}% , ${(hi*100).toFixed(1)}%]`,final:true,underline:true},
    ],result:(lo+hi)/2,label:'IC₉₅%',pct:false};}},
  ic_media:{nombre:'IC para media',general:'IC₉₅% = x̄ ± z · (s / √n)',
    desc:'Rango donde está el verdadero promedio poblacional con 95% de confianza.',
    fields:[{key:'xbar',label:'x̄ — promedio muestral',color:'a',ph:'ej: 6.84'},{key:'s',label:'s — desviación estándar',color:'a',ph:'ej: 1.21'},{key:'n',label:'n — encuestados',color:'b',ph:'ej: 25'},{key:'z',label:'z crítico (95%=1.96)',color:'b',ph:'1.96'}],
    ejemplos:[{label:'De mis datos',auto:true}],
    solve:(v)=>{const se=v.s/Math.sqrt(v.n),me=v.z*se,lo=v.xbar-me,hi=v.xbar+me;return{steps:[
      {tex:`x̄=${v.xbar}  s=${v.s}  n=${v.n}  z=${v.z}`,ann:'Datos del problema'},
      {tex:`SE = ${v.s}/√${v.n} = ${v.s}/${Math.sqrt(v.n).toFixed(3)} = ${se.toFixed(4)}`,ann:'Error estándar'},
      {tex:`ME = ${v.z} × ${se.toFixed(4)} = ${me.toFixed(4)}`,ann:'Margen de error'},
      {tex:`IC = ${v.xbar} ± ${me.toFixed(4)}`,ann:'Aplicar margen al promedio'},
      {tex:`IC = [ ${lo.toFixed(4)} , ${hi.toFixed(4)} ] horas`,ann:'Intervalo final',final:true,underline:true},
    ],result:(lo+hi)/2,label:'IC₉₅%',pct:false};}},
};
let currentICFormula='ic_prop';
function selectICFormula(id,btn){
  currentICFormula=id;
  document.querySelectorAll('#bfs-ic .bfs-btn').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  renderICInputs();clearBoardById('board-ic');
}
function renderICInputs(){
  const f=IC_FORMULAS[currentICFormula];
  document.getElementById('ic-board-desc').textContent=f.desc;
  document.getElementById('ic-board-fields').innerHTML=f.fields.map(fi=>`
    <div class="bi-field"><label><span class="bi-badge ${fi.color==='b'?'b':''}">${fi.color==='a'?'A':'B'}</span> ${fi.label}</label>
    <input type="number" id="icf-${fi.key}" step="0.0001" placeholder="${fi.ph}" value=""></div>`).join('');
  document.getElementById('ic-board-ex').innerHTML=`<span class="bi-ex" onclick="loadICFromData()">Cargar de mis datos</span>`;
}
function loadICFromData(){
  const s=stats();if(!s.n){notify('Primero carga datos');return;}
  if(currentICFormula==='ic_prop'){
    const e=document.getElementById('icf-ph');if(e) e.value=s.pBien.toFixed(4);
    const e2=document.getElementById('icf-n');if(e2) e2.value=s.n;
    const e3=document.getElementById('icf-z');if(e3) e3.value=1.96;
  } else {
    e=document.getElementById('icf-xbar');if(e) e.value=s.mean.toFixed(4);
    e2=document.getElementById('icf-s');if(e2) e2.value=s.std.toFixed(4);
    e3=document.getElementById('icf-n');if(e3) e3.value=s.n;
    const e4=document.getElementById('icf-z');if(e4) e4.value=1.96;
  }
  notify('Datos cargados de la encuesta');
}
function solveIC_board(){
  const f=IC_FORMULAS[currentICFormula];const vals={};
  for(const fi of f.fields){const el=document.getElementById('icf-'+fi.key);if(!el||el.value===''){notify('Completa todos los campos');return;}vals[fi.key]=parseFloat(el.value);}
  const result=f.solve(vals);
  drawOnBoard('board-ic','idle-ic','sdot-ic','stxt-ic','sstep-ic',f.nombre,f.general,result);
}
// alias
function solveIC(){solveIC_board();}

// ════════════════════════════════════════════
// PIZARRA — HIPÓTESIS
// ════════════════════════════════════════════
function solveHyp(){
  const ph=parseFloat(document.getElementById('hyp-ph').value);
  const n=parseFloat(document.getElementById('hyp-n').value);
  const p0=parseFloat(document.getElementById('hyp-p0').value)||0.5;
  const zc=parseFloat(document.getElementById('hyp-alpha').value)||1.96;
  if(isNaN(ph)||isNaN(n)){notify('Completa p̂ y n');return;}
  const z=(ph-p0)/Math.sqrt(p0*(1-p0)/n);
  const reject=Math.abs(z)>zc;
  const result={
    steps:[
      {tex:`n = ${n}  p̂ = ${ph}  p₀ = ${p0}  z_crit = ±${zc}`,ann:'Datos del problema'},
      {tex:`z = (p̂ − p₀) / √(p₀(1−p₀)/n)`,ann:'Fórmula del estadístico Z'},
      {tex:`z = (${ph} − ${p0}) / √(${p0}×${(1-p0).toFixed(2)}/${n})`,ann:'Sustituyo los valores'},
      {tex:`z = ${(ph-p0).toFixed(4)} / ${Math.sqrt(p0*(1-p0)/n).toFixed(4)}`,ann:'Simplifico numerador y denominador'},
      {tex:`z = ${z.toFixed(4)}`,ann:'Estadístico calculado',underline:true},
      {tex:`|z| = ${Math.abs(z).toFixed(4)}  ${reject?'>':'<'}  ${zc}  →  ${reject?'RECHAZAR H₀':'No rechazar H₀'}`,ann:reject?'Cae en zona de rechazo':'No cae en zona de rechazo',final:true},
    ],
    result:z,label:'z_calc',pct:false
  };
  drawOnBoard('board-hyp','idle-hyp','sdot-hyp','stxt-hyp','sstep-hyp','Prueba Z para proporción','z = (p̂ − p₀) / √(p₀(1−p₀)/n)',result);
}

// ════════════════════════════════════════════
// NOTIFY
// ════════════════════════════════════════════
let notifT;
function notify(msg){const el=document.getElementById('notif');el.textContent=msg;el.classList.add('show');clearTimeout(notifT);notifT=setTimeout(()=>el.classList.remove('show'),2200);}

// ════════════════════════════════════════════
// INIT
// ════════════════════════════════════════════
renderDescInputs();
renderProbInputs();
renderICInputs();
calcMuestra();
