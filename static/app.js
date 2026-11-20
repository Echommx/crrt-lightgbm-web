'use strict';
const $ = id => document.getElementById(id);
let cfg = null;
let lastResult = null;
// A loaded long-table payload must survive until Predict is clicked. The
// previous UI displayed the full CSV but silently rebuilt a one-row quick-form
// payload at prediction time. Keeping the active payload preserves all 24-hour
// observations and lets the original trend feature engineering run unchanged.
let activePayload = null;

const binaryKeys = ['antibiotic_exposure_7d','carbapenem_exposure_7d','invasive_ventilation_t0','history_diabetes','history_ckd','history_chronic_liver','history_malignancy','history_chronic_lung','history_heart_failure'];
const numericKeys = ['age','antibiotic_classes_7d','heart_rate','temperature','respiratory_rate','map','spo2','wbc','platelets','lactate','creatinine','glucose'];
const dynamicUnits = {heart_rate:'bpm',temperature:'degC',respiratory_rate:'breaths/min',map:'mmHg',spo2:'%',wbc:'10^9/L',platelets:'10^9/L',lactate:'mmol/L',creatinine:'mg/dL',glucose:'mg/dL'};
const staticFields = ['age','sex','specimen_site','antibiotic_exposure_7d','antibiotic_classes_7d','carbapenem_exposure_7d','invasive_ventilation_t0','history_diabetes','history_ckd','history_chronic_liver','history_malignancy','history_chronic_lung','history_heart_failure'];

function download(name, body, type='application/json') {
  const text = typeof body === 'string' ? body : JSON.stringify(body, null, 2);
  const url = URL.createObjectURL(new Blob([text], {type}));
  const a = document.createElement('a'); a.href=url; a.download=name; a.click();
  setTimeout(() => URL.revokeObjectURL(url), 500);
}
function downloadBlob(name, blob) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a'); a.href=url; a.download=name; a.click();
  setTimeout(() => URL.revokeObjectURL(url), 500);
}
function isoValue(id) {
  const v=$(id).value;
  if (!v) throw Error('Prediction time and ICU start time are required.');
  return v.replace('T',' ')+':00';
}
function numberValue(id, required=true) {
  const raw=$(id).value.trim();
  if (!raw) { if(required) throw Error('Please fill '+id+'.'); return null; }
  const n=Number(raw); if(!Number.isFinite(n)) throw Error(id+' must be numeric.'); return n;
}
function selectValue(id) { return $(id).value; }
function buildPayload() {
  const t0=isoValue('t0'), start=isoValue('start_time');
  const s={event_id:'SIMULATED_WEB_EVENT',patient_id:'SIMULATED_PATIENT',encounter_id:'SIMULATED_ENCOUNTER',t0,start_time:start};
  s.age=numberValue('age'); s.sex=selectValue('sex'); s.specimen_site=selectValue('specimen_site');
  for(const id of binaryKeys) s[id]=Number(selectValue(id));
  s.antibiotic_classes_7d=numberValue('antibiotic_classes_7d');
  const observations=[];
  for(const feature of Object.keys(dynamicUnits)) {
    const value=numberValue(feature);
    observations.push({measured_at:t0,feature,value,unit:dynamicUnits[feature]});
  }
  return {static:s,observations};
}
function clearResult() {
  lastResult=null; $('risk').textContent='—'; $('risk').className='risk neutral'; $('score').textContent='—'; $('result-note').textContent='The result will be shown here after prediction.';
}
function setInput(id, value) {
  if ($(id)) $(id).value = value===null || value===undefined ? '' : value;
}
function fillFromPayload(payload) {
  activePayload=payload;
  const s=payload.static||{}; setInput('t0',String(s.t0||'').replace(' ','T').slice(0,16)); setInput('start_time',String(s.start_time||'').replace(' ','T').slice(0,16));
  for(const id of staticFields) setInput(id,s[id]);
  for(const f of Object.keys(dynamicUnits)) {
    const rows=(payload.observations||[]).filter(o=>o.feature===f); if(rows.length) setInput(f,rows[rows.length-1].value);
  }
  clearResult();
}
function parseCsv(text) {
  const rows=[]; let row=[],cell='',quoted=false; const src=text.replace(/^\uFEFF/,'');
  for(let i=0;i<src.length;i++) { const c=src[i]; if(quoted){if(c==='"'&&src[i+1]==='"'){cell+='"';i++;}else if(c==='"')quoted=false;else cell+=c;} else if(c==='"')quoted=true; else if(c===','){row.push(cell);cell='';} else if(c==='\n'){row.push(cell.replace(/\r$/,''));rows.push(row);row=[];cell='';} else cell+=c; }
  if(cell.length||row.length){row.push(cell);rows.push(row);} if(!rows.length) throw Error('CSV needs a header row.');
  const h=rows.shift().map(x=>x.trim()); if(h.some(x=>!x)||new Set(h).size!==h.length) throw Error('CSV headers must be unique and non-empty.');
  return rows.filter(r=>r.some(x=>x.trim())).map(r=>Object.fromEntries(h.map((k,i)=>[k,r[i]??''])));
}
function csvToPayload(text) {
  const rows=parseCsv(text); const need=['measured_at','feature','value','unit']; const missing=need.filter(k=>!(k in rows[0])); if(missing.length) throw Error('CSV missing: '+missing.join(', '));
  const allStatic=['t0','start_time','age','sex','specimen_site',...binaryKeys,'antibiotic_classes_7d']; const s={event_id:'SIMULATED_WEB_EVENT',patient_id:'SIMULATED_PATIENT',encounter_id:'SIMULATED_ENCOUNTER'};
  for(const k of allStatic){const v=rows.find(r=>String(r[k]??'').trim()!=='')?.[k]; if(v===undefined) continue; if(binaryKeys.includes(k)||k==='age'||k==='antibiotic_classes_7d') s[k]=Number(v); else s[k]=String(v).trim();}
  const obs=rows.map((r,i)=>{for(const k of need)if(!String(r[k]??'').trim())throw Error('CSV row '+(i+2)+' missing '+k); const val=Number(r.value); if(!Number.isFinite(val))throw Error('CSV row '+(i+2)+' value must be numeric.'); return {measured_at:String(r.measured_at).trim(),feature:String(r.feature).trim(),value:val,unit:String(r.unit).trim(),...(String(r.available_at??'').trim()?{available_at:String(r.available_at).trim()}: {})};});
  return {static:s,observations:obs};
}
async function calculate() {
  $('error').textContent=''; const b=$('predict'); b.disabled=true;
  try {
    const payload=activePayload || buildPayload(); const response=await fetch('/api/predict',{method:'POST',headers:{'Content-Type':'application/json','X-CRRT-Token':cfg.token},body:JSON.stringify({model:'LightGBM_trend',payload})});
    const out=await response.json(); if(!response.ok) throw Error(out.error||'Prediction failed.'); lastResult=out;
    $('risk').textContent=out.risk_label; $('risk').className='risk '+(out.binary_prediction===1?'high':'low'); $('score').textContent=(out.score*100).toFixed(1); $('result-note').textContent='Study-defined binary result. Research use only.';
  } catch(e) { $('error').textContent=e.message; } finally { b.disabled=!cfg||!cfg.available; }
}
async function init() {
  try {
    const r=await fetch('/api/config'); if(!r.ok) throw Error('Cannot read local settings.'); cfg=await r.json(); fillFromPayload(cfg.example); $('status').textContent=cfg.available?'Ready':'Model package not found'; $('predict').disabled=!cfg.available;
  } catch(e) { $('error').textContent=e.message; }
}
$('predict').onclick=calculate;
$('download-xlsx').onclick=async()=>{if(!cfg)return;try{const response=await fetch('/api/example.xlsx');if(!response.ok)throw Error('Excel example is unavailable.');downloadBlob('simulated_input.xlsx',await response.blob());}catch(err){$('error').textContent=err.message;}};
$('upload').onchange=async e=>{const f=e.target.files?.[0];if(!f)return;try{if(!f.name.toLowerCase().endsWith('.xlsx'))throw Error('Please upload an Excel workbook (.xlsx).');const response=await fetch('/api/parse-xlsx',{method:'POST',headers:{'Content-Type':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet','X-CRRT-Token':cfg.token},body:await f.arrayBuffer()});const out=await response.json();if(!response.ok)throw Error(out.error||'Excel upload failed.');fillFromPayload(out.payload);$('status').textContent='Loaded '+f.name+' (full trend input)';}catch(err){$('error').textContent=err.message;}};
const clearActiveUpload=()=>{activePayload=null; clearResult(); if(cfg&&cfg.available)$('status').textContent='Quick form ready';};
document.querySelectorAll('input,select').forEach(el=>{el.addEventListener('input',clearActiveUpload);el.addEventListener('change',clearActiveUpload);});
init();
