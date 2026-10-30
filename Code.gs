// MICRO PLASTICS TRAINING PORTAL – sync backend (Google Apps Script)
const ADMIN_KEY = 'CHANGE_THIS_SECRET';   // <-- set your own secret key

function sh(n){const ss=SpreadsheetApp.getActive();return ss.getSheetByName(n)||ss.insertSheet(n);}
function out(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);}

function doGet(e){
  const p=e.parameter, a=p.a;
  if(a==='cfg') return out(getCfg());
  if(a==='prog') return out(getProg(p.u));
  if(a==='all'){ if(p.key!==ADMIN_KEY) return out({error:'key'}); return out(getAll()); }
  return out({error:'bad'});
}
function doPost(e){
  const b=JSON.parse(e.postData.contents);
  const lock=LockService.getScriptLock(); lock.waitLock(15000);
  try{
    if(b.a==='prog'){
      const adm=(b.key===ADMIN_KEY), cur=getProg(b.u);
      if(b.d===null){ if(!adm) return out({error:'key'}); setProg(b.u,null); return out({ok:1}); }
      const nr=(b.d.retrain||0), cr=(cur.retrain||0);
      if(!adm && nr>cr) return out({error:'key'});   // only admin can start a new training cycle
      if(!adm && nr<cr) return out({stale:1});        // ignore old-cycle data from a learner's device
      setProg(b.u,b.d); return out({ok:1});
    }
    if(b.a==='cfg'){ if(b.key!==ADMIN_KEY) return out({error:'key'}); setCfg(b.d); return out({ok:1}); }
  } finally { lock.releaseLock(); }
  return out({error:'bad'});
}
function getCfg(){
  const s=sh('cfg'), n=s.getLastRow(); if(!n) return {};
  const t=s.getRange(1,1,n,1).getValues().map(r=>r[0]).join('');
  try{return JSON.parse(t);}catch(e){return {};}
}
function setCfg(d){
  const s=sh('cfg'); s.clear(); s.getRange('A:A').setNumberFormat('@');
  const t=JSON.stringify(d), rows=[];
  for(let i=0;i<t.length;i+=30000) rows.push([t.substring(i,i+30000)]);
  s.getRange(1,1,rows.length,1).setValues(rows);
}
function getProg(u){
  const s=sh('progress'), v=s.getDataRange().getValues();
  for(const r of v) if(r[0]===u){try{return JSON.parse(r[1]);}catch(e){}}
  return {};
}
function getAll(){
  const v=sh('progress').getDataRange().getValues(), o={};
  v.forEach(r=>{ if(r[0]){try{o[r[0]]=JSON.parse(r[1]);}catch(e){}} });
  return o;
}
function setProg(u,d){
  const s=sh('progress'), v=s.getDataRange().getValues();
  for(let i=0;i<v.length;i++) if(v[i][0]===u){
    if(d===null) s.deleteRow(i+1); else s.getRange(i+1,1,1,3).setValues([[u,JSON.stringify(d),new Date()]]);
    return;
  }
  if(d!==null) s.appendRow([u,JSON.stringify(d),new Date()]);
}
