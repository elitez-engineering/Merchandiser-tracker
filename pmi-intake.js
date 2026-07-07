/* PMI Intake — takes over the manual WhatsApp-to-Excel transcription duty.
 * Parses exported WhatsApp chat logs (.txt/.zip) into per-outlet visit records,
 * lets an admin review/correct/confirm every field before it is saved, then
 * exports a colour-coded .xlsx matching the master "PM Distribution Check" layout.
 *
 * Nothing here is ever silently finalized — every parsed price/feedback value
 * is a suggestion the admin must confirm or edit before export.
 */
(function(){

// ── Week windows for the H2 2026 cycle (from the master Excel's own header text) ──
var PMI_WEEKS=[
  {col:'wk1',label:'Week 1 (22 Jun - 26 Jun)',monday:'2026-06-22'},
  {col:'wk2',label:'Week 2 (29 Jun - 3 Jul)', monday:'2026-06-29'},
  {col:'wk3',label:'Week 3 (6 Jul - 10 Jul)', monday:'2026-07-06'},
  {col:'wk4',label:'Week 4 (13 Jul - 17 Jul)',monday:'2026-07-13'}
];

// Set this once the Cloudflare Worker is deployed (see worker/README.md). Left
// blank means the "Run OCR" button is disabled with an explanatory tooltip.
var PMI_OCR_ENDPOINT='';

var _skuRef=null, _abbrMap=null; // loaded lazily from data/pmi-sku-reference.json
var _queue=[];        // pending records awaiting admin review (Firebase-synced)
var _confirmed={};    // custId -> confirmed record (Firebase-synced)
var _mediaFiles={};   // filename -> {blob, url} extracted from an uploaded .zip
var _activeId=null;   // queue record currently open in the editor
var _pmiBooted=false;

function esc(s){ return String(s).replace(/[.*+?^${}()|[\]\\]/g,'\\$&'); }
function uid(){ return 'pk-'+Date.now().toString(36)+'-'+Math.random().toString(36).slice(2,8); }

// Chat text, filenames and outlet fields all end up in innerHTML — escape before interpolating.
function escHtml(s){
  return String(s==null?'':s).replace(/[&<>"']/g,function(c){
    return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c];
  });
}

function injectStyles(){
  if(document.getElementById('pmiStyles')) return;
  var css=''
    +'.pmi-layout{display:grid;grid-template-columns:340px 1fr;gap:16px;align-items:start}'
    +'@media (max-width:900px){.pmi-layout{grid-template-columns:1fr}}'
    +'.pmi-col{background:#fff;border:1px solid #e5e7eb;border-radius:12px;padding:14px;box-shadow:0 1px 3px rgba(0,0,0,.05)}'
    +'.pmi-h{font-size:12px;font-weight:700;color:#374151;text-transform:uppercase;letter-spacing:.04em;margin-bottom:10px}'
    +'.pmi-qrow{padding:8px 10px;border-radius:8px;cursor:pointer;margin-bottom:6px;border:1.5px solid #eee;font-size:12px}'
    +'.pmi-qrow:hover{border-color:#0071e3}'
    +'.pmi-qrow.active{border-color:#0071e3;background:#f0f7ff}'
    +'.pmi-qrow.bad{border-left:3px solid #e53935}'
    +'.pmi-qrow .id{font-weight:700}'
    +'.pmi-qrow .nm{color:#6b7280;font-size:11px}'
    +'.pmi-field{margin-bottom:10px}'
    +'.pmi-field label{font-size:11px;font-weight:600;color:#555;display:block;margin-bottom:3px}'
    +'.pmi-field input,.pmi-field textarea,.pmi-field select{width:100%;padding:7px 9px;border:1.5px solid #e0e0e0;border-radius:8px;font-size:13px;outline:none;font-family:inherit}'
    +'.pmi-field input:focus,.pmi-field textarea:focus,.pmi-field select:focus{border-color:#0071e3}'
    +'.pmi-warn{color:#e53935;font-size:11px;font-weight:600;margin-top:3px}'
    +'.pmi-ok{color:#34c759;font-size:11px;font-weight:600;margin-top:3px}'
    +'.pmi-ptable{width:100%;border-collapse:collapse;font-size:12px;margin-top:6px}'
    +'.pmi-ptable th{text-align:left;color:#9ca3af;font-size:10px;text-transform:uppercase;padding:4px 6px;border-bottom:1px solid #eee}'
    +'.pmi-ptable td{padding:4px 6px;border-bottom:1px solid #f5f5f5;vertical-align:top}'
    +'.pmi-ptable select,.pmi-ptable input{padding:5px 6px;font-size:12px;border:1px solid #e0e0e0;border-radius:6px}'
    +'.pmi-x{color:#e53935;cursor:pointer;font-weight:700;padding:2px 6px}'
    +'.pmi-src{color:#aaa;font-size:10px}'
    +'.pmi-btnrow{display:flex;gap:8px;margin-top:14px}'
    +'.pmi-btn{padding:8px 14px;border-radius:8px;font-size:12px;font-weight:600;cursor:pointer;border:1.5px solid #e5e7eb;background:#fff}'
    +'.pmi-btn.primary{background:#0071e3;border-color:#0071e3;color:#fff}'
    +'.pmi-btn.danger{background:#fff;border-color:#e53935;color:#e53935}'
    +'.pmi-btn:disabled{opacity:.4;cursor:not-allowed}'
    +'.pmi-att{display:inline-flex;align-items:center;gap:6px;background:#f5f5f7;border-radius:8px;padding:4px 8px;font-size:11px;margin:3px 6px 3px 0}'
    +'.pmi-att img{width:28px;height:28px;object-fit:cover;border-radius:4px}'
    +'.pmi-empty{color:#9ca3af;font-size:12px;padding:20px;text-align:center}';
  var st=document.createElement('style'); st.id='pmiStyles'; st.textContent=css;
  document.head.appendChild(st);
}

function loadSkuRef(){
  if(_skuRef) return Promise.resolve(_skuRef);
  return fetch('data/pmi-sku-reference.json').then(function(r){return r.json();}).then(function(json){
    _skuRef=json;
    _abbrMap={};
    json.skus.forEach(function(s){
      var k=s.abbr.toUpperCase();
      (_abbrMap[k]=_abbrMap[k]||[]).push(s);
    });
    return json;
  });
}

// ── Firebase sync ──
function pmiRef(path){ return dbRef ? dbRef.child('pmi_intake'+(path?'/'+path:'')) : null; }

function attachFirebaseListeners(){
  if(!dbRef || _pmiBooted) return;
  _pmiBooted=true;
  pmiRef('queue').on('value',function(snap){
    var v=snap.val()||{};
    _queue=Object.keys(v).map(function(k){ var r=v[k]; r.id=k; return r; });
    if(document.getElementById('intakePanel').classList.contains('open')) renderIntake();
  });
  pmiRef('confirmed').on('value',function(snap){
    _confirmed=snap.val()||{};
    if(document.getElementById('intakePanel').classList.contains('open')) renderIntake();
  });
}

function saveQueueRecord(rec){ if(pmiRef()) pmiRef('queue/'+rec.id).set(rec); }
function removeQueueRecord(id){ if(pmiRef()) pmiRef('queue/'+id).remove(); }
function saveConfirmed(custId,rec){ if(pmiRef()) pmiRef('confirmed/'+custId).set(rec); }

// ── WhatsApp export parsing ──
var LINE_ANDROID=/^(\d{1,2}\/\d{1,2}\/\d{2,4}),\s*(\d{1,2}:\d{2}\s*(?:[ap]m)?)\s*-\s*([^:]+):\s*([\s\S]*)$/i;
var LINE_IOS=/^\[(\d{1,2}\/\d{1,2}\/\d{2,4}),\s*(\d{1,2}:\d{2}:\d{2}\s*(?:[ap]m)?)\]\s*([^:]+):\s*([\s\S]*)$/i;
var ID_RE=/\b(?:cust\s*id|id)\s*[:\-]?\s*(\d{3,6})\b/i;
var ATTACH_RE=/<attached:\s*([^>]+)>/gi;

function parseMessages(text){
  var lines=text.split(/\r?\n/);
  var messages=[]; var cur=null;
  lines.forEach(function(line){
    var m=line.match(LINE_ANDROID)||line.match(LINE_IOS);
    if(m){
      if(cur) messages.push(cur);
      cur={date:m[1],time:m[2],sender:m[3].trim(),text:m[4]};
    } else if(cur){
      cur.text += '\n'+line;
    }
  });
  if(cur) messages.push(cur);
  return messages;
}

function groupIntoRecords(messages){
  var records=[]; var cur=null;
  messages.forEach(function(msg){
    var idMatch=msg.text.match(ID_RE);
    if(idMatch){
      if(cur) records.push(cur);
      cur={custId:Number(idMatch[1]), messages:[msg]};
    } else if(cur){
      cur.messages.push(msg);
    }
    // messages before the first ID line (chit-chat) are dropped — nothing to attach them to
  });
  if(cur) records.push(cur);
  return records;
}

function extractPriceSuggestions(fullText){
  var out=[];
  if(!_abbrMap) return out;
  var abbrs=Object.keys(_abbrMap).sort(function(a,b){return b.length-a.length;});
  abbrs.forEach(function(abbr){
    var re=new RegExp('(?:^|[\\s,;/])'+esc(abbr)+'[\\s:]+(Y|[0-9]{1,2}(?:\\.[0-9]{1,2})?)(?:\\s*-\\s*([0-9]{1,2}(?:\\.[0-9]{1,2})?))?','gi');
    var m;
    while((m=re.exec(fullText))){
      var candidates=_abbrMap[abbr];
      out.push({
        abbr:abbr,
        candidates:candidates,
        chosenExcelCol: candidates.length===1?candidates[0].excelCol:null,
        value:m[1].toUpperCase()==='Y'?'Y':m[1],
        altValue:m[2]||null,
        source:'chat',
        raw:m[0].trim()
      });
    }
  });
  return out;
}

function stripKnownTokens(text,priceSuggestions){
  var t=text.replace(ID_RE,' ').replace(ATTACH_RE,' ');
  priceSuggestions.forEach(function(p){ t=t.split(p.raw).join(' '); });
  return t.replace(/\s+/g,' ').trim();
}

function buildRecordFromGroup(group){
  var fullText=group.messages.map(function(m){return m.text;}).join('\n');
  var attachments=[]; var am;
  ATTACH_RE.lastIndex=0;
  while((am=ATTACH_RE.exec(fullText))) attachments.push(am[1].trim());
  var prices=extractPriceSuggestions(fullText);
  var feedback=stripKnownTokens(fullText,prices);
  var outlet=(typeof ALL_OUTLETS!=='undefined')?ALL_OUTLETS.find(function(o){return o.i===group.custId;}):null;
  return {
    id:uid(),
    custId:group.custId,
    outletName:outlet?outlet.n:null,
    feedback:feedback,
    prices:prices,
    attachments:attachments.map(function(f){return {filename:f};}),
    rawText:fullText,
    status:'pending'
  };
}

// ── File intake (.txt / .zip) ──
function handleIntakeFile(file){
  var status=document.getElementById('pmiUploadStatus');
  if(status) status.textContent='Reading '+file.name+'…';
  var isZip=/\.zip$/i.test(file.name);
  var afterText=function(text){
    loadSkuRef().then(function(){
      var messages=parseMessages(text);
      var groups=groupIntoRecords(messages);
      var recs=groups.map(buildRecordFromGroup);
      recs.forEach(saveQueueRecord); // Firebase listener repopulates _queue and re-renders
      if(status) status.textContent='Parsed '+recs.length+' outlet-visit record(s) from '+messages.length+' messages. Review each below before it counts.';
    });
  };
  if(isZip){
    if(typeof JSZip==='undefined'){ if(status) status.textContent='JSZip failed to load — check your connection and retry.'; return; }
    JSZip.loadAsync(file).then(function(zip){
      var txtEntry=null;
      zip.forEach(function(path,entry){ if(!entry.dir && /\.txt$/i.test(path) && !txtEntry) txtEntry=entry; });
      var mediaEntries=[];
      zip.forEach(function(path,entry){ if(!entry.dir && /\.(jpe?g|png|webp)$/i.test(path)) mediaEntries.push(entry); });
      var mediaPromise=Promise.all(mediaEntries.map(function(entry){
        return entry.async('blob').then(function(blob){
          _mediaFiles[entry.name.split('/').pop()]={blob:blob,url:URL.createObjectURL(blob)};
        });
      }));
      if(!txtEntry){ if(status) status.textContent='No .txt chat log found inside the zip.'; return; }
      Promise.all([txtEntry.async('string'),mediaPromise]).then(function(res){ afterText(res[0]); });
    }).catch(function(e){ if(status) status.textContent='Could not read zip: '+e.message; });
  } else {
    var reader=new FileReader();
    reader.onload=function(){ afterText(reader.result); };
    reader.readAsText(file);
  }
}

// ── OCR (Cloudflare Worker, Claude vision) ──
function runOcrOnAttachment(recId,filename){
  var media=_mediaFiles[filename];
  if(!media){ alert('This photo was not found in the uploaded zip — re-upload the export with media included.'); return; }
  if(!PMI_OCR_ENDPOINT){ alert('OCR endpoint not configured yet. Deploy worker/ and set PMI_OCR_ENDPOINT in pmi-intake.js.'); return; }
  var rec=_queue.find(function(r){return r.id===recId;}); if(!rec) return;
  var reader=new FileReader();
  reader.onload=function(){
    var b64=reader.result.split(',')[1];
    fetch(PMI_OCR_ENDPOINT,{
      method:'POST',
      headers:{'Content-Type':'application/json'},
      body:JSON.stringify({image:b64,mediaType:media.blob.type||'image/jpeg',skuNames:_skuRef.skus.map(function(s){return {excelCol:s.excelCol,name:s.name};})})
    }).then(function(r){return r.json();}).then(function(json){
      (json.detections||[]).forEach(function(d){
        rec.prices.push({abbr:'(photo)',candidates:[{excelCol:d.excelCol,name:d.name}],chosenExcelCol:d.excelCol,value:d.price||'Y',altValue:null,source:'ocr',raw:'OCR: '+filename});
      });
      saveQueueRecord(rec);
    }).catch(function(e){ alert('OCR request failed: '+e.message); });
  };
  reader.readAsDataURL(media.blob);
}

// ── Review / edit / confirm ──
function selectRecord(id){ _activeId=id; renderIntake(); }

function updateRecordField(id,field,val){
  var rec=_queue.find(function(r){return r.id===id;}); if(!rec) return;
  rec[field]=val;
  if(field==='custId'){
    var outlet=(typeof ALL_OUTLETS!=='undefined')?ALL_OUTLETS.find(function(o){return o.i===Number(val);}):null;
    rec.outletName=outlet?outlet.n:null;
  }
  saveQueueRecord(rec);
}

function updatePriceRow(recId,idx,field,val){
  var rec=_queue.find(function(r){return r.id===recId;}); if(!rec) return;
  var row=rec.prices[idx]; if(!row) return;
  if(field==='chosenExcelCol') row.chosenExcelCol=val?Number(val):null;
  else if(field==='value') row.value=val;
  saveQueueRecord(rec);
}

function removePriceRow(recId,idx){
  var rec=_queue.find(function(r){return r.id===recId;}); if(!rec) return;
  rec.prices.splice(idx,1); saveQueueRecord(rec);
}

function addManualPriceRow(recId,excelCol,value){
  var rec=_queue.find(function(r){return r.id===recId;}); if(!rec) return;
  var sku=_skuRef.skus.find(function(s){return s.excelCol===Number(excelCol);}); if(!sku) return;
  rec.prices.push({abbr:sku.abbr,candidates:[sku],chosenExcelCol:sku.excelCol,value:value||'Y',altValue:null,source:'manual',raw:'manual entry'});
  saveQueueRecord(rec);
}

function discardRecord(id){
  if(!confirm('Discard this parsed record? This cannot be undone.')) return;
  removeQueueRecord(id);
  if(_activeId===id) _activeId=null;
}

function confirmRecord(id){
  var rec=_queue.find(function(r){return r.id===id;}); if(!rec) return;
  if(!rec.custId || !rec.outletName){ alert('Cust ID does not match a known outlet — fix it before confirming.'); return; }
  var unresolved=rec.prices.filter(function(p){return !p.chosenExcelCol;});
  if(unresolved.length){ alert(unresolved.length+' price row(s) still need an exact SKU picked (ambiguous brand shorthand) before you can confirm.'); return; }
  var priceMap={};
  rec.prices.forEach(function(p){ priceMap[p.chosenExcelCol]=p.value; });
  var out={
    custId:rec.custId,
    outletName:rec.outletName,
    feedback:rec.feedback,
    prices:priceMap,
    confirmedAt:new Date().toISOString(),
    confirmedBy:(typeof cUser!=='undefined'&&cUser)||'admin'
  };
  saveConfirmed(rec.custId,out);
  removeQueueRecord(id);
  if(_activeId===id) _activeId=null;
}

// ── Export: colour-coded .xlsx matching the master file layout ──
function weekColumnForOutlet(custId){
  if(typeof visited==='undefined' || typeof visitedMeta==='undefined') return {};
  var status=visited[custId]; var meta=visitedMeta[custId];
  var out={};
  if(status && meta && meta.date && typeof getWeekKey==='function'){
    var wk=getWeekKey(meta.date);
    var match=PMI_WEEKS.find(function(w){return w.monday===wk;});
    if(match) out[match.col]=status;
  }
  return out;
}

function exportIntakeXLSX(){
  if(typeof ExcelJS==='undefined'){ alert('ExcelJS failed to load — check your connection and retry.'); return; }
  loadSkuRef().then(function(){
    var wb=new ExcelJS.Workbook();
    var ws=wb.addWorksheet('Priority1 - OC A');
    var fixedCols=[
      {header:'Idebbf',key:'s',width:10},
      {header:'Class',key:'c',width:9},
      {header:'Cust ID',key:'i',width:10},
      {header:'Name',key:'n',width:32},
      {header:'Combined Address',key:'a',width:44},
      {header:'Postal Code',key:'p',width:11},
      {header:'Status',key:'status',width:8},
      {header:'Week 1 (22 Jun - 26 Jun)',key:'wk1',width:16},
      {header:'Week 2 (29 Jun - 3 Jul)',key:'wk2',width:16},
      {header:'Week 3 (6 Jul - 10 Jul)',key:'wk3',width:16},
      {header:'Week 4 (13 Jul - 17 Jul)',key:'wk4',width:16},
      {header:'Feedback',key:'feedback',width:34}
    ];
    var skuCols=_skuRef.skus.map(function(s){ return {header:s.name,key:'sku'+s.excelCol,width:9}; });
    ws.columns=fixedCols.concat(skuCols);

    var DIST_ARGB={PMI:'FFFF0000',BAT:'FF0070C0',JTI:'FF00B050',OTHER:'FFED7D31'};
    var headerRow=ws.getRow(1);
    fixedCols.forEach(function(c,i){ headerRow.getCell(i+1).font={bold:true}; });
    _skuRef.skus.forEach(function(s,i){
      var cell=headerRow.getCell(fixedCols.length+i+1);
      cell.font={bold:true,color:{argb:'FFFFFFFF'}};
      cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:DIST_ARGB[s.distributor]||DIST_ARGB.OTHER}};
    });

    var outlets=(typeof ALL_OUTLETS!=='undefined')?ALL_OUTLETS:[];
    outlets.forEach(function(o){
      var row={s:o.s,c:o.c,i:o.i,n:o.n,a:o.a,p:o.p};
      Object.assign(row,weekColumnForOutlet(o.i));
      var conf=_confirmed[o.i];
      if(conf){
        row.feedback=conf.feedback;
        Object.keys(conf.prices||{}).forEach(function(col){ row['sku'+col]=conf.prices[col]; });
      }
      ws.addRow(row);
    });

    wb.xlsx.writeBuffer().then(function(buf){
      var blob=new Blob([buf],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
      var a=document.createElement('a');
      a.href=URL.createObjectURL(blob);
      a.download='PMI_Distribution_Export_'+new Date().toISOString().slice(0,10)+'.xlsx';
      document.body.appendChild(a); a.click(); a.remove();
    });
  });
}

// ── Rendering ──
function toggleIntake(){
  injectStyles();
  attachFirebaseListeners();
  var p=document.getElementById('intakePanel');
  if(p.classList.contains('open')){ p.classList.remove('open'); }
  else{ p.classList.add('open'); loadSkuRef().then(renderIntake); }
}

function priceRowHtml(rec,p,idx){
  var opts=p.candidates.map(function(c){
    return '<option value="'+c.excelCol+'"'+(p.chosenExcelCol===c.excelCol?' selected':'')+'>'+escHtml(c.name)+'</option>';
  }).join('');
  var needsPick=p.candidates.length>1;
  return '<tr>'
    +'<td>'+escHtml(p.abbr)+(p.source==='ocr'?' <span class="pmi-src">OCR</span>':p.source==='manual'?' <span class="pmi-src">manual</span>':'')+'</td>'
    +'<td><select onchange="_pmiUpdatePrice(\''+rec.id+'\','+idx+',\'chosenExcelCol\',this.value)">'
      +(needsPick&&!p.chosenExcelCol?'<option value="">— pick exact SKU —</option>':'')+opts+'</select></td>'
    +'<td><input value="'+escHtml(p.value||'')+'" style="width:64px" onchange="_pmiUpdatePrice(\''+rec.id+'\','+idx+',\'value\',this.value)"></td>'
    +'<td class="pmi-src" title="'+escHtml(p.raw||'')+'">'+escHtml((p.raw||'').slice(0,28))+'</td>'
    +'<td><span class="pmi-x" onclick="_pmiRemovePrice(\''+rec.id+'\','+idx+')">&times;</span></td>'
    +'</tr>';
}

function editorHtml(rec){
  var outlet=(typeof ALL_OUTLETS!=='undefined')?ALL_OUTLETS.find(function(o){return o.i===rec.custId;}):null;
  var validity=outlet?'<div class="pmi-ok">&check; '+escHtml(outlet.n)+' ('+escHtml(outlet.c)+', '+escHtml(outlet.a)+')</div>':'<div class="pmi-warn">&#9888; Cust ID not found in outlet list — check for a typo.</div>';
  var attHtml=rec.attachments.map(function(a){
    var media=_mediaFiles[a.filename];
    var safeName=escHtml(a.filename);
    // JSON.stringify gives a properly JS-escaped string literal; escHtml then makes it
    // safe to sit inside the onclick="..." HTML attribute (handles embedded quotes either way).
    var jsArgs=escHtml(JSON.stringify(rec.id)+','+JSON.stringify(a.filename));
    return '<span class="pmi-att">'+(media?'<img src="'+media.url+'">':'📎')+' '+safeName
      +' <span class="pmi-btn" style="padding:2px 6px;font-size:10px" onclick="_pmiRunOcr('+jsArgs+')">Run OCR</span></span>';
  }).join('') || '<span class="pmi-src">no photos referenced in this chat group</span>';

  var priceRows=rec.prices.map(function(p,i){return priceRowHtml(rec,p,i);}).join('');
  var skuOptions=_skuRef.skus.map(function(s){return '<option value="'+s.excelCol+'">'+escHtml(s.name)+'</option>';}).join('');

  return '<div class="pmi-field"><label>Cust ID</label><input value="'+escHtml(rec.custId)+'" onchange="_pmiUpdateField(\''+rec.id+'\',\'custId\',this.value)">'+validity+'</div>'
    +'<div class="pmi-field"><label>Photos referenced</label><div>'+attHtml+'</div></div>'
    +'<div class="pmi-field"><label>Feedback (verbatim, editable)</label><textarea rows="3" onchange="_pmiUpdateField(\''+rec.id+'\',\'feedback\',this.value)">'+escHtml(rec.feedback||'')+'</textarea></div>'
    +'<div class="pmi-field"><label>SKU prices detected — pick exact variant where ambiguous</label>'
      +'<table class="pmi-ptable"><thead><tr><th>Brand</th><th>SKU</th><th>Price</th><th>Source text</th><th></th></tr></thead><tbody>'+priceRows+'</tbody></table>'
      +'<div style="display:flex;gap:6px;margin-top:8px">'
        +'<select id="pmiAddSku_'+rec.id+'" style="flex:1"><option value="">+ add a SKU manually…</option>'+skuOptions+'</select>'
        +'<input id="pmiAddVal_'+rec.id+'" placeholder="price or Y" style="width:80px">'
        +'<span class="pmi-btn" onclick="_pmiAddManual(\''+rec.id+'\')">Add</span>'
      +'</div>'
    +'</div>'
    +'<div class="pmi-btnrow">'
      +'<button class="pmi-btn primary" onclick="_pmiConfirm(\''+rec.id+'\')">&check; Confirm &amp; save</button>'
      +'<button class="pmi-btn danger" onclick="_pmiDiscard(\''+rec.id+'\')">Discard</button>'
    +'</div>';
}

function renderIntake(){
  var content=document.getElementById('intakeContent'); if(!content) return;
  var confirmedCount=Object.keys(_confirmed).length;
  var queueHtml=_queue.length?_queue.map(function(r){
    var bad=!r.custId||!(typeof ALL_OUTLETS!=='undefined'&&ALL_OUTLETS.find(function(o){return o.i===r.custId;}));
    return '<div class="pmi-qrow'+(r.id===_activeId?' active':'')+(bad?' bad':'')+'" onclick="_pmiSelect(\''+r.id+'\')">'
      +'<div class="id">'+escHtml(r.custId||'?')+'</div>'
      +'<div class="nm">'+escHtml(r.outletName||'unknown outlet')+' · '+r.prices.length+' price(s)</div>'
      +'</div>';
  }).join('') : '<div class="pmi-empty">No pending records. Upload a WhatsApp export to start.</div>';

  var active=_queue.find(function(r){return r.id===_activeId;});

  var html='<div class="ana-topbar"><div class="ana-title">📋 PMI Intake — WhatsApp → Master Excel</div>'
    +'<div class="ana-topbar-btns">'
    +'<button class="ana-tbtn blue" onclick="_pmiExport()">↓ Export .xlsx ('+confirmedCount+' confirmed)</button>'
    +'<button class="ana-tbtn" onclick="toggleIntake()">✕ Close</button>'
    +'</div></div>';
  html+='<div class="ana-body">';
  html+='<div class="upload-zone" style="margin-bottom:16px" onclick="document.getElementById(\'pmiFile\').click()">'
    +'📤 Click to upload an exported WhatsApp chat (.txt or .zip with media)'
    +'<input type="file" id="pmiFile" accept=".txt,.zip" style="display:none">'
    +'<div id="pmiUploadStatus" style="font-size:12px;color:#888;margin-top:8px"></div></div>';
  html+='<div class="pmi-layout">'
    +'<div class="pmi-col"><div class="pmi-h">Pending review ('+_queue.length+')</div>'+queueHtml+'</div>'
    +'<div class="pmi-col">'+(active?editorHtml(active):'<div class="pmi-empty">Select a record on the left to review it.</div>')+'</div>'
    +'</div></div>';
  content.innerHTML=html;

  var fileInput=document.getElementById('pmiFile');
  if(fileInput) fileInput.addEventListener('change',function(){ if(this.files[0]) handleIntakeFile(this.files[0]); });
}

// Global bridges for inline handlers
window.toggleIntake=toggleIntake;
window._pmiSelect=selectRecord;
window._pmiUpdateField=updateRecordField;
window._pmiUpdatePrice=updatePriceRow;
window._pmiRemovePrice=removePriceRow;
window._pmiAddManual=function(recId){
  var sel=document.getElementById('pmiAddSku_'+recId), val=document.getElementById('pmiAddVal_'+recId);
  if(sel&&sel.value) addManualPriceRow(recId,sel.value,val?val.value:'Y');
};
window._pmiRunOcr=runOcrOnAttachment;
window._pmiConfirm=confirmRecord;
window._pmiDiscard=discardRecord;
window._pmiExport=exportIntakeXLSX;

})();
