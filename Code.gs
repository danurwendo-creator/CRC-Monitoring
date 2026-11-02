const SPREADSHEET_ID = '1nvp8BsIfIWxftSwAFqjhw223RUhQCiHx8e-KQGmD2RQ';
const SOURCE_SHEET = 'RAW';
const NOTES_SHEET = 'CRC_Notes';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('Mitratel | CRC UMT Monitoring').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getDashboardData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SOURCE_SHEET);
  if (!sh) throw new Error('Tab sumber "' + SOURCE_SHEET + '" tidak ditemukan.');
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return {records:[],summary:emptySummary_(),trend:[],kabupaten:[],notes:{},updatedAt:new Date().toISOString(),sourceSheet:SOURCE_SHEET};
  const headers = values[0].map(v=>String(v||'').trim()), idx={};
  headers.forEach((h,i)=>{if(h&&idx[h]===undefined)idx[h]=i;});
  const get=(row,name)=>idx[name]===undefined?'':row[idx[name]];
  const fmt=v=>v instanceof Date?Utilities.formatDate(v,Session.getScriptTimeZone(),'yyyy-MM-dd'):(v===null||v===undefined?'':String(v).trim());
  const notes=loadNotes_(ss), records=[];
  for(let r=1;r<values.length;r++){
    const row=values[r], ne=fmt(get(row,'ne_name')),fe=fmt(get(row,'fe_name'));
    const nePort=fmt(get(row,'NE Port')),fePort=fmt(get(row,'FE Port'));
    const link=fmt(get(row,'Section Name'))||[fmt(get(row,'NE Site Name')),fmt(get(row,'FE Site Name'))].filter(Boolean).join(' <> ');
    if(!link&&!ne&&!fe)continue;
    const rawP=fmt(get(row,'Priority')).toUpperCase(),pm=rawP.match(/P?([0-3])/),priority=pm?Number(pm[1]):null;
    const sourceStatus=fmt(get(row,'Actual')).toLowerCase()==='close'?'Close':'Open';
    const key=String(r+1), saved=notes[key]||{};
    const rawRecoveryStatus=fmt(get(row,'Recovery Status'));
    const rawDateAction=fmt(get(row,'Recovery Date Action'));
    const rawDateClose=fmt(get(row,'Recovery Date Close'));
    const rawFollowup=fmt(get(row,'Recovery Follow-up / RCA'));
    const status=['Close','Open','In Progress'].includes(rawRecoveryStatus)?rawRecoveryStatus:
      (['Close','Open','In Progress'].includes(saved.status)?saved.status:sourceStatus);
    records.push({
      rowNumber:r+1,key:key,towerNE:fmt(get(row,'Tower ID NE')),towerFE:fmt(get(row,'Tower ID FE')),
      siteNE:fmt(get(row,'NE Site Name')),siteFE:fmt(get(row,'FE Site Name')),neId:ne,feId:fe,
      nePort:nePort,fePort:fePort,link:link,priority:priority,priorityLabel:priority===null?'N/A':'P'+priority,
      sourceStatus:sourceStatus,status:status,recoveryStatus:rawRecoveryStatus,dateAction:rawDateAction||saved.dateAction||'',dateClose:rawDateClose||saved.dateClose||fmt(get(row,'Close Date')),
      taskType:fmt(get(row,'FLP/OSP/FOP')),taskOwner:fmt(get(row,'Task Owner')),ticket:fmt(get(row,'INCM/TR/WO')),
      reqDate:fmt(get(row,'Req Date')),closeDate:fmt(get(row,'Close Date')),aging:fmt(get(row,'Agging')),
      agingCategory:fmt(get(row,'Aging Category')),region:fmt(get(row,'Region')),kab:fmt(get(row,'Kab')),
      vip:fmt(get(row,'VIP GRV')),rsl:fmt(get(row,'REMARKS  IN/OUT')),rslAfter:fmt(get(row,'RSL After')),
      flag:fmt(get(row,'Flag')),originalFlag:fmt(get(row,'Original Flag')),mitigation:fmt(get(row,'Mitigation Remark')),
      remark:fmt(get(row,'Remark')),timeline:fmt(get(row,'Timeline')),propose:fmt(get(row,'Propose')),
      category:fmt(get(row,'Category')),deltaCRC:fmt(get(row,'Delta CRC')),followup:rawFollowup||saved.followup||''
    });
  }
  // Duplicate = same endpoint + port pair, irrespective of direction.
  const dupGroups={};
  records.forEach(x=>{
    const a=(x.neId+'|'+x.nePort).toLowerCase().replace(/\s+/g,'');
    const b=(x.feId+'|'+x.fePort).toLowerCase().replace(/\s+/g,'');
    const key=[a,b].sort().join(' <> ');
    if(a!=='|'&&b!=='|'){(dupGroups[key]||(dupGroups[key]=[])).push(x);}
  });
  Object.keys(dupGroups).forEach(k=>{const group=dupGroups[k];group.forEach(x=>{x.duplicate=group.length>1;x.duplicateCount=group.length;});});
  records.forEach(x=>{if(x.duplicate===undefined){x.duplicate=false;x.duplicateCount=1;}});
  const summary=emptySummary_();summary.total=records.length;
  records.forEach(x=>{if(x.priority!==null)summary['p'+x.priority]++;summary[x.status==='In Progress'?'inProgress':x.status.toLowerCase()]++;if(x.duplicate)summary.duplicate++;});
  const trendMap={},kabMap={};
  records.forEach(x=>{
    const d=x.reqDate||'Tanggal tidak tersedia';
    if(!trendMap[d])trendMap[d]={date:d,total:0,open:0,close:0,inProgress:0};
    trendMap[d].total++;trendMap[d][x.status==='In Progress'?'inProgress':x.status.toLowerCase()]++;
    const kab=x.kab||'Tidak diketahui';kabMap[kab]=(kabMap[kab]||0)+1;
  });
  const trend=Object.keys(trendMap).sort().map(k=>trendMap[k]).slice(-14);
  const kabupaten=Object.keys(kabMap).map(k=>({name:k,total:kabMap[k]})).sort((a,b)=>b.total-a.total);
  return {records:records,summary:summary,trend:trend,kabupaten:kabupaten,notes:notes,updatedAt:new Date().toISOString(),sourceSheet:SOURCE_SHEET,spreadsheetUrl:ss.getUrl()};
}
function emptySummary_(){return {total:0,p0:0,p1:0,p2:0,p3:0,open:0,close:0,inProgress:0,duplicate:0};}
function loadNotes_(ss){
  const sh=ss.getSheetByName(NOTES_SHEET);if(!sh)return{};
  const v=sh.getDataRange().getValues(),out={};
  for(let i=1;i<v.length;i++)if(v[i][0])out[String(v[i][0])]={dateAction:formatSavedDate_(v[i][1]),dateClose:formatSavedDate_(v[i][2]),status:String(v[i][3]||''),followup:String(v[i][4]||'')};
  return out;
}
function formatSavedDate_(v){return v instanceof Date?Utilities.formatDate(v,Session.getScriptTimeZone(),'yyyy-MM-dd'):String(v||'');}
function recoveryColumns_(sh){
  const required=['Recovery Date Action','Recovery Date Close','Recovery Status','Recovery Follow-up / RCA','Recovery Updated At'];
  const lastCol=Math.max(sh.getLastColumn(),1);
  const headers=sh.getRange(1,1,1,lastCol).getDisplayValues()[0].map(v=>String(v||'').trim());
  const found={};
  headers.forEach((h,i)=>{if(h)found[h]=i+1;});
  const missing=required.filter(h=>!found[h]);
  if(missing.length){
    const startCol=lastCol+1;
    if(sh.getMaxColumns()<startCol+missing.length-1){
      sh.insertColumnsAfter(sh.getMaxColumns(),startCol+missing.length-1-sh.getMaxColumns());
    }
    sh.getRange(1,startCol,1,missing.length).setValues([missing]);
    missing.forEach((h,i)=>found[h]=startCol+i);
  }
  return required.map(h=>found[h]);
}

function saveRecovery(key,dateAction,dateClose,status,followup){
  if(!['Close','Open','In Progress'].includes(status))throw new Error('Status tidak valid.');
  const rowNumber=Number(key);
  if(!Number.isInteger(rowNumber)||rowNumber<2)throw new Error('Nomor baris RAW tidak valid.');
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh=ss.getSheetByName(SOURCE_SHEET);
  if(!sh)throw new Error('Tab sumber "'+SOURCE_SHEET+'" tidak ditemukan.');
  if(rowNumber>sh.getLastRow())throw new Error('Baris sumber tidak ditemukan. Muat ulang dashboard.');
  const cols=recoveryColumns_(sh);
  const parseDate=s=>s?new Date(s+'T00:00:00'):'';
  // Jika status diubah menjadi In Progress tanpa mengisi Date Action,
  // catat tanggal hari ini agar titik kurva S tetap terbentuk.
  const today = Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd');
  const effectiveDateAction = dateAction || (status==='In Progress' || status==='Close' ? today : '');
  const effectiveDateClose = dateClose || (status==='Close' ? today : '');
  const values=[parseDate(effectiveDateAction),parseDate(effectiveDateClose),status,followup||'',new Date()];
  cols.forEach((col,i)=>sh.getRange(rowNumber,col).setValue(values[i]));
  SpreadsheetApp.flush();

  // Read back the saved values so the client only refreshes after RAW is updated.
  const savedStatus = String(sh.getRange(rowNumber, cols[2]).getDisplayValue() || '').trim();
  if (savedStatus !== status) {
    throw new Error('Verifikasi penyimpanan gagal. Status pada RAW terbaca: ' + savedStatus);
  }
  return {ok:true,row:rowNumber,status:savedStatus,updatedColumns:cols};
}
