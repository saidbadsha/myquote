/*
QUOTATION MAKER — GOOGLE APPS SCRIPT
-------------------------------------
1) Google Sheet banayein.
2) Extensions > Apps Script open karein.
3) Is complete code ko paste karein.
4) Deploy > New deployment > Web app
   Execute as: Me
   Who has access: Anyone
5) /exec URL ko quotation_maker_pro.html ke Settings > Google Apps Script Web App URL me paste karein.

Sheet automatically "Quotations" and "ItemCatalog" sheets create karegi.
*/

const SS = SpreadsheetApp.getActiveSpreadsheet();
const QUOTE_SHEET = 'Quotations';
const ITEM_SHEET = 'ItemCatalog';

function setup_() {
  const q = SS.getSheetByName(QUOTE_SHEET) || SS.insertSheet(QUOTE_SHEET);
  const i = SS.getSheetByName(ITEM_SHEET) || SS.insertSheet(ITEM_SHEET);

  const qh = ['QuoteNo','Date','Company','Phone','Email','Address','GSTIN','Client','ClientMobile','Validity','Tax','DefaultProfit','Mode','SimpleMode','Notes','Total','ItemsJSON','UpdatedAt'];
  if (q.getLastRow() === 0) q.appendRow(qh);
  const ih = ['ItemName','CostRate','DefaultProfit','UpdatedAt'];
  if (i.getLastRow() === 0) i.appendRow(ih);
}

function doGet(e) {
  setup_();
  return json_({ok:true,message:'Quotation Maker Google Sheet API is running'});
}

function doPost(e) {
  try {
    setup_();
    const data = JSON.parse(e.postData.contents || '{}');
    if (data.action === 'ping') return json_({ok:true,message:'Connected'});
    if (data.action === 'getItems') return json_({ok:true,items:getItems_()});
    if (data.action === 'saveQuotation') return json_(saveQuotation_(data.quotation));
    if (data.action === 'fetchQuotations') return json_(fetchQuotations_(data.quoteNo || '', data.name || '', data.mobile || ''));
    return json_({ok:false,error:'Unknown action'});
  } catch(err) {
    return json_({ok:false,error:String(err)});
  }
}

function saveQuotation_(q) {
  const sh = SS.getSheetByName(QUOTE_SHEET);
  const values = sh.getDataRange().getValues();
  const items = q.items || [];
  const total = items.reduce((s,x)=>s + Number(x.amount || 0),0) * (1 + Number(q.tax||0)/100);
  const itemsJSON = JSON.stringify(items);
  const row = [
    q.quoteNo,q.date,q.company,q.phone,q.email,q.address,q.gst,q.client,q.clientMobile,
    q.validity,q.tax,q.profit,q.mode,q.simpleMode,q.notes,total,itemsJSON,new Date()
  ];

  let found = -1;
  for (let r=1;r<values.length;r++) {
    if (String(values[r][0]).trim() === String(q.quoteNo).trim()) { found=r+1; break; }
  }
  if (found > 0) sh.getRange(found,1,1,row.length).setValues([row]);
  else sh.appendRow(row);

  const itemSh = SS.getSheetByName(ITEM_SHEET);
  const catalog = getItems_();
  items.forEach(x=>{
    const name=String(x.name||'').trim();
    if (!name) return;
    let rowNo=-1;
    for(let r=1;r<itemSh.getLastRow();r++){
      if(String(itemSh.getRange(r+1,1).getValue()).trim().toLowerCase()===name.toLowerCase()){rowNo=r+1;break;}
    }
    const itemRow=[name,Number(x.c||0),Number(x.p||0),new Date()];
    if(rowNo>0)itemSh.getRange(rowNo,1,1,4).setValues([itemRow]);
    else itemSh.appendRow(itemRow);
  });
  return {ok:true,message:'Quotation saved',total:total};
}

function getItems_() {
  const sh=SS.getSheetByName(ITEM_SHEET);
  if(sh.getLastRow()<2)return [];
  return sh.getRange(2,1,sh.getLastRow()-1,4).getValues()
    .filter(r=>String(r[0]).trim())
    .map(r=>({name:String(r[0]),cost:Number(r[1]||0),defaultProfit:Number(r[2]||0)}));
}

function fetchQuotations_(no,name,mobile) {
  const sh=SS.getSheetByName(QUOTE_SHEET);
  if(sh.getLastRow()<2)return {ok:true,rows:[]};
  const rows=sh.getRange(2,1,sh.getLastRow()-1,18).getValues();
  const out=[];
  rows.forEach(r=>{
    const qno=String(r[0]||'').trim(), mob=String(r[8]||'').replace(/\D/g,'');
    const noMatch=no && qno.toLowerCase().includes(String(no).toLowerCase());
    const nameMatch=name && String(r[7]||'').toLowerCase().includes(String(name).toLowerCase());
    const mobMatch=mobile && mob.includes(String(mobile).replace(/\D/g,''));
    if((no && noMatch)||(name && nameMatch)||(mobile && mobMatch)) {
      let items=[];try{items=JSON.parse(r[15]||'[]')}catch(e){}
      out.push({quoteNo:r[0],date:r[1],company:r[2],phone:r[3],email:r[4],address:r[5],gst:r[6],
        client:r[7],clientMobile:r[8],validity:r[9],tax:r[10],profit:r[11],mode:r[12],simpleMode:r[13],notes:r[14],
        total:r[15],items:items});
    }
  });
  return {ok:true,rows:out};
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
