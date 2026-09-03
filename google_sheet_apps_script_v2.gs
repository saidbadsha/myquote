/*
QUOTATION & INVOICE MAKER — GOOGLE APPS SCRIPT v3
---------------------------------------------------
Google Sheets as the single data source for:
  - Company Details
  - Invoices
  - Quotations
  - Item Catalog (with saved rates)
  - Document Counters (auto-numbering)

Setup:
1) Create a Google Sheet.
2) Extensions > Apps Script open karein.
3) Is complete code ko paste karein.
4) Deploy > New deployment > Web app
   Execute as: Me
   Who has access: Anyone
5) /exec URL ko index.html ke Settings me paste karein.

Sheets automatically create hongi:
  - CompanyDetails
  - Invoices
  - Quotations
  - ItemCatalog
  - DocCounters
*/

const SS = SpreadsheetApp.getActiveSpreadsheet();

function setup_() {
  // CompanyDetails sheet
  let cd = SS.getSheetByName('CompanyDetails');
  if (!cd) {
    cd = SS.insertSheet('CompanyDetails');
    cd.appendRow(['CompanyName','Phone','Email','Address','GSTIN','Currency','UpdatedAt']);
  }

  // Invoices sheet
  let inv = SS.getSheetByName('Invoices');
  if (!inv) {
    inv = SS.insertSheet('Invoices');
    inv.appendRow([
      'DocNo','Date','Company','Phone','Email','Address','GSTIN','Client','ClientMobile',
      'Validity','Tax','DefaultProfit','Mode','SimpleMode','Notes','Total','ItemsJSON','UpdatedAt'
    ]);
  }

  // Quotations sheet
  let qt = SS.getSheetByName('Quotations');
  if (!qt) {
    qt = SS.insertSheet('Quotations');
    qt.appendRow([
      'DocNo','Date','Company','Phone','Email','Address','GSTIN','Client','ClientMobile',
      'Validity','Tax','DefaultProfit','Mode','SimpleMode','Notes','Total','ItemsJSON','UpdatedAt'
    ]);
  }

  // ItemCatalog sheet
  let ic = SS.getSheetByName('ItemCatalog');
  if (!ic) {
    ic = SS.insertSheet('ItemCatalog');
    ic.appendRow(['ItemName','CostRate','DefaultProfit','LastRate','UpdatedAt']);
  }

  // DocCounters sheet
  let dc = SS.getSheetByName('DocCounters');
  if (!dc) {
    dc = SS.insertSheet('DocCounters');
    dc.appendRow(['DocType','Year','LastNumber']);
  }
}

function doGet(e) {
  setup_();
  return json_({ok:true,message:'Quotation & Invoice Maker API v3 is running'});
}

function doPost(e) {
  try {
    setup_();
    const data = JSON.parse(e.postData.contents || '{}');
    switch(data.action) {
      case 'ping':
        return json_({ok:true,message:'Connected'});
      case 'getItems':
        return json_({ok:true,items:getItems_()});
      case 'saveDocument':
        return json_(saveDocument_(data.doc));
      case 'fetchDocuments':
        return json_(fetchDocuments_(data.docNo||'', data.name||'', data.mobile||'', data.docType||''));
      case 'getCompany':
        return json_(getCompany_());
      case 'saveCompany':
        return json_(saveCompany_(data.company));
      case 'getNextNumber':
        return json_(getNextNumber_(data.docType||'quotation'));
      default:
        return json_({ok:false,error:'Unknown action: '+data.action});
    }
  } catch(err) {
    return json_({ok:false,error:String(err)});
  }
}

// ─── Company Details ───

function getCompany_() {
  const sh = SS.getSheetByName('CompanyDetails');
  if (sh.getLastRow() < 2) return {ok:true,company:null};
  const r = sh.getRange(2,1,1,7).getValues()[0];
  return {ok:true,company:{
    name:String(r[0]||''), phone:String(r[1]||''), email:String(r[2]||''),
    address:String(r[3]||''), gst:String(r[4]||''), currency:String(r[5]||'₹')
  }};
}

function saveCompany_(c) {
  const sh = SS.getSheetByName('CompanyDetails');
  const row = [c.name||'',c.phone||'',c.email||'',c.address||'',c.gst||'',c.currency||'₹',new Date()];
  if (sh.getLastRow() < 2) sh.appendRow(row);
  else sh.getRange(2,1,1,7).setValues([row]);
  return {ok:true,message:'Company details saved'};
}

// ─── Document Save (Invoice or Quotation) ───

function saveDocument_(d) {
  if (!d) return {ok:false,error:'No document data'};
  const type = (d.docType || 'quotation').toLowerCase();
  const sheetName = type === 'invoice' ? 'Invoices' : 'Quotations';
  const sh = SS.getSheetByName(sheetName);
  const values = sh.getDataRange().getValues();

  const items = d.items || [];
  const subtotal = items.reduce((s,x) => s + Number(x.amount || 0), 0);
  const total = subtotal * (1 + Number(d.tax || 0) / 100);

  const row = [
    d.docNo, d.date, d.company, d.phone, d.email, d.address, d.gst,
    d.client, d.clientMobile, d.validity, d.tax, d.profit, d.mode,
    d.simpleMode, d.notes, total, JSON.stringify(items), new Date()
  ];

  // Find existing record by DocNo
  let found = -1;
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][0]).trim() === String(d.docNo).trim()) {
      found = r + 1;
      break;
    }
  }
  if (found > 0) sh.getRange(found, 1, 1, row.length).setValues([row]);
  else sh.appendRow(row);

  // Update ItemCatalog with rates
  updateItemCatalog_(items);

  return {ok:true, message:(type === 'invoice' ? 'Invoice' : 'Quotation')+' saved', total:total};
}

// ─── Item Catalog ───

function updateItemCatalog_(items) {
  const sh = SS.getSheetByName('ItemCatalog');
  items.forEach(x => {
    const name = String(x.name||'').trim();
    if (!name) return;
    let rowNo = -1;
    for (let r = 1; r < sh.getLastRow(); r++) {
      if (String(sh.getRange(r+1,1).getValue()).trim().toLowerCase() === name.toLowerCase()) {
        rowNo = r + 1;
        break;
      }
    }
    // Save rate as LastRate (column 4) along with cost and profit
    const itemRow = [name, Number(x.c||0), Number(x.p||0), Number(x.rate||0), new Date()];
    if (rowNo > 0) sh.getRange(rowNo, 1, 1, 5).setValues([itemRow]);
    else sh.appendRow(itemRow);
  });
}

function getItems_() {
  const sh = SS.getSheetByName('ItemCatalog');
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow()-1, 5).getValues()
    .filter(r => String(r[0]).trim())
    .map(r => ({
      name: String(r[0]),
      cost: Number(r[1]||0),
      defaultProfit: Number(r[2]||0),
      lastRate: Number(r[3]||0)
    }));
}

// ─── Auto-Numbering ───

function getNextNumber_(docType) {
  const type = (docType || 'quotation').toLowerCase();
  const year = new Date().getFullYear();
  const prefix = type === 'invoice' ? 'INV' : 'QT';
  const sh = SS.getSheetByName('DocCounters');
  const values = sh.getDataRange().getValues();

  let lastNum = 0;
  let counterRow = -1;

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][0]).toLowerCase() === type && Number(values[r][1]) === year) {
      lastNum = Number(values[r][2] || 0);
      counterRow = r + 1;
      break;
    }
  }

  const nextNum = lastNum + 1;
  const docNo = prefix + '-' + year + '-' + String(nextNum).padStart(3, '0');

  if (counterRow > 0) {
    sh.getRange(counterRow, 3).setValue(nextNum);
  } else {
    sh.appendRow([type, year, nextNum]);
  }

  return {ok:true, docNo:docNo, nextNumber:nextNum};
}

// ─── Fetch Documents ───

function fetchDocuments_(docNo, name, mobile, docType) {
  const results = [];

  // Search in Invoices (if no type filter, or type is 'invoice')
  if (!docType || docType === 'invoice') {
    results.push(...searchSheet_('Invoices', 'invoice', docNo, name, mobile));
  }
  // Search in Quotations (if no type filter, or type is 'quotation')
  if (!docType || docType === 'quotation') {
    results.push(...searchSheet_('Quotations', 'quotation', docNo, name, mobile));
  }

  return {ok:true, rows:results};
}

function searchSheet_(sheetName, docType, docNo, name, mobile) {
  const sh = SS.getSheetByName(sheetName);
  if (sh.getLastRow() < 2) return [];

  const rows = sh.getRange(2, 1, sh.getLastRow()-1, 18).getValues();
  const out = [];

  rows.forEach(r => {
    const dNo = String(r[0]||'').trim();
    const mob = String(r[8]||'').replace(/\D/g,'');

    const noMatch = docNo && dNo.toLowerCase().includes(String(docNo).toLowerCase());
    const nameMatch = name && String(r[7]||'').toLowerCase().includes(String(name).toLowerCase());
    const mobMatch = mobile && mob.includes(String(mobile).replace(/\D/g,''));

    if ((docNo && noMatch) || (name && nameMatch) || (mobile && mobMatch)) {
      let items = [];
      try { items = JSON.parse(r[16]||'[]'); } catch(e) {}
      out.push({
        docType: docType,
        docNo: r[0], date: r[1], company: r[2], phone: r[3],
        email: r[4], address: r[5], gst: r[6],
        client: r[7], clientMobile: r[8], validity: r[9],
        tax: r[10], profit: r[11], mode: r[12], simpleMode: r[13],
        notes: r[14], total: r[15], items: items
      });
    }
  });

  return out;
}

// ─── Utility ───

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
