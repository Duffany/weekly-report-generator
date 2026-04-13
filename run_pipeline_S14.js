/**
 * Node.js runner for S14 — reads from "S13-S14 RUN/"
 * RepSKU source: Weekly Report Stock S13 + FULL FFM 3.xlsx
 * Output       : S13-S14 RUN/pipeline_output_S14.xlsx
 *
 * Same column adjustments as S13 script (new repSku layout, Jumia "prix", marge blank).
 */

const XLSX    = require('xlsx');
const ExcelJS = require('exceljs');
const fs      = require('fs');
const path    = require('path');

const DIR = path.join(__dirname, 'S13-S14 RUN');

const CATEGORIE_MAP = {
  'ADULTE - EROTIQUE':                          'Autres',
  'AMENAGEMENT URBAIN - VOIRIE':                'Autres',
  'ANIMALERIE':                                 'Bricolage Jardin Animalerie',
  'APICULTURE':                                 'Autres',
  'ARME DE COMBAT - ARME DE SPORT':             'Autres',
  'ART DE LA TABLE - ARTICLES CULINAIRES':      'Maison',
  'ARTICLES POUR FUMEUR':                       'Autres',
  'AUTO - MOTO':                                'Autres',
  'BAGAGERIE':                                  'Mode',
  'BATEAU MOTEUR - VOILIER':                    'Autres',
  'BIJOUX -  LUNETTES - MONTRES':               'Mode',
  'BRICOLAGE - OUTILLAGE - QUINCAILLERIE':      'Bricolage Jardin Animalerie',
  'CHAUSSURES - ACCESSOIRES':                   'Mode',
  'COFFRET CADEAU BOX':                         'Autres',
  'CONDITIONNEMENT':                            'Autres',
  'DECO - LINGE - LUMINAIRE':                   'Maison',
  'DROGUERIE':                                  'Bricolage Jardin Animalerie',
  'DVD - BLU-RAY':                              'TV Son',
  'ELECTROMENAGER':                             'PEM',
  'ELECTRONIQUE':                               'Autres',
  'EPICERIE SALEE':                             'Autres',
  'EPICERIE SUCREE':                            'Autres',
  'FUNERAIRE':                                  'Autres',
  'HYGIENE - BEAUTE - PARFUM':                  'Beaute',
  'INFORMATIQUE':                               'Informatique & gaming',
  'INSTRUMENTS DE MUSIQUE':                     'TV Son',
  'JARDIN - PISCINE':                           'Bricolage Jardin Animalerie',
  'JEUX - JOUETS':                              'Bebe - Jouet',
  'JEUX VIDEO':                                 'Informatique & gaming',
  'LIBRAIRIE':                                  'Autres',
  'LITERIE':                                    'Maison',
  'LOGISTIQUE':                                 'Autres',
  'LOISIRS CREATIFS - BEAUX ARTS - PAPETERIE':  'Autres',
  'MANUTENTION':                                'Autres',
  'MATERIEL DE BUREAU':                         'Autres',
  'MATERIEL MEDICAL':                           'Beaute',
  'MERCERIE':                                   'Autres',
  'MEUBLE':                                     'Maison',
  'MUSIQUE':                                    'TV Son',
  'OFFRES PARTENAIRES':                         'Autres',
  'PARAPHARMACIE':                              'Beaute',
  'PHOTO - OPTIQUE':                            'TV Son',
  'POINT DE VENTE - COMMERCE - ADMINISTRATION': 'Autres',
  'PRODUITS FRAIS':                             'Autres',
  'PRODUITS SURGELES':                          'Autres',
  'PUERICULTURE':                               'Bebe - Jouet',
  'SONO - DJ':                                  'TV Son',
  'SPORT':                                      'Sport',
  'TATOUAGE - PIERCING':                        'Autres',
  'TELEPHONIE - GPS':                           'Tel',
  'TENUE PROFESSIONNELLE':                      'Autres',
  'TV - VIDEO - SON':                           'TV Son',
  'VETEMENTS - LINGERIE':                       'Mode',
  'VIN - ALCOOL - LIQUIDE':                     'Autres',
};

const CATEGORIES_ORDER = [
  'Maison', 'Beaute', 'PEM', 'Bricolage Jardin Animalerie', 'Mode',
  'Sport', 'Autres', 'Tel', 'TV Son', 'Bebe - Jouet', 'Informatique & gaming'
];

const VG_SECTIONS = [
  { name: 'Scope total Retail',                      extra: '' },
  { name: 'Produits Non BO',                         extra: `,'Report Retail'!$AJ:$AJ,"Retail Non BO"` },
  { name: 'Produits KO Prix Jumia',                  extra: `,'Report Retail'!$AG:$AG,"prix \u00e0 revoir"` },
  { name: 'Age stock >180 jours',                    extra: `,'Report Retail'!$N:$N,">"&180` },
  { name: 'Couverture >120 jours et Age > 60 jours', extra: `,'Report Retail'!$R:$R,">"&120,'Report Retail'!$N:$N,">"&60` },
  { name: 'Low views (PV L3M <200)',                 extra: `,'Report Retail'!$T:$T,"<"&200` },
];

function parseNum(v) {
  if (v === '' || v == null) return 0;
  const n = parseFloat(String(v).replace(',', '.').replace(/\s/g, ''));
  return isNaN(n) ? 0 : n;
}
function roundN(n, d) { const f = Math.pow(10, d); return Math.round(n * f) / f; }

function normalizeGtin(raw) {
  if (raw === null || raw === undefined || raw === '') return '';
  if (typeof raw === 'number') return String(Math.round(raw));
  const s = String(raw).trim();
  if (/^[0-9.]+[eE][+\-]?[0-9]+$/.test(s)) return String(Math.round(parseFloat(s)));
  return s.replace(/\.0+$/, '').replace(/^0+(\d)/, '$1');
}

function normalizeType(t) {
  const lo = String(t).toLowerCase().trim();
  if (lo === '1p local b1') return '1P Local B1';
  if (lo === '1p local b2') return '1P LOCAL B2';
  if (lo === '1p chine')    return '1P Chine';
  return t;
}

function colIdx(headers, ...candidates) {
  const h = headers.map(x => String(x).trim().toLowerCase());
  for (const c of candidates) {
    const exact = h.indexOf(c.toLowerCase());
    if (exact !== -1) return exact;
  }
  for (const c of candidates) {
    const partial = h.findIndex(x => x.includes(c.toLowerCase()));
    if (partial !== -1) return partial;
  }
  return -1;
}

function mapCategory(n1Raw, rawCat) {
  for (const src of [n1Raw, rawCat]) {
    if (!src) continue;
    const up = String(src).toUpperCase().trim();
    if (CATEGORIE_MAP[up]) return CATEGORIE_MAP[up];
    for (const [key, val] of Object.entries(CATEGORIE_MAP)) {
      if (up.includes(key)) return val;
    }
  }
  return 'Autres';
}

function readWb(filePath) {
  return XLSX.read(fs.readFileSync(filePath), { type: 'buffer', cellDates: false });
}

function sheetToArrays(wb, sheetName) {
  const name = sheetName || wb.SheetNames[0];
  const ws   = wb.Sheets[name];
  if (!ws) throw new Error(`Sheet "${name}" not found`);
  const all = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  return { headers: all[0] || [], rows: all.slice(1) };
}

const S = {
  hFill:  { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF000000' } },
  hFont:  { bold: true, color: { argb: 'FFFFFFFF' }, size: 11, name: 'Calibri' },
  hAlign: { horizontal: 'center', vertical: 'middle', wrapText: true },
  hBorderRow1: { top:{style:'medium'}, bottom:{style:'thin'},   left:{style:'medium'}, right:{style:'medium'} },
  hBorderRow2: { top:{style:'thin'},   bottom:{style:'thin'},   left:{style:'medium'}, right:{style:'medium'} },
  hBorderRow3: { top:{style:'thin'},   bottom:{style:'medium'}, left:{style:'medium'}, right:{style:'medium'} },
  dAlign: { horizontal: 'center', vertical: 'middle' },
  dBorder:{ top:{style:'thin'}, bottom:{style:'thin'}, left:{style:'thin'}, right:{style:'thin'} },
  yFill:  { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } },
};
const YELLOW_COLS = new Set([8, 9, 10, 11, 19, 23, 24, 25, 26, 28, 31, 33, 34]);

function styleCell(cell, value, fill, font, alignment, border) {
  cell.value = value; cell.fill = fill; cell.font = font;
  cell.alignment = alignment; cell.border = border;
}
function colLetterToNum(col) {
  let n = 0;
  for (let i = 0; i < col.length; i++) n = n * 26 + col.charCodeAt(i) - 64;
  return n;
}
function numToColLetter(n) {
  let s = '';
  while (n > 0) { s = String.fromCharCode(((n - 1) % 26) + 65) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function buildVueGlobale(ws, totalRetailRows) {
  ws.getColumn(1).width = 32;
  for (let c = 2; c <= 37; c++) ws.getColumn(c).width = 10;
  ws.getRow(1).height = 36; ws.getRow(2).height = 20; ws.getRow(3).height = 18;

  ws.mergeCells('A1:A3');
  const sectionStarts = ['B','H','N','T','Z','AF'];
  sectionStarts.forEach((col, i) => {
    const endColFull = i < 4 ? String.fromCharCode(col.charCodeAt(0) + 5) : (i === 4 ? 'AE' : 'AK');
    ws.mergeCells(`${col}1:${endColFull}1`);
    const c0 = colLetterToNum(col);
    ws.mergeCells(`${numToColLetter(c0)}2:${numToColLetter(c0+1)}2`);
    ws.mergeCells(`${numToColLetter(c0+2)}2:${numToColLetter(c0+3)}2`);
    ws.mergeCells(`${numToColLetter(c0+4)}2:${numToColLetter(c0+5)}2`);
  });

  styleCell(ws.getCell('A1'), 'Catégorie', S.hFill, S.hFont, S.hAlign, S.hBorderRow1);
  VG_SECTIONS.forEach((sec, i) => styleCell(ws.getCell(`${sectionStarts[i]}1`), sec.name, S.hFill, S.hFont, S.hAlign, S.hBorderRow1));
  sectionStarts.forEach(col => {
    const c0 = colLetterToNum(col);
    ['B1','B2','Total'].forEach((lbl, j) => styleCell(ws.getCell(`${numToColLetter(c0+j*2)}2`), lbl, S.hFill, S.hFont, S.hAlign, S.hBorderRow2));
  });
  styleCell(ws.getCell('A3'), 'Catégorie', S.hFill, S.hFont, S.hAlign, S.hBorderRow3);
  sectionStarts.forEach(col => {
    const c0 = colLetterToNum(col);
    for (let j = 0; j < 6; j++) styleCell(ws.getCell(`${numToColLetter(c0+j)}3`), j%2===0?'# SKU':'Stock', S.hFill, S.hFont, S.hAlign, S.hBorderRow3);
  });

  CATEGORIES_ORDER.forEach((cat, idx) => {
    const r = idx + 4;
    const row = ws.getRow(r); row.height = 15;
    const catCell = row.getCell(1);
    catCell.value = cat; catCell.alignment = S.dAlign; catCell.border = S.dBorder; catCell.font = {size:11,name:'Calibri'};
    sectionStarts.forEach((col, si) => {
      const sec = VG_SECTIONS[si]; const c0 = colLetterToNum(col);
      const skuF  = `=COUNTIFS('Report Retail'!$C:$C,"Retail",'Report Retail'!$A:$A,"${cat}",'Report Retail'!$D:$D,"1P Local B1"${sec.extra})`;
      const stkF  = `=SUMIFS('Report Retail'!$M:$M,'Report Retail'!$C:$C,"Retail",'Report Retail'!$A:$A,"${cat}",'Report Retail'!$D:$D,"1P Local B1"${sec.extra})`;
      const skuF2 = `=COUNTIFS('Report Retail'!$C:$C,"Retail",'Report Retail'!$A:$A,"${cat}",'Report Retail'!$D:$D,"1P LOCAL B2"${sec.extra})`;
      const stkF2 = `=SUMIFS('Report Retail'!$M:$M,'Report Retail'!$C:$C,"Retail",'Report Retail'!$A:$A,"${cat}",'Report Retail'!$D:$D,"1P LOCAL B2"${sec.extra})`;
      const skuFT = `=COUNTIFS('Report Retail'!$C:$C,"Retail",'Report Retail'!$A:$A,"${cat}"${sec.extra})`;
      const stkFT = `=SUMIFS('Report Retail'!$M:$M,'Report Retail'!$C:$C,"Retail",'Report Retail'!$A:$A,"${cat}"${sec.extra})`;
      [[skuF,stkF],[skuF2,stkF2],[skuFT,stkFT]].forEach(([sf,tf], j) => {
        const cSku = row.getCell(c0+j*2); const cStk = row.getCell(c0+j*2+1);
        cSku.value={formula:sf}; cStk.value={formula:tf};
        [cSku,cStk].forEach(c=>{c.alignment=S.dAlign;c.border=S.dBorder;c.font={size:11,name:'Calibri'};c.numFmt='#,##0';});
      });
    });
  });

  const totalR = CATEGORIES_ORDER.length + 4;
  const tRow = ws.getRow(totalR); tRow.height = 15;
  const tCell = tRow.getCell(1);
  tCell.value='Total général'; tCell.alignment=S.dAlign; tCell.border=S.dBorder; tCell.font={bold:true,size:11,name:'Calibri'};
  sectionStarts.forEach(col => {
    const c0 = colLetterToNum(col);
    for (let j=0;j<6;j++) {
      const colL = numToColLetter(c0+j);
      const c = tRow.getCell(c0+j);
      c.value={formula:`=SUM(${colL}4:${colL}${3+CATEGORIES_ORDER.length})`};
      c.alignment=S.dAlign; c.border=S.dBorder; c.font={bold:true,size:11,name:'Calibri'}; c.numFmt='#,##0';
    }
  });
}

const RR_HEADERS = [
  'Catégorie Revue','Categorie','Type de Stock','Type','Owner',
  'ProductId','gtin','title','Marque','N1','N2','N3',
  'Stock','Age stock Moyen','Tranche Age','VALEUR PV HT',
  'Moyenne de vente','Couverture (Jrs)','Tranche de couverture',
  'PV L3M','Tranche de PV','CR L3M','Tranche de CR',
  'IS L3M','GMV L3M','Marge L3M','Cout unitaire',
  'ASP L3M','Prix Live','Marge live','Statut marge',
  'Prix Jumia','Check Prix','Lien Jumia','Vendeur BO','Check BO','Animation commerciale'
];

function buildReportRetail(ws, retailRows) {
  const widths = [22,22,12,14,10,16,16,40,18,22,22,22,10,14,12,14,14,14,18,10,12,10,12,10,12,12,12,10,12,12,12,12,18,40,20,16,40];
  widths.forEach((w,i)=>{ws.getColumn(i+1).width=w;});
  const hRow = ws.getRow(1); hRow.height=20;
  RR_HEADERS.forEach((h,i)=>{
    const c=hRow.getCell(i+1); c.value=h; c.font={bold:true,size:11,name:'Calibri'};
    c.alignment={horizontal:'center',vertical:'middle',wrapText:true}; c.border=S.dBorder;
    if(YELLOW_COLS.has(i)) c.fill=S.yFill;
  });

  retailRows.forEach((row,idx)=>{
    const r=idx+2; const eRow=ws.getRow(r); eRow.height=15;
    const vals=[
      row.catRevue,row.rawCat,row.typeV,row.srcType,row.owner,
      row.pid,row.gtin,row.title,row.marque,row.n1,row.n2,row.n3,
      row.stock,row.age,null,row.valeur,null,null,null,row.pv,null,null,null,
      row.is,row.gmv,row.margePct,row.coutU,null,row.prixLive,null,null,row.prixJ,null,row.lienJ,row.vendeurBO,null,null,
    ];
    vals.forEach((v,i)=>{
      if(v!==null){const c=eRow.getCell(i+1);c.value=v;if(typeof v==='number'&&i>=12)c.numFmt='#,##0.##';}
    });
    eRow.getCell(15).value={formula:`IF(N${r}<=60,"0-60jrs",IF(N${r}<=120,"60-120jrs",IF(N${r}<=180,"120-180jrs",">180jrs")))`};
    eRow.getCell(17).value={formula:`IFERROR(X${r}/90,0)`};
    eRow.getCell(18).value={formula:`IFERROR(ROUND(M${r}/Q${r},0),"")`};
    eRow.getCell(19).value={formula:`IF(R${r}="","",IF(R${r}<=60,"0-60jrs",IF(R${r}<=120,"60-120jrs",IF(R${r}<=180,"120-180jrs",">180jrs"))))`};
    eRow.getCell(21).value={formula:`IF(T${r}<50,"<50",IF(T${r}<200,"50-200",IF(T${r}<500,"200-500",IF(T${r}<1000,"500-1000",">=1000"))))`};
    eRow.getCell(22).value={formula:`IFERROR(X${r}/T${r},"")`};
    eRow.getCell(23).value={formula:`IFERROR(IF(V${r}<0.01,"<1%",IF(V${r}<0.03,"1-3%",IF(V${r}<0.1,"3-10%",">10%"))),"")`};
    eRow.getCell(28).value={formula:`IFERROR(Y${r}/X${r},"")`};
    eRow.getCell(30).value={formula:`IFERROR((AC${r}/1.2-AA${r})/(AC${r}/1.2),"")`};
    eRow.getCell(33).value={formula:`IFERROR(IF(AF${r}="","",IF(AF${r}<AC${r},"prix \u00e0 revoir",IF(AC${r}>AB${r},"Prix Supérieur à ASP",""))),"")` };
    eRow.getCell(36).value={formula:`IF(AND(C${r}="Retail",UPPER(AI${r})<>"MARJANEMALL"),"Retail Non BO","")`};
    eRow.getCell(37).value={formula:`IF(AND(N${r}>180,R${r}>120),"Baisse de prix permanente ou retour fournisseur",IF(AND(N${r}>60,R${r}>120),"Baisse de prix permanente",IF(AND(R${r}>60,V${r}>0.05),"Vente Flash","")))`};
  });
  ws.autoFilter={from:'A1',to:'AK1'};
}

async function main() {
  console.log('Reading files...');

  // Base Retail type affectation (optional fallback for types not in repSku)
  const baseRetailMap = new Map();
  try {
    const brWb = readWb(path.join(DIR, 'base retail.xlsx'));
    const brSheet = brWb.SheetNames.find(n => n.trim().toLowerCase() === 'base retail');
    if (brSheet) {
      const { headers: brH, rows: brRows } = sheetToArrays(brWb, brSheet);
      const brGtin = colIdx(brH, 'GTIN_octopia', 'gtin', 'EAN');
      const brPid  = colIdx(brH, 'productid', 'product_id', 'SKU');
      const brType = colIdx(brH, 'Type', 'type');
      for (const row of brRows) {
        const type = normalizeType(String(row[brType] || '').trim());
        if (!type) continue;
        const gtin = normalizeGtin(row[brGtin]);
        const pid  = String(row[brPid] || '').trim();
        if (gtin && !baseRetailMap.has(gtin)) baseRetailMap.set(gtin, type);
        if (pid  && !baseRetailMap.has(pid))  baseRetailMap.set(pid,  type);
      }
      console.log(`Base retail map: ${baseRetailMap.size} entries`);
    }
  } catch(e) { console.warn('Base retail load failed:', e.message); }

  // RepSKU from S13 weekly report (same new layout: col0=SKU, col1=GTIN, col2=Type, col5=Owner)
  const repSkuMap = new Map();
  try {
    const tplWb = readWb(path.join(DIR, 'Weekly Report Stock S13 + FULL FFM 3.xlsx'));
    const { headers: repH, rows: rskuRows } = sheetToArrays(tplWb, 'Repartition sku 1P');
    const isNewLayout = repH.some(h => String(h).toLowerCase().includes('gtin_octopia') || String(h).toLowerCase().includes('gtin octopia'));
    const gtinCol  = isNewLayout ? 1 : 0;
    const typeCol  = isNewLayout ? 2 : 1;
    const ownerCol = isNewLayout ? 5 : 4;
    const pidCol   = isNewLayout ? 0 : -1;
    for (const row of rskuRows) {
      const type = normalizeType(String(row[typeCol] || '').trim());
      const ownerCode = String(row[ownerCol] || '').trim();
      if (!type) continue;
      const gtin = normalizeGtin(row[gtinCol]);
      if (gtin) repSkuMap.set(gtin, { type, owner: ownerCode });
      if (pidCol >= 0) {
        const pid = String(row[pidCol] || '').trim();
        if (pid && !repSkuMap.has(pid)) repSkuMap.set(pid, { type, owner: ownerCode });
      }
    }
    console.log(`RepSKU map: ${repSkuMap.size} entries from Repartition sku 1P (layout: ${isNewLayout ? 'new' : 'old'})`);
    // Fallback: also read Report Retail from S-1 for products not in repSku
    try {
      const { rows: rrRows } = sheetToArrays(tplWb, 'Report Retail');
      let added = 0;
      for (const row of rrRows) {
        const type = normalizeType(String(row[3] || '').trim());
        if (!type) continue;
        const owner = String(row[4] || '').trim();
        const pid   = String(row[5] || '').trim();
        const gtin  = normalizeGtin(row[6]);
        if (gtin && !repSkuMap.has(gtin)) { repSkuMap.set(gtin, { type, owner }); added++; }
        if (pid  && !repSkuMap.has(pid))  { repSkuMap.set(pid,  { type, owner }); added++; }
      }
      console.log(`+ ${added} additional entries from Report Retail S-1. Total: ${repSkuMap.size}`);
    } catch(e2) { console.warn('Report Retail fallback failed:', e2.message); }
  } catch(e) { console.warn('RepSKU load failed:', e.message); }

  // Stock
  const stockWb = readWb(path.join(DIR, 'stock 1.xlsx'));
  const { headers: sH, rows: stockRows } = sheetToArrays(stockWb);
  const SC = {
    cat:   colIdx(sH,'nom_categorie','categorie'),
    pid:   colIdx(sH,'ProductId','product_id','SKU'),
    gtin:  colIdx(sH,'gtin','GTIN','ean'),
    typeV: colIdx(sH,'type vendeur','type_vendeur'),
    stock: colIdx(sH,'stock_dispo','stock'),
    title: colIdx(sH,'title','titre'),
    age:   colIdx(sH,'age_stock','age'),
    val:   colIdx(sH,'Valeur','valeur'),
  };
  console.log(`Stock: ${stockRows.length} rows. Cols:`, SC);

  // Conso
  const consoWb = readWb(path.join(DIR, 'conso 1304.xlsx'));
  const { headers: cH, rows: consoRows } = sheetToArrays(consoWb);
  const CC = {
    n1:    colIdx(cH,'Label_N1','N1','label_n1'),
    n2:    colIdx(cH,'Label_N2','N2'),
    n3:    colIdx(cH,'LABEL_N3','Label_N3','N3'),
    pid:   colIdx(cH,'ProductId','product_id','Id'),
    shop:  colIdx(cH,'shopname','Shopname','shop_name','vendeur'),
    rank:  colIdx(cH,'BestOfferRank','bestofferrank','rank'),
    price: colIdx(cH,'Price','OfferPrice','offerprice','prix'),
    brand: colIdx(cH,'brandlabel','Brandlabel','brand','marque'),
  };
  console.log(`Conso: ${consoRows.length} rows. Cols:`, CC);

  const consoMap = new Map();
  for (const row of consoRows) {
    const k = String(row[CC.pid]||'').trim();
    if (!k) continue;
    const rank = parseNum(row[CC.rank]) || 9999;
    if (!consoMap.has(k) || rank < consoMap.get(k)._rank) {
      const entry = row.slice(); entry._rank = rank;
      consoMap.set(k, entry);
    }
  }
  console.log(`Conso map: ${consoMap.size} entries`);

  // L3M
  const l3mWb = readWb(path.join(DIR, 'Report L3M 1404.xlsx'));
  const { headers: lH, rows: l3mRows } = sheetToArrays(l3mWb);
  const LC = {
    pid:   colIdx(lH,'SKU','ProductId','sku'),
    pv:    colIdx(lH,'PV','Views','page_views'),
    is_:   colIdx(lH,'IS','is_l3m'),
    gmv:   colIdx(lH,'GMV','gmv'),
    marge: -1, // left blank per manager instruction
    coutU: colIdx(lH,'cout unitaire','Cout unitaire','cout_unitaire'),
  };
  console.log(`L3M: ${l3mRows.length} rows. Cols:`, LC);

  const l3mMap = new Map();
  for (const row of l3mRows) {
    const k = String(row[LC.pid]||'').trim();
    if (!k) continue;
    const pv=parseNum(row[LC.pv]), is_=parseNum(row[LC.is_]), gmv=parseNum(row[LC.gmv]);
    const marge=0, coutU=LC.coutU>=0?parseNum(row[LC.coutU]):0;
    if (l3mMap.has(k)) { const e=l3mMap.get(k); e.pv+=pv; e.is+=is_; e.gmv+=gmv; }
    else l3mMap.set(k, {pv, is:is_, gmv, marge, coutU});
  }
  console.log(`L3M map: ${l3mMap.size} products`);

  // Jumia
  const jumiaWb = readWb(path.join(DIR, 'IP.xlsx'));
  const { headers: jH, rows: jumiaRows } = sheetToArrays(jumiaWb);
  const JC = {
    pid:  colIdx(jH,'sku','SKU','ProductId'),
    px:   colIdx(jH,'prix_jumia','Prix_Jumia','prix','Prix'),
    lien: colIdx(jH,'Lien_du_produit','lien','link'),
    ean:  colIdx(jH,'mon_ean','ean','EAN'),
  };
  console.log(`Jumia: ${jumiaRows.length} rows. Cols:`, JC);

  const jumiaMap=new Map(), jumiaEanMap=new Map();
  for (const row of jumiaRows) {
    const k=String(row[JC.pid]||'').trim();
    if(k&&!jumiaMap.has(k)) jumiaMap.set(k,row);
    const eanRaw=row[JC.ean];
    if(eanRaw!==null&&eanRaw!==undefined&&eanRaw!==''){
      const eanKey=typeof eanRaw==='number'?String(Math.round(eanRaw)):String(eanRaw).trim();
      if(eanKey&&!jumiaEanMap.has(eanKey)) jumiaEanMap.set(eanKey,row);
    }
  }
  console.log(`Jumia map: ${jumiaMap.size} (EAN: ${jumiaEanMap.size})`);

  console.log('Building retail rows...');
  const retailRows=[];
  let matchC=0,matchL=0,matchJ=0;

  for (const sRow of stockRows) {
    const pid=String(sRow[SC.pid]||'').trim();
    if(!pid) continue;
    const gtin=String(sRow[SC.gtin]||'').trim();
    const rawCat=String(sRow[SC.cat]||'')||'(vide)';
    const typeV=String(sRow[SC.typeV]||'');
    const stockQty=parseNum(sRow[SC.stock]);
    const age=parseNum(sRow[SC.age]);
    const valeur=parseNum(sRow[SC.val]);
    const title=String(sRow[SC.title]||'');
    const gtinNorm=normalizeGtin(sRow[SC.gtin]);
    const repEntry=repSkuMap.get(gtinNorm)||repSkuMap.get(pid);
    const srcType=repEntry?normalizeType(repEntry.type):(baseRetailMap.get(gtinNorm)||baseRetailMap.get(pid)||'');
    const srcTypeLo=srcType.toLowerCase();
    const owner=srcTypeLo==='1p local b1'?'MB':srcTypeLo==='1p chine'?'SA':srcTypeLo==='1p local b2'?(repEntry.owner||''):'';
    const cRow=consoMap.get(pid);
    if(cRow) matchC++;
    const n1=cRow?String(cRow[CC.n1]||''):'';
    const n2=cRow?String(cRow[CC.n2]||''):'';
    const n3=cRow?String(cRow[CC.n3]||''):'';
    const marque=cRow?String(cRow[CC.brand]||''):'';
    const prixLive=cRow?parseNum(cRow[CC.price]):0;
    const vendeurBO=cRow?String(cRow[CC.shop]||''):'';
    const lDat=l3mMap.get(pid);
    if(lDat) matchL++;
    const pv=lDat?lDat.pv:0;
    const is_=lDat?lDat.is:0;
    const gmv=lDat?roundN(lDat.gmv,2):0;
    const margePct=0;
    const coutU=lDat?lDat.coutU:0;
    const jRow=jumiaMap.get(pid)||jumiaEanMap.get(gtinNorm);
    if(jRow) matchJ++;
    const prixJ=jRow?parseNum(jRow[JC.px]):'';
    const lienJ=jRow?String(jRow[JC.lien]||''):'';
    const catRevue=mapCategory(rawCat,n1);
    retailRows.push({catRevue,rawCat,typeV,srcType,owner,pid,gtin,title,marque,n1,n2,n3,stock:stockQty,age,valeur,pv,is:is_,gmv,margePct,coutU,prixLive,vendeurBO,prixJ,lienJ});
  }

  console.log(`Built ${retailRows.length} rows | Conso ${matchC} | L3M ${matchL} | Jumia ${matchJ}`);

  console.log('Writing Excel...');
  const wb=new ExcelJS.Workbook();
  wb.creator='Marjane Mall Report Generator';
  buildVueGlobale(wb.addWorksheet('Vue globale',{properties:{tabColor:{argb:'FF00B050'}}}),retailRows.length);
  buildReportRetail(wb.addWorksheet('Report Retail',{properties:{tabColor:{argb:'FF00B050'}}}),retailRows);
  const outPath=path.join(DIR,'pipeline_output_S14.xlsx');
  await wb.xlsx.writeFile(outPath);
  console.log(`Done → ${outPath}`);
}

main().catch(e=>{console.error(e);process.exit(1);});
