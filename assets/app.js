/* ============================================================
   Weekly Report Generator — Marjane Mall
   SheetJS (reading) + ExcelJS (formatted output + live formulas)
   ============================================================ */

// ── Category mapping (exact copy of "Categorie revue" sheet) ─
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
  'Maison', 'Beaute', 'Bricolage Jardin Animalerie', 'PEM', 'Mode',
  'Tel', 'Sport', 'Autres', 'TV Son', 'Bebe - Jouet', 'Informatique & gaming'
];

// Vue globale: 6 sections with filter criteria applied to Report Retail sheet
// columns: C=Type de Stock, A=Catégorie Revue, D=Type sourcing, M=Stock,
//          N=Age stock Moyen, R=Couverture (Jrs), T=PV L3M, AG=Check Prix, AJ=Check BO
const VG_SECTIONS = [
  { name: 'Scope total Retail',                      extra: '' },
  { name: 'Produits Non BO',                         extra: `,'Report Retail'!$AJ:$AJ,"Retail Non BO"` },
  { name: 'Produits KO Prix Jumia',                  extra: `,'Report Retail'!$AG:$AG,"prix \u00e0 revoir"` },
  { name: 'Age stock >180 jours',                    extra: `,'Report Retail'!$N:$N,">"&180` },
  { name: 'Couverture >120 jours et Age > 60 jours', extra: `,'Report Retail'!$R:$R,">"&120,'Report Retail'!$N:$N,">"&60` },
  { name: 'Low views (PV L3M <200)',                 extra: `,'Report Retail'!$T:$T,"<"&200` },
];

// ── Styles ────────────────────────────────────────────────────
const S = {
  hFill:  { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF000000' } },
  hFont:  { bold: true, color: { argb: 'FFFFFFFF' }, size: 11, name: 'Calibri' },
  hAlign: { horizontal: 'center', vertical: 'middle', wrapText: true },
  hBorderRow1: { top:{style:'medium'}, bottom:{style:'thin'},  left:{style:'medium'}, right:{style:'medium'} },
  hBorderRow2: { top:{style:'thin'},   bottom:{style:'thin'},  left:{style:'medium'}, right:{style:'medium'} },
  hBorderRow3: { top:{style:'thin'},   bottom:{style:'medium'},left:{style:'medium'}, right:{style:'medium'} },
  dAlign: { horizontal: 'center', vertical: 'middle' },
  dBorder:{ top:{style:'thin'}, bottom:{style:'thin'}, left:{style:'thin'}, right:{style:'thin'} },
  yFill:  { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } },
  numFmt: { numFmt: '#,##0' },
};

// Yellow columns in Report Retail (0-indexed col positions: I=8,J=9,K=10,L=11,T=19,X=23,Y=24,Z=25,AA=26,AC=28,AF=31,AH=33,AI=34)
const YELLOW_COLS = new Set([8, 9, 10, 11, 19, 23, 24, 25, 26, 28, 31, 33, 34]);

// ── State ─────────────────────────────────────────────────────
const files = { stock: null, conso: null, l3m: null, jumia: null, template: null };
let outputBlob = null;
let outputFilename = '';

// ── DOM ready ─────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('footer-year').textContent = new Date().getFullYear();
  document.getElementById('reportDate').valueAsDate = new Date();
  setupUploadZones();
  document.getElementById('btn-process').addEventListener('click', runProcess);
  document.getElementById('btn-download').addEventListener('click', downloadReport);
  document.getElementById('btn-reset').addEventListener('click', resetAll);
});

function setupUploadZones() {
  document.querySelectorAll('.upload-zone').forEach(zone => {
    const key   = zone.dataset.key;
    const input = zone.querySelector('.file-input');
    const label = zone.querySelector('.file-name');
    input.addEventListener('change', () => handleFile(input.files[0], key, zone, label));
    zone.addEventListener('dragover',  e => { e.preventDefault(); zone.classList.add('drag-over'); });
    zone.addEventListener('dragleave', ()  => zone.classList.remove('drag-over'));
    zone.addEventListener('drop', e => {
      e.preventDefault(); zone.classList.remove('drag-over');
      handleFile(e.dataTransfer.files[0], key, zone, label);
    });
  });
}
function handleFile(file, key, zone, label) {
  if (!file) return;
  files[key] = file; label.textContent = file.name; zone.classList.add('done'); checkReady();
}
function checkReady() {
  const weekOk  = !!document.getElementById('weekNum').value;
  const filesOk = files.stock && files.conso && files.l3m && files.jumia;
  document.getElementById('btn-process').disabled = !(weekOk && filesOk);
}
['weekNum', 'year', 'reportDate'].forEach(id =>
  document.getElementById(id).addEventListener('input', checkReady));

// ── Logging / progress ─────────────────────────────────────────
function log(msg, type = '') {
  const block = document.getElementById('log-block');
  block.classList.remove('hidden');
  const line = document.createElement('div');
  line.className = 'log-line' + (type ? ' ' + type : '');
  line.textContent = `[${new Date().toLocaleTimeString()}] ${msg}`;
  block.appendChild(line);
  block.scrollTop = block.scrollHeight;
}
function setProgress(pct, label) {
  document.getElementById('progress-bar').style.width = pct + '%';
  document.getElementById('progress-label').textContent = label;
}
const yield_ = () => new Promise(r => setTimeout(r, 0));

// ── SheetJS helpers ───────────────────────────────────────────
function readWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      try { resolve(XLSX.read(e.target.result, { type: 'array', cellDates: true })); }
      catch(err) { reject(err); }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}
function sheetToArrays(wb, sheetName) {
  const name = sheetName || wb.SheetNames[0];
  const ws   = wb.Sheets[name];
  if (!ws) throw new Error(`Feuille "${name}" introuvable.`);
  const all = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  return { headers: all[0] || [], rows: all.slice(1) };
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

// ── Data helpers ──────────────────────────────────────────────
function parseNum(v)       { if (v===''||v==null) return 0; const n=parseFloat(String(v).replace(',','.').replace(/\s/g,'')); return isNaN(n)?0:n; }
function roundN(n, d)      { const f=Math.pow(10,d); return Math.round(n*f)/f; }
function pct(a, b)         { return b ? Math.round(a/b*100)+'%' : '0%'; }

function mapCategory(n1Raw, rawCat) {
  for (const src of [n1Raw, rawCat]) {
    if (!src) continue;
    const up = String(src).toUpperCase().trim();
    if (CATEGORIE_MAP[up]) return CATEGORIE_MAP[up];
    for (const [key, val] of Object.entries(CATEGORIE_MAP)) {
      if (up.includes(key) || key.includes(up)) return val;
    }
  }
  return 'Autres';
}

function trancheAge(days) {
  const d = parseNum(days);
  if (d <= 60)  return '0-60jrs';
  if (d <= 120) return '60-120jrs';
  if (d <= 180) return '120-180jrs';
  return '>180jrs';
}

// ── Normalize a GTIN/barcode to a canonical string ────────────
// Handles: integers, floats (1.234E+12), strings with decimals, leading zeros
function normalizeGtin(raw) {
  if (raw === null || raw === undefined || raw === '') return '';
  // If it's already a number (from SheetJS numeric cell), convert directly
  if (typeof raw === 'number') {
    // Use Math.round to avoid floating-point artifacts (e.g. 9900006446318.001)
    return String(Math.round(raw));
  }
  const s = String(raw).trim();
  // Handle scientific notation strings like "9.9E+12" or "9.90000644631800E+12"
  if (/^[0-9.]+[eE][+\-]?[0-9]+$/.test(s)) {
    return String(Math.round(parseFloat(s)));
  }
  // Remove trailing .0 or .000
  return s.replace(/\.0+$/, '');
}

// ── Build Repartition SKU map from template ───────────────────
async function buildRepSkuMap(templateFile) {
  if (!templateFile) return new Map();
  try {
    const wb  = await readWorkbook(templateFile);
    const ws  = wb.Sheets['Repartition sku 1P'];
    if (!ws) { log('Repartition sku 1P sheet not found in template', 'warn'); return new Map(); }
    const { rows } = sheetToArrays(wb, 'Repartition sku 1P');
    const map = new Map();
    for (const row of rows) {
      const gtin = normalizeGtin(row[0]);
      const type = String(row[1] || '').trim();
      if (gtin && type) {
        map.set(gtin, type);
      }
    }
    log(`Repartition SKU map: ${map.size} entrées (B1/B2 sourcing)`, 'info');
    return map;
  } catch(e) {
    log('Erreur lecture template: ' + e.message, 'warn');
    return new Map();
  }
}

// ── MAIN PROCESS ──────────────────────────────────────────────
async function runProcess() {
  const btn = document.getElementById('btn-process');
  btn.disabled = true;
  document.getElementById('progress-block').classList.remove('hidden');
  document.getElementById('log-block').innerHTML = '';
  document.getElementById('log-block').classList.add('hidden');
  document.getElementById('result-block').classList.add('hidden');
  document.getElementById('validation-msg').classList.add('hidden');

  const weekNum = document.getElementById('weekNum').value;
  const year    = document.getElementById('year').value;
  outputFilename = `Weekly_Report_Stock_S${weekNum}_${year}.xlsx`;

  try {
    // ── 1. Template → Repartition SKU ─────────────────────────
    setProgress(3, 'Lecture du fichier modèle (B1/B2)…');
    await yield_();
    const repSkuMap = await buildRepSkuMap(files.template);
    const hasRepSku = repSkuMap.size > 0;
    if (!hasRepSku) log('Sans fichier modèle: B1/B2 indisponibles, seuls les totaux seront calculés.', 'warn');

    // ── 2. Stock ───────────────────────────────────────────────
    setProgress(8, 'Lecture Stock…');
    log(`Stock: ${files.stock.name}`);
    await yield_();
    const stockWb = await readWorkbook(files.stock);
    const { headers: sH, rows: stockRows } = sheetToArrays(stockWb);
    log(`Stock: ${stockRows.length.toLocaleString()} lignes`, 'info');

    const SC = {
      cat:   colIdx(sH, 'nom_categorie', 'categorie') !== -1 ? colIdx(sH, 'nom_categorie', 'categorie') : 0,
      pid:   colIdx(sH, 'ProductId', 'product_id', 'SKU') !== -1 ? colIdx(sH, 'ProductId', 'product_id', 'SKU') : 1,
      gtin:  colIdx(sH, 'gtin', 'GTIN', 'ean') !== -1 ? colIdx(sH, 'gtin', 'GTIN') : 2,
      typeV: colIdx(sH, 'type vendeur', 'type_vendeur') !== -1 ? colIdx(sH, 'type vendeur', 'type_vendeur') : 5,
      stock: colIdx(sH, 'stock_dispo', 'stock') !== -1 ? colIdx(sH, 'stock_dispo', 'stock') : 7,
      title: colIdx(sH, 'title', 'titre') !== -1 ? colIdx(sH, 'title', 'titre') : 8,
      age:   colIdx(sH, 'age_stock', 'age') !== -1 ? colIdx(sH, 'age_stock', 'age') : 10,
      val:   colIdx(sH, 'Valeur', 'valeur') !== -1 ? colIdx(sH, 'Valeur', 'valeur') : 11,
    };
    await yield_();

    // ── 3. Conso ───────────────────────────────────────────────
    setProgress(18, 'Lecture Conso…');
    log(`Conso: ${files.conso.name}`);
    await yield_();
    const consoWb = await readWorkbook(files.conso);
    const { headers: cH, rows: consoRows } = sheetToArrays(consoWb);
    log(`Conso: ${consoRows.length.toLocaleString()} lignes`, 'info');

    // Actual conso columns: Unnamed:0, Label_N1, Label_N2, LABEL_N3, ProductId, ParentID, gtin,
    //   title, SellerId, shopname, status, BestOfferRank, Price, OriginPrice, SupplyMode, brandlabel...
    const CC = {
      n1:    colIdx(cH, 'Label_N1','N1','label_n1') !== -1 ? colIdx(cH,'Label_N1','N1','label_n1') : 1,
      n2:    colIdx(cH, 'Label_N2','N2') !== -1 ? colIdx(cH,'Label_N2','N2') : 2,
      n3:    colIdx(cH, 'LABEL_N3','Label_N3','N3') !== -1 ? colIdx(cH,'LABEL_N3','Label_N3','N3') : 3,
      pid:   colIdx(cH, 'ProductId','product_id','Id') !== -1 ? colIdx(cH,'ProductId','product_id','Id') : 4,
      shop:  colIdx(cH, 'shopname','Shopname','shop_name','vendeur') !== -1 ? colIdx(cH,'shopname','Shopname','shop_name','vendeur') : 9,
      rank:  colIdx(cH, 'BestOfferRank','bestofferrank','rank') !== -1 ? colIdx(cH,'BestOfferRank','bestofferrank','rank') : 11,
      price: colIdx(cH, 'Price','OfferPrice','offerprice','prix') !== -1 ? colIdx(cH,'Price','OfferPrice','offerprice','prix') : 12,
      brand: colIdx(cH, 'brandlabel','Brandlabel','brand','marque') !== -1 ? colIdx(cH,'brandlabel','Brandlabel','brand','marque') : 15,
    };
    // Build consoMap: for each ProductId, prefer the BestOfferRank=1 row (the actual BO seller)
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
    log(`Conso map: ${consoMap.size.toLocaleString()} entrées`);
    await yield_();

    // ── 4. L3M ────────────────────────────────────────────────
    setProgress(40, 'Lecture L3M…');
    log(`L3M: ${files.l3m.name}`);
    await yield_();
    const l3mWb = await readWorkbook(files.l3m);
    const { headers: lH, rows: l3mRows } = sheetToArrays(l3mWb);
    log(`L3M: ${l3mRows.length.toLocaleString()} lignes`, 'info');

    // Actual L3M columns: SKU(0), Title(1), Seller BO(2), Seller Max IS(3), Brand(4),
    //   N1(5)...PV(10), IS(20), GMV(27), marge l3M(56), cout unitaire(57)
    const LC = {
      pid:   colIdx(lH,'SKU','ProductId','sku') !== -1 ? colIdx(lH,'SKU','ProductId','sku') : 0,
      pv:    colIdx(lH,'PV','Views','page_views') !== -1 ? colIdx(lH,'PV','Views','page_views') : 10,
      is_:   colIdx(lH,'IS','is_l3m') !== -1 ? colIdx(lH,'IS','is_l3m') : 20,
      gmv:   colIdx(lH,'GMV','gmv') !== -1 ? colIdx(lH,'GMV','gmv') : 27,
      marge: colIdx(lH,'marge l3M','Marge L3M','marge_l3m') !== -1 ? colIdx(lH,'marge l3M','Marge L3M','marge_l3m') : 56,
      coutU: colIdx(lH,'cout unitaire','Cout unitaire','cout_unitaire') !== -1 ? colIdx(lH,'cout unitaire','Cout unitaire','cout_unitaire') : 57,
    };
    const l3mMap = new Map();
    for (const row of l3mRows) {
      const k = String(row[LC.pid]||'').trim();
      if (!k) continue;
      const pv=parseNum(row[LC.pv]), is_=parseNum(row[LC.is_]), gmv=parseNum(row[LC.gmv]);
      const marge=parseNum(row[LC.marge]), coutU=parseNum(row[LC.coutU]);
      if (l3mMap.has(k)) {
        const e=l3mMap.get(k);
        e.pv+=pv; e.is+=is_; e.gmv+=gmv;
      } else {
        l3mMap.set(k, {pv, is:is_, gmv, marge, coutU});
      }
    }
    log(`L3M map: ${l3mMap.size.toLocaleString()} produits`);
    await yield_();

    // ── 5. Jumia ──────────────────────────────────────────────
    setProgress(62, 'Lecture IP Jumia…');
    log(`Jumia: ${files.jumia.name}`);
    await yield_();
    const jumiaWb = await readWorkbook(files.jumia);
    const { headers: jH, rows: jumiaRows } = sheetToArrays(jumiaWb);
    log(`Jumia: ${jumiaRows.length.toLocaleString()} lignes`, 'info');
    // From workflow: C=sku(key=col2), K=prix_jumia(col10), N=Lien(col13), F=mon_ean(col5)
    const JC = {
      pid:  colIdx(jH,'sku','SKU','ProductId') !== -1 ? colIdx(jH,'sku','SKU','ProductId') : 2,
      px:   colIdx(jH,'prix_jumia','Prix_Jumia') !== -1 ? colIdx(jH,'prix_jumia','Prix_Jumia') : 10,
      lien: colIdx(jH,'Lien_du_produit','lien','link') !== -1 ? colIdx(jH,'Lien_du_produit','lien') : 13,
      ean:  colIdx(jH,'mon_ean','ean','EAN') !== -1 ? colIdx(jH,'mon_ean','ean','EAN') : 5,
    };
    const jumiaMap    = new Map();  // by SKU
    const jumiaEanMap = new Map();  // by EAN (fallback)
    for (const row of jumiaRows) {
      const k = String(row[JC.pid]||'').trim();
      if (k && !jumiaMap.has(k)) jumiaMap.set(k, row);
      const eanRaw = row[JC.ean];
      if (eanRaw !== null && eanRaw !== undefined && eanRaw !== '') {
        const eanKey = typeof eanRaw === 'number' ? String(Math.round(eanRaw)) : String(eanRaw).trim();
        if (eanKey && !jumiaEanMap.has(eanKey)) jumiaEanMap.set(eanKey, row);
      }
    }
    log(`Jumia map: ${jumiaMap.size.toLocaleString()} entrées (EAN: ${jumiaEanMap.size.toLocaleString()})`);
    await yield_();

    // ── 6. Build Report Retail rows ───────────────────────────
    setProgress(75, 'Fusion des données…');
    log('Construction des lignes Report Retail…');
    await yield_();

    const retailRows = [];
    let matchC=0, matchL=0, matchJ=0, matchT=0;

    for (const sRow of stockRows) {
      const pid      = String(sRow[SC.pid]||'').trim();
      const gtin     = String(sRow[SC.gtin]||'').trim();
      const rawCat   = String(sRow[SC.cat]||'');
      const typeV    = String(sRow[SC.typeV]||'');
      const stockQty = parseNum(sRow[SC.stock]);
      const age      = parseNum(sRow[SC.age]);
      const valeur   = parseNum(sRow[SC.val]);
      const title    = String(sRow[SC.title]||'');

      // Sourcing type from Repartition sku 1P (normalize GTIN to avoid float/scientific notation mismatches)
      const gtinNorm = normalizeGtin(sRow[SC.gtin]);
      const srcType = repSkuMap.get(gtinNorm) || repSkuMap.get(pid) || '';
      if (srcType) matchT++;

      // Owner
      const owner = srcType === '1P Local B1' ? 'MB' : srcType === '1P Chine' ? 'SA' : '';

      // Conso
      const cRow   = consoMap.get(pid);
      if (cRow) matchC++;
      const n1     = cRow ? String(cRow[CC.n1]||'')    : '';
      const n2     = cRow ? String(cRow[CC.n2]||'')    : '';
      const n3     = cRow ? String(cRow[CC.n3]||'')    : '';
      const marque = cRow ? String(cRow[CC.brand]||'') : '';
      const prixLive = cRow ? parseNum(cRow[CC.price]) : 0;
      const vendeurBO= cRow ? String(cRow[CC.shop]||'')  : '';

      // L3M
      const lDat = l3mMap.get(pid);
      if (lDat) matchL++;
      const pv        = lDat ? lDat.pv           : 0;
      const is_       = lDat ? lDat.is           : 0;
      const gmv       = lDat ? roundN(lDat.gmv,2): 0;
      const margePct  = lDat ? lDat.marge        : 0;  // already a decimal (e.g. 0.0922)
      const coutU     = lDat ? lDat.coutU        : 0;

      // Jumia — lookup by SKU first, then fallback to normalized GTIN/EAN
      const jRow  = jumiaMap.get(pid) || jumiaEanMap.get(gtinNorm);
      if (jRow) matchJ++;
      const prixJ = jRow ? parseNum(jRow[JC.px])        : 0;
      const lienJ = jRow ? String(jRow[JC.lien]||'')    : '';

      // Category — use Stock nom_categorie first (proven correct by Python),
      // conso n1 only as fallback (can mismatch if product has different Label_N1 in marketplace)
      const catRevue = mapCategory(rawCat, n1);

      retailRows.push({
        catRevue, rawCat, typeV, srcType, owner,
        pid, gtin, title,
        marque, n1, n2, n3,
        stock: stockQty, age, valeur,
        pv, is: is_, gmv, margePct, coutU,
        prixLive, vendeurBO, prixJ, lienJ,
      });
    }

    log(`Fusion: ${retailRows.length.toLocaleString()} lignes`, 'info');
    log(`  Sourcing B1/B2: ${pct(matchT,retailRows.length)} | Conso: ${pct(matchC,retailRows.length)} | L3M: ${pct(matchL,retailRows.length)} | Jumia: ${pct(matchJ,retailRows.length)}`);
    await yield_();

    // ── 7. Generate Excel with ExcelJS ─────────────────────────
    setProgress(85, 'Génération Excel formaté…');
    log('Construction du classeur avec formules et mise en forme…');
    await yield_();

    outputBlob = await buildExcelOutput(retailRows, weekNum, year);

    setProgress(100, 'Rapport généré !');
    log(`✅ ${outputFilename}`, 'info');

    document.getElementById('result-title').textContent = `Rapport S${weekNum}_${year} généré`;
    document.getElementById('result-summary').textContent =
      `${retailRows.length.toLocaleString()} produits · ` +
      `Sourcing ${pct(matchT,retailRows.length)} · ` +
      `Conso ${pct(matchC,retailRows.length)} · L3M ${pct(matchL,retailRows.length)} · Jumia ${pct(matchJ,retailRows.length)}`;
    document.getElementById('result-block').classList.remove('hidden');

  } catch (err) {
    log(`ERREUR: ${err.message}`, 'error');
    console.error(err);
    setProgress(0, 'Erreur.');
    const msg = document.getElementById('validation-msg');
    msg.textContent = `Erreur: ${err.message}`;
    msg.classList.remove('hidden');
  } finally {
    btn.disabled = false;
  }
}

// ── Excel generation ──────────────────────────────────────────
async function buildExcelOutput(retailRows, weekNum, year) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Marjane Mall Report Generator';
  wb.created = new Date();

  // Sheet 1: Vue globale
  const vgWs = wb.addWorksheet('Vue globale', { properties: { tabColor: { argb: 'FF00B050' } } });
  buildVueGlobale(vgWs, retailRows.length);

  // Sheet 2: Report Retail (data + formulas)
  const rrWs = wb.addWorksheet('Report Retail', { properties: { tabColor: { argb: 'FF00B050' } } });
  buildReportRetail(rrWs, retailRows);

  const buffer = await wb.xlsx.writeBuffer();
  return new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

// ── Vue globale ───────────────────────────────────────────────
function buildVueGlobale(ws, totalRetailRows) {
  // Column widths
  ws.getColumn(1).width = 32;
  for (let c = 2; c <= 37; c++) ws.getColumn(c).width = 10;

  // Row 1 height
  ws.getRow(1).height = 36;
  ws.getRow(2).height = 20;
  ws.getRow(3).height = 18;

  // ── Merge layout ────────────────────────────────────────────
  ws.mergeCells('A1:A3'); // "Catégorie" header spans rows 1-3

  const sectionStarts = ['B','H','N','T','Z','AF'];
  sectionStarts.forEach((col, i) => {
    const endCol = String.fromCharCode(col.charCodeAt(0) + 5);
    const endColFull = i < 4
      ? String.fromCharCode(col.charCodeAt(0) + 5)      // single letter: B→G, H→M, N→S, T→Y
      : (i === 4 ? 'AE' : 'AK');                        // Z→AE, AF→AK

    // Row 1: section header (6 cols each)
    ws.mergeCells(`${col}1:${endColFull}1`);

    // Row 2: B1(2cols), B2(2cols), Total(2cols)
    const c0 = colLetterToNum(col);
    ws.mergeCells(`${numToColLetter(c0)}2:${numToColLetter(c0+1)}2`);
    ws.mergeCells(`${numToColLetter(c0+2)}2:${numToColLetter(c0+3)}2`);
    ws.mergeCells(`${numToColLetter(c0+4)}2:${numToColLetter(c0+5)}2`);
  });

  // ── Row 1: section headers ───────────────────────────────────
  styleCell(ws.getCell('A1'), 'Catégorie', S.hFill, S.hFont, S.hAlign, S.hBorderRow1);
  VG_SECTIONS.forEach((sec, i) => {
    const col = sectionStarts[i];
    const c = ws.getCell(`${col}1`);
    styleCell(c, sec.name, S.hFill, S.hFont, S.hAlign, S.hBorderRow1);
  });

  // ── Row 2: B1 / B2 / Total ──────────────────────────────────
  sectionStarts.forEach(col => {
    const c0 = colLetterToNum(col);
    ['B1','B2','Total'].forEach((lbl, j) => {
      const c = ws.getCell(`${numToColLetter(c0 + j*2)}2`);
      styleCell(c, lbl, S.hFill, S.hFont, S.hAlign, S.hBorderRow2);
    });
  });

  // ── Row 3: #SKU / Stock col headers ──────────────────────────
  styleCell(ws.getCell('A3'), 'Catégorie', S.hFill, S.hFont, S.hAlign, S.hBorderRow3);
  sectionStarts.forEach(col => {
    const c0 = colLetterToNum(col);
    for (let j = 0; j < 6; j++) {
      const lbl = j % 2 === 0 ? '# SKU' : 'Stock';
      const c = ws.getCell(`${numToColLetter(c0+j)}3`);
      styleCell(c, lbl, S.hFill, S.hFont, S.hAlign, S.hBorderRow3);
    }
  });

  // ── Data rows (4–14): one per category ──────────────────────
  const dataRows = [...CATEGORIES_ORDER]; // 11 categories
  const N = totalRetailRows; // number of data rows in Report Retail

  dataRows.forEach((cat, i) => {
    const row = 4 + i;
    // Column A: category name
    const aCell = ws.getCell(`A${row}`);
    aCell.value = cat;
    aCell.font  = { size: 11, name: 'Calibri' };
    aCell.alignment = S.dAlign;
    aCell.border = S.dBorder;

    // For each of the 6 sections → 6 formula cells (B1 #SKU, B1 Stock, B2 #SKU, B2 Stock, Total #SKU, Total Stock)
    VG_SECTIONS.forEach((sec, si) => {
      const c0 = colLetterToNum(sectionStarts[si]);
      const extra = sec.extra;

      // B1 #SKU
      setVGCell(ws, numToColLetter(c0),   row, mkCountIf(cat, '1P Local B1',   extra));
      // B1 Stock
      setVGCell(ws, numToColLetter(c0+1), row, mkSumIf(cat,   '1P Local B1',   extra));
      // B2 #SKU
      setVGCell(ws, numToColLetter(c0+2), row, mkCountIf(cat, '1P LOCAL B2',   extra));
      // B2 Stock
      setVGCell(ws, numToColLetter(c0+3), row, mkSumIf(cat,   '1P LOCAL B2',   extra));
      // Total #SKU (no type filter — always correct)
      setVGCell(ws, numToColLetter(c0+4), row, mkCountIf(cat, null,             extra));
      // Total Stock
      setVGCell(ws, numToColLetter(c0+5), row, mkSumIf(cat,   null,             extra));
    });
  });

  // ── Row 15: Total général ─────────────────────────────────────
  const totalRow = 4 + dataRows.length; // row 15
  const aTotal   = ws.getCell(`A${totalRow}`);
  aTotal.value   = 'Total général';
  aTotal.font    = { bold: true, size: 11, name: 'Calibri' };
  aTotal.alignment = S.dAlign;
  aTotal.border  = S.dBorder;

  sectionStarts.forEach(col => {
    const c0 = colLetterToNum(col);
    for (let j = 0; j < 6; j++) {
      const cellRef = numToColLetter(c0+j);
      const c = ws.getCell(`${cellRef}${totalRow}`);
      c.value     = { formula: `SUM(${cellRef}4:${cellRef}${totalRow-1})` };
      c.numFmt    = '#,##0';
      c.font      = { bold: true, size: 11, name: 'Calibri' };
      c.alignment = S.dAlign;
      c.border    = S.dBorder;
    }
  });
}

// ── Formula builders for Vue globale ──────────────────────────
function mkCountIf(cat, type, extra) {
  const rr    = `'Report Retail'`;
  const catCr = `${rr}!$A:$A,"${cat}"`;
  const typeCr= type ? `,${rr}!$D:$D,"${type}"` : '';
  return `COUNTIFS(${rr}!$C:$C,"Retail",${catCr}${typeCr}${extra})`;
}
function mkSumIf(cat, type, extra) {
  const rr    = `'Report Retail'`;
  const catCr = `${rr}!$A:$A,"${cat}"`;
  const typeCr= type ? `,${rr}!$D:$D,"${type}"` : '';
  return `SUMIFS(${rr}!$M:$M,${rr}!$C:$C,"Retail",${catCr}${typeCr}${extra})`;
}
function setVGCell(ws, col, row, formula) {
  const c    = ws.getCell(`${col}${row}`);
  c.value    = { formula };
  c.numFmt   = '#,##0';
  c.alignment = S.dAlign;
  c.border   = S.dBorder;
  c.font     = { size: 11, name: 'Calibri' };
}

// ── Report Retail ─────────────────────────────────────────────
const RR_HEADERS = [
  'Catégorie Revue','Categorie','Type de Stock','Type','Owner',
  'ProductId','gtin','title',
  'Marque','N1','N2','N3',
  'Stock','Age stock Moyen','Tranche Age','VALEUR PV HT',
  'Moyenne de vente','Couverture (Jrs)','Tranche de couverture',
  'PV L3M','Tranche de PV','CR L3M','Tranche de CR',
  'IS L3M','GMV L3M','Marge L3M','Cout unitaire',
  'ASP L3M','Prix Live','Marge live','Statut marge',
  'Prix Jumia','Check Prix','Lien Jumia','Vendeur BO','Check BO','Animation commerciale'
];

function buildReportRetail(ws, retailRows) {
  // Column widths
  const widths = [22,22,12,14,10,16,16,40,18,22,22,22,10,14,12,14,14,14,18,10,12,10,12,10,12,12,12,10,12,12,12,12,18,40,20,16,40];
  widths.forEach((w, i) => { ws.getColumn(i+1).width = w; });

  // Header row
  const hRow = ws.getRow(1);
  hRow.height = 20;
  RR_HEADERS.forEach((h, i) => {
    const c = hRow.getCell(i+1);
    c.value     = h;
    c.font      = { bold: true, size: 11, name: 'Calibri' };
    c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    c.border    = S.dBorder;
    if (YELLOW_COLS.has(i)) c.fill = S.yFill;
  });

  // Data rows
  retailRows.forEach((row, idx) => {
    const r   = idx + 2; // Excel row number
    const eRow = ws.getRow(r);
    eRow.height = 15;

    // Pre-computed values (from external lookups)
    const vals = [
      row.catRevue,  row.rawCat,    row.typeV,    row.srcType,  row.owner,
      row.pid,       row.gtin,      row.title,
      row.marque,    row.n1,        row.n2,        row.n3,
      row.stock,     row.age,
      null, // O: Tranche Age — formula
      row.valeur,
      null, // Q: Moyenne de vente — formula
      null, // R: Couverture — formula
      null, // S: Tranche couverture — formula
      row.pv,
      null, // U: Tranche PV — formula
      null, // V: CR L3M — formula
      null, // W: Tranche CR — formula
      row.is,        row.gmv,       row.margePct,  row.coutU,
      null, // AB: ASP — formula
      row.prixLive,
      null, // AD: Marge live — formula
      null, // AE: Statut marge — formula (needs sourcing)
      row.prixJ,
      null, // AG: Check Prix — formula
      row.lienJ,     row.vendeurBO,
      null, // AJ: Check BO — formula
      null, // AK: Animation — formula
    ];

    // Write pre-computed values
    vals.forEach((v, i) => {
      if (v !== null) {
        const c = eRow.getCell(i+1);
        c.value = v;
        if (typeof v === 'number' && i >= 12) c.numFmt = '#,##0.##';
      }
    });

    // Write formulas for computed columns
    // O (15): Tranche Age
    eRow.getCell(15).value = { formula: `IF(N${r}<=60,"0-60jrs",IF(N${r}<=120,"60-120jrs",IF(N${r}<=180,"120-180jrs",">180jrs")))` };
    // Q (17): Moyenne de vente
    eRow.getCell(17).value = { formula: `IFERROR(X${r}/90,0)` };
    // R (18): Couverture
    eRow.getCell(18).value = { formula: `IFERROR(ROUND(M${r}/Q${r},0),"")` };
    // S (19): Tranche couverture
    eRow.getCell(19).value = { formula: `IF(R${r}="","",IF(R${r}<=60,"0-60jrs",IF(R${r}<=120,"60-120jrs",IF(R${r}<=180,"120-180jrs",">180jrs"))))` };
    // U (21): Tranche PV
    eRow.getCell(21).value = { formula: `IF(T${r}<50,"<50",IF(T${r}<200,"50-200",IF(T${r}<500,"200-500",IF(T${r}<1000,"500-1000",">=1000"))))` };
    // V (22): CR L3M
    eRow.getCell(22).value = { formula: `IFERROR(X${r}/T${r},"")` };
    // W (23): Tranche CR
    eRow.getCell(23).value = { formula: `IFERROR(IF(V${r}<0.01,"<1%",IF(V${r}<0.03,"1-3%",IF(V${r}<0.1,"3-10%",">10%"))),"")` };
    // AB (28): ASP L3M
    eRow.getCell(28).value = { formula: `IFERROR(Y${r}/X${r},"")` };
    // AD (30): Marge live  = (Prix_Live/1.2 - Cout_unit) / (Prix_Live/1.2)
    eRow.getCell(30).value = { formula: `IFERROR((AC${r}/1.2-AA${r})/(AC${r}/1.2),"")` };
    // AG (33): Check Prix
    eRow.getCell(33).value = { formula: `IFERROR(IF(AF${r}="","",IF(AF${r}<AC${r},"prix \u00e0 revoir",IF(AC${r}>AB${r},"Prix Sup\u00e9rieur \u00e0 ASP",""))),"")` };
    // AJ (36): Check BO — Non-BO when vendeur is anything other than MARJANEMALL (includes empty = not in conso)
    eRow.getCell(36).value = { formula: `IF(AND(C${r}="Retail",UPPER(AI${r})<>"MARJANEMALL"),"Retail Non BO","")` };
    // AK (37): Animation commerciale
    eRow.getCell(37).value = { formula: `IF(AND(N${r}>180,R${r}>120),"Baisse de prix permanente ou retour fournisseur",IF(AND(N${r}>60,R${r}>120),"Baisse de prix permanente",IF(AND(R${r}>60,V${r}>0.05),"Vente Flash","")))` };
  });

  // Auto-filter on header row
  ws.autoFilter = { from: 'A1', to: `AK1` };
}

// ── ExcelJS style helper ──────────────────────────────────────
function styleCell(cell, value, fill, font, alignment, border) {
  cell.value     = value;
  cell.fill      = fill;
  cell.font      = font;
  cell.alignment = alignment;
  cell.border    = border;
}

// ── Column letter ↔ number conversion ────────────────────────
function colLetterToNum(col) {
  let n = 0;
  for (let i = 0; i < col.length; i++) {
    n = n * 26 + (col.charCodeAt(i) - 64);
  }
  return n;
}
function numToColLetter(n) {
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

// ── Download / Reset ──────────────────────────────────────────
function downloadReport() {
  if (!outputBlob) return;
  const url = URL.createObjectURL(outputBlob);
  const a   = document.createElement('a');
  a.href    = url;
  a.download = outputFilename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function resetAll() {
  Object.keys(files).forEach(k => files[k] = null);
  document.querySelectorAll('.upload-zone').forEach(z => {
    z.classList.remove('done');
    z.querySelector('.file-name').textContent = '';
    z.querySelector('.file-input').value = '';
  });
  ['progress-block','log-block','result-block'].forEach(id =>
    document.getElementById(id).classList.add('hidden'));
  document.getElementById('btn-process').disabled = true;
  outputBlob = null;
}
