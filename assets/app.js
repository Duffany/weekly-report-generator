/* ============================================================
   Weekly Report Generator — Marjane Mall
   Full rewrite: exact Vue globale layout + 37-col Report Retail
   ============================================================ */

// ── Category mapping (from hidden "Categorie revue" sheet) ───
const CATEGORIE_MAP = {
  'HYGIENE - BEAUTE - PARFUM':           'Beaute',
  'MATERIEL MEDICAL':                    'Beaute',
  'PARAPHARMACIE':                       'Beaute',
  'JEUX - JOUETS':                       'Bebe - Jouet',
  'PUERICULTURE':                        'Bebe - Jouet',
  'ANIMALERIE':                          'Bricolage Jardin Animalerie',
  'BRICOLAGE - OUTILLAGE':               'Bricolage Jardin Animalerie',
  'DROGUERIE':                           'Bricolage Jardin Animalerie',
  'JARDIN - PISCINE':                    'Bricolage Jardin Animalerie',
  'ELECTROMENAGER':                      'PEM',
  'INFORMATIQUE':                        'Informatique & gaming',
  'JEUX VIDEO':                          'Informatique & gaming',
  'ART DE LA TABLE':                     'Maison',
  'DECO - LINGE - LUMINAIRE':            'Maison',
  'DECO - LINGE':                        'Maison',
  'LITERIE':                             'Maison',
  'MEUBLE':                              'Maison',
  'BAGAGERIE':                           'Mode',
  'BIJOUX':                              'Mode',
  'CHAUSSURES':                          'Mode',
  'VETEMENTS':                           'Mode',
  'SPORT':                               'Sport',
  'TELEPHONIE - GPS':                    'Tel',
  'INSTRUMENTS DE MUSIQUE':              'TV Son',
  'PHOTO':                               'TV Son',
  'SONO':                                'TV Son',
  'TV-VIDEO-SON':                        'TV Son',
  'DVD':                                 'TV Son',
  'MUSIQUE':                             'TV Son',
};

const CATEGORIES_ORDER = [
  'Maison', 'Beaute', 'Bricolage Jardin Animalerie', 'PEM', 'Mode',
  'Tel', 'Sport', 'Autres', 'TV Son', 'Bebe - Jouet', 'Informatique & gaming'
];

// ── State ────────────────────────────────────────────────────
const files = { stock: null, conso: null, l3m: null, jumia: null, template: null };
let outputWorkbook = null;

// ── DOM ready ────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('footer-year').textContent = new Date().getFullYear();
  document.getElementById('reportDate').valueAsDate = new Date();
  setupUploadZones();
  document.getElementById('btn-process').addEventListener('click', runProcess);
  document.getElementById('btn-download').addEventListener('click', downloadReport);
  document.getElementById('btn-reset').addEventListener('click', resetAll);
});

// ── Upload zones ─────────────────────────────────────────────
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

// ── Logging / progress ────────────────────────────────────────
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

// ── Parse Excel → raw arrays (header row + data rows) ────────
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

// Returns { headers:[], rows:[[]] } — positional, safe for large files
function sheetToArrays(wb, sheetName) {
  const name = sheetName || wb.SheetNames[0];
  const ws   = wb.Sheets[name];
  if (!ws) throw new Error(`Feuille "${name}" introuvable.`);
  const all     = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const headers = all[0] || [];
  return { headers, rows: all.slice(1) };
}

// Find column index by name (case-insensitive, partial match fallback)
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

// ── Helpers ───────────────────────────────────────────────────
function parseNum(v) {
  if (v === '' || v == null) return 0;
  const n = parseFloat(String(v).replace(',', '.').replace(/\s/g, ''));
  return isNaN(n) ? 0 : n;
}
function roundN(n, d) { const f = Math.pow(10, d); return Math.round(n * f) / f; }
function pct(a, b)    { return b ? Math.round(a/b*100) + '%' : '0%'; }

function mapCategory(n1Raw, rawCat) {
  // Try N1 from Conso first, then raw stock category
  for (const src of [n1Raw, rawCat]) {
    if (!src) continue;
    const up = String(src).toUpperCase().trim();
    if (CATEGORIE_MAP[up]) return CATEGORIE_MAP[up];
    // Partial match
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
function trancheCoverage(days) { return trancheAge(days); } // same thresholds

function tranchePV(pv) {
  const v = parseNum(pv);
  if (v < 50)   return '<50';
  if (v < 200)  return '50-200';
  if (v < 500)  return '200-500';
  if (v < 1000) return '500-1000';
  return '≥1000';
}
function trancheCR(cr) {
  const v = parseNum(cr);
  if (v <= 0)    return '';
  if (v < 0.01)  return '<1%';
  if (v < 0.03)  return '1-3%';
  if (v < 0.10)  return '3-10%';
  return '≥10%';
}

// Detect B1 or B2 from stock row zone/empl columns
function detectZone(zoneStr, emplStr) {
  const z = String(zoneStr || '').toUpperCase();
  const e = String(emplStr || '').toUpperCase();
  if (z.includes('B2') || e.includes('B2')) return 'B2';
  if (z.includes('B1') || e.includes('B1')) return 'B1';
  return 'B1'; // default
}

// ── MAIN PROCESS ─────────────────────────────────────────────
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

  try {
    // ── 1. Stock ────────────────────────────────────────────
    setProgress(5, 'Lecture Stock…');
    log(`Stock: ${files.stock.name}`);
    await yield_();
    const stockWb = await readWorkbook(files.stock);
    const { headers: sH, rows: stockRows } = sheetToArrays(stockWb);
    log(`Stock: ${stockRows.length.toLocaleString()} lignes, ${sH.length} colonnes`, 'info');

    // Stock column indices
    const S = {
      cat:      colIdx(sH, 'nom_categorie', 'nom_cat', 'categorie'),
      pid:      colIdx(sH, 'ProductId', 'product_id', 'SKU', 'Id'),
      gtin:     colIdx(sH, 'gtin', 'GTIN', 'ean'),
      statut:   colIdx(sH, 'statut', 'status'),
      seller:   colIdx(sH, 'seller', 'vendeur'),
      typeV:    colIdx(sH, 'type vendeur', 'type_vendeur', 'TypeVendeur'),
      sellerId: colIdx(sH, 'sellerid', 'seller_id'),
      stock:    colIdx(sH, 'stock_dispo', 'stock', 'Stock'),
      title:    colIdx(sH, 'title', 'titre', 'nom_du_produit'),
      dateRec:  colIdx(sH, 'date_recep', 'date_reception'),
      age:      colIdx(sH, 'age_stock', 'age'),
      valeur:   colIdx(sH, 'Valeur', 'valeur', 'value'),
      qty:      colIdx(sH, 'Quantite', 'quantite', 'qty'),
      tranche:  colIdx(sH, 'tranche_age', 'tranche'),
      empl:     colIdx(sH, 'Type Empl', 'type_empl'),
      zone:     colIdx(sH, 'Zone Stockage', 'zone_stockage', 'zone'),
    };
    log(`Stock clé détectée: col[${S.pid}]="${sH[S.pid]}"`);
    await yield_();

    // ── 2. Conso ────────────────────────────────────────────
    setProgress(18, 'Lecture Conso (fichier large, patience…)');
    log(`Conso: ${files.conso.name}`);
    await yield_();
    const consoWb = await readWorkbook(files.conso);
    const { headers: cH, rows: consoRows } = sheetToArrays(consoWb);
    log(`Conso: ${consoRows.length.toLocaleString()} lignes`, 'info');

    // Conso column indices — from workflow: A=N1, B=N2, C=N3, D=ProductId, I=Shopname, L=OfferPrice, O=Brandlabel
    const C = {
      n1:     colIdx(cH, 'Label_N1', 'N1', 'label_n1') !== -1 ? colIdx(cH, 'Label_N1', 'N1', 'label_n1') : 0,
      n2:     colIdx(cH, 'Label_N2', 'N2', 'label_n2') !== -1 ? colIdx(cH, 'Label_N2', 'N2', 'label_n2') : 1,
      n3:     colIdx(cH, 'Label_N3', 'N3', 'label_n3') !== -1 ? colIdx(cH, 'Label_N3', 'N3', 'label_n3') : 2,
      pid:    colIdx(cH, 'ProductId', 'product_id', 'Id') !== -1 ? colIdx(cH, 'ProductId', 'product_id', 'Id') : 3,
      shop:   colIdx(cH, 'Shopname', 'shop_name', 'vendeur', 'seller') !== -1 ? colIdx(cH, 'Shopname', 'shop_name', 'vendeur') : 8,
      price:  colIdx(cH, 'OfferPrice', 'offerprice', 'prix', 'price') !== -1 ? colIdx(cH, 'OfferPrice', 'offerprice', 'prix', 'price') : 11,
      brand:  colIdx(cH, 'Brandlabel', 'brand', 'marque', 'brand_label') !== -1 ? colIdx(cH, 'Brandlabel', 'brand', 'marque') : 14,
    };
    log(`Conso — pid:${C.pid}="${cH[C.pid]}", N1:${C.n1}, brand:${C.brand}="${cH[C.brand]}", price:${C.price}="${cH[C.price]}"`);

    const consoMap = new Map();
    for (const row of consoRows) {
      const k = String(row[C.pid] || '').trim();
      if (k && !consoMap.has(k)) consoMap.set(k, row);
    }
    log(`Conso map: ${consoMap.size.toLocaleString()} entrées`);
    await yield_();

    // ── 3. L3M ─────────────────────────────────────────────
    setProgress(40, 'Lecture L3M (fichier large, patience…)');
    log(`L3M: ${files.l3m.name}`);
    await yield_();
    const l3mWb = await readWorkbook(files.l3m);
    const { headers: lH, rows: l3mRows } = sheetToArrays(l3mWb);
    log(`L3M: ${l3mRows.length.toLocaleString()} lignes, ${lH.length} colonnes`, 'info');

    // L3M columns — positional from workflow formulas (A=0, E=4, F=5, G=6, Q=16, W=22, Z=25)
    // Also try header detection
    const L = {
      pid: colIdx(lH, 'ProductId', 'SKU', 'sku', 'Id') !== -1 ? colIdx(lH, 'ProductId', 'SKU', 'sku') : 0,
      pv:  colIdx(lH, 'Views', 'PV', 'page_views', 'vues') !== -1 ? colIdx(lH, 'Views', 'PV', 'page_views') : 4,
      is_: colIdx(lH, 'Items', 'IS', 'items_sold', 'commandes') !== -1 ? colIdx(lH, 'Items', 'IS', 'items_sold') : 5,
      gmv: colIdx(lH, 'CA ALL', 'GMV', 'Revenue', 'CA', 'ca_all') !== -1 ? colIdx(lH, 'CA ALL', 'GMV', 'Revenue', 'CA') : 6,
      qty: colIdx(lH, 'Quantite', 'Quantity', 'qty') !== -1 ? colIdx(lH, 'Quantite', 'Quantity', 'qty') : 16,
      ca:  colIdx(lH, 'CA Retail', 'CA_HT', 'montant') !== -1 ? colIdx(lH, 'CA Retail', 'CA_HT') : 22,
      marge: colIdx(lH, 'Marge', 'marge_amount', 'Margin') !== -1 ? colIdx(lH, 'Marge', 'marge_amount') : 25,
    };
    log(`L3M — pid:${L.pid}="${lH[L.pid]}", PV:${L.pv}="${lH[L.pv]}", IS:${L.is_}="${lH[L.is_]}", GMV:${L.gmv}="${lH[L.gmv]}"`);

    // Build L3M map — aggregate by ProductId (multiple rows per product)
    const l3mMap = new Map();
    for (const row of l3mRows) {
      const k = String(row[L.pid] || '').trim();
      if (!k) continue;
      const pv   = parseNum(row[L.pv]);
      const is_  = parseNum(row[L.is_]);
      const gmv  = parseNum(row[L.gmv]);
      const qty  = parseNum(row[L.qty]);
      const ca   = parseNum(row[L.ca]);
      const mAmt = parseNum(row[L.marge]);
      if (l3mMap.has(k)) {
        const e = l3mMap.get(k);
        e.pv += pv; e.is += is_; e.gmv += gmv; e.qty += qty; e.ca += ca; e.mAmt += mAmt;
      } else {
        l3mMap.set(k, { pv, is: is_, gmv, qty, ca, mAmt });
      }
    }
    // Post-compute ratios
    for (const [, v] of l3mMap) {
      v.margePct     = v.ca  > 0 ? v.mAmt / v.ca  : 0;
      v.coutUnitaire = v.qty > 0 ? (v.ca - v.mAmt) / v.qty : 0;
    }
    log(`L3M map: ${l3mMap.size.toLocaleString()} produits agrégés`);
    await yield_();

    // ── 4. Jumia ────────────────────────────────────────────
    setProgress(62, 'Lecture IP Jumia…');
    log(`Jumia: ${files.jumia.name}`);
    await yield_();
    const jumiaWb = await readWorkbook(files.jumia);
    const { headers: jH, rows: jumiaRows } = sheetToArrays(jumiaWb);
    log(`Jumia: ${jumiaRows.length.toLocaleString()} lignes`, 'info');

    // Jumia columns — from workflow: C=sku(key), K=prix_jumia, N=Lien
    const J = {
      pid:  colIdx(jH, 'sku', 'SKU', 'ProductId', 'mon_ean') !== -1 ? colIdx(jH, 'sku', 'SKU', 'ProductId') : 2,
      px:   colIdx(jH, 'prix_jumia', 'Prix_Jumia', 'price_jumia') !== -1 ? colIdx(jH, 'prix_jumia', 'Prix_Jumia') : 10,
      lien: colIdx(jH, 'Lien_du_produit', 'lien', 'link', 'url') !== -1 ? colIdx(jH, 'Lien_du_produit', 'lien', 'link') : 13,
      views:colIdx(jH, 'views', 'Views') !== -1 ? colIdx(jH, 'views', 'Views') : 12,
    };
    log(`Jumia — pid:${J.pid}="${jH[J.pid]}", prix:${J.px}="${jH[J.px]}", lien:${J.lien}="${jH[J.lien]}"`);

    const jumiaMap = new Map();
    for (const row of jumiaRows) {
      const k = String(row[J.pid] || '').trim();
      if (k && !jumiaMap.has(k)) jumiaMap.set(k, row);
    }
    log(`Jumia map: ${jumiaMap.size.toLocaleString()} entrées`);
    await yield_();

    // ── 5. Build enriched Retail rows (37 cols) ─────────────
    setProgress(75, 'Fusion et calcul des colonnes…');
    log('Fusion en cours…');
    await yield_();

    const retailRows = [];
    let matchC = 0, matchL = 0, matchJ = 0;

    for (const sRow of stockRows) {
      const pid  = String(sRow[S.pid] || '').trim();
      const cRow = consoMap.get(pid);  if (cRow) matchC++;
      const lDat = l3mMap.get(pid);    if (lDat) matchL++;
      const jRow = jumiaMap.get(pid);  if (jRow) matchJ++;

      // ── Stock base ──
      const rawCat  = String(sRow[S.cat] || '');
      const n1      = cRow ? String(cRow[C.n1] || '')    : '';
      const n2      = cRow ? String(cRow[C.n2] || '')    : '';
      const n3      = cRow ? String(cRow[C.n3] || '')    : '';
      const marque  = cRow ? String(cRow[C.brand] || '') : '';
      const prixLive= cRow ? parseNum(cRow[C.price])     : 0;
      const vendeurBO=cRow ? String(cRow[C.shop] || '')  : '';
      const typeV   = String(sRow[S.typeV] || '');
      const stock   = parseNum(sRow[S.stock]);
      const age     = parseNum(sRow[S.age]);
      const valeur  = parseNum(sRow[S.valeur]);

      // ── L3M ──
      const pv      = lDat ? lDat.pv  : 0;
      const is_     = lDat ? lDat.is  : 0;
      const gmv     = lDat ? lDat.gmv : 0;
      const margePct= lDat ? roundN(lDat.margePct * 100, 2)  : 0;
      const coutU   = lDat ? roundN(lDat.coutUnitaire, 2)    : 0;

      // ── Jumia ──
      const prixJ   = jRow ? parseNum(jRow[J.px])           : 0;
      const lienJ   = jRow ? String(jRow[J.lien] || '')     : '';

      // ── Computed ──
      const moyVente = is_ > 0 ? roundN(is_ / 90, 3) : 0;
      const couv     = moyVente > 0 ? roundN(stock / moyVente, 0) : 0;
      const cr       = pv  > 0 ? roundN(is_ / pv, 4)  : 0;
      const asp      = is_ > 0 ? roundN(gmv / is_, 2)  : 0;
      const margeLive= (prixLive > 0 && coutU > 0) ? roundN((prixLive/1.2 - coutU) / (prixLive/1.2), 4) : '';
      const catRevue = mapCategory(n1, rawCat);
      const zone     = detectZone(sRow[S.zone], sRow[S.empl]);

      // ── Alert columns ──
      const checkPrix= (prixJ > 0 && prixLive > 0 && prixJ < prixLive) ? 'prix à revoir' : '';
      const checkBO  = (typeV === 'Retail' && vendeurBO && vendeurBO !== 'Marjanemall') ? 'Retail Non BO' : '';
      const animCom  = computeAnimation(age, couv, cr);

      retailRows.push({
        // A-H: base
        'Catégorie Revue':      catRevue,
        'Categorie':            rawCat,
        'Type de Stock':        typeV,
        'Type':                 '',          // sourcing type — needs Repartition sku 1P
        'Owner':                '',
        'ProductId':            pid,
        'gtin':                 String(sRow[S.gtin] || ''),
        'title':                String(sRow[S.title] || ''),
        // I-L: Conso (yellow)
        'Marque':               marque,
        'N1':                   n1,
        'N2':                   n2,
        'N3':                   n3,
        // M-S: Stock metrics
        'Stock':                stock,
        'Age stock Moyen':      age,
        'Tranche Age':          trancheAge(age),
        'VALEUR PV HT':         valeur,
        'Moyenne de vente':     moyVente,
        'Couverture (Jrs)':     couv,
        'Tranche de couverture':trancheCoverage(couv),
        // T-W: L3M views/CR
        'PV L3M':               pv,
        'Tranche de PV':        tranchePV(pv),
        'CR L3M':               cr,
        'Tranche de CR':        trancheCR(cr),
        // X-AA: L3M sales (yellow)
        'IS L3M':               is_,
        'GMV L3M':              roundN(gmv, 2),
        'Marge L3M':            margePct,
        'Cout unitaire':        coutU,
        // AB-AE: computed
        'ASP L3M':              asp,
        'Prix Live':            prixLive,
        'Marge live':           margeLive,
        'Statut marge':         '',          // needs Repartition sku 1P for sourcing type
        // AF-AK
        'Prix Jumia':           prixJ,
        'Check Prix':           checkPrix,
        'Lien Jumia':           lienJ,
        'Vendeur BO':           vendeurBO,
        'Check BO':             checkBO,
        'Animation commerciale':animCom,
        // Internal (for Vue globale computation, stripped from sheet)
        _zone:    zone,
        _typeV:   typeV,
        _catRevue:catRevue,
        _pv:      pv,
        _age:     age,
        _couv:    couv,
        _prixLive:prixLive,
        _prixJ:   prixJ,
        _vendeurBO:vendeurBO,
        _checkBO: checkBO,
        _checkPrix:checkPrix,
      });
    }
    log(`Fusion: ${retailRows.length.toLocaleString()} lignes`, 'info');
    log(`  Conso: ${pct(matchC, retailRows.length)} | L3M: ${pct(matchL, retailRows.length)} | Jumia: ${pct(matchJ, retailRows.length)}`);
    await yield_();

    // ── 6. Build output workbook ─────────────────────────────
    setProgress(88, 'Génération du classeur Excel…');
    log('Construction Vue globale…');
    await yield_();

    const wb = XLSX.utils.book_new();

    // Sheet 1 — Vue globale (exact layout)
    const vgSheet = buildVueGlobaleSheet(retailRows, weekNum, year);
    XLSX.utils.book_append_sheet(wb, vgSheet, 'Vue globale');

    // Sheet 2 — Report Retail (37 cols, internal _ fields stripped)
    log('Construction Report Retail…');
    await yield_();
    const cleanRows = retailRows.map(r => {
      const out = {};
      for (const [k, v] of Object.entries(r)) {
        if (!k.startsWith('_')) out[k] = v;
      }
      return out;
    });
    const retailSheet = XLSX.utils.json_to_sheet(cleanRows);
    retailSheet['!cols'] = buildRetailColWidths();
    XLSX.utils.book_append_sheet(wb, retailSheet, 'Report Retail');

    // Sheet 3 — Stock raw
    const stockSheet = XLSX.utils.json_to_sheet(
      stockRows.map(r => Object.fromEntries(sH.map((h, i) => [h || `Col${i}`, r[i]])))
    );
    XLSX.utils.book_append_sheet(wb, stockSheet, 'Stock');

    outputWorkbook = wb;
    outputWorkbook._filename = `Weekly_Report_Stock_S${weekNum}_${year}.xlsx`;
    setProgress(100, 'Rapport généré !');
    log(`✅ Fichier prêt: ${outputWorkbook._filename}`, 'info');

    // ── Show result ─────────────────────────────────────────
    document.getElementById('result-title').textContent = `Rapport S${weekNum}_${year} généré`;
    document.getElementById('result-summary').textContent =
      `${retailRows.length.toLocaleString()} produits · ` +
      `Conso ${pct(matchC, retailRows.length)} · ` +
      `L3M ${pct(matchL, retailRows.length)} · ` +
      `Jumia ${pct(matchJ, retailRows.length)}`;
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

// ── Animation commerciale (simplified) ──────────────────────
function computeAnimation(age, couv, cr) {
  if (age > 180 && couv > 120) return 'Baisse de prix permanente ou retour fournisseur';
  if (age > 60  && couv > 120) return 'Baisse de prix permanente';
  if (couv > 60 && cr > 0.05)  return 'Vente Flash';
  return '';
}

// ── Build Vue globale sheet ───────────────────────────────────
function buildVueGlobaleSheet(retailRows, weekNum, year) {
  // Only Retail rows (exclude FFM for Vue globale)
  const retail = retailRows.filter(r => r._typeV === 'Retail');
  // If no type vendeur distinction → use all
  const base   = retail.length > 0 ? retail : retailRows;

  // 6 sections: [name, filterFn]
  const SECTIONS = [
    ['Scope total Retail',                       r => true],
    ['Produits Non BO',                          r => !!r._checkBO],
    ['Produits KO Prix Jumia',                   r => !!r._checkPrix],
    ['Age stock >180 jours',                     r => r._age > 180],
    ['Couverture >120 jours et Age > 60 jours',  r => r._couv > 120 && r._age > 60],
    ['Low views (PV L3M <200)',                   r => r._pv < 200],
  ];

  // For each section + each category → compute [skuB1, stockB1, skuB2, stockB2, skuTotal, stockTotal]
  function aggCat(rows, filterFn, catName) {
    const filtered = rows.filter(filterFn);
    const catRows  = catName === 'Total général' ? filtered : filtered.filter(r => r._catRevue === catName);
    let skuB1=0, stB1=0, skuB2=0, stB2=0;
    for (const r of catRows) {
      if (r._zone === 'B1') { skuB1++; stB1 += parseNum(r['Stock']); }
      else                  { skuB2++; stB2 += parseNum(r['Stock']); }
    }
    return [skuB1, stB1, skuB2, stB2, skuB1+skuB2, stB1+stB2];
  }

  // Build AoA (array of arrays)
  const aoa = [];

  // Row 0 (Excel row 1): Section headers
  const row0 = new Array(37).fill('');
  row0[0]  = 'Catégorie';
  row0[1]  = SECTIONS[0][0];
  row0[7]  = SECTIONS[1][0];
  row0[13] = SECTIONS[2][0];
  row0[19] = SECTIONS[3][0];
  row0[25] = SECTIONS[4][0];
  row0[31] = SECTIONS[5][0];
  aoa.push(row0);

  // Row 1 (Excel row 2): B1 / B2 / Total sub-headers
  const row1 = new Array(37).fill('');
  for (let s = 0; s < 6; s++) {
    const offset = 1 + s * 6;
    row1[offset + 0] = 'B1';
    row1[offset + 2] = 'B2';
    row1[offset + 4] = 'Total';
  }
  aoa.push(row1);

  // Row 2 (Excel row 3): #SKU / Stock labels
  const row2 = new Array(37).fill('');
  row2[0] = 'Catégorie';
  for (let s = 0; s < 6; s++) {
    const offset = 1 + s * 6;
    row2[offset + 0] = '#SKU';
    row2[offset + 1] = 'Stock';
    row2[offset + 2] = '#SKU';
    row2[offset + 3] = 'Stock';
    row2[offset + 4] = '#SKU';
    row2[offset + 5] = 'Stock';
  }
  aoa.push(row2);

  // Data rows: one per category + Total général
  const allCats = [...CATEGORIES_ORDER, 'Total général'];
  for (const cat of allCats) {
    const dataRow = new Array(37).fill('');
    dataRow[0] = cat;
    for (let s = 0; s < 6; s++) {
      const [skuB1, stB1, skuB2, stB2, skuT, stT] = aggCat(base, SECTIONS[s][1], cat);
      const offset = 1 + s * 6;
      dataRow[offset + 0] = skuB1;
      dataRow[offset + 1] = stB1;
      dataRow[offset + 2] = skuB2;
      dataRow[offset + 3] = stB2;
      dataRow[offset + 4] = skuT;
      dataRow[offset + 5] = stT;
    }
    aoa.push(dataRow);
  }

  const ws = XLSX.utils.aoa_to_sheet(aoa);

  // Merged cells
  ws['!merges'] = [
    // Row 1: category header spans 3 rows
    { s:{r:0,c:0}, e:{r:2,c:0} },
    // Row 1: 6 section headers (each spans 6 cols)
    { s:{r:0,c:1},  e:{r:0,c:6}  },
    { s:{r:0,c:7},  e:{r:0,c:12} },
    { s:{r:0,c:13}, e:{r:0,c:18} },
    { s:{r:0,c:19}, e:{r:0,c:24} },
    { s:{r:0,c:25}, e:{r:0,c:30} },
    { s:{r:0,c:31}, e:{r:0,c:36} },
    // Row 2: B1/B2/Total sub-headers (each spans 2 cols)
    { s:{r:1,c:1},  e:{r:1,c:2}  }, { s:{r:1,c:3},  e:{r:1,c:4}  }, { s:{r:1,c:5},  e:{r:1,c:6}  },
    { s:{r:1,c:7},  e:{r:1,c:8}  }, { s:{r:1,c:9},  e:{r:1,c:10} }, { s:{r:1,c:11}, e:{r:1,c:12} },
    { s:{r:1,c:13}, e:{r:1,c:14} }, { s:{r:1,c:15}, e:{r:1,c:16} }, { s:{r:1,c:17}, e:{r:1,c:18} },
    { s:{r:1,c:19}, e:{r:1,c:20} }, { s:{r:1,c:21}, e:{r:1,c:22} }, { s:{r:1,c:23}, e:{r:1,c:24} },
    { s:{r:1,c:25}, e:{r:1,c:26} }, { s:{r:1,c:27}, e:{r:1,c:28} }, { s:{r:1,c:29}, e:{r:1,c:30} },
    { s:{r:1,c:31}, e:{r:1,c:32} }, { s:{r:1,c:33}, e:{r:1,c:34} }, { s:{r:1,c:35}, e:{r:1,c:36} },
  ];

  // Column widths
  ws['!cols'] = [
    { wch: 30 }, // A: Catégorie
    ...Array(36).fill({ wch: 10 }),
  ];
  ws['!cols'][0] = { wch: 30 };

  return ws;
}

// ── Report Retail column widths ───────────────────────────────
function buildRetailColWidths() {
  // A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, Q, R, S, T, U, V, W, X, Y, Z, AA, AB, AC, AD, AE, AF, AG, AH, AI, AJ, AK
  return [
    { wch: 22 }, // A Catégorie Revue
    { wch: 22 }, // B Categorie
    { wch: 10 }, // C Type de Stock
    { wch: 14 }, // D Type
    { wch: 10 }, // E Owner
    { wch: 16 }, // F ProductId
    { wch: 16 }, // G gtin
    { wch: 40 }, // H title
    { wch: 18 }, // I Marque
    { wch: 22 }, // J N1
    { wch: 22 }, // K N2
    { wch: 22 }, // L N3
    { wch: 10 }, // M Stock
    { wch: 14 }, // N Age stock Moyen
    { wch: 14 }, // O Tranche Age
    { wch: 14 }, // P VALEUR PV HT
    { wch: 14 }, // Q Moyenne de vente
    { wch: 14 }, // R Couverture (Jrs)
    { wch: 18 }, // S Tranche de couverture
    { wch: 10 }, // T PV L3M
    { wch: 12 }, // U Tranche de PV
    { wch: 10 }, // V CR L3M
    { wch: 12 }, // W Tranche de CR
    { wch: 10 }, // X IS L3M
    { wch: 12 }, // Y GMV L3M
    { wch: 12 }, // Z Marge L3M
    { wch: 12 }, // AA Cout unitaire
    { wch: 10 }, // AB ASP L3M
    { wch: 12 }, // AC Prix Live
    { wch: 12 }, // AD Marge live
    { wch: 12 }, // AE Statut marge
    { wch: 12 }, // AF Prix Jumia
    { wch: 16 }, // AG Check Prix
    { wch: 40 }, // AH Lien Jumia
    { wch: 20 }, // AI Vendeur BO
    { wch: 16 }, // AJ Check BO
    { wch: 40 }, // AK Animation commerciale
  ];
}

// ── Download ──────────────────────────────────────────────────
function downloadReport() {
  if (!outputWorkbook) return;
  XLSX.writeFile(outputWorkbook, outputWorkbook._filename);
}

// ── Reset ─────────────────────────────────────────────────────
function resetAll() {
  Object.keys(files).forEach(k => files[k] = null);
  document.querySelectorAll('.upload-zone').forEach(z => {
    z.classList.remove('done');
    z.querySelector('.file-name').textContent = '';
    z.querySelector('.file-input').value = '';
  });
  document.getElementById('progress-block').classList.add('hidden');
  document.getElementById('log-block').classList.add('hidden');
  document.getElementById('result-block').classList.add('hidden');
  document.getElementById('btn-process').disabled = true;
  outputWorkbook = null;
}
