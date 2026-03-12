/* ============================================================
   Weekly Report Generator — Marjane Mall
   Pure browser-side processing with SheetJS
   ============================================================ */

// ── State ───────────────────────────────────────────────────
const files = { stock: null, conso: null, l3m: null, jumia: null, template: null };
let outputWorkbook = null;

// ── DOM ready ────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('footer-year').textContent = new Date().getFullYear();

  // Set today's date as default
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
      e.preventDefault();
      zone.classList.remove('drag-over');
      handleFile(e.dataTransfer.files[0], key, zone, label);
    });
  });
}

function handleFile(file, key, zone, label) {
  if (!file) return;
  files[key] = file;
  label.textContent = file.name;
  zone.classList.add('done');
  checkReady();
}

function checkReady() {
  const weekOk = !!document.getElementById('weekNum').value;
  const filesOk = files.stock && files.conso && files.l3m && files.jumia;
  document.getElementById('btn-process').disabled = !(weekOk && filesOk);
}
['weekNum', 'year', 'reportDate'].forEach(id =>
  document.getElementById(id).addEventListener('input', checkReady)
);

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

// ── Read file as ArrayBuffer → SheetJS workbook ───────────────
function readWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'array', cellDates: true });
        resolve(wb);
      } catch (err) { reject(err); }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

// ── Sheet → JSON (first sheet by default) ─────────────────────
function sheetToJson(wb, sheetName) {
  const name = sheetName || wb.SheetNames[0];
  const ws   = wb.Sheets[name];
  if (!ws) throw new Error(`Feuille "${name}" introuvable dans le fichier.`);
  return XLSX.utils.sheet_to_json(ws, { defval: '' });
}

// ── Detect column name case-insensitively ─────────────────────
function findCol(row, candidates) {
  const keys = Object.keys(row);
  for (const c of candidates) {
    const found = keys.find(k => k.trim().toLowerCase() === c.toLowerCase());
    if (found) return found;
  }
  // Partial match fallback
  for (const c of candidates) {
    const found = keys.find(k => k.trim().toLowerCase().includes(c.toLowerCase()));
    if (found) return found;
  }
  return null;
}

// ── Build lookup map from array ───────────────────────────────
function buildMap(rows, keyCol) {
  const map = new Map();
  for (const row of rows) {
    const k = String(row[keyCol] || '').trim();
    if (k) map.set(k, row);
  }
  return map;
}

// ── Yield to UI (prevent freeze) ─────────────────────────────
const yield_ = () => new Promise(r => setTimeout(r, 0));

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

    // ── 1. Parse Stock ──────────────────────────────────────
    setProgress(5, 'Lecture du fichier Stock…');
    log(`Lecture Stock: ${files.stock.name}`);
    await yield_();
    const stockWb   = await readWorkbook(files.stock);
    const stockRows = sheetToJson(stockWb);
    log(`Stock: ${stockRows.length.toLocaleString()} lignes chargées`, 'info');
    await yield_();

    // Detect Stock key column
    const sampleStock = stockRows[0] || {};
    const stockIdCol  = findCol(sampleStock, ['ProductId', 'product_id', 'SKU', 'Id']) || Object.keys(sampleStock)[1];
    log(`Stock — colonne clé détectée: "${stockIdCol}"`);

    // ── 2. Parse Conso ──────────────────────────────────────
    setProgress(20, 'Lecture du fichier Conso (large fichier, patience…)');
    log(`Lecture Conso: ${files.conso.name}`);
    await yield_();
    const consoWb   = await readWorkbook(files.conso);
    const consoRows = sheetToJson(consoWb);
    log(`Conso: ${consoRows.length.toLocaleString()} lignes chargées`, 'info');
    await yield_();

    // Detect Conso columns
    const sampleConso = consoRows[0] || {};
    const consoIdCol  = findCol(sampleConso, ['ProductId', 'product_id', 'Id', 'SKU', 'Sku']) || Object.keys(sampleConso)[3];
    const consoN1Col  = findCol(sampleConso, ['Label_N1', 'N1', 'Categorie_N1', 'label_n1']);
    const consoN2Col  = findCol(sampleConso, ['Label_N2', 'N2', 'label_n2']);
    const consoN3Col  = findCol(sampleConso, ['Label_N3', 'N3', 'label_n3']);
    const consoBrandCol   = findCol(sampleConso, ['Brandlabel', 'Brand', 'Marque', 'brandlabel', 'brand_label']);
    const consoPriceCol   = findCol(sampleConso, ['OfferPrice', 'Prix', 'Price', 'price', 'offerprice']);
    const consoVendeurCol = findCol(sampleConso, ['Shopname', 'Vendeur', 'seller', 'shop_name', 'shopname']);
    log(`Conso — clé: "${consoIdCol}", N1: "${consoN1Col}", Marque: "${consoBrandCol}", Prix: "${consoPriceCol}"`);

    const consoMap = buildMap(consoRows, consoIdCol);
    log(`Conso map: ${consoMap.size.toLocaleString()} entrées indexées`);
    await yield_();

    // ── 3. Parse L3M ───────────────────────────────────────
    setProgress(45, 'Lecture du fichier L3M (large fichier, patience…)');
    log(`Lecture L3M: ${files.l3m.name}`);
    await yield_();
    const l3mWb   = await readWorkbook(files.l3m);
    const l3mRows = sheetToJson(l3mWb);
    log(`L3M: ${l3mRows.length.toLocaleString()} lignes chargées`, 'info');
    await yield_();

    // Detect L3M columns
    const sampleL3m = l3mRows[0] || {};
    const l3mIdCol  = findCol(sampleL3m, ['ProductId', 'product_id', 'SKU', 'sku', 'Id']) || Object.keys(sampleL3m)[0];
    const l3mPVCol  = findCol(sampleL3m, ['Views', 'PV', 'page_views', 'pageviews']) || Object.keys(sampleL3m)[4];
    const l3mISCol  = findCol(sampleL3m, ['Items', 'IS', 'items_sold', 'quantity_sold']) || Object.keys(sampleL3m)[5];
    const l3mGMVCol = findCol(sampleL3m, ['CA', 'GMV', 'Revenue', 'revenue', 'ca_all', 'CA ALL']) || Object.keys(sampleL3m)[6];
    const l3mQtyCol = findCol(sampleL3m, ['Quantite', 'Quantity', 'Qte', 'qty']) || Object.keys(sampleL3m)[16];
    const l3mCACol  = findCol(sampleL3m, ['CA_total', 'CA Retail', 'montant_ca']) || Object.keys(sampleL3m)[22];
    const l3mMargeAmtCol = findCol(sampleL3m, ['Marge', 'marge_amount', 'Margin', 'margin']) || Object.keys(sampleL3m)[25];
    log(`L3M — clé: "${l3mIdCol}", PV: "${l3mPVCol}", IS: "${l3mISCol}", GMV: "${l3mGMVCol}"`);

    // Build L3M aggregated map (sum by ProductId since there may be multiple rows per product)
    const l3mMap = new Map();
    for (const row of l3mRows) {
      const k = String(row[l3mIdCol] || '').trim();
      if (!k) continue;

      const pv  = parseNum(row[l3mPVCol]);
      const is_ = parseNum(row[l3mISCol]);
      const gmv = parseNum(row[l3mGMVCol]);
      const qty = parseNum(row[l3mQtyCol]);
      const ca  = parseNum(row[l3mCACol]);
      const mAmt = parseNum(row[l3mMargeAmtCol]);

      if (l3mMap.has(k)) {
        const e = l3mMap.get(k);
        e.pv  += pv;  e.is += is_;  e.gmv += gmv;
        e.qty += qty; e.ca += ca;   e.margeAmt += mAmt;
      } else {
        l3mMap.set(k, { pv, is: is_, gmv, qty, ca, margeAmt: mAmt });
      }
    }

    // Compute Marge% and Coût unitaire
    for (const [, v] of l3mMap) {
      v.margePct     = v.ca   > 0 ? v.margeAmt / v.ca   : 0;
      v.coutUnitaire = v.qty  > 0 ? (v.ca - v.margeAmt) / v.qty : 0;
    }
    log(`L3M map: ${l3mMap.size.toLocaleString()} produits agrégés`);
    await yield_();

    // ── 4. Parse Jumia ──────────────────────────────────────
    setProgress(65, 'Lecture du fichier IP Jumia…');
    log(`Lecture IP Jumia: ${files.jumia.name}`);
    await yield_();
    const jumiaWb   = await readWorkbook(files.jumia);
    const jumiaRows = sheetToJson(jumiaWb);
    log(`Jumia: ${jumiaRows.length.toLocaleString()} lignes chargées`, 'info');
    await yield_();

    // Detect Jumia columns
    const sampleJumia  = jumiaRows[0] || {};
    const jumiaIdCol   = findCol(sampleJumia, ['sku', 'SKU', 'ProductId', 'mon_ean']) || Object.keys(sampleJumia)[2];
    const jumiaPxCol   = findCol(sampleJumia, ['prix_jumia', 'prix_Jumia', 'Prix_Jumia', 'price_jumia']);
    const jumiaLinkCol = findCol(sampleJumia, ['Lien_du_produit', 'lien', 'link', 'url', 'Lien']);
    const jumiaViewsCol = findCol(sampleJumia, ['views', 'Views']);
    log(`Jumia — clé: "${jumiaIdCol}", Prix: "${jumiaPxCol}", Lien: "${jumiaLinkCol}"`);

    const jumiaMap = buildMap(jumiaRows, jumiaIdCol);
    log(`Jumia map: ${jumiaMap.size.toLocaleString()} entrées indexées`);
    await yield_();

    // ── 5. Merge data (stock as base) ───────────────────────
    setProgress(78, 'Fusion des données…');
    log('Fusion en cours...');
    await yield_();

    const retailRows = [];
    let matchConso = 0, matchL3m = 0, matchJumia = 0;

    for (const sRow of stockRows) {
      const pid = String(sRow[stockIdCol] || '').trim();

      // Conso enrichment
      const cRow = consoMap.get(pid) || {};
      if (pid && consoMap.has(pid)) matchConso++;

      // L3M enrichment
      const lData = l3mMap.get(pid) || {};
      if (pid && l3mMap.has(pid)) matchL3m++;

      // Jumia enrichment
      const jRow = jumiaMap.get(pid) || {};
      if (pid && jumiaMap.has(pid)) matchJumia++;

      retailRows.push({
        // ── Stock base columns ──
        Categorie:        sRow['nom_categorie']  || sRow['nom_cat'] || '',
        ProductId:        pid,
        GTIN:             sRow['gtin']            || sRow['GTIN']    || '',
        Statut:           sRow['statut']          || '',
        Seller:           sRow['seller']          || '',
        'Type Vendeur':   sRow['type vendeur']    || sRow['type_vendeur'] || '',
        SellerId:         sRow['sellerid']        || sRow['seller_id'] || '',
        'Stock Dispo':    sRow['stock_dispo']     || sRow['stock']   || 0,
        Titre:            sRow['title']           || sRow['nom_du_produit'] || '',
        'Date Réception': formatDate(sRow['date_recep'] || sRow['date_reception']),
        'Age Stock (j)':  parseNum(sRow['age_stock']),
        Valeur:           parseNum(sRow['Valeur']  || sRow['valeur']),
        Quantite:         parseNum(sRow['Quantite'] || sRow['quantite']),
        'Tranche Age':    sRow['tranche_age']     || '',
        'Type Empl':      sRow['Type Empl']       || sRow['type_empl'] || '',
        'Zone Stockage':  sRow['Zone Stockage']   || sRow['zone_stockage'] || '',

        // ── Conso enrichment ──
        Marque:           cRow[consoBrandCol]   || '',
        N1:               cRow[consoN1Col]      || '',
        N2:               cRow[consoN2Col]      || '',
        N3:               cRow[consoN3Col]      || '',
        'Prix Live':      parseNum(cRow[consoPriceCol] || 0),
        'Vendeur BO':     cRow[consoVendeurCol] || '',

        // ── L3M enrichment ──
        'PV L3M':         lData.pv            || 0,
        'IS L3M':         lData.is            || 0,
        'GMV L3M':        roundN(lData.gmv    || 0, 2),
        'Marge L3M (%)':  roundN((lData.margePct || 0) * 100, 2),
        'Coût Unitaire':  roundN(lData.coutUnitaire || 0, 2),

        // ── Jumia enrichment ──
        'Prix Jumia':     parseNum(jRow[jumiaPxCol]   || 0),
        'Lien Jumia':     jRow[jumiaLinkCol]           || '',
        'Views Jumia':    parseNum(jRow[jumiaViewsCol] || 0),
      });
    }

    log(`✅ Fusion terminée: ${retailRows.length.toLocaleString()} lignes`);
    log(`   Conso matchés: ${matchConso} / ${stockRows.length}`, matchConso < stockRows.length * 0.5 ? 'warn' : 'info');
    log(`   L3M matchés:   ${matchL3m}  / ${stockRows.length}`,  matchL3m  < stockRows.length * 0.5 ? 'warn' : 'info');
    log(`   Jumia matchés: ${matchJumia} / ${stockRows.length}`, matchJumia < stockRows.length * 0.5 ? 'warn' : 'info');
    await yield_();

    // ── 6. Build Vue globale summary ────────────────────────
    setProgress(88, 'Calcul des KPIs Vue globale…');
    await yield_();
    const summary = buildSummary(retailRows, weekNum, year);

    // ── 7. Generate Excel workbook ──────────────────────────
    setProgress(93, 'Génération du fichier Excel…');
    log('Construction du classeur Excel…');
    await yield_();

    const wb = XLSX.utils.book_new();

    // Sheet 1 — Vue globale
    const summarySheet = XLSX.utils.json_to_sheet(summary);
    styleHeaderRow(summarySheet, summary[0]);
    XLSX.utils.book_append_sheet(wb, summarySheet, 'Vue globale');

    // Sheet 2 — Retail (main enriched data)
    const retailSheet = XLSX.utils.json_to_sheet(retailRows);
    XLSX.utils.book_append_sheet(wb, retailSheet, 'Retail');

    // Sheet 3 — Stock raw
    const stockSheet = XLSX.utils.json_to_sheet(stockRows);
    XLSX.utils.book_append_sheet(wb, stockSheet, 'Stock');

    // Sheet 4 — Jumia raw
    const jumiaSheet = XLSX.utils.json_to_sheet(jumiaRows);
    XLSX.utils.book_append_sheet(wb, jumiaSheet, 'IP Jumia');

    outputWorkbook = wb;
    outputWorkbook._filename = `Weekly_Report_Stock_S${weekNum}_${year}.xlsx`;

    setProgress(100, 'Rapport généré avec succès !');
    log(`Rapport prêt : ${outputWorkbook._filename}`, 'info');

    // ── Show result ─────────────────────────────────────────
    document.getElementById('result-title').textContent =
      `Rapport S${weekNum}_${year} généré avec succès`;
    document.getElementById('result-summary').textContent =
      `${retailRows.length.toLocaleString()} produits · ` +
      `Conso ${pct(matchConso, retailRows.length)} · ` +
      `L3M ${pct(matchL3m, retailRows.length)} · ` +
      `Jumia ${pct(matchJumia, retailRows.length)}`;

    document.getElementById('result-block').classList.remove('hidden');

  } catch (err) {
    log(`ERREUR: ${err.message}`, 'error');
    console.error(err);
    setProgress(0, 'Erreur lors du traitement.');
    const msg = document.getElementById('validation-msg');
    msg.textContent = `Erreur: ${err.message}`;
    msg.classList.remove('hidden');
  } finally {
    btn.disabled = false;
  }
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

// ── Build Vue globale summary ─────────────────────────────────
function buildSummary(rows, weekNum, year) {
  const totalSKU   = rows.length;
  const retailRows = rows.filter(r => r['Type Vendeur'] === 'Retail');
  const ffmRows    = rows.filter(r => r['Type Vendeur'] === 'FFM');

  const sumGMV     = rows.reduce((s, r) => s + (r['GMV L3M']    || 0), 0);
  const sumStock   = rows.reduce((s, r) => s + (r['Stock Dispo'] || 0), 0);
  const sumValeur  = rows.reduce((s, r) => s + (r['Valeur']      || 0), 0);
  const avgMarge   = rows.filter(r => r['Marge L3M (%)'] > 0)
                         .reduce((s, r, _, a) => s + r['Marge L3M (%)'] / a.length, 0);

  const tranche0_30 = rows.filter(r => r['Tranche Age'] === '[0-30]').length;
  const tranche30_60 = rows.filter(r => r['Tranche Age'] === ']30-60]').length;
  const trancheMore90 = rows.filter(r => r['Tranche Age'] === '>90').length;

  return [
    { 'Indicateur': 'Semaine',             'Valeur': `S${weekNum}_${year}` },
    { 'Indicateur': '#SKU Total',          'Valeur': totalSKU },
    { 'Indicateur': '#SKU Retail',         'Valeur': retailRows.length },
    { 'Indicateur': '#SKU FFM',            'Valeur': ffmRows.length },
    { 'Indicateur': 'Stock disponible total', 'Valeur': sumStock },
    { 'Indicateur': 'Valeur stock (MAD)',  'Valeur': roundN(sumValeur, 0) },
    { 'Indicateur': 'GMV L3M total (MAD)', 'Valeur': roundN(sumGMV, 0) },
    { 'Indicateur': 'Marge moyenne L3M (%)', 'Valeur': roundN(avgMarge, 2) },
    { 'Indicateur': 'Age stock [0-30j]',   'Valeur': tranche0_30 },
    { 'Indicateur': 'Age stock ]30-60j]',  'Valeur': tranche30_60 },
    { 'Indicateur': 'Age stock >90j',      'Valeur': trancheMore90 },
  ];
}

// ── Helpers ───────────────────────────────────────────────────
function parseNum(v) {
  if (v === '' || v === null || v === undefined) return 0;
  const n = parseFloat(String(v).replace(',', '.').replace(/\s/g, ''));
  return isNaN(n) ? 0 : n;
}

function roundN(n, decimals) {
  const factor = Math.pow(10, decimals);
  return Math.round(n * factor) / factor;
}

function formatDate(v) {
  if (!v) return '';
  if (v instanceof Date) return v.toLocaleDateString('fr-FR');
  return String(v);
}

function pct(a, b) {
  if (!b) return '0%';
  return Math.round((a / b) * 100) + '%';
}

function styleHeaderRow(ws, sampleRow) {
  // SheetJS doesn't easily support full styling in xlsx format without
  // commercial plugins. Headers are set by json_to_sheet automatically.
  // Column widths: set reasonable defaults
  const keys = Object.keys(sampleRow || {});
  ws['!cols'] = keys.map(() => ({ wch: 18 }));
}
