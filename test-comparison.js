const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

async function readWB(filePath) {
  const buf = fs.readFileSync(filePath);
  return XLSX.read(buf, { type: 'buffer', cellFormula: true });
}

function wsToRows(ws) {
  if (!ws['!ref']) return [];
  const rng = XLSX.utils.decode_range(ws['!ref']), rows = [];
  const mergeMap = {};
  (ws['!merges'] || []).forEach(m => {
    const val = ws[XLSX.utils.encode_cell({ r: m.s.r, c: m.s.c })];
    const v = val ? (val.f ? '=' + val.f : (val.v !== undefined ? String(val.v) : '')) : '';
    for (let r = m.s.r; r <= m.e.r; r++) for (let c = m.s.c; c <= m.e.c; c++) mergeMap[r + ',' + c] = v;
  });
  for (let r = rng.s.r; r <= rng.e.r; r++) {
    const row = [];
    for (let c = rng.s.c; c <= rng.e.c; c++) {
      const key = r + ',' + c;
      if (key in mergeMap) { row.push(mergeMap[key]); continue; }
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      row.push(cell ? (cell.f ? '=' + cell.f : (cell.v !== undefined ? String(cell.v) : '')) : '');
    }
    rows.push(row);
  }
  return rows;
}

function detectKey(rows) {
  if (rows.length <= 1) return null;
  const data = rows.slice(1);
  const numCols = data[0]?.length || 0;
  for (let c = 0; c < numCols; c++) {
    const vals = data.map(r => r[c] ?? '');
    if (new Set(vals).size === vals.length) return [c];
  }
  for (let a = 0; a < numCols; a++) for (let b = a + 1; b < numCols; b++) {
    const vals = data.map(r => (r[a] ?? '') + '|' + (r[b] ?? ''));
    if (new Set(vals).size === vals.length) return [a, b];
  }
  return null;
}

function rowKey(row, keyCols) {
  if (!keyCols) return row.join('\x00');
  return keyCols.map(c => row[c] ?? '').join('\x00');
}

async function compare() {
  const fileA = 'C:\\Benny\\Playground\\xlsx-compare\\samples\\Enrollment Report.xlsx';
  const fileB = 'C:\\Benny\\Playground\\xlsx-compare\\samples\\Enrollment Report Modified.xlsx';

  const wbA = await readWB(fileA);
  const wbB = await readWB(fileB);

  console.log('File A sheets:', wbA.SheetNames);
  console.log('File B sheets:', wbB.SheetNames);
  console.log('');

  const sA = new Set(wbA.SheetNames), sB = new Set(wbB.SheetNames);
  const allSheets = [...new Set([...wbA.SheetNames, ...wbB.SheetNames])];

  allSheets.forEach(name => {
    console.log(`\n=== Sheet: ${name} ===`);
    const rowsA = sA.has(name) ? wsToRows(wbA.Sheets[name]) : [];
    const rowsB = sB.has(name) ? wsToRows(wbB.Sheets[name]) : [];

    console.log(`Rows in A: ${rowsA.length}, Rows in B: ${rowsB.length}`);

    if (rowsA.length === 0) {
      console.log('Sheet added in B');
      return;
    }
    if (rowsB.length === 0) {
      console.log('Sheet deleted (only in A)');
      return;
    }

    const headersA = rowsA[0] || [];
    const headersB = rowsB[0] || [];

    console.log('Headers A:', headersA);
    console.log('Headers B:', headersB);

    const colDefs = [];
    headersA.forEach((h, i) => colDefs.push({ name: h, idxA: i, idxB: -1 }));
    headersB.forEach((h, bi) => {
      const existing = colDefs.find(d => d.name === h && d.idxB === -1 && d.idxA !== -1);
      if (existing) existing.idxB = bi;
      else colDefs.push({ name: h, idxA: -1, idxB: bi });
    });

    const headers = colDefs.map(d => d.name);
    const colMapA = colDefs.map(d => d.idxA);
    const colMapB = colDefs.map(d => d.idxB);

    console.log('Unified headers:', headers);
    console.log('Column mapping A:', colMapA);
    console.log('Column mapping B:', colMapB);

    const unify = (raw, colMap) => colMap.map(i => i === -1 ? '' : (raw[i] ?? ''));
    const sharedUniCols = headers.map((_, c) => c).filter(c => colMapA[c] !== -1 && colMapB[c] !== -1);
    const sharedRowsA = [headersA, ...rowsA.slice(1).map(r => sharedUniCols.map(c => r[colMapA[c]] ?? ''))];
    const keyCols = detectKey(sharedRowsA);
    const keyColsUnified = keyCols ? keyCols.map(k => sharedUniCols[k]) : null;

    console.log('Detected key columns:', keyCols);
    console.log('Key columns unified:', keyColsUnified);

    const mapA = new Map(), mapB = new Map();
    for (let i = 1; i < rowsA.length; i++) {
      const uRow = unify(rowsA[i], colMapA);
      const k = rowKey(uRow, keyColsUnified);
      if (!mapA.has(k)) mapA.set(k, { row: rowsA[i] });
    }
    for (let i = 1; i < rowsB.length; i++) {
      const uRow = unify(rowsB[i], colMapB);
      const k = rowKey(uRow, keyColsUnified);
      if (!mapB.has(k)) mapB.set(k, { row: rowsB[i] });
    }

    let added = 0, deleted = 0, modified = 0;
    mapA.forEach(({ row: rawA }, k) => {
      if (mapB.has(k)) {
        const rowA = unify(rawA, colMapA), rowB = unify(mapB.get(k).row, colMapB);
        const changedCols = [];
        for (let c = 0; c < headers.length; c++) {
          if (rowA[c] !== rowB[c]) changedCols.push(c);
        }
        if (changedCols.length) {
          modified++;
          console.log(`Modified row (key: ${k}):`, changedCols.map(c => `${headers[c]}: "${rowA[c]}" → "${rowB[c]}"`).join(', '));
        }
      } else {
        deleted++;
        console.log(`Deleted row (key: ${k})`);
      }
    });
    mapB.forEach(({ row: rawB }, k) => {
      if (!mapA.has(k)) {
        added++;
        console.log(`Added row (key: ${k})`);
      }
    });

    console.log(`Summary: ${added} added, ${deleted} deleted, ${modified} modified`);
  });
}

compare().catch(console.error);
