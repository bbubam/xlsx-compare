const XLSX = require('xlsx');

function compareExcelFiles(file1, file2) {
  const wb1 = XLSX.readFile(file1);
  const wb2 = XLSX.readFile(file2);
  
  const sheet1 = wb1.Sheets[wb1.SheetNames[0]];
  const sheet2 = wb2.Sheets[wb2.SheetNames[0]];
  
  const data1 = XLSX.utils.sheet_to_json(sheet1);
  const data2 = XLSX.utils.sheet_to_json(sheet2);
  
  const differences = [];
  
  const maxRows = Math.max(data1.length, data2.length);
  for (let i = 0; i < maxRows; i++) {
    const row1 = data1[i] || {};
    const row2 = data2[i] || {};
    
    const allKeys = new Set([...Object.keys(row1), ...Object.keys(row2)]);
    
    for (const key of allKeys) {
      if (row1[key] !== row2[key]) {
        differences.push({
          row: i + 1,
          column: key,
          file1: row1[key],
          file2: row2[key]
        });
      }
    }
  }
  
  return differences;
}

const file1 = process.argv[2];
const file2 = process.argv[3];

if (!file1 || !file2) {
  console.error('Usage: node compare.js <file1> <file2>');
  process.exit(1);
}

const diffs = compareExcelFiles(file1, file2);

if (diffs.length === 0) {
  console.log('Files are identical');
} else {
  console.log(`Found ${diffs.length} differences:\n`);
  diffs.forEach(diff => {
    console.log(`Row ${diff.row}, Column "${diff.column}": "${diff.file1}" vs "${diff.file2}"`);
  });
}
