import XLSX from 'xlsx-js-style';

export function getFirstRow(filePath: string) {
  const wb = XLSX.readFile(filePath);
  const sheetName = wb.SheetNames[0];
  const sheet = wb.Sheets[sheetName];
  const firstRow = XLSX.utils.sheet_to_json(sheet)[0];
  return Object.keys(firstRow as object);
}
