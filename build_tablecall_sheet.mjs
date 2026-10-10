import fs from 'node:fs/promises';
import { SpreadsheetFile, Workbook } from '@oai/artifact-tool';

const outputDir = 'C:/Users/nongy/Downloads/TableCall-main/outputs/tablecall-google-sheet';
const outputPath = `${outputDir}/TableCall_Google_Sheets_Template.xlsx`;
const workbook = Workbook.create();

const menus = [
  ['m1', 'ข้าวกะเพราไก่', 'อาหารจานเดียว', 60, true, ''],
  ['m2', 'ข้าวกะเพราหมู', 'อาหารจานเดียว', 60, true, ''],
  ['m3', 'ข้าวผัดหมู', 'อาหารจานเดียว', 55, true, ''],
  ['m4', 'ข้าวผัดกุ้ง', 'อาหารจานเดียว', 85, true, ''],
  ['m5', 'ผัดไทยกุ้งสด', 'ก๋วยเตี๋ยว/เส้น', 90, true, ''],
  ['m6', 'ก๋วยเตี๋ยวต้มยำหมู', 'ก๋วยเตี๋ยว/เส้น', 55, true, ''],
  ['m7', 'ก๋วยเตี๋ยวเย็นตาโฟ', 'ก๋วยเตี๋ยว/เส้น', 70, true, ''],
  ['m8', 'ยำวุ้นเส้นทะเล', 'ยำ/น้ำพริก', 150, true, ''],
  ['m9', 'ส้มตำไทย', 'ยำ/น้ำพริก', 55, true, ''],
  ['m10', 'ปอเปี๊ยะทอด', 'ของทอด', 60, true, ''],
  ['m11', 'ปีกไก่ทอด', 'ของทอด', 120, true, ''],
  ['m12', 'บัวลอย', 'ของหวาน', 45, true, ''],
  ['m13', 'ไอติมกะทิ', 'ของหวาน', 45, true, ''],
  ['m14', 'น้ำเปล่า', 'เครื่องดื่ม', 15, true, ''],
  ['m15', 'โค้ก', 'เครื่องดื่ม', 25, true, ''],
  ['m16', 'ชาเย็น', 'เครื่องดื่ม', 35, true, ''],
  ['m17', 'น้ำมะนาว', 'เครื่องดื่ม', 35, true, ''],
  ['m18', 'เนื้อสไลซ์ (ชาบู)', 'หมูกระทะ-ชาบู', 99, true, ''],
  ['m19', 'หมูสไลซ์ (ชาบู)', 'หมูกระทะ-ชาบู', 79, true, ''],
  ['m20', 'ผักรวม', 'ผัก/เห็ด', 45, true, ''],
];

const tables = [
  ['S1', 'S', 4, true, ''], ['S2', 'S', 4, true, ''], ['S3', 'S', 4, true, ''], ['S4', 'S', 6, true, ''],
  ['S5', 'S', 4, true, ''], ['S6', 'S', 4, true, ''], ['S7', 'S', 4, true, ''], ['S8', 'S', 4, true, ''],
  ['A1', 'A', 4, true, ''], ['A2', 'A', 4, true, ''], ['A3', 'A', 2, true, ''], ['A4', 'A', 4, true, ''],
  ['A5', 'A', 6, true, ''], ['A6', 'A', 4, true, ''],
  ['B1', 'B', 4, true, ''], ['B2', 'B', 4, true, ''], ['B3', 'B', 4, true, ''], ['B4', 'B', 4, true, ''],
  ['VIP 1', 'VIP', 10, true, ''], ['VIP 2', 'VIP', 12, true, ''],
];

const sheetConfigs = [
  { name: 'Menu', headers: ['id', 'name', 'category', 'price', 'available', 'updated_at'], rows: menus, table: 'MenuCatalog', priceCol: 'D' },
  { name: 'POS_Tables', headers: ['number', 'zone', 'seats', 'active', 'updated_at'], rows: tables, table: 'POSTables' },
  { name: 'POSOrders', headers: ['id', 'table_number', 'items_json', 'note', 'total', 'status', 'created_at'], rows: [] },
  { name: 'POSState', headers: ['key', 'state_json', 'updated_at'], rows: [] },
  { name: 'Tables', headers: ['id', 'number', 'seats', 'qr_token', 'is_active', 'sort_order', 'created_at', 'updated_at'], rows: [] },
  { name: 'Requests', headers: ['id', 'table_id', 'table_number', 'type', 'status', 'created_at', 'handled_at', 'handled_by', 'note'], rows: [] },
  { name: 'Feedback', headers: ['id', 'table_id', 'table_number', 'rating', 'suggestions', 'created_at'], rows: [] },
];

for (const config of sheetConfigs) {
  const sheet = workbook.worksheets.add(config.name);
  sheet.showGridLines = false;
  const data = [config.headers, ...config.rows];
  const lastCol = String.fromCharCode(64 + config.headers.length);
  const range = sheet.getRange(`A1:${lastCol}${data.length}`);
  range.values = data;
  const header = sheet.getRange(`A1:${lastCol}1`);
  header.format.fill = '#3730A3';
  header.format.font = { bold: true, color: '#FFFFFF', size: 11 };
  header.format.rowHeight = 26;
  header.format.verticalAlignment = 'center';
  range.format.autofitColumns();
  range.format.autofitRows();
  sheet.freezePanes.freezeRows(1);
  if (config.priceCol && config.rows.length) sheet.getRange(`${config.priceCol}2:${config.priceCol}${data.length}`).format.numberFormat = '#,##0';
  if (config.table && config.rows.length) sheet.tables.add(`A1:${lastCol}${data.length}`, true, config.table);
}

workbook.recalculate();
const summary = await workbook.inspect({ kind: 'sheet,table', maxChars: 3000, tableMaxRows: 3, tableMaxCols: 6 });
console.log(summary.ndjson);
const menuPreview = await workbook.render({ sheetName: 'Menu', range: 'A1:F12', scale: 1, format: 'png' });
await fs.mkdir(outputDir, { recursive: true });
await fs.writeFile(`${outputDir}/Menu_preview.png`, new Uint8Array(await menuPreview.arrayBuffer()));
const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(outputPath);
console.log(`Saved ${outputPath}`);
