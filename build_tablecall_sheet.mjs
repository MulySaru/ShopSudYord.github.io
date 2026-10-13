import fs from 'node:fs/promises';
import { SpreadsheetFile, Workbook } from '@oai/artifact-tool';

const outputDir = 'C:/Users/nongy/Downloads/TableCall-main/outputs/rong-rua-menu-sheet';
const outputPath = `${outputDir}/RongRua_Menu_for_Google_Sheets.xlsx`;
const workbook = Workbook.create();

const soupOptions = ['น้ำตกเข้มข้น', 'ต้มยำ', 'น้ำตกต้มยำ', 'น้ำข้น (ไม่มีเลือด)', 'หม่าล่า']
  .map((name, index) => ({ id: `soup_${index + 1}`, name, available: true }));
const noodleOptions = ['เส้นเล็ก', 'เส้นใหญ่', 'หมี่ขาว', 'บะหมี่', 'วุ้นเส้น', 'มาม่า', 'เกาเหลา (ไม่ใส่เส้น)']
  .map((name, index) => ({ id: `noodle_${index + 1}`, name, available: true }));
const riceAddons = [
  { id: 'addon_fried_egg', name: 'ไข่ดาว', price: 10, available: true },
  { id: 'addon_omelet', name: 'ไข่เจียว', price: 15, available: true },
  { id: 'addon_extra', name: 'พิเศษ', price: 10, available: true },
];
const noodleIds = new Set(['rr_noodle_pork', 'rr_noodle_mixed_pork', 'rr_noodle_beef', 'rr_noodle_signature']);
const riceIds = new Set(['rr_rice_pork_basil', 'rr_rice_beef_basil', 'rr_rice_pork_garlic', 'rr_rice_beef_garlic', 'rr_dry_pork_porridge', 'rr_dry_beef_porridge']);
const menuSource = [
  ['rr_noodle_pork', 'ก๋วยเตี๋ยว/เกาเหลาหมู', 'ก๋วยเตี๋ยว/เส้น', 60],
  ['rr_noodle_mixed_pork', 'ก๋วยเตี๋ยว/เกาเหลารวมหมู', 'ก๋วยเตี๋ยว/เส้น', 70],
  ['rr_noodle_beef', 'ก๋วยเตี๋ยว/เกาเหลาเนื้อ', 'ก๋วยเตี๋ยว/เส้น', 70],
  ['rr_noodle_signature', 'ก๋วยเตี๋ยว/เกาเหลารวมรงค์เรือ (หมู+เนื้อ)', 'ก๋วยเตี๋ยว/เส้น', 109],
  ['rr_pork_liver', 'ตับหมูลวก', 'ลวกจิ้ม', 30],
  ['rr_stewed_ribs', 'ซี่โครงหมูตุ๋น', 'ลวกจิ้ม', 40],
  ['rr_blanch_morning_glory', 'ผักบุ้งลวก', 'ลวกจิ้ม', 10],
  ['rr_blanch_pork', 'หมูลวก', 'ลวกจิ้ม', 40],
  ['rr_stewed_pork', 'หมูตุ๋น', 'ลวกจิ้ม', 40],
  ['rr_stewed_beef', 'เนื้อตุ๋น', 'ลวกจิ้ม', 50],
  ['rr_blanch_pork_ball', 'ลูกชิ้นหมูลวก', 'ลวกจิ้ม', 30],
  ['rr_blanch_beef_ball', 'ลูกชิ้นเนื้อลวก', 'ลวกจิ้ม', 40],
  ['rr_blanch_tendon', 'เอ็นแก้วลวก', 'ลวกจิ้ม', 50],
  ['rr_blanch_beef', 'เนื้อลวก', 'ลวกจิ้ม', 50],
  ['rr_rice_pork_basil', 'ข้าวผัดกะเพราหมูสับ/หมูชิ้น', 'อาหารจานเดียว', 55],
  ['rr_rice_beef_basil', 'ข้าวผัดกะเพราเนื้อสับ/เนื้อชิ้น', 'อาหารจานเดียว', 65],
  ['rr_rice_pork_garlic', 'ข้าวหมูกระเทียม', 'อาหารจานเดียว', 55],
  ['rr_rice_beef_garlic', 'ข้าวเนื้อกระเทียม', 'อาหารจานเดียว', 65],
  ['rr_dry_pork_porridge', 'ข้าวต้มแห้งหมู', 'อาหารจานเดียว', 60],
  ['rr_dry_beef_porridge', 'ข้าวต้มแห้งเนื้อ', 'อาหารจานเดียว', 70],
  ['rr_rice', 'ข้าวเปล่า', 'อาหารจานเดียว', 10],
  ['rr_fried_wonton', 'เกี๊ยวทอด', 'ของทานเล่น', 10],
  ['rr_pork_crackling', 'แคบหมู', 'ของทานเล่น', 20],
  ['rr_pork_rind', 'กากหมูเจียว', 'ของทานเล่น', 25],
  ['rr_thai_tea', 'ชาเย็น', 'เครื่องดื่ม', 25],
  ['rr_chrysanthemum', 'เก๊กฮวย', 'เครื่องดื่ม', 20],
  ['rr_roselle', 'กระเจี๊ยบ', 'เครื่องดื่ม', 20],
  ['rr_oleang', 'โอเลี้ยง', 'เครื่องดื่ม', 20],
  ['rr_longan', 'ลำไย', 'เครื่องดื่ม', 20],
  ['rr_soda', 'น้ำอัดลม', 'เครื่องดื่ม', 15],
  ['rr_oishi_honey_lemon', 'โออิชิน้ำผึ้งมะนาว', 'เครื่องดื่ม', 15],
  ['rr_water', 'น้ำเปล่า', 'เครื่องดื่ม', 10],
  ['rr_thai_coconut_cups', 'ขนมถ้วยไทยโบราณ (ชุด)', 'ของหวาน', 10],
];

const menuRows = menuSource.map(([id, name, category, price]) => {
  const noodle = noodleIds.has(id);
  const addons = riceIds.has(id) ? riceAddons : [];
  return [id, name, category, price, true,
    noodle ? JSON.stringify(noodleOptions) : '[]',
    noodle ? JSON.stringify(soupOptions) : '[]',
    JSON.stringify(addons), addons.length > 0, ''];
});

const configs = [
  { name: 'Menu', headers: ['id', 'name', 'category', 'price', 'available', 'noodle_options_json', 'soup_options_json', 'addons_json', 'addons_configured', 'updated_at'], rows: menuRows, tableName: 'MenuCatalog', priceCol: 'D', widths: [210, 390, 180, 90, 100, 330, 300, 260, 140, 150] },
  { name: 'หมวดหมู่', headers: ['category', 'sort_order', 'available'], rows: ['ก๋วยเตี๋ยว/เส้น', 'ลวกจิ้ม', 'อาหารจานเดียว', 'ของทานเล่น', 'เครื่องดื่ม', 'ของหวาน'].map((v, i) => [v, i + 1, true]), tableName: 'MenuCategories', widths: [240, 110, 110] },
  { name: 'ตัวเลือกเส้น', headers: ['id', 'name', 'available'], rows: noodleOptions.map(o => [o.id, o.name, o.available]), tableName: 'NoodleOptions', widths: [160, 260, 110] },
  { name: 'น้ำซุป', headers: ['id', 'name', 'available'], rows: soupOptions.map(o => [o.id, o.name, o.available]), tableName: 'SoupOptions', widths: [160, 280, 110] },
  { name: 'ท็อปปิ้งข้าว', headers: ['id', 'name', 'price_baht', 'available'], rows: riceAddons.map(o => [o.id, o.name, o.price, o.available]), tableName: 'RiceAddons', priceCol: 'C', widths: [190, 180, 140, 110] },
  { name: 'เริ่มต้นใช้งาน', headers: ['ขั้นตอน', 'รายละเอียด'], rows: [
    [1, 'อัปโหลดไฟล์นี้ไปที่ Google Drive แล้วเปิดด้วย Google Sheets'],
    [2, 'แท็บ Menu จัดตามหัวคอลัมน์ที่ Apps Script ใช้ อ่านได้โดยตรงเมื่อใช้ action syncMenu'],
    [3, 'หากต้องการให้เว็บใช้เมนูชุดนี้ ให้เปิด POS ด้วยบัญชีเจ้าของร้าน แล้วกด “ใช้เมนูตามรูป” หรือซิงก์เมนูจาก Apps Script'],
    [4, 'ตรวจชื่อเมนูและราคาให้ตรงกับป้ายร้านก่อนเริ่มใช้งานจริง'],
    [5, 'available ใช้ TRUE/FALSE; ราคาเป็นบาท; ช่อง *_json เป็นข้อมูล JSON สำหรับเว็บ'],
  ], widths: [100, 900] },
];

for (const config of configs) {
  const sheet = workbook.worksheets.add(config.name);
  sheet.showGridLines = false;
  const data = [config.headers, ...config.rows];
  const lastCol = String.fromCharCode(64 + config.headers.length);
  const lastRow = data.length;
  const range = sheet.getRange(`A1:${lastCol}${lastRow}`);
  range.values = data;
  range.format.verticalAlignment = 'center';
  const header = sheet.getRange(`A1:${lastCol}1`);
  header.format.fill = '#3730A3';
  header.format.font = { bold: true, color: '#FFFFFF', size: 11 };
  header.format.rowHeight = 28;
  sheet.freezePanes.freezeRows(1);
  config.widths.forEach((width, i) => { sheet.getRange(`${String.fromCharCode(65 + i)}:${String.fromCharCode(65 + i)}`).format.columnWidth = width; });
  if (config.priceCol && config.rows.length) sheet.getRange(`${config.priceCol}2:${config.priceCol}${lastRow}`).format.numberFormat = '#,##0.00';
  if (config.tableName && config.rows.length) sheet.tables.add(`A1:${lastCol}${lastRow}`, true, config.tableName);
  if (config.name === 'Menu') sheet.getRange(`A2:J${lastRow}`).format.rowHeight = 32;
}

workbook.recalculate();
const inspection = await workbook.inspect({ kind: 'sheet,table', maxChars: 4000, tableMaxRows: 2, tableMaxCols: 10 });
console.log(inspection.ndjson);
await fs.mkdir(outputDir, { recursive: true });
const preview = await workbook.render({ sheetName: 'Menu', range: 'A1:J8', scale: 1, format: 'png' });
await fs.writeFile(`${outputDir}/Menu_preview.png`, new Uint8Array(await preview.arrayBuffer()));
const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(outputPath);
console.log(`Saved ${outputPath}`);
