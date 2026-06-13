/**
 * 江苏省考公务员+事业单位筛岗网站 — 数据管道
 */
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const OUTPUT_FILE = path.join(__dirname, 'index.html');

// ============ 公务员源文件 ============
const DATA_DIR_GWY = 'c:/Users/Administrator/Desktop/江苏省2026年度考试录用公务员职位详情表';
const GWY_FILES = [
  '江苏省2026年度考试录用公务员省级机关职位表.xls',
  '江苏省2026年度地方统计系统考录职位简介表.xls',
  '江苏省2026年度监狱、戒毒系统考录职位简介表.xls',
  '江苏省2026年度考试录用公务员各地职位表/01-南京市.xls',
  '江苏省2026年度考试录用公务员各地职位表/02-无锡市.xls',
  '江苏省2026年度考试录用公务员各地职位表/03-徐州市.xls',
  '江苏省2026年度考试录用公务员各地职位表/04-常州市.xls',
  '江苏省2026年度考试录用公务员各地职位表/05-苏州市.xls',
  '江苏省2026年度考试录用公务员各地职位表/06-南通市.xls',
  '江苏省2026年度考试录用公务员各地职位表/07-连云港市.xls',
  '江苏省2026年度考试录用公务员各地职位表/08-淮安市.xls',
  '江苏省2026年度考试录用公务员各地职位表/09-盐城市.xls',
  '江苏省2026年度考试录用公务员各地职位表/10-扬州市.xls',
  '江苏省2026年度考试录用公务员各地职位表/11-镇江市.xls',
  '江苏省2026年度考试录用公务员各地职位表/12-泰州市.xls',
  '江苏省2026年度考试录用公务员各地职位表/13-宿迁市.xls',
];

// ============ 事业单位源文件 ============
const SYDW_FILES = [
  { file: 'c:/Users/Administrator/Desktop/2026年江苏省省属事业单位统一公开招聘岗位表.xls', sheet: '附件1', city: '江苏省', cityCode: '320000' },
];

// ============ 公务员列名缩写 ============
const GWY_KEY_MAP = {
  '隶属  关系': 'a', '地区  代码': 'ac', '地区  名称': 'an',
  '单位代码': 'uc', '单位名称': 'un', '职位代码': 'pc',
  '职位名称': 'pn', '职位简介': 'pd', '考试类别': 'e',
  '开考比例': 'r', '招考人数': 'n', '学　历': 'ed', '专　业': 'm', '其　它': 'o',
};

// ============ 专业大类关键词 ============
const MAJOR_CATEGORIES = [
  '不限', '法律类', '计算机类', '财务财会类', '审计类', '经济类',
  '中文文秘类', '公共管理类', '社会政治类', '教育类',
  '城建规划类', '建筑工程类', '交通工程类', '水利工程类', '土地管理类',
  '环境保护类', '化学工程类', '医药化工类', '食品工程类',
  '农业类', '林业类', '畜牧养殖类',
  '医学类', '药学类', '公共卫生类',
  '电子信息类', '机电控制类', '机械工程类', '能源动力类',
  '安全生产类', '统计类',
  '外国语言文学类', '艺术类', '公安类',
  '军事学类', '体育类',
];

const MAJOR_ALIASES = {
  '计算机（软件）类': '计算机类',
  '计算机(软件)类': '计算机类',
  '计算机（网络管理）类': '计算机类',
  '计算机(网络管理)类': '计算机类',
};

// ============ 解析公务员 Excel ============
function parseGwyFile(filePath) {
  const wb = XLSX.readFile(filePath);
  const ws = wb.Sheets[wb.SheetNames[0]];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  let headerIdx = -1;
  for (let i = 0; i < Math.min(5, raw.length); i++) {
    if (raw[i].some(c => String(c).includes('隶属'))) {
      headerIdx = i;
      break;
    }
  }
  if (headerIdx === -1) return [];

  const headers = raw[headerIdx];
  const rows = raw.slice(headerIdx + 1);

  return rows
    .filter(r => r[4] && String(r[4]).trim())
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => {
        const key = GWY_KEY_MAP[h] || h;
        let val = String(r[i] || '').trim().replace(/\s+/g, ' ');
        val = val.replace(/，/g, ',').replace(/；/g, ';').replace(/：/g, ':');
        obj[key] = val;
      });
      obj.n = parseInt(obj.n) || 0;
      obj._t = 'gwy'; // 公务员
      return obj;
    });
}

// ============ 解析事业单位 Excel ============
function parseSydwFile(filePath, sheetName, defaultCity, defaultCityCode) {
  const wb = XLSX.readFile(filePath);
  const ws = wb.Sheets[sheetName];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  // Row 0: title, Row 1: column headers A, Row 2: column headers B
  // Data starts from Row 3
  if (raw.length < 4) return [];

  const results = [];
  let lastDept = '';
  let lastUnitName = '';
  let lastUnitCode = '';
  let lastFunding = '';

  for (let i = 3; i < raw.length; i++) {
    const r = raw[i];
    if (!r || !r[5]) continue; // skip rows without position name

    const seq = String(r[0] || '').trim();
    const dept = String(r[1] || '').trim() || lastDept;
    const unitName = String(r[2] || '').trim() || lastUnitName;
    const unitCode = String(r[3] || '').trim() || lastUnitCode;
    const funding = String(r[4] || '').trim() || lastFunding;
    const posName = String(r[5] || '').trim();
    const posCode = String(r[6] || '').trim();
    const posCat = String(r[7] || '').trim();
    const posDesc = String(r[8] || '').trim();
    const recruitCount = parseInt(r[9]) || 0;
    const examRatio = String(r[10] || '').trim();
    const checkRatio = String(r[11] || '').trim();
    const education = String(r[12] || '').trim();
    const major = String(r[13] || '').trim();
    const targetCandidate = String(r[14] || '').trim();
    const otherConditions = String(r[15] || '').trim();
    const examFormat = String(r[16] || '').trim();
    const otherNotes = String(r[17] || '').trim();
    const contact = String(r[18] || '').trim();

    if (seq) { lastDept = dept; lastUnitName = unitName; lastUnitCode = unitCode; lastFunding = funding; }

    // Map to common schema
    const obj = {
      a: '省',                       // 隶属关系: all provincial
      ac: defaultCityCode,
      an: defaultCity,              // 地区: 江苏省/省属
      uc: unitCode,
      un: unitName,
      pc: posCode,
      pn: posName,
      pd: posDesc,
      e: posCat,                    // 岗位类别 (管理类/其他专技类/法律类 etc.)
      r: examRatio,
      n: recruitCount,
      ed: education,
      m: major,
      o: (otherConditions + ' | ' + otherNotes + ' | ' + targetCandidate).replace(/\s*\|\s*$/, ''),
      _t: 'sydw',                   // 事业单位
      _dept: dept,                  // 主管部门
      _fund: funding,               // 经费来源
      _target: targetCandidate,     // 招聘对象
      _examFmt: examFormat,         // 考试形式
      _contact: contact,            // 联系电话
    };

    results.push(obj);
  }

  return results;
}

// ============ 专业归类 ============
function tagMajorCategories(majorStr) {
  if (!majorStr || majorStr === '不限') return ['不限'];
  const cleaned = majorStr
    .replace(/（/g, '(').replace(/）/g, ')')
    .replace(/（除[^）]+外）/g, '')
    .replace(/\(除[^)]+外\)/g, '');
  const tags = [];
  for (const cat of MAJOR_CATEGORIES) {
    if (cleaned.includes(cat)) tags.push(cat);
  }
  for (const [alias, target] of Object.entries(MAJOR_ALIASES)) {
    if (cleaned.includes(alias) && !tags.includes(target)) tags.push(target);
  }
  return tags.length > 0 ? tags : [majorStr.substring(0, 30)];
}

// ============ 主流程 ============
console.log('🔨 开始构建数据管道...\n');

let allData = [];
const fileStats = [];

// 公务员
for (const file of GWY_FILES) {
  const fullPath = path.join(DATA_DIR_GWY, file);
  if (!fs.existsSync(fullPath)) { console.log(`  ⚠️ 跳过: ${file}`); continue; }
  const rows = parseGwyFile(fullPath);
  const recruits = rows.reduce((s, r) => s + r.n, 0);
  const srcName = path.basename(file, '.xls').replace('江苏省2026年度考试录用公务员', '').replace('职位表', '').replace('职位简介表', '').replace('各地', '').replace('考录', '');
  rows.forEach(r => { r.mt = tagMajorCategories(r.m); r._src = srcName; });
  allData = allData.concat(rows);
  fileStats.push({ type: '公务员', file: srcName, positions: rows.length, recruits });
  console.log(`  ✅ 公务员 ${srcName}: ${rows.length} 职位, ${recruits} 人`);
}

// 事业单位
for (const cfg of SYDW_FILES) {
  if (!fs.existsSync(cfg.file)) { console.log(`  ⚠️ 跳过: ${cfg.file}`); continue; }
  const rows = parseSydwFile(cfg.file, cfg.sheet, cfg.city, cfg.cityCode);
  const recruits = rows.reduce((s, r) => s + r.n, 0);
  rows.forEach(r => { r.mt = tagMajorCategories(r.m); r._src = '省属事业单位'; });
  allData = allData.concat(rows);
  fileStats.push({ type: '事业单位', file: '省属事业单位', positions: rows.length, recruits });
  console.log(`  ✅ 事业单位 省属: ${rows.length} 职位, ${recruits} 人`);
}

// ============ 去重 ============
const seen = new Set();
const deduped = [];
let dupCount = 0;
for (const r of allData) {
  const key = `${r.ac}|${r.uc}|${r.pc}|${r._t}`;
  if (seen.has(key)) { dupCount++; }
  else { seen.add(key); deduped.push(r); }
}
if (dupCount > 0) console.log(`\n  ⚠️ 发现 ${dupCount} 条重复，已去重`);
allData = deduped;

const totalRecruits = allData.reduce((s, r) => s + r.n, 0);
console.log(`\n📊 总计: ${allData.length} 职位, ${totalRecruits} 人`);
console.log(`   公务员: ${allData.filter(r=>r._t==='gwy').length} 职位, ${allData.filter(r=>r._t==='gwy').reduce((s,r)=>s+r.n,0)} 人`);
console.log(`   事业单位: ${allData.filter(r=>r._t==='sydw').length} 职位, ${allData.filter(r=>r._t==='sydw').reduce((s,r)=>s+r.n,0)} 人`);

// ============ 统计 ============
const cities = {};
allData.forEach(r => {
  const city = r.an;
  if (!cities[city]) cities[city] = { name: city, count: 0, recruits: 0 };
  cities[city].count++; cities[city].recruits += r.n;
});

const majorStats = {};
allData.forEach(r => {
  (r.mt || []).forEach(tag => {
    if (!majorStats[tag]) majorStats[tag] = { name: tag, count: 0, recruits: 0 };
    majorStats[tag].count++; majorStats[tag].recruits += r.n;
  });
});

const eduList = [...new Set(allData.map(r => r.ed))].filter(Boolean).sort();
const affilList = [...new Set(allData.map(r => r.a))].filter(Boolean).sort();
const examCats = [...new Set(allData.map(r => r.e))].filter(Boolean).sort();
const maxRecruit = Math.max(...allData.map(r => r.n));

const STATS = {
  totalPositions: allData.length,
  totalRecruits,
  fileStats,
  cities: Object.values(cities).sort((a, b) => a.name.localeCompare(b.name, 'zh')),
  majorCategories: Object.values(majorStats).sort((a, b) => b.count - a.count),
  educationLevels: eduList,
  examCategories: examCats,
  affiliations: affilList,
  maxRecruit,
  hasBusiness: true,
};

// ============ 压缩数据 ============
const compactData = allData.map(r => ({
  a: r.a, an: r.an, un: r.un, pc: r.pc, pn: r.pn, pd: r.pd,
  e: r.e, r: r.r, n: r.n, ed: r.ed, m: r.m, o: r.o,
  mt: r.mt, _s: r._src, _t: r._t,
  _dept: r._dept || '', _fund: r._fund || '', _target: r._target || '',
  _examFmt: r._examFmt || '', _contact: r._contact || '',
}));

const DATA_JSON = JSON.stringify(compactData);
const STATS_JSON = JSON.stringify(STATS);

console.log(`\n📦 数据大小: ${(DATA_JSON.length/1024/1024).toFixed(2)} MB`);
console.log(`📦 统计大小: ${(STATS_JSON.length/1024).toFixed(1)} KB`);

// ============ 生成 HTML ============
const templatePath = path.join(__dirname, 'template.html');
if (fs.existsSync(templatePath)) {
  let html = fs.readFileSync(templatePath, 'utf-8');
  html = html.replace('{{DATA}}', DATA_JSON);
  html = html.replace('{{STATS}}', STATS_JSON);
  fs.writeFileSync(OUTPUT_FILE, html, 'utf-8');
  console.log(`\n✅ 已生成: ${OUTPUT_FILE}`);
} else {
  console.log('\n⚠️ 模板文件 template.html 不存在');
}
