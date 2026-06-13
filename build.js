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
const SYDW_DIR = 'c:/Users/Administrator/Desktop/2026年江苏省事业单位岗位表';

// 省属事业单位（用专用解析器）
const PROVINCE_SYDW_FILE = 'c:/Users/Administrator/Desktop/2026年江苏省事业单位岗位表/2026年江苏省省属事业单位统一公开招聘岗位表.xls';

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

// ============ 批量解析事业单位 Excel（自动识别列名） ============
function parseBatchSydw(filePath, fileName) {
  let wb;
  try { wb = XLSX.readFile(filePath); } catch(e) { return []; }
  const sn = wb.SheetNames[0];
  const ws = wb.Sheets[sn];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  if (raw.length < 3) return [];

  // Find header row(s) - look for row containing key columns
  const KEYWORDS = {
    un: ['招聘单位', '单位名称', '名 称', '名称', '单位'],
    pn: ['岗位名称', '职位名称', '招聘岗位'],
    n: ['招聘人数', '拟招聘人数', '招录人数', '招考人数', '人数'],
    ed: ['学历', '学 历', '学位'],
    m: ['专业', '专 业'],
    dept: ['主管部门'],
    fund: ['经费来源', '经费'],
    target: ['招聘对象'],
    other: ['其他条件', '其它条件', '其他', '其它'],
    desc: ['岗位描述', '岗位简介', '职位简介', '工作内容'],
  };

  // Detect header row
  let headerRow = -1;
  let headerRow2 = -1;
  for (let i = 0; i < Math.min(6, raw.length); i++) {
    const row = raw[i].map(c => String(c || '').replace(/\s/g, ''));
    const hits = Object.values(KEYWORDS).filter(kws => kws.some(k => row.some(c => c.includes(k)))).length;
    if (hits >= 2) { headerRow = i; break; }
  }
  if (headerRow === -1) {
    // Try one more row
    for (let i = 0; i < Math.min(8, raw.length); i++) {
      const row = raw[i].map(c => String(c || '').replace(/\s/g, ''));
      if (row.some(c => c.includes('岗位') || c.includes('专业') || c.includes('学历'))) {
        headerRow = i; break;
      }
    }
  }
  if (headerRow === -1) return [];

  // Check if next row is also a header (merged header pattern)
  const nextRow = raw[headerRow + 1] || [];
  const nextHasCols = nextRow.some(c => {
    const s = String(c || '').replace(/\s/g, '');
    return s.includes('名称') || s.includes('代码') || s.includes('学历') || s.includes('专业') || s.includes('对象') || s.includes('条件');
  });
  if (nextHasCols) headerRow2 = headerRow + 1;

  // Build column mapping
  const mapHeaders = (row) => {
    const mapping = {};
    row.forEach((c, i) => {
      const s = String(c || '').replace(/\s/g, '');
      for (const [key, kws] of Object.entries(KEYWORDS)) {
        if (kws.some(k => s.includes(k))) {
          if (!mapping[key]) mapping[key] = i;
          break;
        }
      }
    });
    return mapping;
  };

  const map1 = mapHeaders(raw[headerRow]);
  const map2 = headerRow2 > -1 ? mapHeaders(raw[headerRow2]) : {};
  const colMap = { ...map1, ...map2 }; // row2 overrides row1

  // Must have at least unit, position name, and recruitment count
  if (!colMap.un || !colMap.pn || !colMap.n) return [];

  // Parse data rows
  const startRow = Math.max(headerRow, headerRow2) + 1;
  const results = [];
  let lastDept = '', lastUnit = '', lastFund = '';

  for (let i = startRow; i < raw.length; i++) {
    const r = raw[i];
    let unitName = String(r[colMap.un] || '').trim();
    let posName = String(r[colMap.pn] || '').trim();
    let recruitCount = parseInt(r[colMap.n]) || 0;

    // Fill merged cells
    if (!unitName) unitName = lastUnit;
    else lastUnit = unitName;
    if (colMap.dept) {
      const dept = String(r[colMap.dept] || '').trim();
      if (dept) lastDept = dept;
    }
    if (colMap.fund) {
      const fund = String(r[colMap.fund] || '').trim();
      if (fund) lastFund = fund;
    }

    if (!unitName || !posName || !recruitCount) continue;

    const major = colMap.m ? String(r[colMap.m] || '').trim() : '';
    const education = colMap.ed ? String(r[colMap.ed] || '').trim() : '';
    const other = colMap.other ? String(r[colMap.other] || '').trim() : '';
    const target = colMap.target ? String(r[colMap.target] || '').trim() : '';
    const desc = colMap.desc ? String(r[colMap.desc] || '').trim() : '';

    // Infer city from filename
    const cityHints = ['南京','苏州','无锡','常州','徐州','南通','连云港','淮安','盐城','扬州','镇江','泰州','宿迁','射阳','阜宁','响水','句容','扬中','如皋','东海','灌云','灌南','沭阳','泗洪','泗阳','江阴','宜兴','昆山','太仓','常熟','张家港','姑苏','吴中','吴江','相城','栖霞','雨花台','江宁','浦口','六合','溧水','高淳','鼓楼','玄武','秦淮','建邺','武进','新北','天宁','钟楼','金坛','海陵','高港','姜堰','兴化','泰兴','靖江','大丰','亭湖','盐都','洪泽','淮安','清江浦','淮阴','涟水','盱眙','金湖','广陵','邗江','江都','高邮','宝应','仪征','京口','润州','丹徒','丹阳','崇川','通州','海门','启东','海安','如东','新沂','邳州','丰县','沛县','睢宁','宿豫','宿城','赣榆','海州','连云'];
    let cityInferred = '江苏省';
    for (const hint of cityHints) {
      if (fileName.includes(hint) || unitName.includes(hint)) {
        // Map county to prefecture city
        const countyMap = {
          '射阳':'盐城市','阜宁':'盐城市','响水':'盐城市','大丰':'盐城市','亭湖':'盐城市','盐都':'盐城市',
          '句容':'镇江市','扬中':'镇江市','丹徒':'镇江市','丹阳':'镇江市','京口':'镇江市','润州':'镇江市',
          '如皋':'南通市','海安':'南通市','海门':'南通市','启东':'南通市','如东':'南通市','崇川':'南通市','通州':'南通市',
          '东海':'连云港市','灌云':'连云港市','灌南':'连云港市','赣榆':'连云港市','海州':'连云港市','连云':'连云港市',
          '沭阳':'宿迁市','泗洪':'宿迁市','泗阳':'宿迁市','宿豫':'宿迁市','宿城':'宿迁市',
          '涟水':'淮安市','盱眙':'淮安市','金湖':'淮安市','洪泽':'淮安市','清江浦':'淮安市','淮阴':'淮安市','淮安':'淮安市',
          '江阴':'无锡市','宜兴':'无锡市',
          '昆山':'苏州市','太仓':'苏州市','常熟':'苏州市','张家港':'苏州市','姑苏':'苏州市','吴中':'苏州市','吴江':'苏州市','相城':'苏州市',
          '兴化':'泰州市','泰兴':'泰州市','靖江':'泰州市','姜堰':'泰州市','海陵':'泰州市','高港':'泰州市',
          '高邮':'扬州市','宝应':'扬州市','仪征':'扬州市','广陵':'扬州市','邗江':'扬州市','江都':'扬州市',
          '新沂':'徐州市','邳州':'徐州市','丰县':'徐州市','沛县':'徐州市','睢宁':'徐州市',
          '武进':'常州市','新北':'常州市','天宁':'常州市','钟楼':'常州市','金坛':'常州市',
          '栖霞':'南京市','雨花台':'南京市','江宁':'南京市','浦口':'南京市','六合':'南京市','溧水':'南京市','高淳':'南京市','鼓楼':'南京市','玄武':'南京市','秦淮':'南京市','建邺':'南京市',
        };
        cityInferred = countyMap[hint] || hint + '市';
        if (hint === '南京') cityInferred = '南京市';
        if (hint === '苏州') cityInferred = '苏州市';
        break;
      }
    }

    const obj = {
      a: '县',
      ac: '',
      an: cityInferred,
      uc: '',
      un: unitName,
      pc: '',
      pn: posName,
      pd: desc,
      e: '',
      r: '',
      n: recruitCount,
      ed: education,
      m: major,
      o: (other + ' | ' + target).replace(/\s*\|\s*$/, ''),
      _t: 'sydw',
      _dept: lastDept,
      _fund: lastFund,
      _target: target,
      _examFmt: '',
      _contact: '',
    };
    results.push(obj);
  }

  return results;
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

// 事业单位 - 省属（专用解析器）
if (fs.existsSync(PROVINCE_SYDW_FILE)) {
  const rows = parseSydwFile(PROVINCE_SYDW_FILE, '附件1', '江苏省', '320000');
  const recruits = rows.reduce((s, r) => s + r.n, 0);
  rows.forEach(r => { r.mt = tagMajorCategories(r.m); r._src = '省属事业单位'; });
  allData = allData.concat(rows);
  fileStats.push({ type: '事业单位', file: '省属事业单位', positions: rows.length, recruits });
  console.log(`  ✅ 事业单位 省属: ${rows.length} 职位, ${recruits} 人`);
}

// 事业单位 - 批量文件夹
if (fs.existsSync(SYDW_DIR)) {
  const batchFiles = fs.readdirSync(SYDW_DIR).filter(f => f.match(/\.(xls|xlsx)$/) || f.endsWith('.xls.xls') || f.endsWith('.xlsx.xlsx'));
  let batchTotal = 0, batchRecruits = 0, batchSkipped = 0;
  for (const f of batchFiles) {
    // Skip省属 file (already parsed with dedicated parser)
    if (f.includes('省属事业单位')) continue;
    const fp = path.join(SYDW_DIR, f);
    try {
      const rows = parseBatchSydw(fp, f);
      if (rows.length === 0) { batchSkipped++; continue; }
      const recruits = rows.reduce((s, r) => s + r.n, 0);
      rows.forEach(r => {
        r.mt = tagMajorCategories(r.m);
        r._src = f.substring(0, 30);
        // Ensure unique key
        r.pc = 'B' + batchTotal + '_' + (rows.indexOf(r));
      });
      allData = allData.concat(rows);
      batchTotal += rows.length;
      batchRecruits += recruits;
    } catch(e) {
      batchSkipped++;
    }
  }
  if (batchTotal > 0) {
    fileStats.push({ type: '事业单位', file: '各市县批量导入', positions: batchTotal, recruits: batchRecruits });
    console.log(`  ✅ 事业单位 批量: ${batchTotal} 职位, ${batchRecruits} 人 (跳过${batchSkipped}个无法解析的文件)`);
  }
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

// ============ 智能城市识别（从单位名称推断实际工作地） ============
const CITY_PATTERNS = [
  { pattern: /南京|江宁|溧水|高淳|浦口|六合|栖霞|雨花|鼓楼|玄武|秦淮|建邺|江北/, city: '南京市' },
  { pattern: /无锡|江阴|宜兴|锡山|惠山|滨湖|梁溪|新吴/, city: '无锡市' },
  { pattern: /徐州|丰县|沛县|睢宁|邳州|新沂|云龙|贾汪|泉山|铜山/, city: '徐州市' },
  { pattern: /常州|溧阳|金坛|武进|新北|天宁|钟楼/, city: '常州市' },
  { pattern: /苏州|昆山|太仓|常熟|张家港|吴江|吴中|姑苏|相城|虎丘|园区/, city: '苏州市' },
  { pattern: /南通|启东|海安|如皋|如东|海门|崇川|通州/, city: '南通市' },
  { pattern: /连云港|东海|灌云|灌南|赣榆|海州|连云(?!港)|徐圩/, city: '连云港市' },
  { pattern: /淮安|涟水|盱眙|金湖|洪泽|清江浦|淮阴/, city: '淮安市' },
  { pattern: /盐城|东台|建湖|射阳|阜宁|滨海|响水|大丰|亭湖|盐都/, city: '盐城市' },
  { pattern: /扬州|仪征|高邮|宝应|邗江|广陵|江都/, city: '扬州市' },
  { pattern: /镇江|丹阳|扬中|句容|丹徒|京口|润州/, city: '镇江市' },
  { pattern: /泰州|兴化|泰兴|靖江|姜堰|海陵|高港/, city: '泰州市' },
  { pattern: /宿迁|沭阳|泗洪|泗阳|宿豫|宿城|湖滨/, city: '宿迁市' },
];

function detectCity(unitName, positionDesc, originalAn) {
  // Only reassign positions currently tagged as provincial or county level
  if (originalAn === '江苏省' || originalAn === '省') {
    for (const { pattern, city } of CITY_PATTERNS) {
      if (pattern.test(unitName) || pattern.test(positionDesc)) {
        return city;
      }
    }
  }
  return originalAn;
}

// Apply city detection
allData.forEach(r => {
  const detected = detectCity(r.un, r.pd, r.an);
  if (detected !== r.an) {
    r.an = detected;
  }
});

// ============ 添加父级城市标签 ============
const CITY_MAP = {
  '南京市鼓楼区':'南京市','南京市玄武区':'南京市','南京市秦淮区':'南京市','南京市建邺区':'南京市','南京市栖霞区':'南京市','南京市雨花台区':'南京市','南京市江宁区':'南京市','南京市浦口区':'南京市','南京市六合区':'南京市','南京市溧水区':'南京市','南京市高淳区':'南京市','南京市江北新区':'南京市','南京市市辖区':'南京市',
  '苏州市吴江区':'苏州市','苏州市吴中区':'苏州市','苏州市姑苏区':'苏州市','苏州市相城区':'苏州市','苏州市虎丘区':'苏州市','苏州工业园区':'苏州市','昆山市':'苏州市','太仓市':'苏州市','常熟市':'苏州市','张家港市':'苏州市',
  '无锡市梁溪区':'无锡市','无锡市锡山区':'无锡市','无锡市惠山区':'无锡市','无锡市滨湖区':'无锡市','无锡市新吴区':'无锡市','江阴市':'无锡市','宜兴市':'无锡市',
  '常州市武进区':'常州市','常州市新北区':'常州市','常州市天宁区':'常州市','常州市钟楼区':'常州市','常州市金坛区':'常州市','常州经开区':'常州市','溧阳市':'常州市',
  '徐州市鼓楼区':'徐州市','徐州市云龙区':'徐州市','徐州市贾汪区':'徐州市','徐州市泉山区':'徐州市','徐州市铜山区':'徐州市','徐州经济技术开发区':'徐州市','新沂市':'徐州市','邳州市':'徐州市','丰县':'徐州市','沛县':'徐州市','睢宁县':'徐州市',
  '南通市崇川区':'南通市','南通市通州区':'南通市','南通市海门区':'南通市','启东市':'南通市','海安市':'南通市','如皋市':'南通市','如东县':'南通市',
  '连云港市海州区':'连云港市','连云港市连云区':'连云港市','连云港市赣榆区':'连云港市','连云港经济技术开发区':'连云港市','连云港徐圩新区':'连云港市','东海县':'连云港市','灌云县':'连云港市','灌南县':'连云港市',
  '淮安市清江浦区':'淮安市','淮安市淮安区':'淮安市','淮安市淮阴区':'淮安市','淮安市洪泽区':'淮安市','淮安经济技术开发区':'淮安市','涟水县':'淮安市','盱眙县':'淮安市','金湖县':'淮安市',
  '盐城市亭湖区':'盐城市','盐城市盐都区':'盐城市','盐城市大丰区':'盐城市','盐城经济技术开发区':'盐城市','东台市':'盐城市','建湖县':'盐城市','射阳县':'盐城市','阜宁县':'盐城市','滨海县':'盐城市','响水县':'盐城市',
  '扬州市广陵区':'扬州市','扬州市邗江区':'扬州市','扬州市江都区':'扬州市','扬州经济技术开发区':'扬州市','扬州市市辖区':'扬州市','仪征市':'扬州市','高邮市':'扬州市','宝应县':'扬州市',
  '镇江市京口区':'镇江市','镇江市润州区':'镇江市','镇江市丹徒区':'镇江市','镇江经济技术开发区':'镇江市','镇江高新技术产业开发区':'镇江市','丹阳市':'镇江市','扬中市':'镇江市','句容市':'镇江市',
  '泰州市海陵区':'泰州市','泰州市高港区':'泰州市','泰州市姜堰区':'泰州市','兴化市':'泰州市','泰兴市':'泰州市','靖江市':'泰州市',
  '宿迁市宿城区':'宿迁市','宿迁市宿豫区':'宿迁市','宿迁市湖滨新区':'宿迁市','沭阳县':'宿迁市','泗洪县':'宿迁市','泗阳县':'宿迁市',
};
allData.forEach(r => {
  r._city = CITY_MAP[r.an] || r.an; // 区县→父级城市，市级直接保留
});

const totalRecruits = allData.reduce((s, r) => s + r.n, 0);
console.log(`\n📊 总计: ${allData.length} 职位, ${totalRecruits} 人`);
console.log(`   公务员: ${allData.filter(r=>r._t==='gwy').length} 职位, ${allData.filter(r=>r._t==='gwy').reduce((s,r)=>s+r.n,0)} 人`);
console.log(`   事业单位: ${allData.filter(r=>r._t==='sydw').length} 职位, ${allData.filter(r=>r._t==='sydw').reduce((s,r)=>s+r.n,0)} 人`);

// ============ 统计 ============
const cities = {};
allData.forEach(r => {
  const city = r._city || r.an;
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
  a: r.a, an: r.an, _city: r._city || r.an, un: r.un, pc: r.pc, pn: r.pn, pd: r.pd,
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
