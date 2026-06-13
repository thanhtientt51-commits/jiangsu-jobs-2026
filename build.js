/**
 * 江苏省考公务员筛岗网站 — 数据管道
 * 读取 16 个 .xls 职位表 → 清洗压缩 → 输出 index.html
 */
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// ============ 配置 ============
const DATA_DIR = 'c:/Users/Administrator/Desktop/江苏省2026年度考试录用公务员职位详情表';
const OUTPUT_FILE = path.join(__dirname, 'index.html');

// 16 个源文件
const FILE_LIST = [
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

// 列名缩写映射
const KEY_MAP = {
  '隶属  关系': 'a',    // affiliation
  '地区  代码': 'ac',   // area code
  '地区  名称': 'an',   // area name
  '单位代码': 'uc',     // unit code
  '单位名称': 'un',     // unit name
  '职位代码': 'pc',     // position code
  '职位名称': 'pn',     // position name
  '职位简介': 'pd',     // position description
  '考试类别': 'e',      // exam category
  '开考比例': 'r',      // ratio
  '招考人数': 'n',      // number recruited
  '学　历': 'ed',       // education
  '专　业': 'm',        // major
  '其　它': 'o',        // other
};

// 专业大类关键词（从官方分类中提取）
const MAJOR_CATEGORIES = [
  '不限', '法律类', '计算机类', '财务财会类', '审计类', '经济类',
  '中文文秘类', '公共管理类', '社会政治类', '教育类',
  '城建规划类', '建筑工程类', '交通工程类', '水利工程类', '土地管理类',
  '环境保护类', '化学工程类', '医药化工类', '食品工程类',
  '农业类', '林业类', '畜牧养殖类',
  '医学类', '药学类', '公共卫生类',
  '电子信息类', '机电控制类', '机械工程类', '能源动力类',
  '安全生产类', '统计类', '审计类',
  '外国语言文学类', '艺术类', '公安类',
  '军事学类', '体育类',
];

// 专业大类别称映射（"计算机（软件）类"、"计算机(网络管理)类" 等 → "计算机类"）
const MAJOR_ALIASES = {
  '计算机（软件）类': '计算机类',
  '计算机(软件)类': '计算机类',
  '计算机（网络管理）类': '计算机类',
  '计算机(网络管理)类': '计算机类',
  '电子信息工程': '电子信息类',
  '财务财会': '财务财会类',
  '财务会计': '财务财会类',
  '会计学': '财务财会类',
  '审计': '审计类',
  '法律': '法律类',
  '法学': '法律类',
  '中文': '中文文秘类',
  '文秘': '中文文秘类',
  '汉语言': '中文文秘类',
  '新闻': '中文文秘类',
};

// ============ 解析 Excel ============
function parseFile(filePath) {
  const wb = XLSX.readFile(filePath);
  const ws = wb.Sheets[wb.SheetNames[0]];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  // 找表头行（包含"隶属"的那行）
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
    .filter(r => r[4] && String(r[4]).trim()) // 有单位名的行
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => {
        const key = KEY_MAP[h] || h;
        let val = String(r[i] || '').trim().replace(/\s+/g, ' ');
        // 统一全角半角
        val = val.replace(/，/g, ',').replace(/；/g, ';').replace(/：/g, ':');
        obj[key] = val;
      });
      // 转换数值
      obj.n = parseInt(obj.n) || 0;
      return obj;
    });
}

// ============ 专业归类 ============
function tagMajorCategories(majorStr) {
  if (!majorStr || majorStr === '不限') return ['不限'];

  const tags = [];
  const cleaned = majorStr
    .replace(/（/g, '(')
    .replace(/）/g, ')')
    .replace(/（除[^）]+外）/g, '')
    .replace(/\(除[^)]+外\)/g, '');

  for (const cat of MAJOR_CATEGORIES) {
    if (cleaned.includes(cat)) {
      tags.push(cat);
    }
  }

  // 检查别称
  for (const [alias, target] of Object.entries(MAJOR_ALIASES)) {
    if (cleaned.includes(alias) && !tags.includes(target)) {
      tags.push(target);
    }
  }

  return tags.length > 0 ? tags : [majorStr.substring(0, 30)];
}

// ============ 主流程 ============
console.log('🔨 开始构建数据管道...\n');

let allData = [];
const fileStats = [];

for (const file of FILE_LIST) {
  const fullPath = path.join(DATA_DIR, file);
  if (!fs.existsSync(fullPath)) {
    console.log(`  ⚠️ 跳过 (不存在): ${file}`);
    continue;
  }

  const rows = parseFile(fullPath);
  const recruits = rows.reduce((s, r) => s + r.n, 0);

  // 标记来源（仅保留文件名关键词）
  const sourceName = path.basename(file, '.xls').replace('江苏省2026年度考试录用公务员', '').replace('职位表', '').replace('职位简介表', '').replace('各地', '').replace('考录', '');

  rows.forEach(r => {
    r._src = sourceName;
    // 专业归类标签
    r.mt = tagMajorCategories(r.m);
  });

  allData = allData.concat(rows);
  fileStats.push({ file: sourceName, positions: rows.length, recruits });

  console.log(`  ✅ ${sourceName}: ${rows.length} 职位, ${recruits} 人`);
}

// ============ 去重检查 ============
const seen = new Set();
const deduped = [];
let dupCount = 0;
for (const r of allData) {
  const key = `${r.ac}|${r.uc}|${r.pc}`;
  if (seen.has(key)) {
    dupCount++;
  } else {
    seen.add(key);
    deduped.push(r);
  }
}
if (dupCount > 0) console.log(`\n  ⚠️ 发现 ${dupCount} 条重复，已去重`);
allData = deduped;

// ============ 生成统计 ============
const totalRecruits = allData.reduce((s, r) => s + r.n, 0);
console.log(`\n📊 总计: ${allData.length} 职位, ${totalRecruits} 人\n`);

// 城市列表（分组）
const cities = {};
allData.forEach(r => {
  const city = r.an;
  if (!cities[city]) cities[city] = { name: city, count: 0, recruits: 0 };
  cities[city].count++;
  cities[city].recruits += r.n;
});

// 专业大类统计
const majorStats = {};
allData.forEach(r => {
  (r.mt || []).forEach(tag => {
    if (!majorStats[tag]) majorStats[tag] = { name: tag, count: 0, recruits: 0 };
    majorStats[tag].count++;
    majorStats[tag].recruits += r.n;
  });
});

// 学历
const eduList = [...new Set(allData.map(r => r.ed))].sort();

// 隶属关系
const affilList = [...new Set(allData.map(r => r.a))].sort();

// 最大招录人数
const maxRecruit = Math.max(...allData.map(r => r.n));

// 构建统计数据
const STATS = {
  totalPositions: allData.length,
  totalRecruits,
  fileStats,
  cities: Object.values(cities).sort((a, b) => a.name.localeCompare(b.name, 'zh')),
  majorCategories: Object.values(majorStats).sort((a, b) => b.count - a.count),
  educationLevels: eduList,
  examCategories: ['A', 'B', 'C'],
  affiliations: affilList,
  maxRecruit,
};

// ============ 最小化数据 ============
// 去掉不用于展示的字段进一步压缩
const compactData = allData.map(r => ({
  a: r.a,        // 隶属关系
  an: r.an,      // 地区名称
  un: r.un,      // 单位名称
  pc: r.pc,      // 职位代码
  pn: r.pn,      // 职位名称
  pd: r.pd,      // 职位简介
  e: r.e,        // 考试类别
  r: r.r,        // 开考比例
  n: r.n,        // 招录人数
  ed: r.ed,      // 学历
  m: r.m,        // 专业
  o: r.o,        // 其它
  mt: r.mt,      // 专业分类标签
  _s: r._src,    // 来源文件
}));

const DATA_JSON = JSON.stringify(compactData);
const STATS_JSON = JSON.stringify(STATS);

console.log(`📦 数据大小: ${(DATA_JSON.length / 1024 / 1024).toFixed(2)} MB`);
console.log(`📦 统计大小: ${(STATS_JSON.length / 1024).toFixed(1)} KB`);

// ============ 读取模板生成 HTML ============
const templatePath = path.join(__dirname, 'template.html');
if (fs.existsSync(templatePath)) {
  let html = fs.readFileSync(templatePath, 'utf-8');
  html = html.replace('{{DATA}}', DATA_JSON);
  html = html.replace('{{STATS}}', STATS_JSON);
  fs.writeFileSync(OUTPUT_FILE, html, 'utf-8');
  console.log(`\n✅ 已生成: ${OUTPUT_FILE}`);
} else {
  console.log('\n⚠️ 模板文件 template.html 不存在，仅输出了 JSON');
  // 输出原始 JSON 备用
  fs.writeFileSync(path.join(__dirname, 'data.json'), DATA_JSON, 'utf-8');
  fs.writeFileSync(path.join(__dirname, 'stats.json'), STATS_JSON, 'utf-8');
  console.log('  数据已存为 data.json 和 stats.json');
}
