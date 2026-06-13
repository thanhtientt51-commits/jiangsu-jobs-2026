const fs = require('fs');
let code = fs.readFileSync('build.js', 'utf-8');

const teacherCode = `
// ============ 教师招聘 ============
const TEACHER_FILE = 'c:/Users/Administrator/Desktop/（6.8更新）2026年江苏省教师招聘信息汇总-276人【公众号：考招信息室 整理】.xlsx';
if (fs.existsSync(TEACHER_FILE)) {
  try {
    const tb = XLSX.readFile(TEACHER_FILE);
    const tws = tb.Sheets[tb.SheetNames[0]];
    const tdata = XLSX.utils.sheet_to_json(tws, { header: 1, defval: '' });
    let tCount = 0, tRecruits = 0;
    for (let i = 4; i < tdata.length; i++) {
      const r = tdata[i];
      if (!r[2] || !String(r[2]).trim()) continue;
      const seq = String(r[0]||'').trim();
      if (!seq) continue;
      const time = String(r[1]||'').trim();
      const name = String(r[2]||'').trim();
      const link = String(r[3]||'').trim();
      const count = parseInt(r[4]) || 0;
      if (!name || !count) continue;

      let city = '江苏省';
      for (const { pattern, city: c } of CITY_PATTERNS) {
        if (pattern.test(name) || pattern.test(link)) { city = c; break; }
      }

      const obj = {
        a: '', an: city, uc: '', un: name,
        pc: 'T' + tCount, pn: name.substring(0, 40),
        pd: time,
        e: '教师岗',
        r: '', n: count,
        ed: '', m: '',
        o: '报名时间: ' + time,
        _t: 'teacher',
        _dept: '', _fund: '', _target: '', _examFmt: '', _contact: '',
        _link: link,
        _src: '教师招聘汇总',
        _city: city,
        mt: ['教师'],
      };
      allData.push(obj);
      tCount++; tRecruits += count;
    }
    fileStats.push({ type: '教师', file: '教师招聘汇总(考招信息室)', positions: tCount, recruits: tRecruits });
    console.log('  ✅ 教师招聘: ' + tCount + ' 条公告, 约' + tRecruits + '人');
  } catch(e) {
    console.log('  ⚠️ 教师招聘解析失败: ' + e.message);
  }
}
`;

const insertPoint = code.indexOf('// ============ 去重 ============');
if (insertPoint > -1) {
  code = code.substring(0, insertPoint) + teacherCode + '\n' + code.substring(insertPoint);
  fs.writeFileSync('build.js', code);
  console.log('✅ Teacher parser inserted');
} else {
  console.log('❌ Insert point not found');
}
