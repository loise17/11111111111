/**
 * 技能自己評価シート（Excel）解析モジュール
 * 列: A=カテゴリ  B=分類  C=名称  D列以降=各社員（名前行の1つ下の行＝更新日）
 * 記号: ☆=得意  ○=一人でできる  △=助けがあればできる  ↑=未経験だがやってみたい  空欄=未経験
 */

const LEVELS = {
  '☆': { level: 3, label: '得意', description: '得意（能带领、能教别人）' },
  '○': { level: 2, label: '可以', description: '一个人能做' },
  '△': { level: 1, label: '需协助', description: '有帮助才能做' },
  '↑': { level: 0, label: '想学', description: '未经验但想尝试' }
};

// 同じ意味で入力されがちな別字形。必要なら文字を足す（例: '◎' を ☆ 扱いにするなど）
const SYMBOL_VARIANTS = {
  '☆': '☆★✩✪⭐',
  '○': '○〇◯⭕',
  '△': '△▲▵',
  '↑': '↑⇧⬆'
};

function normalizeText(value) {
  return String(value ?? '').normalize('NFKC').replace(/\s+/g, ' ').trim();
}

/** 記号セルを正規化。 known=false は「認識できない文字」 */
function normalizeMark(value) {
  const text = String(value ?? '')
    .normalize('NFKC')
    .replace(/[\s\u200b-\u200d\ufeff\ufe0e\ufe0f]/g, '');
  if (!text) return { symbol: '', known: true };
  const first = [...text][0];
  for (const [symbol, chars] of Object.entries(SYMBOL_VARIANTS)) {
    if ([...chars].includes(first)) return { symbol, known: true };
  }
  return { symbol: text, known: false };
}

function toDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString().slice(0, 10);
  if (typeof value === 'number' && value > 20000 && value < 80000) {
    return new Date(Math.round((value - 25569) * 86400000)).toISOString().slice(0, 10); // Excelシリアル値
  }
  const m = String(value ?? '').normalize('NFKC').match(/(\d{4})\D+(\d{1,2})\D+(\d{1,2})/);
  return m ? `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}` : null;
}

function parseSkillSheet(fileBuffer) {
  const XLSX = require('xlsx');
  let workbook;
  try {
    workbook = XLSX.read(fileBuffer, { type: 'buffer' });
  } catch (error) {
    throw new Error(`Excel解析失败: ${error.message}`);
  }
  // 技能表不一定是第一个工作表：取第一个 A 列出现「カテゴリ」的工作表
  for (const sheetName of workbook.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: '', raw: true });
    if (rows.some((row) => normalizeText(row?.[0]) === 'カテゴリ')) {
      return { ...extractEmployeeSkills(rows), sheetName };
    }
  }
  throw new Error('所有工作表里都找不到「カテゴリ」表头行，请确认使用的是技能自己评价表。');
}

function extractEmployeeSkills(rows) {
  // 表头行：A列が「カテゴリ」の行（日付の有無には頼らない）
  const headerIndex = rows.findIndex((row) => normalizeText(row?.[0]) === 'カテゴリ');
  if (headerIndex === -1) throw new Error('找不到表头行（A列应为「カテゴリ」），请确认使用的是技能自己评价表。');

  const header = rows[headerIndex];
  const dateRow = rows[headerIndex + 1] || [];
  const warnings = [];
  const employees = [];
  const seen = new Set();

  for (let col = 3; col < header.length; col++) {
    const name = normalizeText(header[col]);
    if (!name) continue;
    if (seen.has(name)) {
      warnings.push({ type: 'duplicate_employee', message: `员工「${name}」在表头出现多次，只采用第一列。` });
      continue;
    }
    seen.add(name);
    employees.push({ name, col, updatedOn: toDate(dateRow[col]) });
  }
  if (!employees.length) throw new Error('表头行里没有找到员工姓名（应从D列开始）。');

  const skills = new Map();
  const records = new Map();
  const unknownCells = [];
  let band = '';

  for (let r = headerIndex + 1; r < rows.length; r++) {
    const row = rows[r] || [];
    const skillName = String(row[2] ?? '').replace(/\s+/g, ' ').trim();
    const categoryCell = String(row[0] ?? '').replace(/\s+/g, ' ').trim();
    if (!skillName) { // 区切り行（例:「通信技術」の帯）や更新日の行
      if (categoryCell) band = categoryCell;
      continue;
    }
    const category = categoryCell || band || '其他';
    const key = `${normalizeText(category)}|${normalizeText(skillName)}`;
    if (skills.has(key)) {
      warnings.push({ type: 'duplicate_skill', message: `第${r + 1}行「${skillName}」与前面的行重复，后面的内容会覆盖前面的。` });
    } else {
      skills.set(key, { key, category, classification: String(row[1] ?? '').trim(), name: skillName, sortOrder: skills.size });
    }

    for (const emp of employees) {
      const raw = row[emp.col];
      const { symbol, known } = normalizeMark(raw);
      if (!known) {
        unknownCells.push({
          row: r + 1, col: emp.col + 1, employee: emp.name, skill: skillName, value: String(raw),
          codes: [...String(raw)].map((ch) => 'U+' + ch.codePointAt(0).toString(16).toUpperCase().padStart(4, '0'))
        });
        continue;
      }
      if (symbol === '') continue; // 空欄＝未経験は保存しない
      const def = LEVELS[symbol];
      records.set(`${emp.name}\u0000${key}`, {
        employeeName: emp.name, skillKey: key, level: def.level, symbol, description: def.description
      });
    }
  }

  return { headerRow: headerIndex + 1, employees, skills: [...skills.values()], records: [...records.values()], unknownCells, warnings };
}

function generateImportSummary(parsed) {
  const count = (fn) => parsed.records.filter(fn).length;
  return {
    employees: parsed.employees.length,
    skills: parsed.skills.length,
    records: parsed.records.length,
    proficiency: {
      expert: count((r) => r.level === 3),
      capable: count((r) => r.level === 2),
      support: count((r) => r.level === 1),
      wish: count((r) => r.symbol === '↑')
    },
    unknownCount: parsed.unknownCells.length
  };
}

module.exports = { parseSkillSheet, extractEmployeeSkills, generateImportSummary, normalizeMark, normalizeText, LEVELS };
