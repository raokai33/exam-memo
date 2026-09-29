// 生成物 —— 由 tools/bundle.mjs 从 src/** 打包，请勿手改
(function () {
'use strict';
var __MODULES__ = {};
(function () {
// 文本层 —— XML 实体解码、空白折叠、单元格显示文本（设计档 §2.5.3 / §2.5.4）。
// 除本文件写死的规则外，不改动任何字符：不删标点、不做繁简转换、不修错别字。

const NAMED_ENTITIES = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };

// 空位括号形态（§2.6.1）：（） （ ） () ( ) 统一计数
const BLANK_RE = /[（(]\s*[)）]/g;

// ECMA-376 内置数字格式中「纯数值 / 百分比 / 千分位」形态（§2.5.4 第 2 条）
const BUILTIN_NUMERIC_FORMATS = {
  1: '0',
  2: '0.00',
  3: '#,##0',
  4: '#,##0.00',
  9: '0%',
  10: '0.00%',
  11: '0.00E+00',
  37: '#,##0 ;(#,##0)',
  38: '#,##0 ;[Red](#,##0)',
  39: '#,##0.00;(#,##0.00)',
  40: '#,##0.00;[Red](#,##0.00)',
  48: '##0.0E+0',
};

const EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 30); // 1900 日期系统（含 1900 闰年错位）
const EXCEL_MAX_SERIAL = 2958465; // 9999-12-31

/** 单趟扫描的 XML 实体解码（不级联）。 */
function decodeXmlEntities(text) {
  if (text.indexOf('&') < 0) return text;
  return text.replace(/&(#[0-9]+|#x[0-9a-fA-F]+|[a-zA-Z]+);/g, (whole, body) => {
    if (body.charAt(0) === '#') {
      const code = body.charAt(1) === 'x' || body.charAt(1) === 'X'
        ? parseInt(body.slice(2), 16)
        : parseInt(body.slice(1), 10);
      if (!Number.isFinite(code) || code < 0 || code > 0x10ffff) return whole;
      return String.fromCodePoint(code);
    }
    return Object.prototype.hasOwnProperty.call(NAMED_ENTITIES, body) ? NAMED_ENTITIES[body] : whole;
  });
}

/** 空白折叠：\s+（含 NBSP、全角空格、\r\n\t）→ 单个半角空格，并 trim。 */
function collapseWhitespace(text) {
  return text.replace(/\s+/g, ' ').trim();
}

/** 规整：实体解码 + 空白折叠。 */
function tidyText(text) {
  return collapseWhitespace(decodeXmlEntities(String(text)));
}

/** 题干空位括号个数（§2.6.1 的正则）。 */
function countBlanks(stem) {
  const matched = String(stem).match(BLANK_RE);
  return matched ? matched.length : 0;
}

/** Excel 1900 日期系统序列号 → YYYY-MM-DD；越界或非整数 → null。 */
function excelSerialToDate(serial) {
  if (!Number.isInteger(serial) || serial < 1 || serial > EXCEL_MAX_SERIAL) return null;
  const date = new Date(EXCEL_EPOCH_UTC + serial * 86400000);
  const y = date.getUTCFullYear();
  const m = String(date.getUTCMonth() + 1).padStart(2, '0');
  const d = String(date.getUTCDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

/**
 * 数值单元格显示文本（§2.5.4）：三分口径。
 * @param {string} raw `<v>` 原文
 * @param {number} numFmtId 单元格样式指向的数字格式 id
 * @param {boolean} formatDefinedInFile 该 id 是否在本文件 <numFmts> 中有定义
 * @returns {{text: string, formatUnresolved: boolean}}
 */
function formatNumberCell(raw, numFmtId, formatDefinedInFile) {
  const text = String(raw).trim();
  const id = Number.isFinite(numFmtId) ? Math.trunc(numFmtId) : -1;
  const value = Number(text);

  // ① 常规（s 缺省或 numFmtId = 0）→ 原文规整
  if (id <= 0) return { text, formatUnresolved: false };

  if (id < 50) {
    // ② 内置表中的纯数值 / 百分比 / 千分位 → 按内置格式渲染
    if (Number.isFinite(value)) {
      const rendered = renderBuiltinNumber(id, value);
      if (rendered !== null) return { text: rendered, formatUnresolved: false };
    }
    // 内置表内的日期 / 时间 / 货币 / 分数等形态：写死规则未覆盖 → 原值 + 标注，不猜显示形态。
    // （真表零命中：本表带数字格式的 18 个数值单元格只用 9 / 3 / 58 三种 id。）
    return { text, formatUnresolved: true };
  }

  // 文件内自定义格式（本表无）→ 原样 + 标注
  if (formatDefinedInFile) return { text, formatUnresolved: true };

  // ③ 格式串缺失（numFmtId ≥ 50 且文件内无定义）→ 日期优先 + 原值标注（用户裁决 D11 / OQ-1）
  const iso = excelSerialToDate(value);
  if (iso !== null) return { text: `换算日期 ${iso}（原表数值 ${text}）`, formatUnresolved: true };
  return { text, formatUnresolved: true };
}

/**
 * 单元格显示文本（§2.5.3）。
 * @param {{type?: string, value?: string, inlineXml?: string, sharedStrings?: string[],
 *          numFmtId?: number, formatDefinedInFile?: boolean}} cell
 * @returns {{text: string, formatUnresolved: boolean, note: string|null}}
 */
function parseCellText(cell) {
  const shared = cell.sharedStrings || [];
  const type = cell.type || '';
  if (type === 's') {
    const index = Number(cell.value);
    if (!Number.isInteger(index) || index < 0 || index >= shared.length) {
      return { text: '', formatUnresolved: false, note: `共享字符串下标越界：${String(cell.value)}` };
    }
    return { text: tidyText(shared[index]), formatUnresolved: false, note: null };
  }
  if (type === 'inlineStr') {
    return { text: tidyText(extractInlineString(cell.inlineXml || '')), formatUnresolved: false, note: null };
  }
  if (type === 'str') {
    return { text: tidyText(cell.value || ''), formatUnresolved: false, note: null };
  }
  if (type === 'b') {
    return { text: String(cell.value || '').trim(), formatUnresolved: false, note: null };
  }
  const formatted = formatNumberCell(cell.value === undefined || cell.value === null ? '' : cell.value, cell.numFmtId, !!cell.formatDefinedInFile);
  return { text: formatted.text, formatUnresolved: formatted.formatUnresolved, note: null };
}

/** 取内联字符串：<is> 下所有 <t> 顺序拼接，剔除 <rPh> 注音块。 */
function extractInlineString(xml) {
  const cleaned = xml.replace(/<rPh\b[\s\S]*?<\/rPh>/g, '');
  let out = '';
  const re = /<t\b[^>]*>([\s\S]*?)<\/t>/g;
  let m;
  while ((m = re.exec(cleaned)) !== null) out += m[1];
  return out;
}

function renderBuiltinNumber(id, value) {
  const abs = Math.abs(value);
  switch (id) {
    case 1:
      return String(Math.round(value));
    case 2:
      return value.toFixed(2);
    case 3:
      return (value < 0 ? '-' : '') + groupDigits(abs.toFixed(0));
    case 4:
      return (value < 0 ? '-' : '') + groupDigits(abs.toFixed(2));
    case 9:
      return `${Math.round(value * 100)}%`;
    case 10:
      return `${(value * 100).toFixed(2)}%`;
    case 11:
      return scientificText(value, 2, 2);
    case 37:
    case 38:
    case 39:
    case 40: {
      const decimals = id === 39 || id === 40 ? 2 : 0;
      const body = groupDigits(abs.toFixed(decimals));
      return value < 0 ? `(${body})` : body;
    }
    case 48:
      return scientificText(value, 1, 1);
    default:
      return null;
  }
}

function groupDigits(text) {
  const dot = text.indexOf('.');
  const int = dot < 0 ? text : text.slice(0, dot);
  const frac = dot < 0 ? '' : text.slice(dot);
  return int.replace(/\B(?=(\d{3})+(?!\d))/g, ',') + frac;
}

function scientificText(value, decimals, minExpDigits) {
  const sign = value < 0 ? '-' : '';
  const abs = Math.abs(value);
  const exp = abs === 0 ? 0 : Math.floor(Math.log10(abs));
  const mantissa = abs === 0 ? 0 : abs / Math.pow(10, exp);
  const expSign = exp < 0 ? '-' : '+';
  const expDigits = String(Math.abs(exp)).padStart(minExpDigits, '0');
  return `${sign}${mantissa.toFixed(decimals)}E${expSign}${expDigits}`;
}

__MODULES__["src/core/text.mjs"] = { decodeXmlEntities: decodeXmlEntities, collapseWhitespace: collapseWhitespace, tidyText: tidyText, countBlanks: countBlanks, excelSerialToDate: excelSerialToDate, formatNumberCell: formatNumberCell, parseCellText: parseCellText, extractInlineString: extractInlineString, BUILTIN_NUMERIC_FORMATS: BUILTIN_NUMERIC_FORMATS };
}());
(function () {
// 判分层 —— 判分 + 展示 / 比对两条独立路径（设计档 §2.4.1 / §2.8 / KD-12）。
// 铁律：displayText 与 normalizeForCompare 必须分开；归一化只发生在比对内部，绝不改变展示文本。

const { countBlanks } = __MODULES__["src/core/text.mjs"];

const PUNCTUATION_EQUIVALENTS = {
  '，': ',', '、': ',', '。': '.', '；': ';', '：': ':', '！': '!', '？': '?',
  '（': '(', '）': ')', '“': '"', '”': '"', '‘': "'", '’': "'",
};

/**
 * 比对用归一化（§2.8.1）：删空白 + 中英文标点等价 + 全角字母数字 → 半角。
 * @param {string} text
 * @returns {string}
 */
function normalizeForCompare(text) {
  if (text === undefined || text === null) return '';
  let out = toHalfWidth(String(text));
  out = out.replace(/[，、。；：！？（）“”‘’]/g, (c) => PUNCTUATION_EQUIVALENTS[c]);
  return out.replace(/\s+/g, '');
}

/**
 * 展示用答案文本（§2.8.1）—— 任何情况下都是源表原文，不做归一化。
 * @param {{type: string, answer: {raw?: string, display?: string, judge?: string|null}}} question
 * @returns {string}
 */
function displayText(question) {
  const answer = (question && question.answer) || {};
  const raw = answer.raw === undefined || answer.raw === null ? '' : String(answer.raw);
  if (question && question.type === 'judge') {
    if (answer.judge === 'correct') return '正确';
    if (answer.judge === 'wrong') return '错误';
    return raw; // 白名单外：展示原文，只标注不猜
  }
  if (answer.display !== undefined && answer.display !== null && answer.display !== '') return String(answer.display);
  return raw;
}

/**
 * 判分（§2.8.2 / §2.8.3）。
 * @param {object} question
 * @param {{letter?: string, letters?: string[], judge?: string, values?: string[]}} response
 * @returns {{correct: boolean, expectedDisplay: string}}
 */
function grade(question, response) {
  const expectedDisplay = displayText(question);
  const type = question && question.type;
  const answer = (question && question.answer) || {};
  const reply = response || {};

  if (type === 'essay') throw new Error('简答题不判分（简答只进背题模式，D5）');

  if (type === 'single') {
    const letters = answer.letters || [];
    const picked = reply.letter !== undefined ? reply.letter : (reply.letters || [])[0];
    return { correct: letters.length === 1 && picked === letters[0], expectedDisplay };
  }

  if (type === 'multi') {
    const want = Array.from(answer.letters || []).sort();
    const got = Array.from(reply.letters || []).sort();
    const correct = want.length > 0 && want.length === got.length && want.every((l, i) => l === got[i]);
    return { correct, expectedDisplay }; // 少选 / 多选 / 错选一律判错（严格，D4）
  }

  if (type === 'judge') {
    const want = answer.judge;
    const got = reply.judge;
    return { correct: !!want && want === got, expectedDisplay };
  }

  if (type === 'fill') {
    const raw = answer.raw === undefined || answer.raw === null ? '' : String(answer.raw);
    const blanks = Number(answer.blanks) || countBlanks(question.stem || '');
    const values = Array.isArray(reply.values) ? reply.values : [];
    if (raw !== '' && blanks <= 1) {
      return { correct: normalizeForCompare(values[0] || '') === normalizeForCompare(raw), expectedDisplay };
    }
    if (raw !== '') {
      // 第一道闸：整串比对（各空输入按空位顺序直接拼接，不加分隔符）
      const joined = values.join('');
      if (normalizeForCompare(joined) === normalizeForCompare(raw)) return { correct: true, expectedDisplay };
      // 第二道闸：逐空比对（仅当答案成功切分出 n 段时）
      const segments = answer.segments;
      if (Array.isArray(segments) && segments.length === blanks) {
        const ok = segments.every((segment, i) => normalizeForCompare(values[i] || '') === normalizeForCompare(segment));
        if (ok) return { correct: true, expectedDisplay };
      }
    }
    return { correct: false, expectedDisplay };
  }

  throw new Error(`未知题型：${String(type)}`);
}

function toHalfWidth(text) {
  let out = '';
  for (const ch of text) {
    const code = ch.codePointAt(0);
    out += code >= 0xff01 && code <= 0xff5e ? String.fromCharCode(code - 0xfee0) : ch;
  }
  return out;
}

__MODULES__["src/core/grade.mjs"] = { normalizeForCompare: normalizeForCompare, displayText: displayText, grade: grade };
}());
(function () {
// 解压层 —— 能力探测 + deflate-raw 解压（设计档 §2.2.2 / §2.4.1 / KD-11）。
//
// 生产路径只有这一条实现：浏览器与 Node 构建脚本都走 DecompressionStream('deflate-raw')，
// 因此“浏览器里跑的那段代码”在 Node 测试里被真跑，而不是被模拟。
// node:zlib 不得出现在本文件（它只允许出现在测试面，作独立编码器 / 独立复算基准）。

// deflate-raw("") 的字节：固定 Huffman 空块（BFINAL=1, BTYPE=01）。
const EMPTY_DEFLATE_RAW = new Uint8Array([0x03, 0x00]);

/**
 * 解压一段 raw-deflate（ZIP method 8）字节。
 * @param {Uint8Array} bytes
 * @returns {Promise<Uint8Array>}
 */
async function inflateRaw(bytes) {
  const decompressed = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
  const buffer = await new Response(decompressed).arrayBuffer();
  return new Uint8Array(buffer);
}

/**
 * 能力探测：真实解压一段已知字节，而不是只检测 API 是否存在。
 * @returns {Promise<{ok: boolean, reason: string}>}
 */
async function probeInflate() {
  if (typeof DecompressionStream !== 'function') {
    return {
      ok: false,
      reason: '此浏览器不支持 DecompressionStream（需 iOS Safari 16.4 及以上）。请改用单文件形态（题库已在文件内，无需解压）。',
    };
  }
  try {
    const out = await inflateRaw(EMPTY_DEFLATE_RAW);
    if (out.length !== 0) {
      return { ok: false, reason: 'deflate-raw 解压结果异常，本浏览器无法读取 xlsx。请改用单文件形态。' };
    }
    return { ok: true, reason: '' };
  } catch (err) {
    const message = err && err.message ? err.message : String(err);
    return {
      ok: false,
      reason: `deflate-raw 不可用（${message}）。请改用单文件形态（题库已在文件内，无需解压）。`,
    };
  }
}

__MODULES__["src/core/inflate.mjs"] = { inflateRaw: inflateRaw, probeInflate: probeInflate };
}());
(function () {
// 进度层 —— 组序 / 游标 / 正确率 / 错题本 / 掌握度 / 乱序的纯逻辑
// （设计档 §2.4.1 / §2.10.3 / §2.11.1 / §2.11.6 / KD-12 / KD-21）。
// 无任何 IO：可完全单测；存储由 store.mjs 负责。

const MODES = ['memorize', 'practice'];
const MEMORIZE_TYPES = ['all', 'single', 'multi', 'judge', 'fill', 'essay'];
const PRACTICE_TYPES = ['all', 'single', 'multi', 'judge', 'fill', 'essay'];
const QUESTION_TYPES = ['single', 'multi', 'judge', 'fill', 'essay'];
const WRONGBOOK = 'wrongbook';
const NUMBERS = 'numbers';

const MODE_LABELS = { memorize: '背题', practice: '练习' };
const TYPE_LABELS = {
  all: '全部', single: '单选', multi: '多选', judge: '判断', fill: '填空', essay: '简答',
  wrongbook: '错题本', numbers: '数字专项',
};

/** 组键 = `mode|type`（§2.10.3）。 */
function groupKey(mode, type) {
  return `${mode}|${type}`;
}

/** 集合型入口：题集随作答实时变化，不持久化 order / index，游标 = cursorQid（§2.10.3）。 */
const COLLECTION_TYPES = [WRONGBOOK, NUMBERS];

function isCollection(type) {
  return type === WRONGBOOK || type === NUMBERS;
}

/** 入口矩阵：背题 × 6 + 练习 × 6 + 错题本 + 数字专项 = 14 组（§2.11.1）。 */
function entryMatrix() {
  const out = [];
  for (const mode of MODES) {
    const types = mode === 'practice' ? PRACTICE_TYPES : MEMORIZE_TYPES;
    for (const type of types) out.push({ mode, type, groupKey: groupKey(mode, type) });
  }
  for (const type of COLLECTION_TYPES) out.push({ mode: 'practice', type, groupKey: groupKey('practice', type) });
  return out;
}

/**
 * 该入口覆盖的题型集合：两个集合型入口 = 全 5 题型（错题本含简答）；
 * 「练习 × 全部」= 4 类，不含简答（§2.11.1）。
 */
function typesForGroup(mode, type) {
  if (isCollection(type)) return new Set(QUESTION_TYPES);
  if (type === 'all') return new Set(mode === 'practice' ? ['single', 'multi', 'judge', 'fill'] : QUESTION_TYPES);
  return new Set([type]);
}

/** 数字专项判据（R18）：题干或答案列原文含数字，半角 / 全角都算。 */
function isNumberQuestion(question) {
  const stem = question && typeof question.stem === 'string' ? question.stem : '';
  const raw = question && question.answer && typeof question.answer.raw === 'string' ? question.answer.raw : '';
  return /[0-9０-９]/.test(stem) || /[0-9０-９]/.test(raw);
}

/**
 * 选出该入口的题目（顺序 = 源表顺序，questions 已按表顺序 → 行号升序）。
 * @param {Array} questions
 * @param {string} mode
 * @param {string} type
 * @param {Iterable<string>} [wrongQids]
 */
function selectQuestions(questions, mode, type, wrongQids) {
  if (mode === 'practice' && type === WRONGBOOK) {
    const set = wrongQids instanceof Set ? wrongQids : new Set(wrongQids || []);
    return questions.filter((q) => set.has(q.id));
  }
  if (mode === 'practice' && type === NUMBERS) return questions.filter(isNumberQuestion);
  const types = typesForGroup(mode, type);
  return questions.filter((q) => types.has(q.type));
}

function emptyProgress(group) {
  return {
    groupKey: group,
    order: [],
    index: 0,
    cursorQid: null,
    stats: { attempts: 0, correct: 0 },
    response: {},
  };
}

function ensureProgress(progress, group) {
  if (progress && progress.groupKey === group) return progress;
  return emptyProgress(group);
}

/**
 * 记录一次作答：只有【首次作答】进统计（KD-12）；再次作答只返回反馈，不改统计。
 * @returns {{progress: object, first: boolean}}
 */
function recordAnswer(progress, qid, correct) {
  const next = {
    ...progress,
    stats: { ...progress.stats },
    response: { ...progress.response },
    cursorQid: qid,
  };
  if (Object.prototype.hasOwnProperty.call(next.response, qid)) return { progress: next, first: false };
  next.response[qid] = !!correct;
  next.stats.attempts += 1;
  if (correct) next.stats.correct += 1;
  return { progress: next, first: true };
}

/** 正确率（首次作答口径）；无作答 → null。 */
function accuracy(progress) {
  const attempts = progress && progress.stats ? progress.stats.attempts : 0;
  if (!attempts) return null;
  return progress.stats.correct / attempts;
}

function addWrong(wrongSet, qid) {
  const next = new Set(wrongSet || []);
  next.add(qid);
  return next;
}

function removeWrong(wrongSet, qid) {
  const next = new Set(wrongSet || []);
  next.delete(qid);
  return next;
}

/** 答错一次错次 +1（答对不减，R17）；入参不改，返回新对象。 */
function bumpWrongCount(counts, qid) {
  const next = { ...(counts || {}) };
  next[qid] = (Number(next[qid]) || 0) + 1;
  return next;
}

/**
 * 错题本组顺序（R17）：错次降序，平手按源表下标升序 —— 全序确定、可复现。
 * @param {Array<string>} questionIds 源表顺序的题号
 * @param {Set<string>|Iterable<string>} wrongSet
 * @param {object} counts qid → 错次
 */
function wrongbookOrderByWeight(questionIds, wrongSet, counts) {
  const set = wrongSet instanceof Set ? wrongSet : new Set(wrongSet || []);
  const weight = counts || {};
  const sourceIndex = new Map();
  questionIds.forEach((id, i) => {
    if (!sourceIndex.has(id)) sourceIndex.set(id, i);
  });
  return questionIds
    .filter((id) => set.has(id))
    .sort((a, b) => (Number(weight[b]) || 0) - (Number(weight[a]) || 0) || sourceIndex.get(a) - sourceIndex.get(b));
}

/**
 * 确定性乱序（R20）：键 = (hash32(seed|qid), 源表下标) 升序 —— 同种子同序。
 * 集合型入口不参与（题集实时变化，见 orderFor）。
 */
function shuffledOrder(ids, seed) {
  const sourceIndex = new Map();
  ids.forEach((id, i) => {
    if (!sourceIndex.has(id)) sourceIndex.set(id, i);
  });
  return ids.slice().sort((a, b) => {
    const diff = hash32(`${seed}|${a}`) - hash32(`${seed}|${b}`);
    return diff || sourceIndex.get(a) - sourceIndex.get(b);
  });
}

/** FNV-1a 32 位散列（与题库指纹同族，§2.10.3）。 */
function hash32(text) {
  const value = String(text);
  let hash = 2166136261 >>> 0;
  for (let i = 0; i < value.length; i++) {
    hash ^= value.charCodeAt(i);
    hash = Math.imul(hash, 16777619) >>> 0;
  }
  return hash >>> 0;
}

/**
 * 掌握度三分（互斥完备，KD-21 / §2.11.6①）：跨【全部已存进度组】聚合（含两个集合型组）。
 * mastered = 任一组答案为 true；failed = 非 mastered 且任一组为 false；untrained = 任何组都没有该题记录。
 * @param {Array} questions
 * @param {Iterable<object>} progresses 各组的 progress（可含 null）
 * @returns {Map<string, 'mastered'|'failed'|'untrained'>}
 */
function masteryOf(questions, progresses) {
  const seen = new Map();
  for (const progress of progresses || []) {
    const response = progress && progress.response;
    if (!response) continue;
    for (const qid of Object.keys(response)) {
      if (response[qid] === true) seen.set(qid, true);
      else if (response[qid] === false && seen.get(qid) !== true) seen.set(qid, false);
    }
  }
  const out = new Map();
  for (const question of questions || []) {
    const mark = seen.get(question.id);
    out.set(question.id, mark === true ? 'mastered' : mark === false ? 'failed' : 'untrained');
  }
  return out;
}

/** 掌握度计数：total = mastered + failed + untrained（互斥完备，可机器校验）。 */
function masterySummary(questions, progresses) {
  const mastery = masteryOf(questions, progresses);
  let mastered = 0;
  let failed = 0;
  let untrained = 0;
  for (const question of questions || []) {
    const state = mastery.get(question.id);
    if (state === 'mastered') mastered += 1;
    else if (state === 'failed') failed += 1;
    else untrained += 1;
  }
  return { total: (questions || []).length, mastered, failed, untrained };
}

/** 「只练没掌握的」过滤（R16）：剔掉已掌握者，未掌握与未练都留下。 */
function selectUnmastered(questions, mastery) {
  return (questions || []).filter((question) => mastery.get(question.id) !== 'mastered');
}

/** 游标定位：找不到（被移出 / 集合变了）→ 从首题开始。 */
function cursorIndex(order, cursorQid) {
  if (!cursorQid) return 0;
  const index = order.indexOf(cursorQid);
  return index < 0 ? 0 : index;
}

/** 集合型入口不持久化 order / index（§2.10.3）。 */
function progressForStore(progress, group) {
  const type = String(group || '').split('|')[1];
  if (!isCollection(type)) return progress;
  const { order, index, ...rest } = progress;
  return rest;
}

__MODULES__["src/core/progress.mjs"] = { groupKey: groupKey, isCollection: isCollection, entryMatrix: entryMatrix, typesForGroup: typesForGroup, isNumberQuestion: isNumberQuestion, selectQuestions: selectQuestions, emptyProgress: emptyProgress, ensureProgress: ensureProgress, recordAnswer: recordAnswer, accuracy: accuracy, addWrong: addWrong, removeWrong: removeWrong, bumpWrongCount: bumpWrongCount, wrongbookOrderByWeight: wrongbookOrderByWeight, shuffledOrder: shuffledOrder, hash32: hash32, masteryOf: masteryOf, masterySummary: masterySummary, selectUnmastered: selectUnmastered, cursorIndex: cursorIndex, progressForStore: progressForStore, MODES: MODES, MEMORIZE_TYPES: MEMORIZE_TYPES, PRACTICE_TYPES: PRACTICE_TYPES, QUESTION_TYPES: QUESTION_TYPES, WRONGBOOK: WRONGBOOK, NUMBERS: NUMBERS, MODE_LABELS: MODE_LABELS, TYPE_LABELS: TYPE_LABELS, COLLECTION_TYPES: COLLECTION_TYPES };
}());
(function () {
// 存储层 —— 三级后端 + 启动探测 + 降级报告（设计档 §2.4.1 / §2.10）。
// 探测方式是「写—读—比—删」四步，而不是检测 API 是否存在（§2.10.1）。

const BACKENDS = ['indexeddb', 'localstorage', 'memory'];
const DB_NAME = 'exam-memo';
const OBJECT_STORE = 'kv';
const BANK_KEY = 'bank';
const BANK_META_KEY = 'bank:meta';
const BANK_PART_PREFIX = 'bank:part:';

/**
 * 打开存储层（探测顺序 indexeddb → localstorage → memory）。
 * @param {{probeTimeoutMs?: number, indexedDB?: object, localStorage?: object}} [options]
 * @returns {Promise<object>}
 */
async function openStore(options) {
  const opts = options || {};
  const probeTimeoutMs = Number.isFinite(opts.probeTimeoutMs) ? opts.probeTimeoutMs : 3000;
  const idbFactory = pickGlobal(opts, 'indexedDB');
  const ls = pickGlobal(opts, 'localStorage');

  const backends = [
    { name: 'indexeddb', impl: idbFactory ? indexedDbBackend(idbFactory) : null, reason: idbFactory ? '' : '浏览器未提供 IndexedDB' },
    { name: 'localstorage', impl: ls ? localStorageBackend(ls) : null, reason: ls ? '' : '浏览器未提供 localStorage' },
    { name: 'memory', impl: memoryBackend(), reason: '内存后端（会话内有效）' },
  ];

  const notices = [];
  let index = 0;
  for (let i = 0; i < backends.length; i++) {
    const candidate = backends[i];
    if (!candidate.impl) {
      notices.push(`${candidate.name}：${candidate.reason}`);
      continue;
    }
    try {
      await withTimeout(candidate.impl.probe(), probeTimeoutMs, `存储探测超时（${probeTimeoutMs} ms）`);
      index = i;
      if (i > 0) notices.push(`已回退到 ${candidate.name}（${backends[0].name} 不可用）`);
      break;
    } catch (err) {
      notices.push(`${candidate.name} 探测失败：${errorMessage(err)}`);
      index = Math.min(i + 1, backends.length - 1);
    }
  }

  let persisted = false;
  if (backends[index].name === 'indexeddb') persisted = await requestPersist();

  let degraded = null;

  function demote(from, to, err) {
    index = to;
    degraded = `${backends[from].name} 写入失败（${errorMessage(err)}），已降级到 ${backends[to].name}`;
  }

  async function run(operation) {
    let lastError = null;
    for (let i = index; i < backends.length; i++) {
      try {
        const result = await operation(backends[i].impl, backends[i]);
        if (i > index) index = i;
        return result;
      } catch (err) {
        lastError = err;
        if (i + 1 < backends.length) demote(i, i + 1, err);
      }
    }
    throw lastError || new Error('存储写入失败');
  }

  const store = {
    async saveBank(bank) {
      await run((impl, backend) => saveBankRecord(impl, backend.name, bank));
    },
    async loadBank() {
      return await run((impl, backend) => loadBankRecord(impl, backend.name));
    },
    async saveProgress(group, data) {
      await run((impl) => impl.set(`progress:${group}`, data));
    },
    async loadProgress(group) {
      const value = await run((impl) => impl.get(`progress:${group}`));
      return value === undefined ? null : value;
    },
    async saveWrongbook(data) {
      await run((impl) => impl.set('wrongbook', data));
    },
    async loadWrongbook() {
      const value = await run((impl) => impl.get('wrongbook'));
      return value === undefined ? null : value;
    },
    async saveUi(data) {
      await run((impl) => impl.set('ui', data));
    },
    async loadUi() {
      const value = await run((impl) => impl.get('ui'));
      return value === undefined ? null : value;
    },
    async saveKey(key, value) {
      await run((impl) => impl.set(key, value));
    },
    async loadKey(key) {
      const value = await run((impl) => impl.get(key));
      return value === undefined ? null : value;
    },
    async clear() {
      await run(async (impl) => {
        const keys = await impl.keys('');
        for (const key of keys) await impl.del(key);
      });
    },
    describe() {
      const name = backends[index].name;
      return {
        backend: name,
        persistent: name !== 'memory',
        persisted,
        degraded,
        notices: notices.slice(),
      };
    },
  };

  Object.defineProperty(store, 'backend', { get: () => backends[index].name });
  Object.defineProperty(store, 'persisted', { get: () => persisted });
  return store;
}

function pickGlobal(options, key) {
  // 显式传 null 表示「本环境没有这个后端」（测试注入用）；未传则读全局
  if (options && Object.prototype.hasOwnProperty.call(options, key)) return options[key] || null;
  try {
    return globalThis[key] || null;
  } catch (err) {
    return null; // 某些沙箱下访问 globalThis.localStorage 会抛 SecurityError
  }
}

/**
 * 换表：把当前各组进度整体备份到 `progress:<group>@old`（§2.10.3）。
 * @returns {Promise<number>} 实际备份的组数
 */
async function backupProgress(store, groupKeys) {
  let count = 0;
  for (const group of groupKeys || []) {
    try {
      const current = await store.loadProgress(group);
      if (current) {
        await store.saveKey(`progress:${group}@old`, current);
        count += 1;
      }
    } catch (err) {
      console.error('旧进度备份失败', group, err);
    }
  }
  return count;
}

async function saveBankRecord(impl, backendName, bank) {
  if (backendName !== 'localstorage') {
    await impl.set(BANK_KEY, bank);
    await impl.del(BANK_META_KEY);
    return;
  }
  // localStorage 分包写：按表分片（§2.10.2 分片键写死）
  const groups = new Map();
  for (const question of bank.questions || []) {
    if (!groups.has(question.sheet)) groups.set(question.sheet, []);
    groups.get(question.sheet).push(question);
  }
  const order = (bank.sheets || []).map((s) => s.name).filter((name) => groups.has(name));
  for (const name of groups.keys()) if (!order.includes(name)) order.push(name);
  for (let i = 0; i < order.length; i++) await impl.set(`${BANK_PART_PREFIX}${i}`, groups.get(order[i]));
  await impl.set(BANK_META_KEY, {
    schemaVersion: bank.schemaVersion,
    bankFingerprint: bank.bankFingerprint,
    sourceName: bank.sourceName,
    builtAt: bank.builtAt,
    sheets: bank.sheets || [],
    partCount: order.length,
  });
  await impl.del(BANK_KEY);
}

async function loadBankRecord(impl, backendName) {
  if (backendName !== 'localstorage') {
    const value = await impl.get(BANK_KEY);
    return value === undefined ? null : value;
  }
  const meta = await impl.get(BANK_META_KEY);
  if (!meta) return null;
  const questions = [];
  for (let i = 0; i < meta.partCount; i++) {
    const part = await impl.get(`${BANK_PART_PREFIX}${i}`);
    if (Array.isArray(part)) questions.push(...part);
  }
  return {
    schemaVersion: meta.schemaVersion,
    bankFingerprint: meta.bankFingerprint,
    sourceName: meta.sourceName,
    builtAt: meta.builtAt,
    sheets: meta.sheets || [],
    questions,
  };
}

function indexedDbBackend(factory) {
  let dbPromise = null;
  const db = () => {
    if (!dbPromise) dbPromise = idbOpen(factory);
    return dbPromise;
  };
  return {
    async probe() {
      const handle = await db();
      await idbRequest(handle.transaction(OBJECT_STORE, 'readwrite').objectStore(OBJECT_STORE).put(1, '__probe__'));
      const readBack = await idbRequest(handle.transaction(OBJECT_STORE, 'readonly').objectStore(OBJECT_STORE).get('__probe__'));
      if (readBack !== 1) throw new Error('IndexedDB 写后读回不一致');
      await idbRequest(handle.transaction(OBJECT_STORE, 'readwrite').objectStore(OBJECT_STORE).delete('__probe__'));
      return true;
    },
    async get(key) {
      const handle = await db();
      return await idbRequest(handle.transaction(OBJECT_STORE, 'readonly').objectStore(OBJECT_STORE).get(key));
    },
    async set(key, value) {
      const handle = await db();
      await idbRequest(handle.transaction(OBJECT_STORE, 'readwrite').objectStore(OBJECT_STORE).put(value, key));
    },
    async del(key) {
      const handle = await db();
      await idbRequest(handle.transaction(OBJECT_STORE, 'readwrite').objectStore(OBJECT_STORE).delete(key));
    },
    async keys() {
      const handle = await db();
      const all = await idbRequest(handle.transaction(OBJECT_STORE, 'readonly').objectStore(OBJECT_STORE).getAllKeys());
      return (all || []).filter((k) => typeof k === 'string');
    },
  };
}

function idbOpen(factory) {
  return new Promise((resolve, reject) => {
    const request = factory.open(DB_NAME, 1);
    request.onupgradeneeded = () => {
      const handle = request.result;
      if (!handle.objectStoreNames.contains(OBJECT_STORE)) handle.createObjectStore(OBJECT_STORE);
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error || new Error('IndexedDB 打开失败'));
    request.onblocked = () => reject(new Error('IndexedDB 被其他标签页阻塞'));
  });
}

function idbRequest(request) {
  return new Promise((resolve, reject) => {
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error || new Error('IndexedDB 请求失败'));
  });
}

function localStorageBackend(localStorage) {
  return {
    async probe() {
      const key = '__probe__';
      localStorage.setItem(key, '1');
      const readBack = localStorage.getItem(key);
      localStorage.removeItem(key);
      if (readBack !== '1') throw new Error('localStorage 写后读回不一致');
      return true;
    },
    async get(key) {
      const raw = localStorage.getItem(key);
      return raw === null ? undefined : JSON.parse(raw);
    },
    async set(key, value) {
      localStorage.setItem(key, JSON.stringify(value));
    },
    async del(key) {
      localStorage.removeItem(key);
    },
    async keys() {
      const out = [];
      for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        if (typeof key === 'string') out.push(key);
      }
      return out;
    },
  };
}

function memoryBackend() {
  const map = new Map();
  return {
    async probe() {
      map.set('__probe__', 1);
      if (map.get('__probe__') !== 1) throw new Error('内存后端写后读回不一致');
      map.delete('__probe__');
      return true;
    },
    async get(key) {
      return map.get(key);
    },
    async set(key, value) {
      map.set(key, value);
    },
    async del(key) {
      map.delete(key);
    },
    async keys() {
      return Array.from(map.keys());
    },
  };
}

async function requestPersist() {
  try {
    const nav = globalThis.navigator;
    if (nav && nav.storage && typeof nav.storage.persist === 'function') return !!(await nav.storage.persist());
  } catch (err) {
    return false; // 失败不阻塞（§2.2.3 方案 5）
  }
  return false;
}

function withTimeout(promise, ms, message) {
  if (!Number.isFinite(ms) || ms <= 0) return promise;
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error(message)), ms);
    promise.then(
      (value) => { clearTimeout(timer); resolve(value); },
      (err) => { clearTimeout(timer); reject(err); },
    );
  });
}

function errorMessage(err) {
  return err && err.message ? err.message : String(err);
}

__MODULES__["src/core/store.mjs"] = { openStore: openStore, backupProgress: backupProgress, BACKENDS: BACKENDS };
}());
(function () {
// 答案层 —— 列位置字母映射、归一化白名单、多空切分（设计档 §2.4.1 / §2.7）。
// 铁律：字母 ↔ 选项按【列位置】映射，不按“第几个非空选项”编号。

const { countBlanks } = __MODULES__["src/core/text.mjs"];

const CHOICE_SEPARATORS = ['、', '，', ',', '；', ';', '|'];
const FULLWIDTH_LETTERS = /[Ａ-Ｚ]/g;

/**
 * 选择类答案归一化（§2.7.1 第 1 条）。
 * @param {string} answerText
 * @returns {{ok: boolean, letters: string[], changed: boolean}}
 */
function normalizeChoiceAnswer(answerText) {
  const raw = String(answerText === undefined || answerText === null ? '' : answerText);
  let text = raw;
  for (const sep of CHOICE_SEPARATORS) text = text.split(sep).join('');
  text = text.replace(/\s+/g, '').replace(FULLWIDTH_LETTERS, (c) => String.fromCharCode(c.charCodeAt(0) - 0xfee0));
  // 空答案（非空字符一个也没有）不算「含非字母字符」——它是 noAnswer，不是 suspect
  if (text === '') return { ok: true, letters: [], changed: false };
  if (!/^[A-Z]+$/.test(text)) return { ok: false, letters: [], changed: text !== raw };
  const letters = Array.from(new Set(text.split(''))).sort();
  return { ok: true, letters, changed: text !== raw };
}

/**
 * 字母 → 选项（按列位置）。
 * @param {string} answerText 答案原文
 * @param {string[]} optionColumns 选项列组（位置升序）
 * @param {Record<string, string>} row 该行网格
 * @returns {{letters: string[], options: {letter: string, column: string, text: string}[],
 *            outOfRange: string[], normalizedFrom: string|null, invalid: boolean}}
 */
function mapLettersToOptions(answerText, optionColumns, row) {
  const raw = String(answerText === undefined || answerText === null ? '' : answerText);
  const normalized = normalizeChoiceAnswer(raw);
  const columns = Array.isArray(optionColumns) ? optionColumns : [];
  if (!normalized.ok) {
    return { letters: [], options: [], outOfRange: [], normalizedFrom: null, invalid: true };
  }
  const grid = row || {};
  const options = [];
  const outOfRange = [];
  for (const letter of normalized.letters) {
    const position = letter.charCodeAt(0) - 65; // A → 0
    if (position >= columns.length) {
      outOfRange.push(letter); // 序数超出列组长度：不映射、不猜测
      continue;
    }
    const column = columns[position];
    const text = grid[column] === undefined ? '' : grid[column];
    options.push({ letter, column, text });
  }
  return {
    letters: normalized.letters,
    options,
    outOfRange,
    normalizedFrom: normalized.changed && raw !== '' ? raw : null,
    invalid: false,
  };
}

/**
 * 填空多空切分（§2.7.3 第 4 条）。
 * @param {string} answerText
 * @param {number} blankCount
 * @returns {{segments: string[]|null, separator: string|null, reason: string|null}}
 */
function splitBlankAnswer(answerText, blankCount) {
  const text = String(answerText === undefined || answerText === null ? '' : answerText);
  const blanks = Number(blankCount) || 0;
  if (blanks < 2) return { segments: null, separator: null, reason: '空位数少于 2，不切分' };
  // 候选分隔符优先级：；→ 、→ ，/,（`/` 与空格永不作为分隔符）
  const groups = [['；'], ['、'], ['，', ',']];
  for (const group of groups) {
    const separator = group.find((sep) => text.indexOf(sep) >= 0);
    if (!separator) continue;
    const segments = text.split(separator).map((s) => s.trim()).filter((s) => s !== '');
    if (segments.length === blanks) return { segments, separator, reason: null };
    return { segments: null, separator, reason: `答案按「${separator}」分出 ${segments.length} 段，与空位数 ${blanks} 不符` };
  }
  return { segments: null, separator: null, reason: `答案中找不到分隔符（空位数 ${blanks}）` };
}

/**
 * 判断类答案归一化（§2.7.2 白名单，写死三形态）。
 * @param {string} answerText
 * @returns {'correct'|'wrong'|null} null = 白名单外，须标注，不猜
 */
function normalizeJudge(answerText) {
  const text = String(answerText === undefined || answerText === null ? '' : answerText).replace(/\s+/g, '');
  if (text === '√') return 'correct';
  if (text === '×') return 'wrong';
  if (text === 'X') return 'wrong'; // 可安全归一化（§1.6.4 第 3 条）
  return null;
}

/** 题干空位数（复用文本层写死的空位括号正则）。 */
function blankCountOf(stem) {
  return countBlanks(stem);
}

__MODULES__["src/core/answers.mjs"] = { normalizeChoiceAnswer: normalizeChoiceAnswer, mapLettersToOptions: mapLettersToOptions, splitBlankAnswer: splitBlankAnswer, normalizeJudge: normalizeJudge, blankCountOf: blankCountOf };
}());
(function () {
// 识别层 —— 表头行 / 列角色 / 「不确定」判定（设计档 §2.4.1 / §2.6）。
// 判据全部可解释、可写用例；判据不唯一时置 uncertain，交确认页，绝不静默选一个。

const { countBlanks } = __MODULES__["src/core/text.mjs"];

const TYPE_WORDS = [
  { type: 'single', label: '单选题', pattern: /单选/ },
  { type: 'multi', label: '多选题', pattern: /多选/ },
  { type: 'judge', label: '判断题', pattern: /判断/ },
  { type: 'fill', label: '填空题', pattern: /填空/ },
  { type: 'essay', label: '简答题', pattern: /简答/ },
];

const TYPE_LABELS = { single: '单选题', multi: '多选题', judge: '判断题', fill: '填空题', essay: '简答题' };

// 表头关键词集合（写死，§2.6.2）
const HEADER_WORDS = [
  '题型', '题目', '题干', '题号', '选项', '选项A', '选项B', '选项C', '选项D',
  '选项E', '选项F', '选项G', '选项H', '标准答案', '答案', '解析', '备注',
];

// 确认页角色下拉框取值（§2.6.5）
const ROLE_OPTIONS = [
  { value: 'type', label: '题型' },
  { value: 'stem', label: '题干' },
  { value: 'option', label: '选项' },
  { value: 'answer', label: '答案' },
  { value: 'ignore', label: '忽略' },
];

const OPTION_LETTERS = 'ABCDEFGH';

/** 题型词表匹配（允许前后缀，如「单选」「单选题」）。 */
function matchTypeWord(value) {
  const text = String(value === undefined || value === null ? '' : value).trim();
  if (!text) return null;
  for (const word of TYPE_WORDS) if (word.pattern.test(text)) return word.type;
  return null;
}

/**
 * 逐表识别（§2.6）。
 * @param {{name: string, index?: number, rows: Map<number, Record<string,string>>}} sheet
 * @returns {object} SheetDetection
 */
function detectSheet(sheet) {
  const rows = toRowMap(sheet);
  const rowNums = Array.from(rows.keys()).sort((a, b) => a - b);
  const reasons = [];
  let uncertain = false;

  // —— §2.6.2 表头行
  const headerInfo = detectHeaderRow(rows, rowNums);
  if (headerInfo.candidates.length > 1) {
    uncertain = true;
    reasons.push({
      role: 'headerRow',
      candidates: headerInfo.candidates.map(String),
      why: `前 3 行中有 ${headerInfo.candidates.length} 行都像表头行（第 ${headerInfo.candidates.join('、')} 行）`,
    });
  }
  const headerRow = headerInfo.headerRow;
  const dataRowNums = headerRow === null ? rowNums : rowNums.filter((r) => r > headerRow);
  if (dataRowNums.length === 0) {
    uncertain = true;
    reasons.push({ role: 'headerRow', candidates: [], why: '表头行之下没有任何数据行' });
  }

  const profile = buildProfile(rows, dataRowNums);
  const columns = Array.from(profile.values());

  // —— §2.6.3 题型列
  let typeColumn = null;
  const typeCandidates = columns
    .filter((p) => p.ratio >= 0.8 && p.distinct.size <= 8 && p.avgLen <= 4
      && p.values.length > 0 && typeWordRatio(p.values) >= 0.8)
    .sort((a, b) => b.ratio - a.ratio || colIndex(a.col) - colIndex(b.col));
  if (typeCandidates.length === 1) {
    typeColumn = typeCandidates[0].col;
  } else if (typeCandidates.length > 1) {
    if (Math.abs(typeCandidates[0].ratio - typeCandidates[1].ratio) < 1e-9) {
      uncertain = true;
      reasons.push({
        role: 'type',
        candidates: typeCandidates.map((p) => p.col),
        why: `${typeCandidates.length} 列同时像题型列（非空率相同，均为 ${typeCandidates[0].ratio.toFixed(2)}）`,
      });
    } else {
      typeColumn = typeCandidates[0].col;
    }
  }

  const typeValues = typeColumn ? profile.get(typeColumn).values : [];
  const typeHintCounts = new Map();
  for (const value of typeValues) {
    const type = matchTypeWord(value);
    if (type) typeHintCounts.set(type, (typeHintCounts.get(type) || 0) + 1);
  }
  const rowType = typeHintCounts.size
    ? Array.from(typeHintCounts.entries()).sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))[0][0]
    : null;
  const nameType = matchTypeWord(sheet.name);
  if (rowType && nameType && rowType !== nameType) {
    reasons.push({
      role: 'type',
      candidates: typeColumn ? [typeColumn] : [],
      why: `工作表名「${sheet.name}」提示${TYPE_LABELS[nameType]}，但题型列取值为${TYPE_LABELS[rowType]}——以行取值为准`,
    });
  }
  const sheetType = rowType || nameType;
  if (!sheetType) {
    uncertain = true;
    reasons.push({
      role: 'type',
      candidates: typeColumn ? [typeColumn] : [],
      why: '没有题型列，工作表名也无法判定题型，请手工指定题型列',
    });
  }
  if (!typeColumn && nameType) {
    reasons.push({ role: 'type', candidates: [], why: `无题型列，按工作表名「${sheet.name}」判定为${TYPE_LABELS[nameType]}` });
  }

  // —— §2.6.3 答案列（位置最靠右优先，写死）
  let answerColumn = null;
  const answerCandidates = columns
    .filter((p) => p.ratio >= 0.8 && p.values.length > 0 && compatRatio(p.values, sheetType) >= 0.8)
    .sort((a, b) => colIndex(b.col) - colIndex(a.col));
  if (answerCandidates.length === 0) {
    uncertain = true;
    reasons.push({ role: 'answer', candidates: [], why: '没有取值与题型相容的答案列候选' });
  } else {
    answerColumn = answerCandidates[0].col;
    if (answerCandidates.length > 1) {
      reasons.push({
        role: 'answer',
        candidates: answerCandidates.map((p) => p.col),
        why: `${answerCandidates.length} 列与题型相容，按写死规则取位置最靠右的 ${answerColumn} 列`,
      });
    }
  }

  // —— §2.6.3 题干列
  // 角色互斥：答案列已是答案角色，不再作题干候选（§2.6.5 校验规则②「题型 / 题干 / 答案 各至多 1 列」；
  // 若不排除，答案文本较长的表（如简答题）会被误判成题干列）。
  let stemColumn = null;
  const stemCandidates = columns
    .filter((p) => p.ratio >= 0.9 && p.avgLen >= 8 && !p.numericOnly && p.col !== typeColumn && p.col !== answerColumn)
    .sort((a, b) => b.avgLen - a.avgLen || colIndex(a.col) - colIndex(b.col));
  if (stemCandidates.length === 0) {
    uncertain = true;
    reasons.push({ role: 'stem', candidates: [], why: '找不到题干列候选（非空率 ≥ 0.9 且平均长度 ≥ 8 的列）' });
  } else {
    stemColumn = stemCandidates[0].col;
    if (stemCandidates.length > 1 && stemCandidates[1].avgLen >= stemCandidates[0].avgLen * 0.9) {
      uncertain = true;
      reasons.push({
        role: 'stem',
        candidates: stemCandidates.map((p) => p.col),
        why: `题干列有 ${stemCandidates.length} 个候选：${stemCandidates.slice(0, 3).map((p) => p.col).join(' 与 ')}，平均长度比 ${(stemCandidates[1].avgLen / stemCandidates[0].avgLen).toFixed(2)}`,
      });
    }
  }

  // —— §2.6.3 选项列组（题干列与答案列之间的连续区间，长度 2–8）
  let optionColumns = [];
  if (stemColumn && answerColumn) {
    const left = colIndex(stemColumn);
    const right = colIndex(answerColumn);
    if (right > left) {
      const between = columns
        .filter((p) => colIndex(p.col) > left && colIndex(p.col) < right && p.col !== typeColumn)
        .map((p) => p.col)
        .sort((a, b) => colIndex(a) - colIndex(b));
      const contiguous = between.every((col, i) => i === 0 || colIndex(col) - colIndex(between[i - 1]) === 1);
      if (!contiguous) {
        uncertain = true;
        reasons.push({ role: 'option', candidates: between, why: '题干列与答案列之间的选项列不连续' });
      }
      optionColumns = between;
    }
  }
  const isChoice = sheetType === 'single' || sheetType === 'multi';
  if (isChoice && (optionColumns.length < 2 || optionColumns.length > 8)) {
    uncertain = true;
    reasons.push({
      role: 'option',
      candidates: optionColumns,
      why: `题型为${TYPE_LABELS[sheetType]}，但选项列组长度为 ${optionColumns.length}（需 2–8 列）`,
    });
  }

  return {
    name: sheet.name,
    index: sheet.index === undefined ? 0 : sheet.index,
    headerRow,
    typeColumn,
    stemColumn,
    optionColumns,
    answerColumn,
    typeHints: Array.from(typeHintCounts.keys()),
    sheetType,
    uncertain,
    reasons,
    dataRowCount: dataRowNums.length,
    allColumns: columns.map((p) => p.col).sort((a, b) => colIndex(a) - colIndex(b)),
  };
}

/**
 * 用户改映射后的校验（§2.6.5 第 4 条）。
 * 接受两种形态：① 四个角色字段（typeColumn / stemColumn / optionColumns / answerColumn）；
 * ② 逐列角色表 roles = { 列字母: 'type'|'stem'|'option'|'answer'|'ignore' }（确认页的形态，能表达“角色被 2 列占用”）。
 * @returns {{ok: boolean, errors: string[]}}
 */
function validateMapping(mapping) {
  const input = mapping || {};
  const errors = [];
  let typeColumns = [];
  let stemColumns = [];
  let answerColumns = [];
  let optionColumns = [];

  if (input.roles && typeof input.roles === 'object') {
    const entries = Object.entries(input.roles).sort((a, b) => colIndex(a[0]) - colIndex(b[0]));
    for (const [col, role] of entries) {
      if (role === 'type') typeColumns.push(col);
      else if (role === 'stem') stemColumns.push(col);
      else if (role === 'answer') answerColumns.push(col);
      else if (role === 'option') optionColumns.push(col);
    }
  } else {
    if (input.typeColumn) typeColumns = [input.typeColumn];
    if (input.stemColumn) stemColumns = [input.stemColumn];
    if (input.answerColumn) answerColumns = [input.answerColumn];
    optionColumns = Array.isArray(input.optionColumns) ? input.optionColumns.slice() : [];
  }

  if (typeColumns.length > 1) errors.push(`题型角色被 ${typeColumns.length} 列占用`);
  if (stemColumns.length > 1) errors.push(`题干角色被 ${stemColumns.length} 列占用`);
  if (answerColumns.length > 1) errors.push(`答案角色被 ${answerColumns.length} 列占用`);
  if (stemColumns.length === 0) errors.push('缺少题干列');
  if (answerColumns.length === 0) errors.push('缺少答案列');

  const type = input.sheetType || input.type || null;
  if (typeColumns.length === 0 && !type) errors.push('缺少题型列（工作表名也无法判定题型）');
  if (typeColumns.length === 1 && stemColumns[0] === typeColumns[0]) errors.push('题型角色与题干角色被同一列占用');
  if (typeColumns.length === 1 && answerColumns[0] === typeColumns[0]) errors.push('题型角色与答案角色被同一列占用');
  if (stemColumns.length === 1 && answerColumns[0] === stemColumns[0]) errors.push('题干角色与答案角色被同一列占用');

  const overlaps = optionColumns.filter((col) => col === stemColumns[0] || col === answerColumns[0] || col === typeColumns[0]);
  if (overlaps.length > 0) errors.push(`选项列与其他角色重叠：${overlaps.join('、')}`);
  if (optionColumns.length > 8) errors.push(`选项列超过 8 列（当前 ${optionColumns.length} 列）`);

  const isChoice = type === 'single' || type === 'multi';
  if (isChoice && optionColumns.length < 2) errors.push(`选项列不足 2 列（当前 ${optionColumns.length} 列）`);

  return { ok: errors.length === 0, errors };
}

/** 从识别结果导出逐列角色表（确认页初始值）。 */
function mappingFromDetection(detection) {
  const roles = {};
  const seen = new Set();
  const mark = (col, role) => {
    if (!col) return;
    roles[col] = role;
    seen.add(col);
  };
  mark(detection.typeColumn, 'type');
  mark(detection.stemColumn, 'stem');
  for (const col of detection.optionColumns || []) mark(col, 'option');
  mark(detection.answerColumn, 'answer');
  for (const col of detection.allColumns || []) if (!seen.has(col)) roles[col] = 'ignore';
  return roles;
}

function detectHeaderRow(rows, rowNums) {
  const candidates = [];
  for (const rowNum of rowNums.slice(0, 3)) {
    const row = rows.get(rowNum) || {};
    let hits = 0;
    for (const value of Object.values(row)) {
      if (HEADER_WORDS.includes(String(value).trim())) hits += 1;
    }
    if (hits >= 2) candidates.push(rowNum);
  }
  if (candidates.length === 1) return { headerRow: candidates[0], candidates };
  if (candidates.length === 0) return { headerRow: null, candidates };
  return { headerRow: null, candidates };
}

function buildProfile(rows, dataRowNums) {
  const allColumns = new Set();
  for (const rowNum of dataRowNums) {
    const row = rows.get(rowNum);
    if (!row) continue;
    for (const col of Object.keys(row)) allColumns.add(col);
  }
  const profile = new Map();
  for (const col of allColumns) {
    const values = [];
    let blanks = 0;
    for (const rowNum of dataRowNums) {
      const row = rows.get(rowNum);
      const value = row ? row[col] : undefined;
      if (value === undefined || value === '') continue;
      values.push(value);
      blanks += countBlanks(value);
    }
    const total = dataRowNums.length;
    profile.set(col, {
      col,
      total,
      count: values.length,
      ratio: total ? values.length / total : 0,
      values,
      distinct: new Set(values),
      avgLen: values.length ? values.reduce((sum, v) => sum + v.length, 0) / values.length : 0,
      maxLen: values.reduce((max, v) => Math.max(max, v.length), 0),
      blankRatio: values.length ? blanks / values.length : 0,
      numericOnly: values.length > 0 && values.every((v) => /^-?\d+(\.\d+)?$/.test(v)),
    });
  }
  return profile;
}

function typeWordRatio(values) {
  const matched = values.filter((v) => matchTypeWord(v)).length;
  return values.length ? matched / values.length : 0;
}

function compatRatio(values, sheetType) {
  if (!sheetType) return 0;
  const matched = values.filter((v) => isCompatible(v, sheetType)).length;
  return values.length ? matched / values.length : 0;
}

function isCompatible(value, sheetType) {
  const text = String(value).trim();
  if (!text) return false;
  if (sheetType === 'single' || sheetType === 'multi') {
    const compact = text.replace(/[\s、，,；;|]/g, '').replace(/[Ａ-Ｚ]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xfee0));
    return /^[A-H]+$/.test(compact);
  }
  if (sheetType === 'judge') return text === '√' || text === '×' || text === 'X';
  return true; // 填空 / 简答：任意非空
}

function colIndex(letter) {
  let index = 0;
  const text = String(letter || '').toUpperCase();
  for (let i = 0; i < text.length; i++) index = index * 26 + (text.charCodeAt(i) - 64);
  return index;
}

function toRowMap(sheet) {
  if (sheet && sheet.rows instanceof Map) return sheet.rows;
  const rows = new Map();
  const source = (sheet && sheet.rows) || {};
  for (const key of Object.keys(source)) rows.set(Number(key), source[key]);
  return rows;
}

__MODULES__["src/core/detect.mjs"] = { matchTypeWord: matchTypeWord, detectSheet: detectSheet, validateMapping: validateMapping, mappingFromDetection: mappingFromDetection, colIndex: colIndex, TYPE_WORDS: TYPE_WORDS, TYPE_LABELS: TYPE_LABELS, ROLE_OPTIONS: ROLE_OPTIONS, OPTION_LETTERS: OPTION_LETTERS };
}());
(function () {
// ZIP 层 —— 中央目录解析与条目字节取回（设计档 §2.4.1 / §2.5.1）。
// 零依赖：手写 ZIP 结构读取。规则全部写死，结构异常一律抛错，不猜偏移、不静默降级。

const EOCD_SIG = 0x06054b50; // 中央目录结尾记录
const CD_SIG = 0x02014b50; // 中央目录条目
const LFH_SIG = 0x04034b50; // 本地文件头
const EOCD_MIN = 22;
const MAX_COMMENT = 65535;

/**
 * 读取 ZIP 条目表（目录条目已剔除）。
 * @param {Uint8Array} bytes
 * @returns {{name: string, method: number, compSize: number, uncompSize: number, localOffset: number}[]}
 */
function readZipEntries(bytes) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const eocd = findEocd(view, bytes.length);
  if (eocd < 0) throw new Error('ZIP 结构异常：找不到中央目录结尾记录（EOCD），文件可能不是有效的 xlsx');

  const total = view.getUint16(eocd + 10, true);
  const cdSize = view.getUint32(eocd + 12, true);
  const cdOffset = view.getUint32(eocd + 16, true);
  if (total === 0xffff || cdSize === 0xffffffff || cdOffset === 0xffffffff) {
    throw new Error('不支持 ZIP64 格式的 xlsx 文件');
  }

  const decoder = new TextDecoder('utf-8');
  const entries = [];
  let p = cdOffset;
  for (let i = 0; i < total; i++) {
    if (p + 46 > bytes.length || view.getUint32(p, true) !== CD_SIG) {
      throw new Error(`ZIP 结构异常：第 ${i + 1} 个中央目录条目签名不符`);
    }
    const flags = view.getUint16(p + 8, true);
    const method = view.getUint16(p + 10, true);
    const compSize = view.getUint32(p + 20, true);
    const uncompSize = view.getUint32(p + 24, true);
    const nameLen = view.getUint16(p + 28, true);
    const extraLen = view.getUint16(p + 30, true);
    const commentLen = view.getUint16(p + 32, true);
    const localOffset = view.getUint32(p + 42, true);
    const name = decoder.decode(bytes.subarray(p + 46, p + 46 + nameLen));
    if (flags & 0x1) throw new Error(`ZIP 条目已加密，无法读取：${name}`);
    p += 46 + nameLen + extraLen + commentLen;
    // 名列以 / 结尾 = 目录条目，跳过（不计入条目表）
    if (name.endsWith('/')) continue;
    entries.push({ name, method, compSize, uncompSize, localOffset });
  }
  return entries;
}

/**
 * 取回单个条目的解压后字节。
 * @param {Uint8Array} bytes 整个 ZIP 文件
 * @param {{name: string, method: number, compSize: number, uncompSize: number, localOffset: number}} entry
 * @param {(bytes: Uint8Array) => Promise<Uint8Array>} inflateRaw 注入的解压实现
 * @returns {Promise<Uint8Array>}
 */
async function entryBytes(bytes, entry, inflateRaw) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const off = entry.localOffset;
  if (off + 30 > bytes.length || view.getUint32(off, true) !== LFH_SIG) {
    throw new Error(`ZIP 结构异常：条目 ${entry.name} 的本地文件头签名不符`);
  }
  const nameLen = view.getUint16(off + 26, true);
  const extraLen = view.getUint16(off + 28, true);
  const start = off + 30 + nameLen + extraLen;
  const raw = bytes.subarray(start, start + entry.compSize);

  let out;
  if (entry.method === 0) out = raw.slice();
  else if (entry.method === 8) out = await inflateRaw(raw);
  else throw new Error(`不支持的压缩方法 ${entry.method}（条目 ${entry.name}）：只支持 stored(0) 与 deflate(8)`);

  if (entry.uncompSize && out.length !== entry.uncompSize) {
    throw new Error(`条目 ${entry.name} 解压长度不符（实得 ${out.length}，声明 ${entry.uncompSize}）`);
  }
  return out;
}

function findEocd(view, length) {
  const lowest = Math.max(0, length - EOCD_MIN - MAX_COMMENT);
  for (let i = length - EOCD_MIN; i >= lowest; i--) {
    if (view.getUint32(i, true) === EOCD_SIG) return i;
  }
  return -1;
}

__MODULES__["src/core/zip.mjs"] = { readZipEntries: readZipEntries, entryBytes: entryBytes };
}());
(function () {
// 表读取层 —— workbook / rels / sharedStrings / styles / sheet XML → 行网格（设计档 §2.4.1 / §2.5）。
// 输出：每表一行网格 Map<Excel 行号, {列字母 → 显示文本}>；空单元格不入网格。

const { entryBytes, readZipEntries } = __MODULES__["src/core/zip.mjs"];
const { extractInlineString, parseCellText } = __MODULES__["src/core/text.mjs"];

const CHUNK_CELLS = 20000; // 分片单位 = 非空单元格数（§2.5.5）
const MAX_CELLS = 500000; // 超限直接拒绝（§2.2.6）

/**
 * 读取整个 xlsx。
 * @param {Uint8Array} bytes
 * @param {{inflateRaw: (b: Uint8Array) => Promise<Uint8Array>, onProgress?: (p: {processed: number, total: number}) => void}} options
 * @returns {Promise<{sheets: Array<{name: string, index: number, rows: Map<number, Record<string,string>>, cellCount: number, notes: string[], unresolvedCells: Set<string>}>}>}
 */
async function readXlsx(bytes, options) {
  const opts = options || {};
  const inflateRaw = opts.inflateRaw;
  if (typeof inflateRaw !== 'function') throw new Error('readXlsx 需要注入 inflateRaw 解压实现（见 inflate.mjs）');
  const onProgress = typeof opts.onProgress === 'function' ? opts.onProgress : null;
  const decoder = new TextDecoder('utf-8');

  const entries = readZipEntries(bytes);
  const byName = new Map();
  for (const entry of entries) byName.set(entry.name, entry);

  const readEntry = async (name) => {
    const entry = byName.get(name);
    if (!entry) return null;
    return decoder.decode(await entryBytes(bytes, entry, inflateRaw));
  };

  const workbookXml = await readEntry('xl/workbook.xml');
  if (workbookXml === null) throw new Error('xlsx 结构异常：缺少 xl/workbook.xml');
  const relsXml = await readEntry('xl/_rels/workbook.xml.rels');
  if (relsXml === null) throw new Error('xlsx 结构异常：缺少 xl/_rels/workbook.xml.rels（无法定位工作表，不猜 sheetN.xml）');

  const sharedStrings = parseSharedStrings((await readEntry('xl/sharedStrings.xml')) || '');
  const styles = parseStyles((await readEntry('xl/styles.xml')) || '');
  const rels = parseRels(relsXml);
  const declared = parseWorkbookSheets(workbookXml);

  const usedPaths = new Set();
  const sheets = [];
  let globalDone = 0;
  for (const item of declared) {
    if (item.state && item.state !== 'visible') continue; // 隐藏表跳过
    const target = rels.get(item.rid);
    if (!target) throw new Error(`xlsx 结构异常：工作表「${item.name}」在 rels 中没有映射（不猜 sheetN.xml）`);
    const path = resolveTarget(target);
    if (!byName.has(path)) throw new Error(`xlsx 结构异常：工作表文件不存在：${path}`);
    usedPaths.add(path);
    const xml = decoder.decode(await entryBytes(bytes, byName.get(path), inflateRaw));
    const parsed = await parseSheetXml(xml, {
      name: item.name,
      sharedStrings,
      styles,
      onProgress,
      baseDone: globalDone,
    });
    globalDone += parsed.cellCount;
    sheets.push({
      name: item.name,
      index: sheets.length,
      rows: parsed.rows,
      cellCount: parsed.cellCount,
      notes: parsed.notes,
      unresolvedCells: parsed.unresolvedCells,
      path,
    });
  }
  return { sheets, sharedStrings, usedPaths };
}

async function parseSheetXml(xml, ctx) {
  const approxCells = (xml.match(/<c[\s/>]/g) || []).length;
  if (approxCells > MAX_CELLS) {
    throw new Error(`工作表「${ctx.name}」规模超出本程序处理能力（约 ${approxCells} 个单元格 > ${MAX_CELLS}）`);
  }

  const rows = new Map();
  const notes = [];
  const unresolvedCells = new Set();
  let cellCount = 0;
  let chunkCells = 0;
  let pos = 0;

  while (pos < xml.length) {
    const start = findTag(xml, pos, 'row');
    if (start < 0) break;
    const tagEnd = xml.indexOf('>', start);
    if (tagEnd < 0) break;
    const selfClosing = xml.charAt(tagEnd - 1) === '/';
    const attrs = xml.slice(start + 4, selfClosing ? tagEnd - 1 : tagEnd);
    let contentEnd = tagEnd;
    if (!selfClosing) {
      const end = xml.indexOf('</row>', tagEnd);
      contentEnd = end < 0 ? xml.length : end;
    }
    const rowNum = Number(attrValue(attrs, 'r'));
    const content = selfClosing ? '' : xml.slice(tagEnd + 1, contentEnd);

    if (Number.isInteger(rowNum) && content) {
      const row = {};
      const cellRe = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
      let m;
      while ((m = cellRe.exec(content)) !== null) {
        const cellAttrs = m[1] || '';
        const inner = m[2] || '';
        const ref = attrValue(cellAttrs, 'r');
        const col = ref ? ref.replace(/[^A-Z]/g, '') : '';
        if (!col) continue;
        const vMatch = /<v\b[^>]*>([\s\S]*?)<\/v>/.exec(inner);
        const isMatch = /<is\b[^>]*>([\s\S]*?)<\/is>/.exec(inner);
        if (!vMatch && !isMatch) continue;
        const styleIndex = Number(attrValue(cellAttrs, 's'));
        const numFmtId = Number.isInteger(styleIndex) && ctx.styles.cellXfs[styleIndex] !== undefined
          ? ctx.styles.cellXfs[styleIndex]
          : 0;
        const parsed = parseCellText({
          type: attrValue(cellAttrs, 't') || '',
          value: vMatch ? vMatch[1] : '',
          inlineXml: isMatch ? isMatch[1] : '',
          sharedStrings: ctx.sharedStrings,
          numFmtId,
          formatDefinedInFile: ctx.styles.definedIds.has(numFmtId),
        });
        if (parsed.note) notes.push(`R${rowNum}${col}：${parsed.note}`);
        if (parsed.text === '') continue; // 空单元格不入网格
        row[col] = parsed.text;
        cellCount += 1;
        chunkCells += 1;
        if (parsed.formatUnresolved) unresolvedCells.add(`${rowNum}:${col}`);
      }
      if (Object.keys(row).length > 0) rows.set(rowNum, row);
    }

    pos = selfClosing ? tagEnd + 1 : Math.max(contentEnd + '</row>'.length, tagEnd + 1);
    if (chunkCells >= CHUNK_CELLS) {
      chunkCells = 0;
      if (ctx.onProgress) ctx.onProgress({ processed: ctx.baseDone + cellCount, total: ctx.baseDone + approxCells });
      await yieldToEventLoop();
    }
  }
  return { rows, cellCount, notes, unresolvedCells };
}

function yieldToEventLoop() {
  return new Promise((resolve) => setTimeout(resolve, 0));
}

function findTag(xml, from, tag) {
  const needle = `<${tag}`;
  let i = xml.indexOf(needle, from);
  while (i >= 0) {
    const next = xml.charAt(i + needle.length);
    if (next === ' ' || next === '>' || next === '/') return i;
    i = xml.indexOf(needle, i + 1);
  }
  return -1;
}

function attrValue(attrs, name) {
  const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const m = new RegExp(`(?:^|\\s)${escaped}="([^"]*)"`).exec(attrs);
  return m ? m[1] : null;
}

function parseWorkbookSheets(xml) {
  const out = [];
  const re = /<sheet\b[^>]*\/?>/g;
  let m;
  while ((m = re.exec(xml)) !== null) {
    out.push({
      name: attrValue(m[0], 'name'),
      state: attrValue(m[0], 'state'),
      rid: attrValue(m[0], 'r:id'),
    });
  }
  return out;
}

function parseRels(xml) {
  const map = new Map();
  const re = /<Relationship\b[^>]*\/?>/g;
  let m;
  while ((m = re.exec(xml)) !== null) {
    const id = attrValue(m[0], 'Id');
    const target = attrValue(m[0], 'Target');
    if (id && target) map.set(id, target);
  }
  return map;
}

function resolveTarget(target) {
  let t = String(target).replace(/\\/g, '/');
  if (t.startsWith('/')) return t.slice(1);
  if (t.startsWith('xl/')) return t;
  const parts = `xl/${t}`.split('/');
  const out = [];
  for (const part of parts) {
    if (part === '.' || part === '') continue;
    if (part === '..') out.pop();
    else out.push(part);
  }
  return out.join('/');
}

function parseSharedStrings(xml) {
  const out = [];
  const re = /<si\b[^>]*>([\s\S]*?)<\/si>|<si\b[^>]*\/>/g;
  let m;
  while ((m = re.exec(xml)) !== null) out.push(extractInlineString(m[1] || ''));
  return out;
}

function parseStyles(xml) {
  const definedIds = new Set();
  const numFmtRe = /<numFmt\b[^>]*\/?>/g;
  let m;
  while ((m = numFmtRe.exec(xml)) !== null) {
    const id = Number(attrValue(m[0], 'numFmtId'));
    if (Number.isInteger(id)) definedIds.add(id);
  }

  const cellXfs = [];
  const block = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(xml);
  if (block) {
    const xfRe = /<xf\b[^>]*>/g;
    let xf;
    while ((xf = xfRe.exec(block[1])) !== null) {
      const id = Number(attrValue(xf[0], 'numFmtId'));
      cellXfs.push(Number.isInteger(id) ? id : 0);
    }
  }
  return { cellXfs, definedIds };
}

__MODULES__["src/core/xlsx.mjs"] = { readXlsx: readXlsx };
}());
(function () {
// 建库层 —— 答案解读 + 五类数据标记 + 题目对象装配 + 题库指纹（设计档 §2.4.1 / §2.7 / §2.9 / §2.10.3）。

const { mapLettersToOptions, normalizeJudge, splitBlankAnswer } = __MODULES__["src/core/answers.mjs"];
const { colIndex, detectSheet, matchTypeWord, OPTION_LETTERS } = __MODULES__["src/core/detect.mjs"];
const { countBlanks } = __MODULES__["src/core/text.mjs"];
const { readXlsx } = __MODULES__["src/core/xlsx.mjs"];

const SCHEMA_VERSION = 1;

// 标记展示顺序（§2.9 规则 1）
const MARK_ORDER = ['partial', 'outOfRange', 'suspect', 'noAnswer', 'formatUnresolved'];

/**
 * 装配题库（§2.4.1）。
 * @param {Array} sheets readXlsx 的输出
 * @param {Array} detections 与 sheets 等长的 SheetDetection（可带 roles 覆盖 / skip）
 * @param {{sourceName?: string}} [meta]
 * @returns {{questions: Array, fingerprints: object, stats: object}}
 */
function buildBank(sheets, detections, meta) {
  const info = meta || {};
  const questions = [];
  const sheetStats = [];
  const stats = { total: 0, byType: {}, byMark: {}, typeFallbacks: 0, skippedSheets: [] };

  for (let i = 0; i < sheets.length; i++) {
    const sheet = sheets[i];
    const detection = (detections && detections[i]) || null;
    if (!detection || detection.skip) {
      stats.skippedSheets.push(sheet.name);
      sheetStats.push({ name: sheet.name, index: i, count: 0, skipped: true });
      continue;
    }
    const mapping = effectiveMapping(detection);
    const rowNums = Array.from(sheet.rows.keys())
      .sort((a, b) => a - b)
      .filter((r) => mapping.headerRow === null || r > mapping.headerRow);
    let count = 0;
    for (const rowNum of rowNums) {
      const question = buildQuestion(sheet, sheet.rows.get(rowNum), rowNum, mapping, stats);
      if (!question) continue;
      questions.push(question);
      count += 1;
      stats.total += 1;
      stats.byType[question.type] = (stats.byType[question.type] || 0) + 1;
      for (const mark of question.marks) stats.byMark[mark] = (stats.byMark[mark] || 0) + 1;
    }
    sheetStats.push({ name: sheet.name, index: i, count });
  }

  const sourceName = info.sourceName || '';
  const fingerprints = {
    schemaVersion: SCHEMA_VERSION,
    bankFingerprint: fingerprintOf(sourceName, questions),
    sourceName,
    sheets: sheetStats,
    questionCount: questions.length,
  };
  return { questions, fingerprints, stats };
}

/** 组装可持久化的题库记录（store 的 `bank` 键 / 单文件内嵌题库同形）。 */
function makeBankRecord(options) {
  const opts = options || {};
  const fingerprints = opts.fingerprints || {};
  return {
    schemaVersion: SCHEMA_VERSION,
    bankFingerprint: fingerprints.bankFingerprint || fingerprintOf(opts.sourceName || '', opts.questions || []),
    sourceName: opts.sourceName || '',
    builtAt: opts.builtAt || null,
    sheets: fingerprints.sheets || [],
    questions: opts.questions || [],
  };
}

/** 题库内容指纹：表名 + 逐题 id 的散列（§2.10.3）。 */
function fingerprintOf(sourceName, questions) {
  let h = 2166136261 >>> 0;
  const feed = (text) => {
    for (let i = 0; i < text.length; i++) {
      h ^= text.charCodeAt(i);
      h = Math.imul(h, 16777619) >>> 0;
    }
    h ^= 0x1f;
    h = Math.imul(h, 16777619) >>> 0;
  };
  feed(String(sourceName || ''));
  for (const q of questions || []) feed(q.id);
  return h.toString(16).padStart(8, '0');
}

/**
 * 完整导入流水线：字节 → 读表 → 逐表识别 → 建库（PWA 与单文件共用同一条路径，R11/AC-25）。
 * @returns {Promise<{sheets: Array, detections: Array, bank: object, stats: object}>}
 */
async function buildBankFromBytes(bytes, options) {
  const opts = options || {};
  const parsed = await readXlsx(bytes, { inflateRaw: opts.inflateRaw, onProgress: opts.onProgress });
  const detections = parsed.sheets.map((sheet) => detectSheet(sheet));
  const built = buildBank(parsed.sheets, detections, { sourceName: opts.sourceName || '' });
  return {
    sheets: parsed.sheets,
    detections,
    bank: makeBankRecord({
      questions: built.questions,
      fingerprints: built.fingerprints,
      sourceName: opts.sourceName || '',
      builtAt: opts.builtAt === undefined ? null : opts.builtAt,
    }),
    stats: built.stats,
  };
}

function effectiveMapping(detection) {
  let typeColumn = detection.typeColumn || null;
  let stemColumn = detection.stemColumn || null;
  let answerColumn = detection.answerColumn || null;
  let optionColumns = Array.isArray(detection.optionColumns) ? detection.optionColumns.slice() : [];
  if (detection.roles && typeof detection.roles === 'object') {
    typeColumn = null;
    stemColumn = null;
    answerColumn = null;
    optionColumns = [];
    for (const [col, role] of Object.entries(detection.roles)) {
      if (role === 'type' && !typeColumn) typeColumn = col;
      else if (role === 'stem' && !stemColumn) stemColumn = col;
      else if (role === 'answer' && !answerColumn) answerColumn = col;
      else if (role === 'option') optionColumns.push(col);
    }
    optionColumns.sort((a, b) => colIndex(a) - colIndex(b));
  }
  return {
    headerRow: detection.headerRow === undefined ? null : detection.headerRow,
    typeColumn,
    stemColumn,
    answerColumn,
    optionColumns,
    sheetType: detection.sheetType || null,
  };
}

function buildQuestion(sheet, row, rowNum, mapping, stats) {
  if (!row || !mapping.stemColumn) return null;
  const stem = row[mapping.stemColumn] === undefined ? '' : row[mapping.stemColumn];
  if (stem === '') return null; // 全空行（如简答表末行空元素）不是题目

  const typeValue = mapping.typeColumn ? (row[mapping.typeColumn] || '') : '';
  let type = matchTypeWord(typeValue);
  if (!type) {
    type = mapping.sheetType;
    if (typeValue) stats.typeFallbacks += 1; // 题型列取值不在词表内 → 退到工作表题型（不猜）
  }
  if (!type) return null;

  const rawAnswer = mapping.answerColumn ? (row[mapping.answerColumn] || '') : '';
  const marks = [];
  const isChoice = type === 'single' || type === 'multi';

  const options = isChoice
    ? mapping.optionColumns.map((col, i) => ({
      letter: OPTION_LETTERS.charAt(i),
      column: col,
      text: row[col] === undefined ? '' : row[col],
    }))
    : [];

  let answer;
  if (isChoice) {
    const mapped = mapLettersToOptions(rawAnswer, mapping.optionColumns, row);
    const mappedTexts = mapped.options.map((o) => o.text);
    answer = {
      raw: rawAnswer,
      display: mapped.invalid ? rawAnswer : mappedTexts.join(' / '),
      letters: mapped.letters,
    };
    // partial：选项列组内任一列为空（含末列）；或答案字母指向的选项单元格为空
    if (mapping.optionColumns.some((col) => row[col] === undefined || row[col] === '')) marks.push('partial');
    if (mapped.options.some((o) => o.text === '')) marks.push('partial');
    if (mapped.outOfRange.length > 0) marks.push('outOfRange');
    if (mapped.invalid || (type === 'multi' && mapped.letters.length === 1)) marks.push('suspect');
  } else if (type === 'judge') {
    const judge = normalizeJudge(rawAnswer);
    answer = {
      raw: rawAnswer,
      display: judge === 'correct' ? '正确' : judge === 'wrong' ? '错误' : rawAnswer,
      judge,
    };
    if (judge === null) marks.push('suspect');
  } else if (type === 'fill') {
    const blanks = countBlanks(stem);
    const split = blanks >= 2 ? splitBlankAnswer(rawAnswer, blanks) : { segments: null, separator: null };
    answer = { raw: rawAnswer, display: rawAnswer, blanks, segments: split.segments, separator: split.separator };
    if (blanks === 0) marks.push('partial');
    else if (blanks >= 2 && split.segments === null) marks.push('partial');
  } else {
    answer = { raw: rawAnswer, display: rawAnswer };
  }

  if (rawAnswer === '') marks.push('noAnswer');
  if (rowHasUnresolvedFormat(sheet, rowNum, mapping)) marks.push('formatUnresolved');

  const ordered = MARK_ORDER.filter((mark) => marks.includes(mark));
  return {
    id: `${sheet.name}#R${rowNum}`,
    sheet: sheet.name,
    row: rowNum,
    type,
    stem,
    options,
    answer,
    marks: ordered,
    sourceRow: `${sheet.name} R${rowNum}`,
  };
}

function rowHasUnresolvedFormat(sheet, rowNum, mapping) {
  if (!sheet.unresolvedCells || sheet.unresolvedCells.size === 0) return false;
  const columns = [mapping.stemColumn, mapping.answerColumn].concat(mapping.optionColumns || []).filter(Boolean);
  return columns.some((col) => sheet.unresolvedCells.has(`${rowNum}:${col}`));
}

__MODULES__["src/core/bank.mjs"] = { buildBank: buildBank, makeBankRecord: makeBankRecord, fingerprintOf: fingerprintOf, buildBankFromBytes: buildBankFromBytes, SCHEMA_VERSION: SCHEMA_VERSION };
}());
(function () {
// 界面基础件 —— 手写 DOM 构造（设计档 §2.11.4）。
// 铁律：文本一律用 textContent 赋值，绝不用 innerHTML（源表含 “” <> 等字符）。

function el(tag, props, children) {
  const node = document.createElement(tag);
  if (props) {
    for (const key of Object.keys(props)) {
      const value = props[key];
      if (value === undefined || value === null || value === false) continue;
      if (key === 'class') node.className = value;
      else if (key === 'text') node.textContent = String(value);
      else if (key === 'on') {
        for (const eventName of Object.keys(value)) node.addEventListener(eventName, value[eventName]);
      } else if (key === 'dataset') {
        for (const dataKey of Object.keys(value)) node.dataset[dataKey] = String(value[dataKey]);
      } else if (key === 'style') {
        node.setAttribute('style', value);
      } else if (key === 'value') {
        node.value = value;
      } else if (key === 'disabled' || key === 'hidden' || key === 'checked') {
        node[key] = true;
      } else {
        node.setAttribute(key, String(value));
      }
    }
  }
  add(node, children);
  return node;
}

function add(node, children) {
  if (children === undefined || children === null || children === false) return node;
  if (Array.isArray(children)) {
    for (const child of children) add(node, child);
    return node;
  }
  if (typeof children === 'string' || typeof children === 'number') {
    node.appendChild(document.createTextNode(String(children)));
    return node;
  }
  node.appendChild(children);
  return node;
}

function clear(node) {
  if (typeof node.replaceChildren === 'function') node.replaceChildren();
  else while (node.firstChild) node.removeChild(node.firstChild);
  return node;
}

function byId(id) {
  return document.getElementById(id);
}

__MODULES__["src/ui/dom.mjs"] = { el: el, add: add, clear: clear, byId: byId };
}());
(function () {
// 会话持久化与诊断字段 —— 由 app.mjs 迁出（设计档 §2.4.4 / §2.10.3 / §2.11.6）。
// 只经 app.store 的存储接口读写；界面状态由 app.mjs 持有。

const { SCHEMA_VERSION } = __MODULES__["src/core/bank.mjs"];
const { ensureProgress, entryMatrix, progressForStore } = __MODULES__["src/core/progress.mjs"];
const { byId } = __MODULES__["src/ui/dom.mjs"];

/** 读回错题本（题目集合 + 错次）；换表时连同进度一起备份，避免旧册子被后续写入覆盖（KD-13）。 */
async function loadWrongbook(app, backup) {
  if (!app.store) return;
  try {
    const wrong = await app.store.loadWrongbook();
    const qids = wrong && Array.isArray(wrong.qids) ? wrong.qids : [];
    const counts = wrong && wrong.counts && typeof wrong.counts === 'object' ? wrong.counts : {};
    if (backup && qids.length > 0) await app.store.saveKey('wrongbook@old', wrong);
    app.wrong = new Set(qids);
    app.wrongCounts = { ...counts };
  } catch (err) {
    app.wrong = new Set();
    app.wrongCounts = {};
  }
}

/** 预读 14 组进度：进入题卡即用，不必等 IO。 */
async function preloadProgress(app) {
  if (!app.store) return;
  for (const entry of entryMatrix()) {
    try {
      const stored = await app.store.loadProgress(entry.groupKey);
      app.progressCache.set(entry.groupKey, ensureProgress(stored, entry.groupKey));
    } catch (err) {
      app.progressCache.set(entry.groupKey, ensureProgress(null, entry.groupKey));
    }
  }
}

/** 读回界面偏好（开关状态 / 关闭过的提示条）。 */
async function loadUiState(app) {
  if (!app.store) return;
  try {
    app.uiState = await app.store.loadUi();
  } catch (err) {
    app.uiState = null;
  }
}

/** 保存当前组进度；集合型入口不落 order / index（§2.10.3）。 */
function persistProgress(app) {
  if (!app.store || !app.group || !app.progress) return;
  // 落盘形状按 §2.10.3：带上该组当前的 order / index（集合型入口由 progressForStore 剥掉）
  const payload = progressForStore(
    Object.assign({}, app.progress, { order: app.order.slice(), index: app.index }),
    app.group.groupKey,
  );
  app.progressCache.set(app.group.groupKey, app.progress);
  app.store.saveProgress(app.group.groupKey, payload).catch((err) => console.error('进度保存失败', err));
}

function persistWrongbook(app) {
  if (!app.store) return Promise.resolve();
  return app.store.saveWrongbook({ qids: Array.from(app.wrong), counts: { ...app.wrongCounts } }).catch((err) => {
    console.error('错题本保存失败', err);
  });
}

/** 保存界面偏好（开关状态 / 关闭过的提示条）；存储不可用时静默跳过。 */
function saveUiState(app) {
  if (!app.store) return;
  app.store.saveUi(app.uiState).catch((err) => console.error('界面偏好保存失败', err));
}

/** 诊断字段（§2.4.4）：页面不显示但 DOM 中稳定存在，无头实测与真机排障都读它。 */
function writeDiag(app) {
  const node = byId('diag');
  if (!node) return;
  const store = app.storeInfo || { backend: 'memory', persistent: false };
  node.textContent = [
    `version=${SCHEMA_VERSION}`,
    `sheets=${app.bank && app.bank.sheets ? app.bank.sheets.length : 0}`,
    `bank=${app.bank && app.bank.questions ? app.bank.questions.length : 0}`,
    `store=${store.backend}`,
    `persist=${store.persistent ? 'yes' : 'no'}`,
    `import=${Math.round(app.importMs)}`,
    `render=${Math.round(app.renderMs)}`,
    `ready=${Math.round(app.readyMs)}`,
    `inflate=${app.inflate.ok ? 'deflate-raw' : 'unsupported'}`,
  ].join(' ');
}

__MODULES__["src/ui/session.mjs"] = { loadWrongbook: loadWrongbook, preloadProgress: preloadProgress, loadUiState: loadUiState, persistProgress: persistProgress, persistWrongbook: persistWrongbook, saveUiState: saveUiState, writeDiag: writeDiag };
}());
(function () {
// 题库装载（N8 拆分自 app.mjs）：读存储 / 内嵌题库、换表备份与重绘（§2.10.3 / §2.4.2）。
// 装载结束时的整体重绘由调用方传入 render，避免本模块反向依赖 app.mjs。
const { entryMatrix } = __MODULES__["src/core/progress.mjs"];
const { backupProgress } = __MODULES__["src/core/store.mjs"];
const { loadUiState, loadWrongbook, preloadProgress } = __MODULES__["src/ui/session.mjs"];

/** 启动装载：已存题库优先；单文件形态的内嵌题库与已存指纹不一致时按换表处理（§2.10.3）。 */
async function loadBank(app, opts, render) {
  let record = null;
  if (app.store) {
    try { record = await app.store.loadBank(); } catch (err) { record = null; }
  }
  const embedded = opts.bank || globalThis.EXAM_MEMO_BANK || null;
  const embed = embedded && Array.isArray(embedded.questions) ? embedded : null;
  // 单文件形态：内嵌题库就是「本次导入」，与已存题库指纹不一致时按换表处理（§2.10.3）
  const stale = !!(record && embed && embed.bankFingerprint && record.bankFingerprint
    && embed.bankFingerprint !== record.bankFingerprint);

  if (stale) {
    app.bank = record; // 让 adoptBank 看到「已存题库」以便比对与备份
    await adoptBank(app, embed, render);
  } else if (record || embed) {
    app.bank = record || embed;
    indexQuestions(app);
    await loadWrongbook(app, false);
    if (app.store && !record) {
      try { await app.store.saveBank(embed); } catch (err) { console.error('内置题库写入存储失败', err); }
    }
    await preloadProgress(app);
  }
  await loadUiState(app);
}

/** 采用一份题库（导入 / 内嵌）：指纹变了就备份旧进度并如实告知（§2.10.3 / KD-13）。 */
async function adoptBank(app, record, render) {
  const previous = app.bank;
  const changed = !!(previous && previous.bankFingerprint && record.bankFingerprint
    && previous.bankFingerprint !== record.bankFingerprint);
  await loadWrongbook(app, changed);
  if (changed && app.store) {
    const groups = entryMatrix().map((entry) => entry.groupKey);
    const backed = await backupProgress(app.store, groups);
    app.notice = `题库已变化，进度将按题号重新对齐（旧的 ${backed} 组进度已备份为 progress:<组>@old）。`;
  }
  app.bank = record;
  indexQuestions(app);
  if (app.store) {
    try { await app.store.saveBank(record); } catch (err) { console.error('题库保存失败', err); }
  }
  await preloadProgress(app);
  app.view = 'matrix';
  render(app);
}

/** 题目 id → 题目 的索引（游标、错题本与背题揭示都按 id 定位）。 */
function indexQuestions(app) {
  app.questionIndex = new Map(((app.bank && app.bank.questions) || []).map((q) => [q.id, q]));
}

__MODULES__["src/ui/bank-load.mjs"] = { loadBank: loadBank, adoptBank: adoptBank, indexQuestions: indexQuestions };
}());
(function () {
// 滑动手势（R26 / §2.11.7④）：纯函数判据与 DOM 适配器分离。
// 不 preventDefault（不吞页面滚动与输入框手势）、不引第三方库（N5 零依赖）、不用 closest()（DOM 桩不支持）。

/** 触发滑动的最小水平位移（像素）。 */
const SWIPE_MIN_PX = 48;
/** 方向占优比：水平位移须大于垂直位移的这么多倍才判为切题。 */
const SWIPE_DOMINANCE = 1.5;
/** 左缘保护带（像素）：起点在此之内不判右滑（避开 iOS 边缘返回手势）。 */
const SWIPE_EDGE_GUARD_PX = 24;

const EDITABLE_TAGS = ['INPUT', 'TEXTAREA', 'SELECT'];

/**
 * 纯判据：左滑 → `'next'`，右滑 → `'prev'`，其余（不足阈值 / 方向不占优）→ `null`。
 * 判定式：`|dx| ≥ SWIPE_MIN_PX && |dx| > SWIPE_DOMINANCE·|dy|`（AC-37）。
 */
function swipeDirection(dx, dy) {
  if (!Number.isFinite(dx) || !Number.isFinite(dy)) return null;
  if (Math.abs(dx) < SWIPE_MIN_PX) return null;
  if (Math.abs(dx) <= SWIPE_DOMINANCE * Math.abs(dy)) return null;
  return dx < 0 ? 'next' : 'prev';
}

/** 起点是否落在输入控件内：沿 `parentNode` 链逐级查 `tagName`（DOM 桩没有 `closest()`）。 */
function inEditable(node) {
  let current = node;
  while (current) {
    if (EDITABLE_TAGS.indexOf(String(current.tagName || '').toUpperCase()) >= 0) return true;
    current = current.parentNode;
  }
  return false;
}

/** 取触摸列表里第一个触点的坐标（touchstart 用 touches，touchend 用 changedTouches）。 */
function firstPoint(event, key) {
  const list = event ? event[key] : null;
  if (!list || list.length === 0) return null;
  const touch = list[0];
  if (!touch) return null;
  return { x: touch.clientX, y: touch.clientY };
}

/**
 * 把左右滑动切题绑到元素上，返回 `detach()`。
 * 边界（写死）：多指（`touches.length > 1`）忽略；起点在 `input` / `textarea` / `select` 内忽略；
 * 起点距左缘 < `SWIPE_EDGE_GUARD_PX` 不判右滑。
 */
function attachSwipe(el, handlers) {
  const opts = handlers || {};
  const onPrev = typeof opts.onPrev === 'function' ? opts.onPrev : null;
  const onNext = typeof opts.onNext === 'function' ? opts.onNext : null;
  let start = null;
  let detached = false;

  const onStart = (event) => {
    if (detached) return;
    const touches = event ? event.touches : null;
    if (!touches || touches.length !== 1) { start = null; return; } // 多指忽略
    const point = firstPoint(event, 'touches');
    if (!point || inEditable(event.target)) { start = null; return; }
    start = { x: point.x, y: point.y, guarded: point.x < SWIPE_EDGE_GUARD_PX };
  };

  const onEnd = (event) => {
    if (detached) return;
    const from = start;
    start = null;
    if (!from) return;
    const point = firstPoint(event, 'changedTouches');
    if (!point) return;
    const direction = swipeDirection(point.x - from.x, point.y - from.y);
    if (direction === 'prev') {
      if (from.guarded) return; // 左缘保护带：起点贴左缘不判右滑
      if (onPrev) onPrev();
    } else if (direction === 'next' && onNext) {
      onNext();
    }
  };

  el.addEventListener('touchstart', onStart);
  el.addEventListener('touchend', onEnd);
  return function detach() {
    detached = true;
    // DOM 桩没有 removeEventListener：有才调（§2.11.7④）
    if (typeof el.removeEventListener === 'function') {
      el.removeEventListener('touchstart', onStart);
      el.removeEventListener('touchend', onEnd);
    }
  };
}

__MODULES__["src/ui/gesture.mjs"] = { swipeDirection: swipeDirection, attachSwipe: attachSwipe, SWIPE_MIN_PX: SWIPE_MIN_PX, SWIPE_DOMINANCE: SWIPE_DOMINANCE, SWIPE_EDGE_GUARD_PX: SWIPE_EDGE_GUARD_PX };
}());
(function () {
// 题卡 —— 背题（§2.11.2 / R13 / R24：延迟揭示）与练习（§2.11.3 / R14 / R15 / R19 / R26）。

const { blankCountOf } = __MODULES__["src/core/answers.mjs"];
const { TYPE_LABELS } = __MODULES__["src/core/detect.mjs"];
const { displayText } = __MODULES__["src/core/grade.mjs"];
const { MODE_LABELS, TYPE_LABELS: GROUP_TYPE_LABELS } = __MODULES__["src/core/progress.mjs"];
const { add, clear, el } = __MODULES__["src/ui/dom.mjs"];
const { attachSwipe } = __MODULES__["src/ui/gesture.mjs"];

const MARK_LABELS = {
  partial: '数据残缺',
  outOfRange: '答案越界',
  suspect: '题型存疑',
  noAnswer: '无答案',
  formatUnresolved: '数值格式无法还原',
};

/**
 * 渲染一张题卡（每次状态变化整体重画，界面简单、无中间态泄漏）。
 * @param {HTMLElement} container
 * @param {object} view
 */
function renderCard(container, view) {
  const question = view.question;
  if (!question) {
    clear(container);
    add(container, el('main', { class: 'screen' }, [
      el('p', { class: 'empty', text: view.emptyText || '这一组暂时没有题目。' }),
      el('button', { class: 'btn', type: 'button', text: '返回入口', on: { click: view.onExit } }),
    ]));
    return;
  }

  const head = el('header', { class: 'card-head' }, [
    el('div', { class: 'card-meta' }, [
      el('span', { class: 'source-row', text: question.sourceRow }),
      el('span', { class: 'chip chip-type', text: TYPE_LABELS[question.type] || question.type }),
      ...(question.marks || []).map((mark) => el('span', { class: 'chip chip-mark', text: MARK_LABELS[mark] || mark })),
    ]),
    el('div', { class: 'card-actions' }, [
      view.canRemoveWrong
        ? el('button', {
          class: 'btn btn-ghost remove-wrong',
          type: 'button',
          text: '移出本题',
          on: { click: view.onRemoveWrong },
        })
        : null,
      el('button', { class: 'btn btn-ghost', type: 'button', text: '退出', on: { click: view.onExit } }),
    ]),
  ]);

  const body = el('section', { class: 'card-body' }, [el('p', { class: 'stem', text: question.stem })]);
  body.appendChild(view.mode === 'memorize' ? renderMemorizeBody(question, view) : renderPracticeBody(question, view));

  const foot = el('footer', { class: 'card-foot' }, [
    el('button', {
      class: 'btn',
      type: 'button',
      text: '上一题',
      disabled: view.index <= 0,
      on: { click: view.onPrev },
    }),
    el('span', {
      class: 'counter',
      text: `${view.index + 1} / ${view.total}`,
    }),
    el('button', {
      class: 'btn btn-primary',
      type: 'button',
      text: '下一题',
      disabled: view.index >= view.total - 1,
      on: { click: view.onNext },
    }),
  ]);

  const summary = el('p', {
    class: 'card-summary',
    text: `${MODE_LABELS[view.group.mode]} · ${GROUP_TYPE_LABELS[view.group.type] || view.group.type}` +
      (view.accuracy === null ? '' : ` · 正确率 ${(view.accuracy * 100).toFixed(0)}%（${view.answeredCount} 题已答）`),
  });

  const screen = el('main', { class: 'screen' }, [
    summary,
    el('article', { class: 'card' }, [
      head,
      body,
      // 末题答对后不循环，停在本卡并明示整组完成（§2.11.3 / R14）
      view.completed ? el('p', { class: 'banner banner-ok completed', text: '本组已完成' }) : null,
      // 底栏常驻（R25）：`上一题 / 计数 / 下一题`由 style.css 固定到屏幕底部
      el('div', { class: 'card-footer' }, [foot]),
    ]),
  ]);
  // 左右滑动切题（R26）：与「上一题 / 下一题」同一条路径（`move(app, ±1)`，先取消已排程的自动前进）
  attachSwipe(screen, { onPrev: view.onPrev, onNext: view.onNext });
  clear(container);
  add(container, screen);
}

/**
 * 背题题卡（R13 / R24）：首屏只给题干与全部选项原文；延迟揭示到点后再补正确项高亮与答案原文。
 * 没有「看答案」控件、没有翻答案状态；不判分、不记录；未置位（`answerShown !== true`）一律 fail-closed。
 */
function renderMemorizeBody(question, view) {
  const shown = view.answerShown === true; // §2.11.7②：fail-closed
  const nodes = [];
  if (question.type === 'single' || question.type === 'multi') {
    nodes.push(renderOptions(question, view, { interactive: false, showCorrect: shown }));
  } else if (question.type === 'judge') {
    nodes.push(judgeRow(question, { picked: null, interactive: false, showAnswer: shown }));
  }
  if (shown) nodes.push(answerBlock(question));
  return el('div', { class: 'answer-zone' }, nodes);
}

function renderPracticeBody(question, view) {
  const nodes = [];
  const feedback = view.feedback;
  const response = view.response || {};

  if (question.type === 'single' || question.type === 'multi') {
    nodes.push(renderOptions(question, view, { interactive: !feedback, showCorrect: !!feedback }));
  } else if (question.type === 'judge') {
    nodes.push(judgeRow(question, {
      picked: response.judge,
      interactive: !feedback,
      onPick: (payload) => view.onPick(payload),
    }));
  } else if (question.type === 'fill') {
    nodes.push(...fillZone(question, view, feedback));
  } else if (question.type === 'essay') {
    nodes.push(...essayZone(question, view));
  }

  if (question.type === 'multi' && !feedback) {
    nodes.push(el('button', {
      class: 'btn btn-primary',
      type: 'button',
      text: '提交',
      on: { click: () => view.onSubmit({ letters: response.letters || [] }) },
    }));
  }

  if (feedback) {
    nodes.push(el('div', { class: `feedback ${feedback.correct ? 'ok' : 'bad'}` }, [
      el('strong', { text: feedback.correct ? '答对了' : '答错了' }),
      el('p', { class: 'answer-raw', text: `正确答案：${feedback.expectedDisplay}` }),
    ]));
  }
  return el('div', { class: 'answer-zone' }, nodes);
}

/** 填空：每空一个输入框 + 提交（一次交全部空）。 */
function fillZone(question, view, feedback) {
  const nodes = [];
  const values = (view.response || {}).values || [];
  const blanks = Math.max(1, blankCountOf(question.stem) || question.answer.blanks || 1);
  const inputs = [];
  for (let i = 0; i < blanks; i++) {
    const input = el('input', {
      class: 'blank-input',
      type: 'text',
      inputmode: 'text',
      placeholder: `第 ${i + 1} 空`,
      disabled: !!feedback,
    });
    input.value = values[i] || '';
    inputs.push(input);
    nodes.push(el('label', { class: 'blank-row' }, [el('span', { class: 'blank-index', text: `第 ${i + 1} 空` }), input]));
  }
  if (!feedback) {
    nodes.push(el('button', {
      class: 'btn btn-primary',
      type: 'button',
      text: '提交',
      on: { click: () => view.onSubmit({ values: inputs.map((input) => input.value) }) },
    }));
  }
  return nodes;
}

/**
 * 简答（R19）：写答 → 提交 → 参考答案原文 → 自评「会了 / 不会」；
 * 全程不判分（app 层不调 grade()，对错来自自评）。
 */
function essayZone(question, view) {
  const stage = view.essayStage || 'write';
  const nodes = [];
  const input = el('textarea', { class: 'essay-input', rows: '3', disabled: stage !== 'write' });
  input.value = view.essayDraft || '';
  nodes.push(input);

  if (stage === 'write') {
    nodes.push(el('button', {
      class: 'btn btn-primary',
      type: 'button',
      text: '提交',
      on: { click: () => view.onEssaySubmit(input.value) },
    }));
    return nodes;
  }
  if (stage === 'assess') {
    nodes.push(answerBlock(question));
    nodes.push(el('div', { class: 'self-assess' }, [
      el('button', {
        class: 'btn btn-primary assess-yes', type: 'button', text: '会了',
        on: { click: () => view.onSelfAssess(true) },
      }),
      el('button', {
        class: 'btn assess-no', type: 'button', text: '不会',
        on: { click: () => view.onSelfAssess(false) },
      }),
    ]));
  }
  return nodes;
}

/** 答案原文（+ 填空的逐空原文），背题与简答自评都用它。 */
function answerBlock(question) {
  const block = el('div', { class: 'answer-block' }, [
    el('div', { class: 'answer-title', text: '答案' }),
    el('p', { class: 'answer-raw', text: displayText(question) }),
  ]);
  if (question.type === 'fill') {
    const segments = question.answer.segments;
    if (Array.isArray(segments) && segments.length > 0) {
      block.appendChild(el('ol', { class: 'blank-list' }, segments.map((segment, i) => el('li', { text: `第 ${i + 1} 空：${segment}` }))));
    }
  }
  return block;
}

/** 判断题的两个按钮：正确 / 错误（练习态可点即判；背题态 `showAnswer` 为真才标出正确项）。 */
function judgeRow(question, options) {
  const answerJudge = options.showAnswer === false ? null : question.answer.judge;
  const pick = options.interactive ? options.onPick : null;
  return el('div', { class: 'judge-row' }, [
    el('button', {
      class: judgeClass(options.picked, 'correct', answerJudge), type: 'button', text: '正确',
      disabled: !options.interactive, on: pick ? { click: () => pick({ judge: 'correct' }) } : undefined,
    }),
    el('button', {
      class: judgeClass(options.picked, 'wrong', answerJudge), type: 'button', text: '错误',
      disabled: !options.interactive, on: pick ? { click: () => pick({ judge: 'wrong' }) } : undefined,
    }),
  ]);
}

function renderOptions(question, view, options) {
  const response = view.response || {};
  const letters = response.letters || (response.letter ? [response.letter] : []);
  const correctLetters = (question.answer.letters || []).slice();
  const nodes = question.options.map((option) => {
    const classes = ['option'];
    if (letters.indexOf(option.letter) >= 0) classes.push('selected');
    if (options.showCorrect) {
      if (correctLetters.indexOf(option.letter) >= 0) classes.push('correct');
      else if (letters.indexOf(option.letter) >= 0) classes.push('wrong');
    }
    return el('button', {
      class: classes.join(' '),
      type: 'button',
      disabled: !options.interactive,
      on: options.interactive ? { click: () => view.onPick({ letter: option.letter }) } : undefined,
    }, [
      el('span', { class: 'option-letter', text: option.letter }),
      el('span', { class: 'option-text', text: option.text }),
    ]);
  });
  return el('div', { class: 'options' }, nodes);
}

function judgeClass(picked, value, answerJudge) {
  const classes = ['btn'];
  if (picked === value) classes.push('selected');
  if (answerJudge && value === answerJudge) classes.push('correct');
  if (answerJudge && picked === value && value !== answerJudge) classes.push('wrong');
  return classes.join(' ');
}

__MODULES__["src/ui/card.mjs"] = { renderCard: renderCard };
}());
(function () {
// 导入流程与列映射确认页（设计档 §2.6.5 / §2.11.5 / §2.12.1）。
// 识别结果先展示后入库：用户确认之前，存储层不会出现题库记录。

const { buildBank, makeBankRecord } = __MODULES__["src/core/bank.mjs"];
const { colIndex, detectSheet, mappingFromDetection, OPTION_LETTERS, validateMapping, TYPE_LABELS } = __MODULES__["src/core/detect.mjs"];
const { readXlsx } = __MODULES__["src/core/xlsx.mjs"];
const { add, clear, el } = __MODULES__["src/ui/dom.mjs"];

/** 导入流程控制器：inflateRaw / inflateOk / inflateReason / onConfirm / onCancel / onImportDone。 */
function createImportFlow(options) {
  const state = {
    phase: 'pick',
    fileName: '',
    fileSize: 0,
    importMs: 0,
    progress: null,
    error: null,
    sheets: [],
    detections: [],
    roles: [],
    headers: [],
    headerTouched: [],
    skip: [],
    lastSummary: null,
  };
  let container = null;
  let progressNode = null;

  function render(target) {
    if (target) container = target;
    if (!container) return;
    clear(container);
    if (state.phase === 'pick') add(container, renderPick());
    else if (state.phase === 'parsing') add(container, renderParsing());
    else if (state.phase === 'confirm') add(container, renderConfirm());
    else add(container, renderError());
  }

  async function loadFile(file) {
    if (!file) return;
    if (options.inflateOk === false) {
      state.phase = 'error';
      state.error = options.inflateReason || '此浏览器不支持导入所需能力。';
      console.error('导入不可用：', state.error);
      render();
      return;
    }
    state.fileName = file.name || '(未命名)';
    state.fileSize = file.size || 0;
    state.phase = 'parsing';
    state.progress = { processed: 0, total: 0 };
    render();

    const started = now();
    try {
      const buffer = await file.arrayBuffer();
      const bytes = new Uint8Array(buffer);
      const parsed = await readXlsx(bytes, {
        inflateRaw: options.inflateRaw,
        onProgress: (p) => {
          state.progress = p;
          if (progressNode) progressNode.textContent = progressText(p);
        },
      });
      state.sheets = parsed.sheets;
      state.detections = parsed.sheets.map((sheet) => detectSheet(sheet));
      state.roles = state.detections.map((d) => mappingFromDetection(d));
      state.headers = state.detections.map((d) => d.headerRow);
      state.headerTouched = state.detections.map(() => false);
      state.skip = state.detections.map(() => false);
      state.importMs = Math.round(now() - started);
      state.phase = 'confirm';
      render();
      if (options.onImportDone) options.onImportDone(state.importMs);
    } catch (err) {
      state.phase = 'error';
      state.error = errorMessage(err);
      console.error('导入失败', err); // 错误不吞：控制台 + 界面双写（§2.11.5）
      render();
    }
  }

  function detectionWithRoles(index) {
    const detection = state.detections[index];
    return Object.assign({}, detection, {
      roles: state.roles[index],
      headerRow: state.headers[index],
      skip: state.skip[index],
    });
  }

  /** 逐表校验 + 未消解的不确定项 → 待确认计数（§2.6.5 第 3、4 条）。 */
  function evaluate() {
    const perSheet = state.detections.map((detection, index) => {
      if (state.skip[index]) return { skipped: true, errors: [], unresolved: [] };
      const check = validateMapping({
        roles: state.roles[index],
        sheetType: detection.sheetType,
        type: detection.typeColumn,
      });
      const unresolved = (detection.reasons || [])
        .filter((reason) => !isResolved(reason, index))
        .map((reason) => reason.why);
      return { skipped: false, errors: check.errors, unresolved };
    });
    const pending = perSheet.reduce((sum, item) => sum + item.errors.length + item.unresolved.length, 0);
    return { perSheet, pending };
  }

  function isResolved(reason, index) {
    const roles = state.roles[index] || {};
    const counts = countRoles(roles);
    if (reason.role === 'headerRow') return state.headerTouched[index] === true;
    if (reason.role === 'type') return counts.type === 1 || (!counts.type && !!state.detections[index].sheetType);
    if (reason.role === 'stem') return counts.stem === 1;
    if (reason.role === 'answer') return counts.answer === 1;
    if (reason.role === 'option') {
      const type = state.detections[index].sheetType;
      const need = type === 'single' || type === 'multi';
      return need ? counts.option >= 2 && counts.option <= 8 : true;
    }
    return true;
  }

  function confirm() {
    const evaluation = evaluate();
    if (evaluation.pending > 0) return null;
    const detections = state.detections.map((_, index) => detectionWithRoles(index));
    const built = buildBank(state.sheets, detections, { sourceName: state.fileName });
    const record = makeBankRecord({
      questions: built.questions,
      fingerprints: built.fingerprints,
      sourceName: state.fileName,
      builtAt: new Date().toISOString(),
    });
    state.lastSummary = built.stats;
    if (options.onConfirm) options.onConfirm(record, built.stats);
    return record;
  }

  function renderPick() {
    const blocked = options.inflateOk === false;
    const input = el('input', {
      class: 'file-input',
      type: 'file',
      accept: '.xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      disabled: blocked,
      on: {
        change: (event) => {
          const files = event && event.target && event.target.files ? event.target.files : [];
          loadFile(files[0]);
        },
      },
    });
    return el('main', { class: 'screen' }, [
      el('h2', { class: 'screen-title', text: '导入题库' }),
      el('p', { class: 'hint', text: '选择 .xlsx 题库文件；解析与识别都在本机完成，不联网、不上传。' }),
      blocked ? el('p', { class: 'banner banner-bad', text: options.inflateReason }) : null,
      el('div', { class: 'pick-row' }, [input]),
      el('div', { class: 'actions' }, [
        el('button', { class: 'btn', type: 'button', text: '返回', on: { click: options.onCancel } }),
      ]),
    ]);
  }

  function renderParsing() {
    progressNode = el('p', { class: 'progress', text: progressText(state.progress) });
    return el('main', { class: 'screen' }, [
      el('h2', { class: 'screen-title', text: '正在读取题库' }),
      el('p', { class: 'hint', text: state.fileName }),
      progressNode,
    ]);
  }

  function renderError() {
    return el('main', { class: 'screen' }, [
      el('h2', { class: 'screen-title', text: '导入失败' }),
      el('p', { class: 'banner banner-bad', text: state.error || '未知错误' }),
      el('p', { class: 'hint', text: '可执行的下一步：确认文件是 .xlsx 题库；若浏览器不支持导入，请改用单文件形态。' }),
      el('div', { class: 'actions' }, [
        el('button', { class: 'btn', type: 'button', text: '返回', on: { click: options.onCancel } }),
      ]),
    ]);
  }

  function renderConfirm() {
    const evaluation = evaluate();
    let totalQuestions = 0;
    const blocks = state.detections.map((detection, index) => {
      const built = state.skip[index]
        ? { questions: [] }
        : buildBank([state.sheets[index]], [detectionWithRoles(index)], { sourceName: state.fileName });
      totalQuestions += built.questions.length;
      return renderSheetBlock(detection, index, built.questions, evaluation.perSheet[index]);
    });

    const summary = el('p', {
      class: 'file-summary',
      text: `文件：${state.fileName}（${state.fileSize} 字节）· 工作表 ${state.sheets.length} 个 · 识别题目合计 ${totalQuestions} 题`,
    });

    const sheetsLine = el('ul', { class: 'sheet-counts' }, state.detections.map((detection, index) => el('li', {
      text: `${detection.name}：${state.skip[index] ? 0 : countOf(index)} 题`,
    })));

    const button = el('button', {
      class: 'btn btn-primary confirm-button',
      type: 'button',
      text: evaluation.pending > 0 ? `有 ${evaluation.pending} 处待确认` : '确认入库',
      disabled: evaluation.pending > 0,
      on: { click: () => confirm() },
    });

    return el('main', { class: 'screen' }, [
      el('h2', { class: 'screen-title', text: '确认识别结果' }),
      summary,
      sheetsLine,
      el('p', { class: 'hint', text: '逐表确认列映射；识别不确定的地方会标出来，改正后按钮才可用。' }),
      el('div', { class: 'sheet-blocks' }, blocks),
      el('div', { class: 'actions' }, [
        button,
        el('button', { class: 'btn', type: 'button', text: '取消', on: { click: options.onCancel } }),
      ]),
    ]);
  }

  function countOf(index) {
    const built = buildBank([state.sheets[index]], [detectionWithRoles(index)], { sourceName: state.fileName });
    return built.questions.length;
  }

  function renderSheetBlock(detection, index, previewQuestions, evaluation) {
    const block = el('section', { class: 'sheet-block', dataset: { sheet: detection.name } });
    block.appendChild(el('h3', { class: 'sheet-name', text: `${detection.name}　${TYPE_LABELS[detection.sheetType] || '题型未知'}` }));
    block.appendChild(el('p', {
      class: 'sheet-meta',
      text: `表头行：${detection.headerRow === null ? '无表头' : `第 ${detection.headerRow} 行`}` +
        ` · 数据行：${detection.dataRowCount} · 识别题数：${state.skip[index] ? 0 : previewQuestions.length}`,
    }));

    const headerSelect = el('select', {
      class: 'header-select',
      on: {
        change: (event) => {
          const value = event.target.value;
          state.headers[index] = value === '0' ? null : Number(value);
          state.headerTouched[index] = true;
          render();
        },
      },
    });
    ['无表头', '第 1 行', '第 2 行', '第 3 行'].forEach((label, i) => {
      const option = el('option', { value: String(i), text: label });
      if ((state.headers[index] === null ? 0 : state.headers[index]) === i) option.selected = true;
      headerSelect.appendChild(option);
    });
    block.appendChild(el('label', { class: 'field' }, [el('span', { text: '表头行' }), headerSelect]));

    const mappingList = el('div', { class: 'mapping-list' });
    const allColumns = detection.allColumns || [];
    const optionPositions = optionColumnPositions(state.roles[index]);
    for (const column of allColumns) {
      const select = el('select', {
        class: 'role-select',
        dataset: { column },
        on: {
          change: (event) => {
            state.roles[index][column] = roleFromSelectValue(event.target.value);
            render();
          },
        },
      });
      selectOptionsFor(column).forEach((choice) => {
        const option = el('option', { value: choice.value, text: choice.label });
        if (choice.value === selectValueFor(state.roles[index][column], column, optionPositions)) option.selected = true;
        select.appendChild(option);
      });
      mappingList.appendChild(el('label', { class: 'field' }, [el('span', { text: `列 ${column}` }), select]));
    }
    block.appendChild(mappingList);

    if (detection.reasons && detection.reasons.length) {
      block.appendChild(el('ul', { class: 'reasons' }, detection.reasons.map((reason) => el('li', {
        class: isResolved(reason, index) ? 'reason resolved' : 'reason',
        text: `${reason.role}：${reason.why}`,
      }))));
    }

    if (evaluation && evaluation.errors.length) {
      block.appendChild(el('ul', { class: 'errors' }, evaluation.errors.map((message) => el('li', { text: message }))));
    }

    if (!state.skip[index]) {
      const samples = previewQuestions.slice(0, 3);
      const previewList = samples.length
        ? el('div', { class: 'preview' }, samples.map((question) => renderPreviewItem(question)))
        : el('p', { class: 'hint', text: '当前映射下识别不到题目。' });
      block.appendChild(el('div', { class: 'preview-wrap' }, [el('div', { class: 'preview-title', text: '抽样预览（3 条）' }), previewList]));
    }

    const skipBox = el('input', {
      type: 'checkbox',
      class: 'skip-toggle',
      checked: state.skip[index],
      on: {
        change: (event) => {
          state.skip[index] = !!event.target.checked;
          render();
        },
      },
    });
    block.appendChild(el('label', { class: 'field field-skip' }, [skipBox, el('span', { text: '整表跳过' })]));
    return block;
  }

  function renderPreviewItem(question) {
    const nodes = [el('p', { class: 'preview-stem', text: `${question.sourceRow ?? ''} ${question.stem}` })];
    if (question.options.length) {
      nodes.push(el('p', {
        class: 'preview-options',
        text: question.options.map((option) => `${option.letter}. ${option.text}`).join('　'),
      }));
    }
    nodes.push(el('p', { class: 'preview-answer', text: `答案：${question.answer.display || question.answer.raw || '（空）'}` }));
    if (question.marks.length) {
      nodes.push(el('p', { class: 'preview-marks', text: `标记：${question.marks.join('、')}` }));
    }
    return el('div', { class: 'preview-item' }, nodes);
  }

  return {
    state,
    render,
    loadFile,
    confirm,
    evaluate,
    isResolved,
  };
}

function progressText(progress) {
  if (!progress || !progress.total) return '正在读取…';
  return `正在读取… ${progress.processed} / ${progress.total} 个单元格`;
}

function selectOptionsFor() {
  const out = [
    { value: 'type', label: '题型' },
    { value: 'stem', label: '题干' },
  ];
  for (let i = 0; i < OPTION_LETTERS.length; i++) out.push({ value: `opt:${OPTION_LETTERS[i]}`, label: `选项${OPTION_LETTERS[i]}` });
  out.push({ value: 'answer', label: '答案' }, { value: 'ignore', label: '忽略' });
  return out;
}

/** 选项列的角色只记 'option'——列位置按列序确定（§2.6.5 第 4 条②「列序即位置」）。 */
function roleFromSelectValue(value) {
  if (value.indexOf('opt:') === 0) return 'option';
  return value === 'ignore' ? 'ignore' : value;
}

function selectValueFor(role, column, positions) {
  if (role === 'option') return `opt:${OPTION_LETTERS[positions[column] === undefined ? 0 : positions[column]]}`;
  if (role === 'ignore' || !role) return 'ignore';
  return role;
}

function optionColumnPositions(roles) {
  const positions = {};
  let i = 0;
  Object.keys(roles || {})
    .filter((col) => roles[col] === 'option')
    .sort((a, b) => colIndex(a) - colIndex(b))
    .forEach((col) => {
      positions[col] = i;
      i += 1;
    });
  return positions;
}

function countRoles(roles) {
  const counts = { type: 0, stem: 0, answer: 0, option: 0 };
  for (const column of Object.keys(roles || {})) {
    const role = roles[column];
    if (counts[role] !== undefined) counts[role] += 1;
  }
  return counts;
}

function now() {
  if (globalThis.performance && typeof globalThis.performance.now === 'function') return globalThis.performance.now();
  return Date.now();
}

function errorMessage(err) {
  return err && err.message ? err.message : String(err);
}

__MODULES__["src/ui/import.mjs"] = { createImportFlow: createImportFlow, countRoles: countRoles };
}());
(function () {
// 入口 / 组的视图模型 —— 由 app 状态派生「这组有多少题、什么顺序、空态说什么」
// （设计档 §2.11.1 / §2.11.5 / §2.11.6）。由 app.mjs 迁出，守 N8 行数纪律；纯读，不改状态。

const { accuracy, groupKey, isCollection, masteryOf, masterySummary, selectQuestions, selectUnmastered, shuffledOrder, WRONGBOOK, wrongbookOrderByWeight } = __MODULES__["src/core/progress.mjs"];

function uiFlag(app, key) {
  return !!(app.uiState && app.uiState[key]);
}

/** 乱序种子：未生成过（仍为 0）时同种子同序，仍确定（R20）。 */
function seedOf(app) {
  const seed = app.uiState ? app.uiState.shuffleSeed : null;
  return typeof seed === 'number' ? seed : 0;
}

function progressList(app) {
  return Array.from(app.progressCache.values());
}

/** 掌握度聚合口径：跨【全部已存进度组】，含两个集合型组（§2.11.6①）。 */
function currentMastery(app) {
  return masteryOf((app.bank && app.bank.questions) || [], progressList(app));
}

/** 入口 meta 的原料：背题 = 正确率；练习 6 组 = 掌握度三分；集合型 = 题数。 */
function groupSummary(app, mode, type) {
  const questions = (app.bank && app.bank.questions) || [];
  if (!app.bank) {
    if (isCollection(type)) return { total: 0 };
    return mode === 'practice'
      ? { total: 0, mastery: { total: 0, mastered: 0, failed: 0, untrained: 0 } }
      : { total: 0, accuracy: null };
  }
  if (isCollection(type)) return { total: orderFor(app, mode, type).length };
  const selected = selectQuestions(questions, mode, type, app.wrong);
  if (mode === 'practice') {
    return { total: selected.length, mastery: masterySummary(selected, progressList(app)) };
  }
  const progress = app.progressCache.get(groupKey(mode, type));
  return { total: selected.length, accuracy: progress ? accuracy(progress) : null };
}

/** 该入口的题目顺序（源表序为底；练习 6 组可叠加「只练没掌握的」与「乱序」）。 */
function orderFor(app, mode, type) {
  const questions = (app.bank && app.bank.questions) || [];
  if (mode === 'practice' && type === WRONGBOOK) {
    // 错题本 = 错次降序、平手按源表下标升序（R17），不洗牌、不受开关影响
    return wrongbookOrderByWeight(questions.map((q) => q.id), app.wrong, app.wrongCounts);
  }
  if (isCollection(type)) return selectQuestions(questions, mode, type, app.wrong).map((q) => q.id);
  let selected = selectQuestions(questions, mode, type, app.wrong);
  if (mode !== 'practice') return selected.map((q) => q.id);
  if (uiFlag(app, 'practiceUnmasteredOnly')) selected = selectUnmastered(selected, currentMastery(app));
  let ids = selected.map((q) => q.id);
  if (uiFlag(app, 'practiceShuffled')) ids = shuffledOrder(ids, seedOf(app));
  return ids;
}

/** 过滤后为空的练习组要说明「都掌握了」，不是「没有题目」（§2.11.5）。 */
function emptyTextFor(app) {
  if (!app.bank || !app.group || app.order.length > 0) return null;
  const { mode, type } = app.group;
  if (mode !== 'practice' || isCollection(type) || !uiFlag(app, 'practiceUnmasteredOnly')) return null;
  const selected = selectQuestions(app.bank.questions, mode, type, app.wrong);
  return selected.length > 0 ? '这一组的题都已掌握 —— 没有需要重练的题。' : null;
}

__MODULES__["src/ui/groups.mjs"] = { uiFlag: uiFlag, seedOf: seedOf, currentMastery: currentMastery, groupSummary: groupSummary, orderFor: orderFor, emptyTextFor: emptyTextFor };
}());
(function () {
// 入口矩阵与统计（设计档 §2.11.1 / §2.11.6：14 个入口 + 掌握度 meta + 两个开关）。
// R6 / R8 / R9 / R16 / R18 / R20。

const { isCollection, MEMORIZE_TYPES, MODE_LABELS, NUMBERS, PRACTICE_TYPES, TYPE_LABELS, WRONGBOOK } = __MODULES__["src/core/progress.mjs"];
const { add, clear, el } = __MODULES__["src/ui/dom.mjs"];

const MODE_HINTS = {
  memorize: '看题 → 一次给全，不判分',
  practice: '作答 → 立即判对错',
};

/** iOS 且非主屏全屏打开时，提示一次「分享 → 添加到主屏幕」（可关闭，§2.11.4）。 */
function shouldShowInstallHint(ui) {
  if (ui && ui.installHintDismissed) return false;
  const nav = globalThis.navigator || {};
  const ua = String(nav.userAgent || '');
  const isIOS = /iPhone|iPad|iPod/.test(ua) || (/(Macintosh|Mac OS X)/.test(ua) && nav.maxTouchPoints > 1);
  if (!isIOS) return false;
  const standalone = nav.standalone === true
    || (typeof globalThis.matchMedia === 'function' && globalThis.matchMedia('(display-mode: standalone)').matches);
  return !standalone;
}

/**
 * 渲染入口矩阵（背题 × 6 + 练习 × 6 + 错题本 + 数字专项 = 14 组）。
 * @param {HTMLElement} container
 * @param {object} view
 */
function renderEntryMatrix(container, view) {
  clear(container);
  const bank = view.bank;
  const screen = el('main', { class: 'screen' }, [
    el('h1', { class: 'app-title', text: '高级工背题' }),
  ]);

  if (shouldShowInstallHint(view.ui)) {
    screen.appendChild(el('p', { class: 'banner banner-info install-hint' }, [
      el('span', { text: '想要全屏离线使用：Safari 分享 → 「添加到主屏幕」。' }),
      el('button', {
        class: 'btn btn-ghost dismiss-install-hint',
        type: 'button',
        text: '知道了',
        on: { click: view.onDismissInstallHint },
      }),
    ]));
  }

  const store = view.storeInfo || { backend: 'memory', persistent: false };
  screen.appendChild(el('p', {
    class: 'store-status',
    text: `存储：${store.backend}${store.persistent ? '' : '（本次使用不会保存进度）'}`,
  }));
  if (!store.persistent) {
    screen.appendChild(el('p', {
      class: 'banner banner-warn',
      text: '本次使用不会保存进度（此形态下浏览器不允许存储）。功能全部可用。',
    }));
  }
  if (store.degraded) {
    screen.appendChild(el('p', { class: 'banner banner-warn', text: store.degraded }));
  }
  if (view.notice) {
    screen.appendChild(el('p', { class: 'banner banner-warn bank-notice', text: view.notice }));
  }
  if (!view.inflateOk) {
    screen.appendChild(el('p', { class: 'banner banner-bad', text: view.inflateReason }));
  }

  if (!bank) {
    screen.appendChild(el('p', { class: 'empty', text: '还没有题库。请先导入 .xlsx 题库文件，或者使用内置题库的单文件形态。' }));
    screen.appendChild(el('div', { class: 'actions' }, [
      el('button', {
        class: 'btn btn-primary import-button',
        type: 'button',
        text: '导入题库',
        on: { click: view.onImport },
      }),
    ]));
    add(container, screen);
    return;
  }

  for (const mode of ['memorize', 'practice']) {
    const group = el('section', { class: 'matrix-group', dataset: { mode } });
    group.appendChild(el('h2', { class: 'matrix-title', text: `${MODE_LABELS[mode]}　` }));
    group.appendChild(el('p', { class: 'matrix-hint', text: MODE_HINTS[mode] }));
    const types = mode === 'practice' ? PRACTICE_TYPES : MEMORIZE_TYPES;
    const grid = el('div', { class: 'matrix-grid' });
    for (const type of types) {
      grid.appendChild(entryButton({
        mode,
        type,
        summary: view.summaryOf(mode, type),
        onEnter: view.onEnter,
      }));
    }
    group.appendChild(grid);
    // 掌握度开关与两个集合型入口只挂在练习区（§2.11.1 / §2.11.4）
    if (mode === 'practice') group.appendChild(renderPracticeExtras(view));
    screen.appendChild(group);
  }

  screen.appendChild(el('div', { class: 'actions' }, [
    el('button', {
      class: 'btn import-button',
      type: 'button',
      text: '导入题库',
      on: { click: view.onImport },
    }),
  ]));

  add(container, screen);
}

/** 练习区附加区：两个开关 + 错题本（含清空已掌握）+ 数字专项。 */
function renderPracticeExtras(view) {
  const extras = el('div', { class: 'practice-extras' });
  extras.appendChild(el('div', { class: 'switch-row' }, [
    switchButton('只练没掌握的', 'toggle-unmastered', isOn(view, 'practiceUnmasteredOnly'), view.onToggleUnmasteredOnly),
    switchButton('乱序', 'toggle-shuffled', isOn(view, 'practiceShuffled'), view.onToggleShuffled),
  ]));

  const wrongSummary = view.summaryOf('practice', WRONGBOOK);
  extras.appendChild(el('div', { class: 'wrongbook-row' }, [
    entryButton({ mode: 'practice', type: WRONGBOOK, summary: wrongSummary, onEnter: view.onEnter }),
    el('button', {
      class: 'btn btn-ghost clear-mastered',
      type: 'button',
      text: '清空已掌握',
      disabled: wrongSummary.total === 0,
      on: { click: view.onClearMastered },
    }),
  ]));
  extras.appendChild(entryButton({
    mode: 'practice',
    type: NUMBERS,
    summary: view.summaryOf('practice', NUMBERS),
    onEnter: view.onEnter,
  }));
  return extras;
}

/** 开关控件：button + aria-pressed，开启态加 data-on（§2.11.4）。 */
function switchButton(label, className, on, onToggle) {
  return el('button', {
    class: `btn switch ${className}`,
    type: 'button',
    text: label,
    'aria-pressed': on ? 'true' : 'false',
    dataset: on ? { on: 'true' } : {},
    on: { click: onToggle },
  });
}

function isOn(view, key) {
  return !!(view.ui && view.ui[key]);
}

function entryButton(options) {
  const summary = options.summary || { total: 0, accuracy: null };
  const label = TYPE_LABELS[options.type] || options.type;
  // 两个集合型入口即使为空也可进入 —— 进去看到的空态是 AC-16 要求的「明确呈现」
  const disabled = summary.total === 0 && !isCollection(options.type);
  return el('button', {
    class: `btn entry entry-${options.type}`,
    type: 'button',
    dataset: { group: `${options.mode}|${options.type}` },
    disabled,
    on: { click: () => options.onEnter(options.mode, options.type) },
  }, [
    el('span', { class: 'entry-label', text: label }),
    el('span', { class: 'entry-meta', text: entryMeta(summary, options.mode, options.type) }),
  ]);
}

/**
 * 入口 meta：背题区「N 题 · 正确率」、练习 6 组「共 N 题 · 已掌握 a / 未掌握 b / 未练 c」、
 * 集合型入口「N 题」（§2.11.6①）。
 */
function entryMeta(summary, mode, type) {
  if (summary.total === 0) return '暂无题目';
  if (isCollection(type)) return `${summary.total} 题`;
  if (mode === 'practice' && summary.mastery) {
    const mastery = summary.mastery;
    return `共 ${summary.total} 题 · 已掌握 ${mastery.mastered} / 未掌握 ${mastery.failed} / 未练 ${mastery.untrained}`;
  }
  const accuracyText = summary.accuracy === null || summary.accuracy === undefined
    ? '未开始'
    : `正确率 ${(summary.accuracy * 100).toFixed(0)}%`;
  return `${summary.total} 题 · ${accuracyText}`;
}

// 入口矩阵的渲染到此为止：空态与矩阵两态都在上面处理。

__MODULES__["src/ui/stats.mjs"] = { shouldShowInstallHint: shouldShowInstallHint, renderEntryMatrix: renderEntryMatrix };
}());
(function () {
// 界面入口 —— 挂载 / 状态 / 路由与入口接线（设计档 §2.4.2 / §2.11）。
// 会话持久化与诊断字段在 session.mjs；纯逻辑在 core/。
const { displayText, grade } = __MODULES__["src/core/grade.mjs"];
const { inflateRaw, probeInflate } = __MODULES__["src/core/inflate.mjs"];
const { accuracy, addWrong, bumpWrongCount, cursorIndex, ensureProgress, groupKey, recordAnswer, removeWrong, WRONGBOOK } = __MODULES__["src/core/progress.mjs"];
const { openStore } = __MODULES__["src/core/store.mjs"];
const { adoptBank, loadBank } = __MODULES__["src/ui/bank-load.mjs"];
const { renderCard } = __MODULES__["src/ui/card.mjs"];
const { clear } = __MODULES__["src/ui/dom.mjs"];
const { createImportFlow } = __MODULES__["src/ui/import.mjs"];
const { persistProgress, persistWrongbook, saveUiState, writeDiag } = __MODULES__["src/ui/session.mjs"];
const { currentMastery, emptyTextFor, groupSummary, orderFor, uiFlag } = __MODULES__["src/ui/groups.mjs"];
const { renderEntryMatrix } = __MODULES__["src/ui/stats.mjs"];

/** 答对后停留时长（毫秒）：让用户看清对错再走（§2.11.3）。 */
const FEEDBACK_HOLD_MS = 600;

/** 背题延迟揭示时长（毫秒）：先给题干与选项原文，到点后补上正确项与答案原文（R24 / §2.11.7②）。 */
const MEMORIZE_REVEAL_MS = 2000;

/** 挂载界面（打包产物暴露为 window.EXAM_MEMO.mount）。 */
async function mount(rootEl, options) {
  const opts = options || {};
  const started = now();
  const timers = opts.timers || { set: (fn, ms) => setTimeout(fn, ms), clear: (handle) => clearTimeout(handle) };
  const app = {
    root: rootEl,
    view: 'matrix',
    bank: null,
    store: null,
    storeInfo: { backend: 'memory', persistent: false, degraded: null },
    inflate: { ok: false, reason: '' },
    group: null,
    order: [],
    index: 0,
    progress: null,
    progressCache: new Map(),
    wrong: new Set(),
    wrongCounts: {},
    response: null, feedback: null, essayStage: 'write', essayDraft: '', completed: false,
    answerShown: false, // 背题延迟揭示（R24）：每题先给题干与选项原文，到点才置位
    feedbackHoldMs: typeof opts.feedbackHoldMs === 'number' ? opts.feedbackHoldMs : FEEDBACK_HOLD_MS,
    memorizeRevealMs: typeof opts.memorizeRevealMs === 'number' ? opts.memorizeRevealMs : MEMORIZE_REVEAL_MS,
    timers, pendingTimer: null,
    questionIndex: new Map(),
    flow: null, uiState: null, notice: null,
    importMs: 0, renderMs: 0, readyMs: 0,
  };

  writeDiag(app);
  clear(rootEl);

  try { app.inflate = await probeInflate(); } catch (err) { app.inflate = { ok: false, reason: errorMessage(err) }; }
  try { app.store = await openStore({}); app.storeInfo = app.store.describe(); } catch (err) { app.storeInfo = { backend: 'memory', persistent: false, degraded: errorMessage(err) }; }

  await loadBank(app, opts, render);
  app.readyMs = Math.round(now() - started);
  render(app);
  registerServiceWorker();
  return app;
}


function render(app) {
  const started = now();
  if (app.view === 'import' && app.flow) app.flow.render(app.root);
  else if (app.view === 'card' && app.bank) renderQuestion(app);
  else renderMatrix(app);
  app.renderMs = Math.round(now() - started);
  writeDiag(app);
}

function renderMatrix(app) {
  renderEntryMatrix(app.root, {
    bank: app.bank,
    storeInfo: app.storeInfo,
    inflateOk: app.inflate.ok,
    inflateReason: app.inflate.reason,
    summaryOf: (mode, type) => groupSummary(app, mode, type),
    onEnter: (mode, type) => enterGroup(app, mode, type),
    onImport: () => openImport(app),
    onClearMastered: () => clearMastered(app),
    onToggleUnmasteredOnly: () => toggleUnmasteredOnly(app),
    onToggleShuffled: () => toggleShuffled(app),
    ui: app.uiState,
    notice: app.notice,
    onDismissInstallHint: () => dismissInstallHint(app),
  });
}

async function enterGroup(app, mode, type) {
  cancelAdvance(app);
  const key = groupKey(mode, type);
  let stored = app.progressCache.get(key) || null;
  if (!stored && app.store) {
    try { stored = await app.store.loadProgress(key); } catch (err) { stored = null; }
  }
  const progress = ensureProgress(stored, key);
  app.progressCache.set(key, progress);
  app.group = { mode, type, groupKey: key };
  app.order = orderFor(app, mode, type);
  app.progress = progress;
  app.index = cursorIndex(app.order, progress.cursorQid);
  resetQuestionState(app);
  app.view = 'card';
  render(app);
}

/** 离开题卡 / 返回入口：先取消排程，再换屏（§2.11.3 排程取消不变量）。 */
function exitGroup(app) {
  cancelAdvance(app);
  app.view = 'matrix';
  app.group = null;
  render(app);
}

/** 换题 / 切组的统一复位：清作答态、取消在用排程，背题组再起一次延迟揭示（§2.11.2 / R24）。 */
function resetQuestionState(app) {
  app.response = null;
  app.feedback = null;
  app.essayStage = 'write';
  app.essayDraft = '';
  app.completed = false;
  app.answerShown = false; // 每题重新起算（渲染侧 fail-closed：非 true 不显示答案）
  cancelAdvance(app);
  scheduleReveal(app);
}

function currentQuestion(app) {
  const qid = app.order[app.index];
  return qid ? app.questionIndex.get(qid) || null : null;
}

function renderQuestion(app) {
  const question = currentQuestion(app);
  renderCard(app.root, {
    question,
    mode: app.group.mode,
    group: app.group,
    index: app.index, total: app.order.length,
    response: app.response, feedback: app.feedback,
    essayStage: app.essayStage, essayDraft: app.essayDraft,
    completed: app.completed, answerShown: app.answerShown, emptyText: emptyTextFor(app),
    accuracy: app.progress ? accuracy(app.progress) : null,
    answeredCount: app.progress ? app.progress.stats.attempts : 0,
    canRemoveWrong: app.group.mode === 'practice' && app.group.type === WRONGBOOK,
    onExit: () => exitGroup(app),
    onPrev: () => move(app, -1), onNext: () => move(app, 1),
    onPick: (payload) => handlePick(app, payload),
    onSubmit: (payload) => resolveAnswer(app, payload),
    onEssaySubmit: (text) => submitEssay(app, text),
    onSelfAssess: (correct) => selfAssess(app, correct),
    onRemoveWrong: () => removeFromWrongbook(app),
  });
}

function move(app, delta) {
  cancelAdvance(app); // 手点上一题 / 下一题 = 改游标，先取消在用排程
  const next = app.index + delta;
  if (next < 0 || next >= app.order.length) return;
  goTo(app, next);
}

function goTo(app, next) {
  app.index = next;
  resetQuestionState(app);
  const question = currentQuestion(app);
  if (question && app.progress) {
    app.progress = { ...app.progress, cursorQid: question.id };
    persistProgress(app);
  }
  render(app);
}

/** 答对后的自动前进（§2.11.3）：到点才动，任何改 order / 离开题卡的动作都会先取消它。 */
function advance(app) {
  if (app.index + 1 >= app.order.length) return; // 末题答对不循环（D12）
  goTo(app, app.index + 1);
}

function scheduleAdvance(app) {
  cancelAdvance(app);
  app.pendingTimer = app.timers.set(() => {
    app.pendingTimer = null;
    advance(app);
  }, app.feedbackHoldMs);
}

function cancelAdvance(app) {
  if (app.pendingTimer === null || app.pendingTimer === undefined) return;
  app.timers.clear(app.pendingTimer);
  app.pendingTimer = null;
}

/** 背题延迟揭示（R24 / §2.11.7②）：复用 `app.pendingTimer`，换题 / 切组 / 退出都走既有取消路径。 */
function scheduleReveal(app) {
  if (!app.group || app.group.mode !== 'memorize') return; // 与练习态无关：不排程、也不白重画一次
  app.pendingTimer = app.timers.set(() => {
    app.pendingTimer = null;
    app.answerShown = true;
    render(app);
  }, app.memorizeRevealMs);
}

function handlePick(app, payload) {
  if (app.group.mode !== 'practice' || app.feedback) return;
  const question = currentQuestion(app);
  if (!question) return;
  if (question.type === 'multi') {
    const current = (app.response && app.response.letters) || [];
    const next = current.indexOf(payload.letter) >= 0
      ? current.filter((letter) => letter !== payload.letter)
      : current.concat([payload.letter]);
    app.response = { letters: next.slice().sort() };
    render(app);
    return;
  }
  if (question.type === 'single' || question.type === 'judge') resolveAnswer(app, payload);
}

function resolveAnswer(app, response) {
  const question = currentQuestion(app);
  if (!question || app.feedback) return;
  let result;
  try { result = grade(question, response); } catch (err) { console.error('判分失败', err); return; }
  app.response = response;
  finishAnswer(app, question, result.correct, result.expectedDisplay);
}

/** 简答写答提交：只呈现参考答案，对错留给自评（R19，不调 grade()）。 */
function submitEssay(app, text) {
  const question = currentQuestion(app);
  if (!question || question.type !== 'essay' || app.feedback) return;
  app.essayDraft = text;
  app.essayStage = 'assess';
  render(app);
}

/** 简答自评（R19）：会了 = true；不会 = false + 进错题本 + 错次 +1。 */
function selfAssess(app, correct) {
  const question = currentQuestion(app);
  if (!question || question.type !== 'essay' || app.feedback || app.essayStage !== 'assess') return;
  app.essayStage = 'done';
  finishAnswer(app, question, correct, displayText(question));
}

/** 收口一次作答：判分 → 记统计 → 错题本/错次 → 答对排程自动前进（末题答对不循环，D12）。 */
function finishAnswer(app, question, correct, expectedDisplay) {
  app.feedback = { correct, expectedDisplay };
  app.progress = recordAnswer(app.progress, question.id, correct).progress;
  app.progressCache.set(app.group.groupKey, app.progress);
  if (!correct) {
    app.wrong = addWrong(app.wrong, question.id);
    app.wrongCounts = bumpWrongCount(app.wrongCounts, question.id);
    persistWrongbook(app);
  }
  persistProgress(app);
  if (correct) {
    // 末题答对不循环（D12）：停在本卡明示整组完成；否则排程自动前进
    if (app.index >= app.order.length - 1) app.completed = true;
    else scheduleAdvance(app);
  }
  render(app); // 答错也要重画：停在本题并显示正确答案（R15）
}

async function removeFromWrongbook(app) {
  cancelAdvance(app);
  const question = currentQuestion(app);
  if (!question) return;
  app.wrong = removeWrong(app.wrong, question.id);
  const counts = { ...app.wrongCounts }; // 移出时同步删错次，不留孤儿计数（R17）
  delete counts[question.id];
  app.wrongCounts = counts;
  await persistWrongbook(app);
  app.order = orderFor(app, app.group.mode, app.group.type);
  app.index = cursorIndex(app.order, app.progress ? app.progress.cursorQid : null);
  if (app.index >= app.order.length) app.index = Math.max(0, app.order.length - 1);
  resetQuestionState(app);
  render(app);
}

/** 清空已掌握：把已掌握的题从错题本移出（含错次），未掌握的留着。 */
async function clearMastered(app) {
  cancelAdvance(app);
  if (!app.store) return;
  const mastery = currentMastery(app);
  const counts = { ...app.wrongCounts };
  for (const qid of Array.from(app.wrong)) {
    if (mastery.get(qid) !== 'mastered') continue;
    app.wrong = removeWrong(app.wrong, qid);
    delete counts[qid];
  }
  app.wrongCounts = counts;
  await persistWrongbook(app);
  render(app);
}

function toggleUnmasteredOnly(app) {
  cancelAdvance(app);
  app.uiState = Object.assign({}, app.uiState, {
    practiceUnmasteredOnly: !uiFlag(app, 'practiceUnmasteredOnly'),
  });
  saveUiState(app);
  render(app);
}

function toggleShuffled(app) {
  cancelAdvance(app);
  const on = !uiFlag(app, 'practiceShuffled');
  const next = Object.assign({}, app.uiState, { practiceShuffled: on });
  // 种子只在首次开启时生成一次，此后固定（R20）
  if (on && typeof next.shuffleSeed !== 'number') next.shuffleSeed = Math.floor(Math.random() * 0x7fffffff);
  app.uiState = next;
  saveUiState(app);
  render(app);
}

function dismissInstallHint(app) {
  app.uiState = Object.assign({}, app.uiState, { installHintDismissed: true });
  saveUiState(app);
  render(app);
}

function openImport(app) {
  app.flow = createImportFlow({
    inflateRaw,
    inflateOk: app.inflate.ok,
    inflateReason: app.inflate.reason,
    onConfirm: (record) => bankImported(app, record),
    onCancel: () => {
      app.view = 'matrix';
      render(app);
    },
    onImportDone: (ms) => {
      app.importMs = ms;
      writeDiag(app);
    },
  });
  app.view = 'import';
  render(app);
}

async function bankImported(app, record) {
  app.importMs = app.flow && app.flow.state ? app.flow.state.importMs : app.importMs;
  await adoptBank(app, record, render);
}

function registerServiceWorker() {
  const location = globalThis.location;
  const protocol = location && location.protocol ? location.protocol : '';
  // 单文件形态从 file:// 打开：不注册 service worker（§2.12.1）
  if (protocol !== 'http:' && protocol !== 'https:') return;
  const nav = globalThis.navigator;
  if (!nav || !('serviceWorker' in nav)) return;
  try {
    nav.serviceWorker.register('sw.js').catch((err) => console.error('service worker 注册失败', err));
  } catch (err) {
    console.error('service worker 注册失败', err);
  }
}

function now() { return globalThis.performance ? globalThis.performance.now() : Date.now(); }

function errorMessage(err) { return err && err.message ? err.message : String(err); }

__MODULES__["src/ui/app.mjs"] = { mount: mount, MEMORIZE_REVEAL_MS: MEMORIZE_REVEAL_MS };
}());
window.EXAM_MEMO = { mount: __MODULES__["src/ui/app.mjs"].mount, MEMORIZE_REVEAL_MS: __MODULES__["src/ui/app.mjs"].MEMORIZE_REVEAL_MS };
}());
