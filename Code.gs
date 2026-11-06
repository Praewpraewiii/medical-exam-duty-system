/**
 * ระบบสุ่มข้อสอบแพทย์ (แบ่งตามชั้นยศ) + สมัครแพทย์ + รหัสผ่าน + ข้อช้อย/ข้อเขียน
 * Backend: Google Apps Script + Google Sheets
 *
 * ชีตที่ใช้ (สร้างอัตโนมัติด้วยฟังก์ชัน setup):
 *   Users     : ผู้ใช้ + ข้อมูลใบสมัคร + ยศ + รหัสผ่าน (hash)
 *   Exams     : ตั้งค่าข้อสอบแต่ละยศ (จำนวนข้อช้อย/ข้อเขียน, % ผ่าน, เวลา, ยศถัดไป)
 *   Questions : คลังข้อสอบ  Type = CHOICE (ช้อย) หรือ WRITTEN (เขียน), Points = คะแนนต่อข้อ
 *   Results   : ผลสอบรวม
 *   Answers   : คำตอบข้อเขียน — ระบบให้คะแนนอัตโนมัติจากช่อง Keywords ของข้อสอบ
 *               ผู้บริหารยืนยัน/แก้คะแนนในหน้าเว็บ (หรือกรอกช่อง Score ในชีต
 *               แล้วกดเมนู "ระบบสอบ → คำนวณผลข้อเขียนที่ตรวจแล้ว")
 *
 * ระดับสิทธิ์:
 *   SUPER (ผู้พัฒนา)  : ทำได้ทุกอย่าง + จัดการผู้ใช้ / สิทธิ์ / ชุดข้อสอบ ผ่านหน้าเว็บ
 *   ADMIN             : ยศที่ผู้พัฒนาให้สิทธิ์ไว้ในหน้า "จัดการสิทธิ์" (ชีต Permissions)
 *                       เพิ่ม/แก้ข้อสอบ และตรวจข้อเขียน
 *   ผู้ใช้ทั่วไป       : ทำข้อสอบตามยศ
 *   ทุกการกระทำของผู้บริหารและผู้พัฒนาจะถูกบันทึกในชีต AuditLog
 *
 * รหัสผ่าน:
 *   - PasswordHash ว่าง = ใช้รหัสใน Script Property INITIAL_PASSWORD และต้องเปลี่ยนตอนเข้าสู่ระบบ
 *   - รีเซ็ตรหัสผ่าน: ลบค่าในช่อง PasswordHash / บังคับเปลี่ยน: MustChange = TRUE
 */

const SH = { USERS: 'Users', EXAMS: 'Exams', QUESTIONS: 'Questions', RESULTS: 'Results', ANSWERS: 'Answers',
             PERMS: 'Permissions', AUDIT: 'AuditLog', ATTENDANCE: 'Attendance', PAYCFG: 'PayrollConfig', PAYRATES: 'PayrollRates' };
const APPLICANT_RANK = 'ผู้สมัคร';
const PENDING = 'รอตรวจข้อเขียน';
const REVIEW_AUTO = 'อัตโนมัติ (รอยืนยัน)';
const REVIEW_DONE = 'ยืนยันแล้ว';
const SESSION_SECONDS = 21600;
const MAX_FAILS = 5;
const LOCK_SECONDS = 300;
const MAX_ANSWER_CHARS = 3000;
const SUPER_RANKS = ['ผู้พัฒนา'];                   // Super Admin (กำหนดในโค้ด กันล็อกตัวเองออก)
const ADMIN_RANKS = [];   // ไม่มียศ ADMIN เริ่มต้น — DEV001 เป็นคนให้สิทธิ์เองในหน้าเว็บ
const DEFAULT_TIME_MIN = 60;    // เวลาสอบเมื่อไม่ได้ตั้งในชีต Exams (นาที)

const PW_RULE = /^(?=.*\d)(?=.*[a-z])(?=.*[A-Z])(?=.*[!-\/:-@\[-`{-~])[!-~]{6,10}$/;
const DISCORD_RULE = /^https?:\/\/(www\.|ptb\.|canary\.)?(discord(app)?\.com|discord\.gg)\/\S+$/i;

/* ---------- หน้าเว็บ + เมนูในชีต ---------- */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ระบบสอบแพทย์')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('ระบบสอบ')
    .addItem('คำนวณผลข้อเขียนที่ตรวจแล้ว', 'gradeWritten')
    .addToUi();
}

/* ---------- เข้าสู่ระบบ / รหัสผ่าน ---------- */

function login(code, password) {
  const c = norm_(code);
  password = String(password || '');
  if (!c || !password) return { ok: false, message: 'กรุณากรอกรหัสพนักงานและรหัสผ่าน' };

  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('fail_' + c) || 0);
  if (fails >= MAX_FAILS) {
    return { ok: false, message: 'กรอกรหัสผ่านผิดเกิน ' + MAX_FAILS + ' ครั้ง กรุณารอ 5 นาทีแล้วลองใหม่' };
  }

  const user = findUser_(c);
  if (!user || !checkPassword_(user, password)) {
    cache.put('fail_' + c, String(fails + 1), LOCK_SECONDS);
    return { ok: false, message: 'รหัสพนักงานหรือรหัสผ่านไม่ถูกต้อง' };
  }
  cache.remove('fail_' + c);

  if (mustChange_(user)) {
    return { ok: true, mustChange: true, token: issueToken_(c, 'change'), name: String(user.Name) };
  }
  return Object.assign({ ok: true, token: issueToken_(c, 'auth') }, homeData_(user));
}

function initialPassword_() {
  return PropertiesService.getScriptProperties().getProperty('INITIAL_PASSWORD') || '';
}

function changePassword(token, newPassword, currentPassword) {
  const tok = readToken_(token);
  if (!tok) return { ok: false, expired: true, message: 'เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const users = table_(SH.USERS);
    const user = users.rows.find(r => norm_(r.EmployeeID) === tok.code);
    if (!user) return { ok: false, expired: true, message: 'ไม่พบผู้ใช้' };

    if (tok.type === 'auth' && !checkPassword_(user, currentPassword)) {
      return { ok: false, message: 'รหัสผ่านปัจจุบันไม่ถูกต้อง' };
    }
    const err = validatePassword_(newPassword);
    if (err) return { ok: false, message: err };
    if (checkPassword_(user, newPassword)) return { ok: false, message: 'รหัสผ่านใหม่ต้องไม่ซ้ำกับรหัสผ่านเดิม' };

    const hash = hash_(user.EmployeeID, newPassword);
    setCell_(users, user, 'PasswordHash', hash);
    setCell_(users, user, 'MustChange', false);
    user.PasswordHash = hash;
    user.MustChange = false;

    CacheService.getScriptCache().remove('tok_' + token);
    return Object.assign({ ok: true, token: issueToken_(tok.code, 'auth') }, homeData_(user));
  } finally {
    lock.releaseLock();
  }
}

function getHome(token) {
  const tok = readToken_(token, ['auth']);
  if (!tok) return { ok: false, expired: true, message: 'เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่' };
  const user = findUser_(tok.code);
  if (!user) return { ok: false, expired: true, message: 'ไม่พบผู้ใช้' };
  return Object.assign({ ok: true }, homeData_(user));
}

/* ---------- สมัครแพทย์ ---------- */

/**
 * f = { nameIC, nameOC, ageIC, discord, reason }
 * สมัครเสร็จจะได้ token ชนิด reg เพื่อเข้าทำข้อสอบเข้าแพทย์ได้ทันที
 */
function register(f) {
  f = f || {};
  const nameIC = String(f.nameIC || '').trim();
  const nameOC = String(f.nameOC || '').trim();
  const ageIC = Number(String(f.ageIC || '').trim());
  const discord = String(f.discord || '').trim();
  const reason = String(f.reason || '').trim();

  if (nameIC.length < 2 || nameIC.length > 60) return { ok: false, message: 'กรุณากรอกชื่อ IC (2–60 ตัวอักษร)' };
  if (nameOC.length < 2 || nameOC.length > 60) return { ok: false, message: 'กรุณากรอกชื่อ OC (2–60 ตัวอักษร)' };
  if (!Number.isInteger(ageIC) || ageIC < 1 || ageIC > 150) return { ok: false, message: 'กรุณากรอกอายุ IC เป็นตัวเลข' };
  if (!DISCORD_RULE.test(discord) || discord.length > 200) {
    return { ok: false, message: 'ลิงก์ Discord ไม่ถูกต้อง (เช่น https://discord.com/users/123456789)' };
  }
  if (reason.length < 10 || reason.length > 1000) return { ok: false, message: 'กรุณาบอกเหตุผลที่อยากสมัครแพทย์ (10–1000 ตัวอักษร)' };
  const initialPassword = initialPassword_();
  if (!initialPassword) return { ok:false, message:'ระบบยังไม่ได้ตั้งรหัสเริ่มต้น กรุณาติดต่อผู้ดูแลระบบ' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const t = table_(SH.USERS);
    const key = discordKey_(discord);
    if (t.rows.some(r => r.Discord && discordKey_(r.Discord) === key)) {
      return { ok: false, message: 'ลิงก์ Discord นี้เคยสมัครแล้ว กรุณาเข้าสู่ระบบด้วยรหัสเดิม' };
    }

    const used = new Set(t.rows.map(r => norm_(r.EmployeeID)));
    let code;
    do { code = 'APP' + Math.floor(100000 + Math.random() * 900000); } while (used.has(code));

    appendObj_(t, {
      EmployeeID: code, Name: safe_(nameIC), NameOC: safe_(nameOC), AgeIC: ageIC,
      Discord: safe_(discord), Reason: safe_(reason), Rank: APPLICANT_RANK,
      CreatedAt: new Date(), PasswordHash: '', MustChange: true
    });
    return { ok: true, code: code, defaultPassword: initialPassword, token: issueToken_(code, 'reg') };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- ข้อสอบ ---------- */

function startExam(token) {
  const tok = readToken_(token, ['auth', 'reg']);
  if (!tok) return { ok: false, expired: true, message: 'เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่' };

  const user = findUser_(tok.code);
  if (!user) return { ok: false, message: 'ไม่พบผู้ใช้' };
  if (tok.type === 'auth' && mustChange_(user)) {
    return { ok: false, expired: true, message: 'กรุณาเปลี่ยนรหัสผ่านก่อนเข้าสอบ' };
  }

  const exam = findExam_(user.Rank);
  if (!exam) return { ok: false, message: 'ยศ "' + user.Rank + '" ยังไม่มีข้อสอบ' };

  const pool = questionsFor_(user.Rank);
  const choicePool = pool.filter(q => qType_(q) === 'CHOICE');
  const writtenPool = pool.filter(q => qType_(q) === 'WRITTEN');
  const nC = Math.min(count_(exam.NumQuestions, choicePool.length), choicePool.length);
  const nW = Math.min(count_(exam.NumWritten, writtenPool.length), writtenPool.length);
  if (!nC && !nW) return { ok: false, message: 'ยังไม่มีข้อสอบในคลังสำหรับยศนี้' };

  // สุ่มข้อช้อยก่อน แล้วตามด้วยข้อเขียน
  const picked = shuffle_(choicePool).slice(0, nC).concat(shuffle_(writtenPool).slice(0, nW));

  // สลับลำดับตัวเลือกของแต่ละข้อต่อผู้สอบ (เก็บตัวแปลงไว้ที่เซิร์ฟเวอร์เพื่อใช้ตรวจ)
  const LABELS = ['A', 'B', 'C', 'D'];
  const imgCache = {};
  const items = [], questions = [];
  picked.forEach(q => {
    const id = String(q.ID), type = qType_(q), pts = points_(q);
    const item = { id, type, pts };
    const out = { id, type, points: pts, text: String(q.Question), choices: [], image: imageData_(q.Image, imgCache) };
    if (type === 'CHOICE') {
      const keys = shuffle_(LABELS.filter(k => str_(q[k]).trim() !== ''));
      item.map = {};
      out.choices = keys.map((k, i) => { item.map[LABELS[i]] = k; return { key: LABELS[i], text: String(q[k]) }; });
    }
    items.push(item); questions.push(out);
  });

  const examToken = Utilities.getUuid();
  const session = {
    code: norm_(user.EmployeeID),
    rank: String(user.Rank),
    examName: String(exam.ExamName),
    passPercent: Number(exam.PassPercent) || 0,
    timeMin: timeOf_(exam),
    showScore: showScore_(exam),
    items: items,
    start: Date.now()
  };
  CacheService.getScriptCache().put('exam_' + examToken, JSON.stringify(session), SESSION_SECONDS);

  return { ok: true, examToken: examToken, exam: session.examName, timeMin: session.timeMin, questions: questions };
}

function submitExam(examToken, answers) {
  const cache = CacheService.getScriptCache();
  const raw = cache.get('exam_' + examToken);
  if (!raw) return { ok: false, message: 'เซสชันสอบหมดอายุ หรือส่งคำตอบไปแล้ว' };
  cache.remove('exam_' + examToken);

  const s = JSON.parse(raw);
  answers = answers || {};

  const qmap = {};
  table_(SH.QUESTIONS).rows.forEach(q => { qmap[String(q.ID)] = q; });

  let cScore = 0, cMax = 0, wMax = 0, wAuto = 0;
  const written = [];
  s.items.forEach(it => {
    const q = qmap[it.id];
    if (it.type === 'CHOICE') {
      cMax += it.pts;
      const given = norm_(answers[it.id]);
      const original = it.map ? (it.map[given] || '') : given;   // แปลงตัวเลือกที่สลับแล้วกลับเป็นตัวเลือกเดิม
      if (q && original && norm_(q.Answer) === original) cScore += it.pts;
    } else {
      wMax += it.pts;
      const answer = String(answers[it.id] == null ? '' : answers[it.id]).trim().slice(0, MAX_ANSWER_CHARS);
      const g = autoGrade_(answer, q ? q.Keywords : '', it.pts);
      wAuto += g.score;
      written.push({ id: it.id, pts: it.pts, question: q ? String(q.Question) : '', answer, g });
    }
  });

  const total = cMax + wMax;
  const late = s.timeMin > 0 && (Date.now() - s.start) > (s.timeMin * 60 + 60) * 1000;
  const hasWritten = wMax > 0;
  const allKeyed = written.every(w => w.g.has);
  // ข้อเขียนที่ตั้งคีย์เวิร์ดครบ → คิดคะแนนอัตโนมัติทันที (รอผู้บริหารยืนยัน)
  // มีข้อที่ไม่มีคีย์เวิร์ด → รอตรวจ (แสดงคะแนนเบื้องต้น)
  const pending = hasWritten && !late && !allKeyed;
  const score = cScore + wAuto;
  const percent = total ? Math.round(score * 100 / total) : 0;
  const passed = !late && !pending && meetsPass_(score, total, s.passPercent);
  const review = hasWritten && !late ? REVIEW_AUTO : '';
  const resultId = 'R' + Date.now().toString(36).toUpperCase() + Math.floor(Math.random() * 1000);

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  let promotedTo = '';
  try {
    const users = table_(SH.USERS);
    const user = users.rows.find(r => norm_(r.EmployeeID) === s.code);
    const name = user ? user.Name : '';

    // ผลที่ประกาศได้ทันที (ไม่มีข้อเขียน หรือส่งเกินเวลา) → เลื่อนยศ + ส่ง Discord ตอนนี้เลย
    const final = !pending && review === '';
    if (final && passed && user && !hasWritten) promotedTo = promote_(users, user, s.rank);
    const discord = final ? notifyDiscord_({ EmployeeID: s.code, Name: name, Rank: s.rank, ExamName: s.examName },
      { score, total, percent, passed, passPercent: s.passPercent, promotedTo }, { user }) : '';

    appendObj_(table_(SH.RESULTS), {
      Timestamp: new Date(), ResultID: resultId, EmployeeID: s.code, Name: name,
      Rank: s.rank, ExamName: s.examName, DiscordStatus: discord,
      ChoiceScore: cScore, ChoiceMax: cMax, WrittenScore: hasWritten ? wAuto : '', WrittenMax: wMax,
      Score: pending ? '' : score, Total: total, Percent: pending ? '' : percent,
      Result: pending ? PENDING : (passed ? 'ผ่าน' : 'ไม่ผ่าน'),
      Review: review, Note: late ? 'ส่งเกินเวลา' : (promotedTo ? 'เลื่อนยศเป็น ' + promotedTo : ''), UpdatedAt: new Date()
    });

    if (written.length) {
      const ans = table_(SH.ANSWERS);
      written.forEach(w => appendObj_(ans, {
        Timestamp: new Date(), ResultID: resultId, EmployeeID: s.code, Name: name,
        QuestionID: w.id, Question: safe_(w.question), Answer: safe_(w.answer),
        MaxPoints: w.pts, AutoScore: w.g.has ? w.g.score : '',
        Matched: w.g.has ? w.g.matched.length + '/' + w.g.total + (w.g.matched.length ? ': ' + w.g.matched.join(', ') : '') : 'ไม่มีคีย์เวิร์ด',
        Score: w.g.has ? w.g.score : ''
      }));
    }

  } finally {
    lock.releaseLock();
  }

  // ชุดข้อสอบที่ไม่เปิดให้เห็นคะแนน: ส่งกลับแค่ว่าส่งสำเร็จ (ไม่ส่งคะแนนไปที่หน้าเว็บ)
  if (!s.showScore) return { ok: true, hidden: true };

  return {
    ok: true, pending, late, passed, promotedTo, auto: review === REVIEW_AUTO && !pending,
    choiceScore: cScore, choiceMax: cMax, writtenMax: wMax, writtenAuto: wAuto,
    score, total, percent, passPercent: s.passPercent,
    viewToken: issueResultToken_(s.code, resultId)
  };
}

/** ให้คะแนนข้อเขียนจากคีย์เวิร์ด — คั่นคีย์เวิร์ดด้วย , หรือขึ้นบรรทัดใหม่, คำพ้องใช้ / เช่น CPR/ปั๊มหัวใจ */
function autoGrade_(answer, keywords, pts) {
  const groups = String(keywords == null ? '' : keywords).split(/[,\n]/).map(x => x.trim()).filter(Boolean);
  if (!groups.length) return { has: false, score: 0, matched: [], total: 0 };
  const squash = x => String(x).toLowerCase().replace(/\s+/g, '');
  const a = squash(answer);
  const matched = groups.filter(g => g.split('/').map(squash).filter(Boolean).some(k => a.indexOf(k) !== -1));
  const score = Math.round(pts * matched.length / groups.length * 2) / 2;
  return { has: true, score, matched, total: groups.length };
}

/* ---------- แจ้งผลบนหน้าเว็บ ---------- */

function issueResultToken_(code, resultId) {
  const t = Utilities.getUuid();
  CacheService.getScriptCache().put('res_' + t, JSON.stringify({ code: norm_(code), resultId }), SESSION_SECONDS);
  return t;
}

/** ผ่านเมื่อคะแนนถึงเกณฑ์จริง (ไม่ปัดเศษ) เช่น 40 ข้อ เกณฑ์ 70% ต้องได้ ≥ 28 */
function meetsPass_(score, total, passPercent) {
  return total > 0 && score * 100 >= passPercent * total - 1e-9;
}

function timeOf_(exam) { const t = Number(exam && exam.TimeLimitMin); return t > 0 ? t : DEFAULT_TIME_MIN; }
function showScore_(exam) { return !!exam && truthy_(exam.ShowScore); }

let EXAM_CACHE_ = null; // อ่านชีต Exams ครั้งเดียวต่อการเรียก
function examCfg_(rank) {
  if (!EXAM_CACHE_) {
    EXAM_CACHE_ = {};
    table_(SH.EXAMS).rows.forEach(e => { EXAM_CACHE_[norm_(e.Rank)] = { pass: Number(e.PassPercent) || 0, show: showScore_(e) }; });
  }
  return EXAM_CACHE_[norm_(rank)] || { pass: 0, show: false };
}
function passPercentOf_(rank) { return examCfg_(rank).pass; }

function publicResult_(r) {
  const pending = String(r.Result).trim() === PENDING;
  return {
    resultId: str_(r.ResultID), exam: str_(r.ExamName), date: fmtDate_(r.Timestamp),
    result: str_(r.Result), review: str_(r.Review), pending,
    choiceScore: Number(r.ChoiceScore) || 0, choiceMax: Number(r.ChoiceMax) || 0,
    writtenScore: str_(r.WrittenScore) === '' ? null : Number(r.WrittenScore), writtenMax: Number(r.WrittenMax) || 0,
    score: str_(r.Score) === '' ? null : Number(r.Score), total: Number(r.Total) || 0,
    percent: str_(r.Percent) === '' ? null : Number(r.Percent),
    passPercent: passPercentOf_(r.Rank),
    note: str_(r.Note), updatedAt: fmtDate_(r.UpdatedAt)
  };
}

/** หน้าผลสอบเรียกซ้ำเป็นระยะ เพื่อแสดงผลทันทีเมื่อผู้บริหารยืนยัน/แก้คะแนน */
function checkResult(viewToken) {
  const raw = CacheService.getScriptCache().get('res_' + viewToken);
  if (!raw) return { ok: false, message: 'หมดเวลาติดตามผล — เข้าสู่ระบบเพื่อดูผลสอบ' };
  const v = JSON.parse(raw);
  const res = table_(SH.RESULTS);
  const r = res.rows.find(x => str_(x.ResultID) === v.resultId);
  if (!r) return { ok: false, message: 'ไม่พบผลสอบ' };
  if (!examCfg_(r.Rank).show) return { ok: false, message: 'ชุดข้อสอบนี้ไม่เปิดให้ดูคะแนน' };
  if (str_(r.Review) === REVIEW_DONE) setCell_(res, r, 'SeenAt', new Date());
  return Object.assign({ ok: true }, publicResult_(r));
}

/** ผลสอบที่ถูกยืนยัน/แก้คะแนนแล้วแต่ผู้สอบยังไม่เห็น */
function noticesFor_(code) {
  const c = norm_(code);
  return table_(SH.RESULTS).rows.filter(r => {
    if (norm_(r.EmployeeID) !== c || str_(r.Review) !== REVIEW_DONE || !examCfg_(r.Rank).show) return false;
    const up = r.UpdatedAt instanceof Date ? r.UpdatedAt.getTime() : 0;
    const seen = r.SeenAt instanceof Date ? r.SeenAt.getTime() : 0;
    return up > seen;
  }).map(publicResult_);
}

function getNotices(token) {
  const tok = readToken_(token, ['auth']);
  if (!tok) return { ok: false, expired: true, message: 'เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่' };
  return { ok: true, notices: noticesFor_(tok.code) };
}

function ackNotices(token, resultIds) {
  const tok = readToken_(token, ['auth']);
  if (!tok) return { ok: false, expired: true };
  const ids = new Set((resultIds || []).map(String));
  const res = table_(SH.RESULTS);
  res.rows.filter(r => norm_(r.EmployeeID) === tok.code && ids.has(str_(r.ResultID)))
    .forEach(r => setCell_(res, r, 'SeenAt', new Date()));
  return { ok: true };
}

/**
 * แอดมิน: กรอกคะแนนข้อเขียนในชีต Answers (ช่อง Score) ให้ครบทุกข้อของผลสอบนั้น
 * แล้วกดเมนู ระบบสอบ → คำนวณผลข้อเขียนที่ตรวจแล้ว
 */
function gradeWritten() {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let done = 0, waiting = 0;
  try {
    const res = table_(SH.RESULTS), ans = table_(SH.ANSWERS), users = table_(SH.USERS);
    res.rows.filter(r => String(r.Result).trim() === PENDING).forEach(r => {
      const list = ans.rows.filter(a => String(a.ResultID) === String(r.ResultID));
      const unscored = list.some(a => String(a.Score).trim() === '' || isNaN(Number(a.Score)));
      if (!list.length || unscored) { waiting++; return; }
      const w = list.reduce((sum, a) => sum + Math.max(0, Math.min(Number(a.Score), Number(a.MaxPoints) || 0)), 0);
      finalize_(res, users, r, w, 'ตรวจจากชีต');
      done++;
    });
  } finally {
    lock.releaseLock();
  }
  SpreadsheetApp.getUi().alert('คำนวณผลแล้ว ' + done + ' รายการ' +
    (waiting ? '\nยังรอให้คะแนนข้อเขียนไม่ครบ ' + waiting + ' รายการ' : ''));
}

/** รวมคะแนนข้อเขียนเข้ากับข้อช้อย ตัดสินผล และ stamp ว่ายืนยันแล้ว (ผู้สอบจะเห็นผลทันที) */
function finalize_(res, users, r, writtenScore, grader) {
  const total = (Number(r.ChoiceMax) || 0) + (Number(r.WrittenMax) || 0);
  const score = (Number(r.ChoiceScore) || 0) + writtenScore;
  const percent = total ? Math.round(score * 100 / total) : 0;
  const exam = findExam_(r.Rank);
  const passPercent = exam ? Number(exam.PassPercent) || 0 : 0;
  const passed = meetsPass_(score, total, passPercent);
  const wasDone = str_(r.Review) === REVIEW_DONE;

  setCell_(res, r, 'WrittenScore', writtenScore);
  setCell_(res, r, 'Score', score);
  setCell_(res, r, 'Percent', percent);
  setCell_(res, r, 'Result', passed ? 'ผ่าน' : 'ไม่ผ่าน');
  setCell_(res, r, 'Review', REVIEW_DONE);
  setCell_(res, r, 'GradedBy', grader || '');
  setCell_(res, r, 'UpdatedAt', new Date());

  let promotedTo = '';
  if (passed) {
    const user = users.rows.find(u => norm_(u.EmployeeID) === norm_(r.EmployeeID));
    promotedTo = user ? promote_(users, user, r.Rank) : '';
    if (promotedTo) setCell_(res, r, 'Note', 'เลื่อนยศเป็น ' + promotedTo);
  }
  const out = { score, total, percent, passed, passPercent, promotedTo, regraded: wasDone };
  const user = users.rows.find(u => norm_(u.EmployeeID) === norm_(r.EmployeeID));
  const discord = notifyDiscord_(r, out, { user, grader, regraded: wasDone });
  if (discord) setCell_(res, r, 'DiscordStatus', discord + ' · ' + fmtDate_(new Date()));
  out.discord = discord;
  return out;
}

/* ---------- ผู้บริหาร: จัดการข้อสอบ / ตรวจข้อเขียน ---------- */

function rankKey_(v) { return String(v == null ? '' : v).replace(/[\s.]/g, ''); }

function adminRanks_() {
  if (!ss_().getSheetByName(SH.PERMS)) return ADMIN_RANKS.map(rankKey_);
  return table_(SH.PERMS).rows.filter(r => norm_(r.Level) === 'ADMIN').map(r => rankKey_(r.Rank));
}

/** 2 = SUPER, 1 = ADMIN, 0 = ผู้ใช้ทั่วไป */
function levelOf_(rank) {
  const k = rankKey_(rank);
  if (!k) return 0;
  if (SUPER_RANKS.map(rankKey_).indexOf(k) !== -1) return 2;
  return adminRanks_().indexOf(k) !== -1 ? 1 : 0;
}
function isAdmin_(rank) { return levelOf_(rank) >= 1; }
function isSuper_(rank) { return levelOf_(rank) >= 2; }

function whoOf_(user) { return str_(user.Name) + ' (' + str_(user.EmployeeID) + ')'; }

/** บันทึกการกระทำลงชีต AuditLog */
function audit_(by, action, target, detail) {
  try {
    const ss = ss_();
    if (!ss.getSheetByName(SH.AUDIT)) {
      const sh = ss.insertSheet(SH.AUDIT);
      sh.getRange(1, 1, 1, 5).setValues([['Timestamp', 'By', 'Action', 'Target', 'Detail']])
        .setFontWeight('bold').setBackground('#d9f2ef');
      sh.setFrozenRows(1);
    }
    appendObj_(table_(SH.AUDIT), { Timestamp: new Date(), By: by, Action: action, Target: safe_(target), Detail: safe_(detail || '') });
  } catch (e) { /* log ไม่ได้ไม่ควรทำให้งานหลักล้ม */ }
}

/** ตรวจว่า token นี้เป็นผู้บริหารจริง (เช็กทุกครั้งที่เรียก) */
function requireAdmin_(token) {
  const tok = readToken_(token, ['auth']);
  if (!tok) return { err: { ok: false, expired: true, message: 'เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่' } };
  const user = findUser_(tok.code);
  if (!user || mustChange_(user) || !isAdmin_(user.Rank)) {
    return { err: { ok: false, message: 'คุณไม่มีสิทธิ์ใช้งานส่วนนี้' } };
  }
  return { user };
}

function adminGetQuestions(token) {
  const adm = requireAdmin_(token);
  if (adm.err) return adm.err;
  return {
    ok: true,
    sets: table_(SH.EXAMS).rows.map(e => ({ rank: str_(e.Rank), name: str_(e.ExamName) })),
    questions: table_(SH.QUESTIONS).rows.map(q => ({
      id: str_(q.ID), rank: str_(q.ExamRank), type: qType_(q), points: points_(q),
      question: str_(q.Question), A: str_(q.A), B: str_(q.B), C: str_(q.C), D: str_(q.D),
      answer: norm_(q.Answer), keywords: str_(q.Keywords), image: str_(q.Image),
      createdBy: str_(q.CreatedBy), createdAt: fmtDate_(q.CreatedAt),
      updatedBy: str_(q.UpdatedBy), updatedAt: fmtDate_(q.UpdatedAt)
    }))
  };
}

/** เพิ่มข้อสอบใหม่ (ไม่มี id) หรือแก้ไขข้อเดิม (มี id) */
function adminSaveQuestion(token, q) {
  const adm = requireAdmin_(token);
  if (adm.err) return adm.err;
  q = q || {};

  const v = buildQuestion_(q);
  if (v.err) return { ok: false, message: v.err };
  const data = v.data, rank = data.ExamRank, type = data.Type;

  // Stamp ผู้บันทึกและเวลาลงหลังบ้าน
  const who = adm.user.Name + ' (' + adm.user.EmployeeID + ')';
  const now = new Date();

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const t = table_(SH.QUESTIONS);
    let id = String(q.id || '').trim();
    if (id) {
      const row = t.rows.find(r => str_(r.ID) === id);
      if (!row) return { ok: false, message: 'ไม่พบข้อสอบ ' + id + ' (อาจถูกลบไปแล้ว)' };
      Object.assign(data, { UpdatedBy: who, UpdatedAt: now });
      Object.keys(data).forEach(k => setCell_(t, row, k, data[k]));
    } else {
      Object.assign(data, { CreatedBy: who, CreatedAt: now });
      const used = new Set(t.rows.map(r => str_(r.ID)));
      let n = t.rows.reduce((m, r) => { const x = /^Q(\d+)$/.exec(str_(r.ID)); return x ? Math.max(m, Number(x[1])) : m; }, 0);
      do { n++; id = 'Q' + String(n).padStart(3, '0'); } while (used.has(id));
      appendObj_(t, Object.assign({ ID: id }, data));
    }
    audit_(who, q.id ? 'แก้ไขข้อสอบ' : 'เพิ่มข้อสอบ', id, rank + ' · ' + type);
    return { ok: true, id: id, stamp: who + ' · ' + fmtDate_(now) };
  } finally {
    lock.releaseLock();
  }
}

function adminDeleteQuestion(token, id) {
  const adm = requireAdmin_(token);
  if (adm.err) return adm.err;
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const t = table_(SH.QUESTIONS);
    const row = t.rows.find(r => str_(r.ID) === String(id));
    if (!row) return { ok: false, message: 'ไม่พบข้อสอบนี้' };
    t.sh.deleteRow(row._row);
    audit_(whoOf_(adm.user), 'ลบข้อสอบ', String(id), str_(row.Question).slice(0, 100));
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/** สร้างข้อมูลผลสอบ 1 รายการ พร้อมคำตอบข้อเขียน สำหรับหน้าผู้บริหาร */
function resultDetail_(r, ansRows, users, qmap) {
  const u = users.find(x => norm_(x.EmployeeID) === norm_(r.EmployeeID));
  return Object.assign(publicResult_(r), {
    employeeId: str_(r.EmployeeID), name: str_(r.Name), rank: str_(r.Rank), gradedBy: str_(r.GradedBy),
    applicant: u ? { nameOC: str_(u.NameOC), age: str_(u.AgeIC), discord: str_(u.Discord), reason: str_(u.Reason) } : null,
    answers: ansRows.filter(a => String(a.ResultID) === String(r.ResultID)).map(a => ({
      qid: str_(a.QuestionID), question: str_(a.Question), answer: str_(a.Answer),
      max: Number(a.MaxPoints) || 0, score: str_(a.Score), auto: str_(a.AutoScore), matched: str_(a.Matched),
      keywords: qmap[str_(a.QuestionID)] ? str_(qmap[str_(a.QuestionID)].Keywords) : '',
      image: qmap[str_(a.QuestionID)] ? str_(qmap[str_(a.QuestionID)].Image) : ''
    }))
  });
}

function needsReview_(r) {
  return (Number(r.WrittenMax) || 0) > 0 &&
    (String(r.Result).trim() === PENDING || str_(r.Review) === REVIEW_AUTO);
}

/** รายการที่ต้องตรวจ/ยืนยันข้อเขียน (เก่าสุดก่อน) */
function adminGetPending(token) {
  const adm = requireAdmin_(token);
  if (adm.err) return adm.err;
  const ans = table_(SH.ANSWERS).rows, users = table_(SH.USERS).rows;
  const qmap = {}; table_(SH.QUESTIONS).rows.forEach(q => { qmap[str_(q.ID)] = q; });
  const items = table_(SH.RESULTS).rows.filter(needsReview_).slice(0, 100)
    .map(r => resultDetail_(r, ans, users, qmap));
  return { ok: true, items: items };
}

/** คะแนนย้อนหลังทั้งหมด (ใหม่สุดก่อน) */
function adminGetHistory(token) {
  const adm = requireAdmin_(token);
  if (adm.err) return adm.err;
  const rows = table_(SH.RESULTS).rows.slice().reverse().slice(0, 1000);
  return {
    ok: true,
    sets: table_(SH.EXAMS).rows.map(e => str_(e.ExamName)),
    items: rows.map(r => Object.assign(publicResult_(r), {
      employeeId: str_(r.EmployeeID), name: str_(r.Name), rank: str_(r.Rank), gradedBy: str_(r.GradedBy)
    }))
  };
}

function adminGetResult(token, resultId) {
  const adm = requireAdmin_(token);
  if (adm.err) return adm.err;
  const r = table_(SH.RESULTS).rows.find(x => str_(x.ResultID) === String(resultId));
  if (!r) return { ok: false, message: 'ไม่พบผลสอบ' };
  const qmap = {}; table_(SH.QUESTIONS).rows.forEach(q => { qmap[str_(q.ID)] = q; });
  return Object.assign({ ok: true }, resultDetail_(r, table_(SH.ANSWERS).rows, table_(SH.USERS).rows, qmap));
}

/**
 * บันทึกคะแนนข้อเขียนของผลสอบ 1 รายการ แล้วตัดสินผล — scores = { QuestionID: คะแนน }
 * regrade = true เมื่อแก้คะแนนผลที่ยืนยันไปแล้ว (จากหน้าคะแนนย้อนหลัง)
 */
function adminGrade(token, resultId, scores, regrade) {
  const adm = requireAdmin_(token);
  if (adm.err) return adm.err;
  scores = scores || {};

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const res = table_(SH.RESULTS), ans = table_(SH.ANSWERS), users = table_(SH.USERS);
    const r = res.rows.find(x => str_(x.ResultID) === String(resultId));
    if (!r) return { ok: false, message: 'ไม่พบผลสอบนี้' };
    if (!regrade && !needsReview_(r)) return { ok: false, message: 'ผลสอบนี้ถูกยืนยันไปแล้ว', done: true };

    const list = ans.rows.filter(a => String(a.ResultID) === String(resultId));
    if (!list.length) return { ok: false, message: 'ผลสอบนี้ไม่มีข้อเขียน' };
    const vals = [];
    for (const a of list) {
      const raw = scores[str_(a.QuestionID)];
      const v = Number(raw);
      const max = Number(a.MaxPoints) || 0;
      if (raw === '' || raw == null || isNaN(v) || v < 0 || v > max) {
        return { ok: false, message: 'คะแนนข้อ ' + a.QuestionID + ' ต้องอยู่ระหว่าง 0–' + max };
      }
      vals.push(v);
    }
    const before = str_(r.Score) + '/' + str_(r.Total) + ' ' + str_(r.Result);
    let w = 0;
    list.forEach((a, i) => { setCell_(ans, a, 'Score', vals[i]); w += vals[i]; });

    const out = finalize_(res, users, r, w, whoOf_(adm.user));
    audit_(whoOf_(adm.user), out.regraded ? 'แก้คะแนนข้อเขียน' : 'ยืนยันคะแนนข้อเขียน', str_(r.EmployeeID) + ' / ' + resultId,
      (out.regraded ? before + ' → ' : '') + out.score + '/' + out.total + ' ' + (out.passed ? 'ผ่าน' : 'ไม่ผ่าน'));
    return Object.assign({ ok: true }, out);
  } finally {
    lock.releaseLock();
  }
}






/* ---------- แจ้งผลเข้า Discord (Webhook) ---------- */

const WEBHOOK_RULE = /^https:\/\/(?:ptb\.|canary\.)?discord(?:app)?\.com\/api\/webhooks\/\d+\/[\w-]+$/;

function discordWebhook_() { return PropertiesService.getScriptProperties().getProperty('DISCORD_WEBHOOK') || ''; }
function sendDiscord_(exam) { return !!exam && truthy_(exam.SendDiscord); }

/** ดึง Discord user ID จากลิงก์โปรไฟล์ หรือเลข ID ตรงๆ */
function discordUserId_(v) {
  v = str_(v).trim();
  const m = /discord(?:app)?\.com\/users\/(\d{15,22})/i.exec(v) || /^<?@?!?(\d{15,22})>?$/.exec(v);
  return m ? m[1] : '';
}

function postDiscord_(payload) {
  const url = discordWebhook_();
  if (!url) return 'ยังไม่ได้ตั้ง Webhook';
  try {
    const res = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    return code >= 200 && code < 300 ? 'ส่งแล้ว' : 'ส่งไม่สำเร็จ (' + code + ')';
  } catch (e) {
    return 'ส่งไม่สำเร็จ (' + String(e.message || e).slice(0, 80) + ')';
  }
}

/**
 * ส่งผลสอบที่ประกาศแล้วเข้าห้อง Discord — เฉพาะชุดข้อสอบที่เปิด "ส่งผลเข้า Discord"
 * r = ข้อมูลผลสอบ { EmployeeID, Name, Rank, ExamName }, out = { score, total, percent, passed, passPercent, promotedTo }
 * คืนสถานะไว้บันทึกในช่อง DiscordStatus ('' = ชุดนี้ไม่ได้เปิดส่ง)
 */
function notifyDiscord_(r, out, opts) {
  opts = opts || {};
  const exam = findExam_(r.Rank);
  if (!sendDiscord_(exam)) return '';
  const user = opts.user || findUser_(r.EmployeeID);
  const uid = user ? discordUserId_(user.Discord) : '';
  const fields = [
    { name: 'ผู้สอบ', value: (str_(r.Name) || '-') + ' (' + str_(r.EmployeeID) + ')', inline: true },
    { name: 'ยศตอนสอบ', value: str_(r.Rank) || '-', inline: true },
    { name: 'คะแนน', value: out.score + '/' + out.total + ' (' + out.percent + '%)', inline: true },
    { name: 'เกณฑ์ผ่าน', value: out.passPercent + '%', inline: true },
    { name: 'ผล', value: out.passed ? '✅ ผ่าน' : '❌ ไม่ผ่าน', inline: true }
  ];
  if (out.promotedTo) fields.push({ name: 'เลื่อนยศ', value: '🎉 ' + out.promotedTo, inline: true });
  return postDiscord_({
    username: 'ระบบสอบแพทย์',
    content: uid ? '<@' + uid + '>' : '',
    allowed_mentions: { users: uid ? [uid] : [] },
    embeds: [{
      title: (opts.regraded ? '✏️ แก้ไขผลสอบ · ' : '📋 ประกาศผลสอบ · ') + str_(r.ExamName),
      color: out.passed ? 0x1d8a4a : 0xc0392b,
      fields: fields,
      footer: { text: opts.grader ? 'ยืนยันผลโดย ' + opts.grader : 'ตรวจอัตโนมัติ' },
      timestamp: new Date().toISOString()
    }]
  });
}

/* ---------- รูปภาพประกอบข้อสอบ (เก็บใน Google Drive ของเจ้าของระบบ ไม่แชร์สาธารณะ) ---------- */

const IMAGE_FOLDER_NAME = 'ระบบสอบแพทย์ - รูปข้อสอบ';
const ATTENDANCE_FOLDER_NAME = 'ระบบสอบแพทย์ - หลักฐานลงเวลา';
const IMAGE_ID_RULE = /^[A-Za-z0-9_-]{20,}$/;

function ensureImageFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('IMG_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { /* ถูกลบไปแล้ว → สร้างใหม่ */ } }
  const folder = DriveApp.createFolder(IMAGE_FOLDER_NAME);
  props.setProperty('IMG_FOLDER_ID', folder.getId());
  return folder;
}

function ensureAttendanceFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('ATTENDANCE_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { /* folder removed; recreate */ } }
  const folder = DriveApp.createFolder(ATTENDANCE_FOLDER_NAME);
  props.setProperty('ATTENDANCE_FOLDER_ID', folder.getId());
  return folder;
}

/** รูปเป็น data URI เพื่อแสดงในหน้าเว็บโดยไม่ต้องเปิดแชร์ไฟล์ */
function imageData_(id, cache) {
  id = str_(id).trim();
  if (!id) return '';
  cache = cache || {};
  if (cache[id] !== undefined) return cache[id];
  try {
    const blob = DriveApp.getFileById(id).getBlob();
    cache[id] = 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
  } catch (e) { cache[id] = ''; }
  return cache[id];
}

/** รับรหัสไฟล์ หรือ ลิงก์ Google Drive แล้วคืนรหัสไฟล์ */
function driveId_(v) {
  v = String(v == null ? '' : v).trim();
  if (!v) return '';
  const m = /\/d\/([A-Za-z0-9_-]{20,})/.exec(v) || /[?&]id=([A-Za-z0-9_-]{20,})/.exec(v);
  if (m) return m[1];
  return IMAGE_ID_RULE.test(v) ? v : '#';
}

function adminUploadImage(token, f) {
  const adm = requireAdmin_(token);
  if (adm.err) return adm.err;
  f = f || {};
  const m = /^data:(image\/(?:png|jpeg|gif|webp));base64,([A-Za-z0-9+/=]+)$/.exec(String(f.dataUrl || ''));
  if (!m) return { ok: false, message: 'รองรับเฉพาะไฟล์ภาพ PNG, JPG, GIF หรือ WEBP' };
  if (m[2].length > 6 * 1024 * 1024) return { ok: false, message: 'ไฟล์ภาพใหญ่เกินไป (ไม่เกินประมาณ 4 MB)' };
  const ext = { 'image/png': 'png', 'image/jpeg': 'jpg', 'image/gif': 'gif', 'image/webp': 'webp' }[m[1]];
  const base = String(f.name || 'image').replace(/\.[^.]+$/, '').replace(/[^\wก-๙ -]/g, '').slice(0, 60) || 'image';
  const name = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss') + '_' + base + '.' + ext;
  const file = ensureImageFolder_().createFile(Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], name));
  file.setDescription('อัปโหลดโดย ' + whoOf_(adm.user));
  audit_(whoOf_(adm.user), 'อัปโหลดรูปข้อสอบ', file.getId(), name);
  return { ok: true, id: file.getId() };
}

function adminGetImage(token, id) {
  const adm = requireAdmin_(token);
  if (adm.err) return adm.err;
  const data = imageData_(id, {});
  return data ? { ok: true, data } : { ok: false, message: 'เปิดรูปไม่ได้ (ไฟล์อาจถูกลบหรือไม่มีสิทธิ์เข้าถึง)' };
}

/** ตรวจและแปลงข้อมูลข้อสอบ 1 ข้อ — ใช้ทั้งฟอร์มในเว็บและการนำเข้าจาก Excel */
function buildQuestion_(q) {
  const rank = String(q.rank || '').trim();
  if (!rank || !findExam_(rank)) return { err: 'ไม่พบชุดข้อสอบของยศ "' + rank + '" ในชีต Exams' };
  const type = q.type === 'WRITTEN' ? 'WRITTEN' : 'CHOICE';
  const points = Number(q.points);
  if (!(points > 0 && points <= 100)) return { err: 'คะแนนต่อข้อต้องอยู่ระหว่าง 1–100' };
  const question = String(q.question == null ? '' : q.question).trim();
  if (question.length < 3 || question.length > 2000) return { err: 'กรุณากรอกคำถาม (3–2000 ตัวอักษร)' };

  const ch = {};
  ['A', 'B', 'C', 'D'].forEach(k => { ch[k] = type === 'CHOICE' ? String(q[k] == null ? '' : q[k]).trim() : ''; });
  let answer = '';
  if (type === 'CHOICE') {
    if (!ch.A || !ch.B) return { err: 'ข้อช้อยต้องมีตัวเลือกอย่างน้อย A และ B' };
    if (['A', 'B', 'C', 'D'].some(k => ch[k].length > 500)) return { err: 'ตัวเลือกยาวเกิน 500 ตัวอักษร' };
    answer = norm_(q.answer);
    if (!ch[answer]) return { err: 'คำตอบที่ถูกต้องเป็น A–D และต้องเป็นตัวเลือกที่มีข้อความ' };
  }
  const keywords = type === 'WRITTEN' ? String(q.keywords == null ? '' : q.keywords).trim() : '';
  if (keywords.length > 1000) return { err: 'คีย์เวิร์ดยาวเกิน 1000 ตัวอักษร' };
  const image = driveId_(q.image);
  if (image === '#') return { err: 'รูปภาพต้องเป็นลิงก์หรือรหัสไฟล์ Google Drive' };

  return { data: {
    Image: image,
    ExamRank: rank, Type: type, Points: points, Question: safe_(question),
    A: safe_(ch.A), B: safe_(ch.B), C: safe_(ch.C), D: safe_(ch.D), Answer: answer, Keywords: safe_(keywords)
  } };
}

/* ---------- นำเข้าข้อสอบแบบ Batch จาก Excel (ADMIN ขึ้นไป) ---------- */

const IMPORT_MAX_ROWS = 500;

/**
 * rows = [{ row, id, set, type, points, question, A, B, C, D, answer, keywords }] (หน้าเว็บอ่านจาก Excel แล้วส่งมา)
 * - มี id ที่ตรงกับข้อสอบเดิม → อัปเดตข้อนั้น / ไม่มี id → เพิ่มเป็นข้อใหม่
 * - set = ยศ หรือชื่อชุดข้อสอบ (ว่าง = ใช้ defaultRank)
 * - dryRun = true → ตรวจสอบอย่างเดียว ยังไม่บันทึก
 * ถ้ามีแถวไหนผิด จะไม่บันทึกเลยสักแถว
 */
function adminImportQuestions(token, rows, defaultRank, dryRun) {
  const adm = requireAdmin_(token);
  if (adm.err) return adm.err;
  rows = Array.isArray(rows) ? rows : [];
  if (!rows.length) return { ok: false, message: 'ไม่พบข้อมูลในไฟล์' };
  if (rows.length > IMPORT_MAX_ROWS) return { ok: false, message: 'นำเข้าได้ครั้งละไม่เกิน ' + IMPORT_MAX_ROWS + ' ข้อ (ไฟล์นี้มี ' + rows.length + ' แถว)' };

  const exams = table_(SH.EXAMS).rows;
  const resolveSet = v => {
    const k = rankKey_(v);
    if (!k) return '';
    const e = exams.find(x => rankKey_(x.Rank) === k) || exams.find(x => rankKey_(x.ExamName) === k);
    return e ? str_(e.Rank) : '#' + String(v).trim();
  };
  const parseType = (v, a) => qType_({ Type: v, A: a });
  const thaiAns = { 'ก': 'A', 'ข': 'B', 'ค': 'C', 'ง': 'D' };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const t = table_(SH.QUESTIONS);
    const byId = {};
    t.rows.forEach(r => { byId[str_(r.ID).toUpperCase()] = r; });

    const seen = {};
    const preview = [], prepared = [];
    rows.forEach((x, i) => {
      const rowNo = Number(x.row) || i + 2;
      const id = String(x.id == null ? '' : x.id).trim();
      let rank = resolveSet(x.set) || String(defaultRank || '').trim();
      let err = '';
      if (rank.charAt(0) === '#') { err = 'ไม่พบชุดข้อสอบ "' + rank.slice(1) + '"'; rank = ''; }
      const type = parseType(x.type, x.A);
      const ansRaw = String(x.answer == null ? '' : x.answer).trim();
      const q = {
        rank, type, question: x.question, A: x.A, B: x.B, C: x.C, D: x.D, keywords: x.keywords, image: x.image,
        points: String(x.points == null ? '' : x.points).trim() === '' ? (type === 'WRITTEN' ? 5 : 1) : x.points,
        answer: thaiAns[ansRaw] || ansRaw
      };
      const existing = id ? byId[id.toUpperCase()] : null;
      if (!err && id && !existing) err = 'ไม่พบข้อสอบรหัส ' + id + ' (ถ้าจะเพิ่มข้อใหม่ ให้เว้นช่อง ID ว่าง)';
      if (!err && id) {
        if (seen[id.toUpperCase()]) err = 'รหัส ' + id + ' ซ้ำกับแถว ' + seen[id.toUpperCase()];
        seen[id.toUpperCase()] = rowNo;
      }
      let v = null;
      if (!err) { v = buildQuestion_(q); if (v.err) err = v.err; }
      if (!err && v.data.Image) {
        try { DriveApp.getFileById(v.data.Image).getName(); }
        catch (e) { err = 'เปิดไฟล์รูปไม่ได้ — ต้องเป็นไฟล์ที่บัญชีเจ้าของระบบเข้าถึงได้'; }
      }
      preview.push({
        row: rowNo, action: existing ? 'อัปเดต' : 'เพิ่ม', id: existing ? str_(existing.ID) : '',
        rank: rank, type: type, points: Number(q.points) || '',
        question: String(x.question == null ? '' : x.question).slice(0, 120), error: err
      });
      if (!err) prepared.push({ existing, data: v.data });
    });

    const errors = preview.filter(p => p.error).length;
    const add = preview.filter(p => !p.error && p.action === 'เพิ่ม').length;
    const update = preview.filter(p => !p.error && p.action === 'อัปเดต').length;
    if (dryRun || errors) return { ok: true, dryRun: true, preview, errors, add, update };

    // ---- บันทึกจริง ----
    const who = whoOf_(adm.user), now = new Date();
    const batch = 'IMP-' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
    ['Image', 'Keywords', 'CreatedBy', 'CreatedAt', 'UpdatedBy', 'UpdatedAt', 'ImportBatch'].forEach(h => {
      if (t.headers.indexOf(h) === -1) addHeader_(t, h);
    });
    const width = t.headers.length;
    const idx = h => t.headers.indexOf(h);

    // อัปเดตข้อเดิม (เขียนทั้งแถวครั้งเดียว)
    prepared.filter(p => p.existing).forEach(p => {
      const range = t.sh.getRange(p.existing._row, 1, 1, width);
      const vals = range.getValues()[0];
      Object.assign(p.data, { UpdatedBy: who, UpdatedAt: now, ImportBatch: batch });
      Object.keys(p.data).forEach(k => { vals[idx(k)] = p.data[k]; });
      range.setValues([vals]);
    });

    // เพิ่มข้อใหม่ (เขียนเป็นก้อนเดียว)
    const used = new Set(t.rows.map(r => str_(r.ID)));
    let n = t.rows.reduce((m, r) => { const m2 = /^Q(\d+)$/.exec(str_(r.ID)); return m2 ? Math.max(m, Number(m2[1])) : m; }, 0);
    const newIds = [];
    const newRows = prepared.filter(p => !p.existing).map(p => {
      let id;
      do { n++; id = 'Q' + String(n).padStart(3, '0'); } while (used.has(id));
      used.add(id); newIds.push(id);
      const d = Object.assign({ ID: id }, p.data, { CreatedBy: who, CreatedAt: now, ImportBatch: batch });
      return t.headers.map(h => d[h] === undefined ? '' : d[h]);
    });
    if (newRows.length) {
      t.sh.getRange(t.sh.getLastRow() + 1, 1, newRows.length, width).setValues(newRows);
    }

    audit_(who, 'นำเข้าข้อสอบจาก Excel', batch, 'เพิ่ม ' + newRows.length + ' · อัปเดต ' + update);
    return { ok: true, dryRun: false, batch, add: newRows.length, update, newIds };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- รายงานคะแนน (ADMIN ขึ้นไป) ---------- */

/**
 * f = { from: 'yyyy-MM-dd', to: 'yyyy-MM-dd', exam: ชื่อชุดข้อสอบ|'' , status: 'ผ่าน'|'ไม่ผ่าน'|'wait'|'' }
 * กรองตามวันที่ส่งข้อสอบ (รวมทั้งวัน from และวัน to)
 */
function queryReport_(f) {
  f = f || {};
  const tz = Session.getScriptTimeZone();
  const re = /^\d{4}-\d{2}-\d{2}$/;
  if (!re.test(String(f.from)) || !re.test(String(f.to))) return { err: 'กรุณาเลือกวันที่เริ่มต้นและวันที่สิ้นสุด' };
  const start = Utilities.parseDate(f.from + ' 00:00:00', tz, 'yyyy-MM-dd HH:mm:ss').getTime();
  const end = Utilities.parseDate(f.to + ' 23:59:59', tz, 'yyyy-MM-dd HH:mm:ss').getTime();
  if (start > end) return { err: 'วันที่เริ่มต้นต้องไม่เกินวันที่สิ้นสุด' };

  const waiting = r => String(r.Result).trim() === PENDING || str_(r.Review) === REVIEW_AUTO;
  const rows = table_(SH.RESULTS).rows.filter(r => {
    if (!(r.Timestamp instanceof Date)) return false;
    const t = r.Timestamp.getTime();
    if (t < start || t > end) return false;
    if (f.exam && str_(r.ExamName) !== f.exam) return false;
    if (f.status === 'wait') return waiting(r);
    if (f.status) return !waiting(r) && str_(r.Result) === f.status;
    return true;
  }).sort((a, b) => b.Timestamp - a.Timestamp);

  const items = rows.map(r => Object.assign(publicResult_(r), {
    employeeId: str_(r.EmployeeID), name: str_(r.Name), rank: str_(r.Rank), gradedBy: str_(r.GradedBy),
    day: Utilities.formatDate(r.Timestamp, tz, 'yyyy-MM-dd'), waiting: waiting(r)
  }));

  const sum = list => {
    const decided = list.filter(x => !x.waiting);
    const pass = decided.filter(x => x.result === 'ผ่าน').length;
    const pcts = decided.map(x => x.percent).filter(v => v != null);
    return {
      count: list.length, pass, fail: decided.length - pass, wait: list.length - decided.length,
      passRate: decided.length ? Math.round(pass * 100 / decided.length) : 0,
      avg: pcts.length ? Math.round(pcts.reduce((a, b) => a + b, 0) / pcts.length * 10) / 10 : null,
      max: pcts.length ? Math.max.apply(null, pcts) : null,
      min: pcts.length ? Math.min.apply(null, pcts) : null,
      people: new Set(list.map(x => x.employeeId)).size
    };
  };
  const group = key => {
    const m = {};
    items.forEach(x => { (m[x[key]] = m[x[key]] || []).push(x); });
    return Object.keys(m).sort().map(k => Object.assign({ key: k }, sum(m[k])));
  };

  return { items, summary: sum(items), byExam: group('exam'), byDay: group('day') };
}

function adminReport(token, f) {
  const adm = requireAdmin_(token);
  if (adm.err) return adm.err;
  const q = queryReport_(f);
  if (q.err) return { ok: false, message: q.err };
  return {
    ok: true, summary: q.summary, byExam: q.byExam, byDay: q.byDay,
    items: q.items.slice(0, 3000), total: q.items.length,
    sets: table_(SH.EXAMS).rows.map(e => str_(e.ExamName))
  };
}

/** ส่งออกรายงานเป็นชีตใหม่ในไฟล์หลังบ้าน (ชื่อเดิมจะถูกสร้างใหม่ทับ) */
function adminExportReport(token, f) {
  const adm = requireAdmin_(token);
  if (adm.err) return adm.err;
  const q = queryReport_(f);
  if (q.err) return { ok: false, message: q.err };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = ss_();
    const name = ('Report ' + f.from + ' ถึง ' + f.to + (f.exam ? ' ' + f.exam : '')).slice(0, 90);
    const old = ss.getSheetByName(name);
    if (old) ss.deleteSheet(old);
    const sh = ss.insertSheet(name);
    const S = q.summary;
    const head = [
      ['รายงานคะแนนสอบ', '', '', ''],
      ['ช่วงวันที่', f.from + ' ถึง ' + f.to, 'ชุดข้อสอบ', f.exam || 'ทั้งหมด'],
      ['สถานะ', f.status === 'wait' ? 'รอตรวจ/รอยืนยัน' : (f.status || 'ทั้งหมด'), 'ส่งออกโดย', whoOf_(adm.user) + ' · ' + fmtDate_(new Date())],
      ['จำนวนครั้งที่สอบ', S.count, 'จำนวนผู้สอบ (คน)', S.people],
      ['ผ่าน', S.pass, 'ไม่ผ่าน', S.fail],
      ['รอตรวจ/รอยืนยัน', S.wait, 'อัตราผ่าน', S.passRate + '%'],
      ['คะแนนเฉลี่ย', S.avg == null ? '-' : S.avg + '%', 'สูงสุด / ต่ำสุด', S.max == null ? '-' : S.max + '% / ' + S.min + '%']
    ];
    sh.getRange(1, 1, head.length, 4).setValues(head);
    sh.getRange(1, 1).setFontWeight('bold').setFontSize(14);
    sh.getRange(2, 1, head.length - 1, 1).setFontWeight('bold');
    sh.getRange(2, 3, head.length - 1, 1).setFontWeight('bold');

    const cols = ['วันเวลา', 'รหัส', 'ชื่อ', 'ยศตอนสอบ', 'ชุดข้อสอบ', 'ข้อช้อย', 'ข้อเขียน', 'คะแนนรวม', 'คะแนนเต็ม', '%', 'ผล', 'สถานะตรวจ', 'ผู้ตรวจ', 'หมายเหตุ'];
    const start = head.length + 2;
    const data = q.items.map(x => [
      x.date, x.employeeId, x.name, x.rank, x.exam,
      x.choiceScore + '/' + x.choiceMax,
      x.writtenMax ? (x.writtenScore == null ? '-' : x.writtenScore) + '/' + x.writtenMax : '-',
      x.score == null ? '' : x.score, x.total, x.percent == null ? '' : x.percent,
      x.pending ? 'รอตรวจ' : x.result, x.review, x.gradedBy, x.note
    ]);
    sh.getRange(start, 1, 1, cols.length).setValues([cols]).setFontWeight('bold').setBackground('#d9f2ef');
    if (data.length) sh.getRange(start + 1, 1, data.length, cols.length).setNumberFormat('@').setValues(data);
    sh.setFrozenRows(start);
    sh.autoResizeColumns(1, cols.length);

    audit_(whoOf_(adm.user), 'ส่งออกรายงาน', name, data.length + ' รายการ');
    return { ok: true, sheet: name, rows: data.length, url: ss.getUrl() + '#gid=' + sh.getSheetId() };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- Super Admin (ผู้พัฒนา): ผู้ใช้ / สิทธิ์ / ชุดข้อสอบ ---------- */

function requireSuper_(token) {
  const tok = readToken_(token, ['auth']);
  if (!tok) return { err: { ok: false, expired: true, message: 'เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่' } };
  const user = findUser_(tok.code);
  if (!user || mustChange_(user) || !isSuper_(user.Rank)) {
    return { err: { ok: false, message: 'เฉพาะผู้พัฒนาเท่านั้น' } };
  }
  return { user };
}

/** ข้อมูลทั้งหมดสำหรับหน้าผู้พัฒนา */
function superGetAll(token) {
  const sup = requireSuper_(token);
  if (sup.err) return sup.err;

  const users = table_(SH.USERS).rows.map(u => ({
    id: str_(u.EmployeeID), name: str_(u.Name), nameOC: str_(u.NameOC), age: str_(u.AgeIC),
    discord: str_(u.Discord), rank: str_(u.Rank), createdAt: fmtDate_(u.CreatedAt),
    hasPassword: !!str_(u.PasswordHash).trim(), mustChange: mustChange_(u), level: levelOf_(u.Rank)
  }));
  const perms = ss_().getSheetByName(SH.PERMS)
    ? table_(SH.PERMS).rows.filter(r => norm_(r.Level) === 'ADMIN')
        .map(r => ({ rank: str_(r.Rank), by: str_(r.UpdatedBy), at: fmtDate_(r.UpdatedAt) }))
    : ADMIN_RANKS.map(r => ({ rank: r, by: '', at: '' }));
  const pool = {};
  table_(SH.QUESTIONS).rows.forEach(q => {
    if (!String(q.Question).trim()) return;
    const k = norm_(q.ExamRank);
    pool[k] = pool[k] || { c: 0, w: 0 };
    pool[k][qType_(q) === 'CHOICE' ? 'c' : 'w']++;
  });
  const exams = table_(SH.EXAMS).rows.map(e => ({
    poolChoice: (pool[norm_(e.Rank)] || {}).c || 0, poolWritten: (pool[norm_(e.Rank)] || {}).w || 0,
    rank: str_(e.Rank), name: str_(e.ExamName), numChoice: str_(e.NumQuestions), numWritten: str_(e.NumWritten),
    pass: str_(e.PassPercent), next: str_(e.NextRank), auto: truthy_(e.AutoPromote),
    time: timeOf_(e), showScore: showScore_(e), sendDiscord: sendDiscord_(e)
  }));

  const ranks = new Set(SUPER_RANKS);
  perms.forEach(p => ranks.add(p.rank));
  exams.forEach(e => { ranks.add(e.rank); if (e.next) ranks.add(e.next); });
  users.forEach(u => { if (u.rank) ranks.add(u.rank); });

  const hook = discordWebhook_();
  const discord = { configured: !!hook, masked: hook ? hook.replace(/\/[\w-]+$/, '/••••••' + hook.slice(-4)) : '' };
  return { ok: true, me: str_(sup.user.EmployeeID), users, perms, exams, discord, superRanks: SUPER_RANKS, ranks: Array.from(ranks) };
}

/** เพิ่มผู้ใช้ (isNew) หรือแก้ชื่อ/ยศ */
function superSaveUser(token, u) {
  const sup = requireSuper_(token);
  if (sup.err) return sup.err;
  u = u || {};
  if (u.isNew && !initialPassword_()) return { ok:false, message:'ยังไม่ได้ตั้ง Script Property INITIAL_PASSWORD ใน Apps Script' };
  const id = String(u.id || '').trim();
  const name = String(u.name || '').trim();
  const rank = String(u.rank || '').trim();
  if (!/^[A-Za-z0-9_-]{2,20}$/.test(id)) return { ok: false, message: 'รหัสพนักงานใช้ได้เฉพาะ A–Z, 0–9, _ และ - (2–20 ตัว)' };
  if (name.length < 2 || name.length > 60) return { ok: false, message: 'กรุณากรอกชื่อ (2–60 ตัวอักษร)' };
  if (!rank || rank.length > 40) return { ok: false, message: 'กรุณากรอกยศ' };
  const discordIn = String(u.discord || '').trim();
  if (discordIn && !discordUserId_(discordIn)) {
    return { ok: false, message: 'Discord ต้องเป็นลิงก์โปรไฟล์ (https://discord.com/users/...) หรือเลข User ID' };
  }
  const discordVal = discordIn ? 'https://discord.com/users/' + discordUserId_(discordIn) : '';

  const me = norm_(sup.user.EmployeeID);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const t = table_(SH.USERS);
    const row = t.rows.find(r => norm_(r.EmployeeID) === norm_(id));
    if (u.isNew) {
      if (row) return { ok: false, message: 'รหัส ' + id + ' มีอยู่แล้ว' };
      appendObj_(t, { EmployeeID: id, Name: safe_(name), Rank: rank, Discord: discordVal, CreatedAt: new Date(), PasswordHash: '', MustChange: true });
      audit_(whoOf_(sup.user), 'เพิ่มผู้ใช้', id, name + ' · ' + rank);
    } else {
      if (!row) return { ok: false, message: 'ไม่พบผู้ใช้ ' + id };
      if (norm_(id) === me && !isSuper_(rank)) return { ok: false, message: 'ไม่สามารถลดสิทธิ์ของตัวเองได้ (กันล็อกตัวเองออกจากระบบ)' };
      const old = str_(row.Rank);
      setCell_(t, row, 'Name', safe_(name));
      setCell_(t, row, 'Rank', rank);
      if (u.discord !== undefined) setCell_(t, row, 'Discord', discordVal);
      audit_(whoOf_(sup.user), 'แก้ไขผู้ใช้', id, old === rank ? 'ชื่อ: ' + name : 'ยศ: ' + old + ' → ' + rank);
    }
    return { ok: true, defaultPassword: u.isNew ? initialPassword_() : '' };
  } finally {
    lock.releaseLock();
  }
}

/** รีเซ็ตรหัสผ่านไปเป็นรหัสตั้งต้นใน Script Properties และบังคับเปลี่ยนตอนเข้าสู่ระบบ */
function superResetPassword(token, id) {
  const sup = requireSuper_(token);
  if (sup.err) return sup.err;
  if (norm_(id) === norm_(sup.user.EmployeeID)) return { ok: false, message: 'เปลี่ยนรหัสของตัวเองที่ปุ่ม "เปลี่ยนรหัสผ่าน" ในหน้าหลัก' };
  const initialPassword = initialPassword_();
  if (!initialPassword) return { ok:false, message:'ยังไม่ได้ตั้ง Script Property INITIAL_PASSWORD ใน Apps Script' };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const t = table_(SH.USERS);
    const row = t.rows.find(r => norm_(r.EmployeeID) === norm_(id));
    if (!row) return { ok: false, message: 'ไม่พบผู้ใช้' };
    setCell_(t, row, 'PasswordHash', '');
    setCell_(t, row, 'MustChange', true);
    CacheService.getScriptCache().remove('fail_' + norm_(id));
    audit_(whoOf_(sup.user), 'รีเซ็ตรหัสผ่าน', String(id), '');
    return { ok: true, defaultPassword: initialPassword };
  } finally {
    lock.releaseLock();
  }
}

function superDeleteUser(token, id) {
  const sup = requireSuper_(token);
  if (sup.err) return sup.err;
  if (norm_(id) === norm_(sup.user.EmployeeID)) return { ok: false, message: 'ไม่สามารถลบบัญชีของตัวเองได้' };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const t = table_(SH.USERS);
    const row = t.rows.find(r => norm_(r.EmployeeID) === norm_(id));
    if (!row) return { ok: false, message: 'ไม่พบผู้ใช้' };
    t.sh.deleteRow(row._row);
    audit_(whoOf_(sup.user), 'ลบผู้ใช้', String(id), str_(row.Name) + ' · ' + str_(row.Rank));
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/** ให้/ถอนสิทธิ์ผู้บริหารของยศ — level: 'ADMIN' หรือ 'NONE' */
function superSetPerm(token, rank, level) {
  const sup = requireSuper_(token);
  if (sup.err) return sup.err;
  rank = String(rank || '').trim();
  if (!rank || rank.length > 40) return { ok: false, message: 'กรุณากรอกชื่อยศ' };
  if (SUPER_RANKS.map(rankKey_).indexOf(rankKey_(rank)) !== -1) return { ok: false, message: 'ยศผู้พัฒนามีสิทธิ์สูงสุดอยู่แล้ว' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    ensurePermsSheet_();
    const t = table_(SH.PERMS);
    const row = t.rows.find(r => rankKey_(r.Rank) === rankKey_(rank));
    const who = whoOf_(sup.user);
    if (level === 'ADMIN') {
      if (row) { setCell_(t, row, 'Level', 'ADMIN'); setCell_(t, row, 'UpdatedBy', who); setCell_(t, row, 'UpdatedAt', new Date()); }
      else appendObj_(t, { Rank: rank, Level: 'ADMIN', UpdatedBy: who, UpdatedAt: new Date() });
      audit_(who, 'ให้สิทธิ์ผู้บริหาร', rank, '');
    } else {
      if (row) t.sh.deleteRow(row._row);
      audit_(who, 'ถอนสิทธิ์ผู้บริหาร', rank, '');
    }
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/** ตั้ง/ลบ Webhook ของห้อง Discord ที่จะแสดงผลสอบ */
function superSetDiscord(token, url) {
  const sup = requireSuper_(token);
  if (sup.err) return sup.err;
  url = String(url || '').trim();
  const props = PropertiesService.getScriptProperties();
  if (!url) {
    props.deleteProperty('DISCORD_WEBHOOK');
    audit_(whoOf_(sup.user), 'ลบ Discord Webhook', '', '');
    return { ok: true };
  }
  if (!WEBHOOK_RULE.test(url)) return { ok: false, message: 'ลิงก์ Webhook ไม่ถูกต้อง (ต้องขึ้นต้นด้วย https://discord.com/api/webhooks/...)' };
  props.setProperty('DISCORD_WEBHOOK', url);
  audit_(whoOf_(sup.user), 'ตั้ง Discord Webhook', '', '…' + url.slice(-6));
  return { ok: true };
}

function superTestDiscord(token) {
  const sup = requireSuper_(token);
  if (sup.err) return sup.err;
  const status = postDiscord_({
    username: 'ระบบสอบแพทย์',
    embeds: [{ title: '🔔 ทดสอบการเชื่อมต่อ', description: 'ห้องนี้จะแสดงผลสอบของชุดข้อสอบที่เปิด "ส่งผลเข้า Discord"',
               color: 0x0f8b7d, footer: { text: 'ทดสอบโดย ' + whoOf_(sup.user) }, timestamp: new Date().toISOString() }]
  });
  return status === 'ส่งแล้ว' ? { ok: true } : { ok: false, message: status };
}

function ensurePermsSheet_() {
  const ss = ss_();
  if (ss.getSheetByName(SH.PERMS)) return;
  const sh = ss.insertSheet(SH.PERMS);
  sh.getRange(1, 1, 1, 4).setValues([['Rank', 'Level', 'UpdatedBy', 'UpdatedAt']]).setFontWeight('bold').setBackground('#d9f2ef');
  if (ADMIN_RANKS.length) sh.getRange(2, 1, ADMIN_RANKS.length, 4).setValues(ADMIN_RANKS.map(r => [r, 'ADMIN', 'setup', new Date()]));
  sh.setFrozenRows(1);
}

/** เพิ่มหรือแก้ตั้งค่าชุดข้อสอบ (ชีต Exams) */
function superSaveExam(token, e) {
  const sup = requireSuper_(token);
  if (sup.err) return sup.err;
  e = e || {};
  const rank = String(e.rank || '').trim();
  const name = String(e.name || '').trim();
  const numC = String(e.numChoice == null ? '' : e.numChoice).trim();
  const numW = String(e.numWritten == null ? '' : e.numWritten).trim();
  const pass = Number(e.pass);
  const next = String(e.next || '').trim();
  if (!rank || rank.length > 40) return { ok: false, message: 'กรุณากรอกยศที่ใช้ชุดข้อสอบนี้' };
  if (name.length < 2 || name.length > 80) return { ok: false, message: 'กรุณากรอกชื่อชุดข้อสอบ' };
  const okNum = v => v === '' || (/^\d+$/.test(v) && Number(v) <= 200);
  if (!okNum(numC) || !okNum(numW)) return { ok: false, message: 'จำนวนข้อต้องเป็นตัวเลข 0–200 (เว้นว่าง = ใช้ทั้งหมดในคลัง)' };
  if (!(pass >= 0 && pass <= 100)) return { ok: false, message: 'เกณฑ์ผ่านต้องอยู่ระหว่าง 0–100%' };
  const time = Number(e.time);
  if (!(Number.isInteger(time) && time >= 1 && time <= 300)) return { ok: false, message: 'เวลาสอบต้องเป็นตัวเลข 1–300 นาที' };
  if (numC === '0' && numW === '0') return { ok: false, message: 'ต้องมีข้อช้อยหรือข้อเขียนอย่างน้อย 1 ข้อ' };

  const data = { ExamName: safe_(name), NumQuestions: numC === '' ? '' : Number(numC), NumWritten: numW === '' ? '' : Number(numW),
                 PassPercent: pass, TimeLimitMin: time, NextRank: next, AutoPromote: !!e.auto, ShowScore: !!e.showScore,
                 SendDiscord: !!e.sendDiscord };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const t = table_(SH.EXAMS);
    const row = t.rows.find(r => norm_(r.Rank) === norm_(rank));
    if (e.isNew) {
      if (row) return { ok: false, message: 'ยศ "' + rank + '" มีชุดข้อสอบอยู่แล้ว' };
      appendObj_(t, Object.assign({ Rank: rank }, data));
      audit_(whoOf_(sup.user), 'เพิ่มชุดข้อสอบ', rank, name);
    } else {
      if (!row) return { ok: false, message: 'ไม่พบชุดข้อสอบของยศ "' + rank + '"' };
      Object.keys(data).forEach(k => setCell_(t, row, k, data[k]));
      audit_(whoOf_(sup.user), 'แก้ไขชุดข้อสอบ', rank, name + ' · ช้อย ' + (numC || 'ทั้งหมด') + ' · เขียน ' + (numW || 'ทั้งหมด') +
        ' · ผ่าน ' + pass + '% · ' + time + ' นาที · ' + (e.showScore ? 'แสดงคะแนน' : 'ซ่อนคะแนน'));
    }
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- ตัวช่วย: รหัสผ่านและ token ---------- */

function validatePassword_(pw) {
  pw = String(pw || '');
  const miss = [];
  if (pw.length < 6 || pw.length > 10) miss.push('ความยาว 6–10 ตัว');
  if (!/\d/.test(pw)) miss.push('ตัวเลข');
  if (!/[a-z]/.test(pw)) miss.push('ตัวอักษรภาษาอังกฤษตัวเล็ก');
  if (!/[A-Z]/.test(pw)) miss.push('ตัวอักษรภาษาอังกฤษตัวใหญ่');
  if (!/[!-\/:-@\[-`{-~]/.test(pw)) miss.push('อักขระพิเศษ');
  if (/[^!-~]/.test(pw)) miss.push('ห้ามมีช่องว่างหรืออักษรภาษาไทย');
  if (miss.length || !PW_RULE.test(pw)) return 'รหัสผ่านยังไม่ครบเงื่อนไข: ' + miss.join(', ');
  return '';
}

function checkPassword_(user, pw) {
  const h = String(user.PasswordHash || '').trim();
  if (!h) { const initialPassword = initialPassword_(); return !!initialPassword && String(pw) === initialPassword; }
  return h === hash_(user.EmployeeID, pw);
}

function mustChange_(user) {
  return !String(user.PasswordHash || '').trim() || truthy_(user.MustChange);
}

function hash_(code, pw) {
  const props = PropertiesService.getScriptProperties();
  let pepper = props.getProperty('PW_PEPPER');
  if (!pepper) { pepper = Utilities.getUuid() + Utilities.getUuid(); props.setProperty('PW_PEPPER', pepper); }
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    pepper + '|' + norm_(code) + '|' + String(pw),
    Utilities.Charset.UTF_8);
  return 'sha256$' + bytes.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

function issueToken_(code, type) {
  const t = Utilities.getUuid();
  CacheService.getScriptCache().put('tok_' + t, JSON.stringify({ code: norm_(code), type: type }), SESSION_SECONDS);
  return t;
}

function readToken_(token, types) {
  if (!token) return null;
  const raw = CacheService.getScriptCache().get('tok_' + token);
  if (!raw) return null;
  const tok = JSON.parse(raw);
  if (types && types.indexOf(tok.type) === -1) return null;
  return tok;
}

/* ---------- ตัวช่วยทั่วไป ---------- */

function homeData_(user) {
  const exam = findExam_(user.Rank);
  let nC = 0, nW = 0;
  if (exam) {
    const pool = questionsFor_(user.Rank);
    const c = pool.filter(q => qType_(q) === 'CHOICE').length;
    const w = pool.length - c;
    nC = Math.min(count_(exam.NumQuestions, c), c);
    nW = Math.min(count_(exam.NumWritten, w), w);
  }
  return {
    user: { id: String(user.EmployeeID), name: String(user.Name), rank: String(user.Rank) },
    exam: exam ? {
      name: String(exam.ExamName), choice: nC, written: nW,
      pass: Number(exam.PassPercent) || 0,
      time: timeOf_(exam),
      showScore: showScore_(exam),
      next: String(exam.NextRank || '')
    } : null,
    history: historyFor_(user.EmployeeID, isAdmin_(user.Rank)),
    isAdmin: isAdmin_(user.Rank),
    isSuper: isSuper_(user.Rank),
    pending: isAdmin_(user.Rank) ? table_(SH.RESULTS).rows.filter(needsReview_).length : 0,
    notices: noticesFor_(user.EmployeeID)
  };
}

function promote_(users, user, rankAtExam) {
  const exam = findExam_(rankAtExam);
  if (exam && exam.NextRank && truthy_(exam.AutoPromote) && norm_(user.Rank) === norm_(rankAtExam)) {
    setCell_(users, user, 'Rank', exam.NextRank);
    user.Rank = exam.NextRank;
    return String(exam.NextRank);
  }
  return '';
}

function qType_(q) {
  const t = norm_(q.Type);
  if (['WRITTEN', 'W', 'เขียน', 'ข้อเขียน', 'อัตนัย'].indexOf(t) !== -1) return 'WRITTEN';
  if (['CHOICE', 'C', 'ช้อย', 'ปรนัย'].indexOf(t) !== -1) return 'CHOICE';
  return String(q.A == null ? '' : q.A).trim() ? 'CHOICE' : 'WRITTEN';
}

function str_(v) { return v == null ? '' : String(v); }

function fmtDate_(v) {
  return v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') : str_(v);
}

function points_(q) { const p = Number(q.Points); return p > 0 ? p : 1; }

/** ช่องว่าง = ใช้ทั้งหมดที่มีในคลัง */
function count_(v, all) { return String(v == null ? '' : v).trim() === '' ? all : Math.max(0, Number(v) || 0); }

/** กันข้อความที่ผู้ใช้พิมพ์ถูกตีความเป็นสูตรในชีต */
function safe_(s) { s = String(s == null ? '' : s); return /^[=+\-@]/.test(s) ? "'" + s : s; }

function discordKey_(s) { return String(s).trim().toLowerCase().replace(/^https?:\/\/(www\.|ptb\.|canary\.)?/, '').replace(/\/+$/, ''); }

/** ไฟล์ Google Sheet หลังบ้าน */
function ss_() {
  const active = SpreadsheetApp.getActive();
  if (active) return active;
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  throw new Error('สคริปต์นี้ยังไม่ได้ผูกกับ Google Sheet — ให้เปิด Apps Script จากเมนู "ส่วนขยาย → Apps Script" ในไฟล์ Google Sheet แล้ววางโค้ดที่นั่น');
}

let AUTO_SETUP_DONE_ = false;
function table_(name) {
  let sh = ss_().getSheetByName(name);
  if (!sh && !AUTO_SETUP_DONE_) {
    // ยังไม่ได้รัน setup → สร้างชีตที่ขาดให้อัตโนมัติ (ไม่ทับข้อมูลเดิม)
    AUTO_SETUP_DONE_ = true;
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try { setup(); } finally { lock.releaseLock(); }
    sh = ss_().getSheetByName(name);
  }
  if (!sh) throw new Error('ไม่พบชีต "' + name + '" — กรุณารันฟังก์ชัน setup ในหน้า Apps Script');
  const values = sh.getDataRange().getValues();
  const headers = values.shift().map(h => String(h).trim());
  const rows = values
    .map((r, i) => {
      const o = { _row: i + 2 };
      headers.forEach((h, j) => { o[h] = r[j]; });
      return o;
    })
    .filter(o => headers.some(h => String(o[h]).trim() !== ''));
  return { sh, headers, rows };
}

function addHeader_(t, header) {
  const col = t.sh.getLastColumn() + 1;
  t.sh.getRange(1, col).setValue(header).setFontWeight('bold').setBackground('#d9f2ef');
  t.headers.push(header);
  return col;
}

function setCell_(t, row, header, value) {
  let col = t.headers.indexOf(header) + 1;
  if (!col) col = addHeader_(t, header);
  t.sh.getRange(row._row, col).setValue(value);
}

function appendObj_(t, obj) {
  Object.keys(obj).forEach(k => { if (t.headers.indexOf(k) === -1) addHeader_(t, k); });
  t.sh.appendRow(t.headers.map(h => (obj[h] === undefined ? '' : obj[h])));
}

function norm_(v) { return String(v == null ? '' : v).trim().toUpperCase(); }

function truthy_(v) { return v === true || /^(true|yes|y|1|ใช่)$/i.test(String(v).trim()); }

function findUser_(code) {
  const c = norm_(code);
  if (!c) return null;
  return table_(SH.USERS).rows.find(r => norm_(r.EmployeeID) === c) || null;
}

function findExam_(rank) {
  const r = norm_(rank);
  return table_(SH.EXAMS).rows.find(e => norm_(e.Rank) === r) || null;
}

function questionsFor_(rank) {
  const r = norm_(rank);
  return table_(SH.QUESTIONS).rows.filter(q => norm_(q.ExamRank) === r && String(q.Question).trim() !== '');
}

/** ประวัติสอบของตัวเอง — ผู้ใช้ทั่วไปเห็นเฉพาะชุดข้อสอบที่เปิดให้ดูคะแนน */
function historyFor_(code, all) {
  const c = norm_(code);
  return table_(SH.RESULTS).rows
    .filter(r => norm_(r.EmployeeID) === c && (all || examCfg_(r.Rank).show))
    .slice(-5).reverse()
    .map(r => {
      const pending = String(r.Result).trim() === PENDING;
      return {
        date: fmtDate_(r.Timestamp),
        exam: String(r.ExamName),
        score: pending ? 'ช้อย ' + r.ChoiceScore + '/' + r.ChoiceMax : r.Score + '/' + r.Total,
        result: String(r.Result) + (str_(r.Review) === REVIEW_AUTO && !pending ? ' (รอยืนยัน)' : '')
      };
    });
}

function shuffle_(arr) {
  const a = arr.slice();
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}

/* ---------- ติดตั้ง: รันฟังก์ชันนี้ (รันซ้ำได้ ไม่ทับข้อมูลเดิม) ---------- */

const SAMPLE_WRITTEN = [
  ['W001', 'ผู้สมัคร', 'WRITTEN', 5, 'อธิบายขั้นตอนการช่วยเหลือผู้บาดเจ็บจากอุบัติเหตุรถชน ตั้งแต่ไปถึงที่เกิดเหตุจนส่งตัวถึงโรงพยาบาล', 'ความปลอดภัย/ปลอดภัย, ประเมิน, เรียก/แจ้ง, ห้ามเลือด, ส่งตัว/ส่งโรงพยาบาล'],
  ['W002', 'ผู้สมัคร', 'WRITTEN', 5, 'ถ้ามีผู้บาดเจ็บหลายคนพร้อมกันแต่มีแพทย์อยู่คนเดียว คุณจะจัดลำดับการรักษาอย่างไร เพราะอะไร', 'คัดแยก/triage, อาการหนัก/วิกฤต, ก่อน, เร่งด่วน'],
  ['W003', 'ผู้สมัคร', 'WRITTEN', 5, 'ในความคิดของคุณ แพทย์ที่ดีควรมีคุณสมบัติอะไรบ้าง และคุณมีคุณสมบัติข้อใด']
];

function setup() {
  AUTO_SETUP_DONE_ = true;
  const ss = ss_();
  if (!initialPassword_()) throw new Error('ก่อนติดตั้ง ให้ตั้ง Script Property ชื่อ INITIAL_PASSWORD เป็นรหัสผ่านตั้งต้นที่คาดเดายาก');
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());
  const make = (name, headers, rows) => {
    let sh = ss.getSheetByName(name);
    if (sh) {
      const have = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0].map(String);
      headers.filter(h => have.indexOf(h) === -1).forEach(h => {
        const col = sh.getLastColumn() + 1;
        sh.getRange(1, col).setValue(h).setFontWeight('bold').setBackground('#d9f2ef');
      });
      return false;
    }
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#d9f2ef');
    if (rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, headers.length);
    return true;
  };

  make(SH.USERS,
    ['EmployeeID', 'Name', 'NameOC', 'AgeIC', 'Discord', 'Reason', 'Rank', 'CreatedAt', 'PasswordHash', 'MustChange'], [
    // เริ่มต้นมีแค่ผู้พัฒนา (DEV001) — รหัสตั้งต้นเก็บใน Script Property; DEV001 เป็นคนสร้างผู้ใช้/สิทธิ์อื่น
    ['DEV001', 'ผู้พัฒนาระบบ', '', '', '', '', SUPER_RANKS[0], new Date(), '', true]
  ]);

  const examsCreated = make(SH.EXAMS,
    ['Rank', 'ExamName', 'NumQuestions', 'NumWritten', 'PassPercent', 'TimeLimitMin', 'NextRank', 'AutoPromote', 'ShowScore', 'SendDiscord'], [
    ['ผู้สมัคร', 'ข้อสอบเข้าแพทย์', 5, 2, 60, 60, 'แพทย์ใหม่', false, false, false],
    ['แพทย์ใหม่', 'ข้อสอบเลื่อนยศเป็นแพทย์', 40, 0, 70, 60, 'แพทย์', false, true, true]
  ]);
  // ชีต Exams เดิม: ตั้งค่า ShowScore (แพทย์ใหม่ = แสดง) และเวลาสอบ 60 นาทีให้ครั้งเดียว
  const props = PropertiesService.getScriptProperties();
  if (!examsCreated && !props.getProperty('MIGRATED_EXAM_V2')) {
    const et = table_(SH.EXAMS);
    et.rows.forEach(e => {
      if (str_(e.ShowScore).trim() === '') setCell_(et, e, 'ShowScore', rankKey_(e.Rank) === rankKey_('แพทย์ใหม่'));
      setCell_(et, e, 'TimeLimitMin', 60);
    });
  }
  props.setProperty('MIGRATED_EXAM_V2', '1');
  // ชีต Exams เดิม: เปิดส่ง Discord ให้ทุกชุด ยกเว้นชุดผู้สมัคร (ครั้งเดียว)
  if (!examsCreated && !props.getProperty('MIGRATED_DISCORD')) {
    const et = table_(SH.EXAMS);
    et.rows.forEach(e => { if (str_(e.SendDiscord).trim() === '') setCell_(et, e, 'SendDiscord', rankKey_(e.Rank) !== rankKey_(APPLICANT_RANK)); });
  }
  props.setProperty('MIGRATED_DISCORD', '1');

  const Q = ['ID', 'ExamRank', 'Type', 'Points', 'Question', 'A', 'B', 'C', 'D', 'Answer',
    'Keywords', 'CreatedBy', 'CreatedAt', 'UpdatedBy', 'UpdatedAt', 'ImportBatch', 'Image'];
  const c = (id, rank, q, a, b, cc, d, ans) => [id, rank, 'CHOICE', 1, q, a, b, cc, d, ans, '', 'setup', new Date(), '', '', '', ''];
  const w = r => [r[0], r[1], r[2], r[3], r[4], '', '', '', '', '', r[5] || '', 'setup', new Date(), '', '', '', ''];
  const created = make(SH.QUESTIONS, Q, [
    c('Q001', 'ผู้สมัคร', 'พบผู้ป่วยหมดสติและไม่หายใจ ควรทำอะไรเป็นอันดับแรก', 'ให้ดื่มน้ำ', 'เริ่ม CPR และเรียกความช่วยเหลือ', 'พาไปนอนพัก', 'รอดูอาการ', 'B'),
    c('Q002', 'ผู้สมัคร', 'อัตราการกดหน้าอกขณะทำ CPR ในผู้ใหญ่ที่เหมาะสม (ครั้ง/นาที)', '60–80', '100–120', '140–160', '40–60', 'B'),
    c('Q003', 'ผู้สมัคร', 'แผลเลือดออกมาก ควรห้ามเลือดด้วยวิธีใดก่อน', 'กดที่แผลโดยตรง', 'ล้างด้วยแอลกอฮอล์', 'ปล่อยให้หยุดเอง', 'ทาขี้ผึ้ง', 'A'),
    c('Q004', 'ผู้สมัคร', 'ข้อใดไม่ใช่สัญญาณชีพ', 'ชีพจร', 'ความดันโลหิต', 'อัตราการหายใจ', 'กรุ๊ปเลือด', 'D'),
    c('Q005', 'ผู้สมัคร', 'ผู้ป่วยสงสัยกระดูกคอหัก ควรทำอย่างไร', 'ให้ลุกนั่ง', 'ยึดตรึงศีรษะและคอให้นิ่ง', 'หมุนคอเพื่อตรวจ', 'ยกขึ้นทันที', 'B'),
    c('Q006', 'ผู้สมัคร', 'แผลไฟลวกเล็กน้อย ควรปฐมพยาบาลอย่างไร', 'ทายาสีฟัน', 'ประคบน้ำแข็งโดยตรง', 'ราดน้ำสะอาดอุณหภูมิห้องนานๆ', 'เจาะตุ่มน้ำ', 'C'),
    w(SAMPLE_WRITTEN[0]), w(SAMPLE_WRITTEN[1]), w(SAMPLE_WRITTEN[2]),
    c('Q101', 'แพทย์ใหม่', 'อาการสำคัญของภาวะช็อกจากการเสียเลือด', 'ชีพจรเร็ว ความดันต่ำ', 'ชีพจรช้า ความดันสูง', 'ไข้สูง', 'ผื่นแดงทั้งตัว', 'A'),
    c('Q102', 'แพทย์ใหม่', 'การคัดแยกผู้ป่วย (Triage) ป้ายสีแดงหมายถึง', 'เสียชีวิตแล้ว', 'วิกฤต ต้องรักษาทันที', 'บาดเจ็บเล็กน้อย', 'รอได้', 'B'),
    c('Q103', 'แพทย์ใหม่', 'คะแนน GCS ต่ำสุดคือเท่าใด', '0', '1', '3', '5', 'C'),
    c('Q104', 'แพทย์ใหม่', 'ยาอันดับแรกในภาวะแพ้รุนแรง (Anaphylaxis)', 'Paracetamol', 'Epinephrine', 'Amoxicillin', 'Omeprazole', 'B'),
    c('Q105', 'แพทย์ใหม่', 'ค่า SpO2 ปกติในผู้ใหญ่', '70–80%', '85–90%', '95–100%', '50–60%', 'C'),
    c('Q106', 'แพทย์ใหม่', 'ABC ในการประเมินผู้ป่วยฉุกเฉินย่อมาจาก', 'Airway Breathing Circulation', 'Ambulance Bed Care', 'Assess Bandage Call', 'Alert Blood Check', 'A')
  ]);

  // ชีตข้อสอบเดิมที่ยังไม่มีข้อเขียน: เติมข้อเขียนตัวอย่างให้
  if (!created) {
    const t = table_(SH.QUESTIONS);
    const hasWritten = t.rows.some(q => norm_(q.ExamRank) === norm_(APPLICANT_RANK) && qType_(q) === 'WRITTEN');
    const ids = new Set(t.rows.map(q => String(q.ID)));
    if (!hasWritten) {
      t.rows.forEach(q => { if (!String(q.Type).trim()) setCell_(t, q, 'Type', 'CHOICE'); });
      SAMPLE_WRITTEN.filter(r => !ids.has(r[0])).forEach(r =>
        appendObj_(t, { ID: r[0], ExamRank: r[1], Type: r[2], Points: r[3], Question: r[4], Keywords: r[5] || '' }));
    }
  }

  make(SH.RESULTS, ['Timestamp', 'ResultID', 'EmployeeID', 'Name', 'Rank', 'ExamName',
    'ChoiceScore', 'ChoiceMax', 'WrittenScore', 'WrittenMax', 'Score', 'Total', 'Percent', 'Result', 'Review', 'Note', 'GradedBy', 'UpdatedAt', 'SeenAt', 'DiscordStatus'], []);

  ensurePermsSheet_();
  try { ensureImageFolder_(); } catch (e) { /* ยังไม่ได้อนุญาตสิทธิ์ Drive — รัน setup จากหน้า Apps Script เพื่ออนุญาต */ }
  make(SH.AUDIT, ['Timestamp', 'By', 'Action', 'Target', 'Detail'], []);

  // ถ้ายังไม่มีผู้พัฒนาเลย สร้างบัญชี DEV001 ให้ (รหัสตั้งต้นอยู่ใน Script Property)
  const ut = table_(SH.USERS);
  if (!ut.rows.some(u => isSuper_(u.Rank)) && !ut.rows.some(u => norm_(u.EmployeeID) === 'DEV001')) {
    appendObj_(ut, { EmployeeID: 'DEV001', Name: 'ผู้พัฒนาระบบ', Rank: SUPER_RANKS[0], CreatedAt: new Date(), PasswordHash: '', MustChange: true });
  }

  make(SH.ANSWERS, ['Timestamp', 'ResultID', 'EmployeeID', 'Name', 'QuestionID', 'Question', 'Answer', 'MaxPoints', 'AutoScore', 'Matched', 'Score'], []);

  make(SH.ATTENDANCE, ['Date', 'EmployeeID', 'Name', 'ClockIn', 'ClockOut', 'RoundedIn', 'RoundedOut', 'BreakHours', 'NetHours', 'ClockInPhoto', 'ClockOutPhoto', 'UpdatedAt'], []);
  make(SH.PAYCFG, ['Key', 'Value'], [
    ['BreakIntervals', '12:00-13:00,17:00-18:00,22:00-23:00'], ['MinimumHoursPerDay', 3],
    ['DailyMinimumEnabled', true], ['MinimumHoursPerWeek', 21], ['WeeklyMinimumEnabled', true],
    ['OvertimeThresholdHours', 21], ['OvertimeBonusPercent', 10],
    ['RoundMinutes', 30], ['Timezone', 'Asia/Bangkok']
  ]);
  const pct = table_(SH.PAYCFG), haveCfg = new Set(pct.rows.map(r => str_(r.Key)));
  [['BreakIntervals','12:00-13:00,17:00-18:00,22:00-23:00'],['MinimumHoursPerDay',3],['DailyMinimumEnabled',true],
    ['MinimumHoursPerWeek',21],['WeeklyMinimumEnabled',true],['OvertimeThresholdHours',21],['OvertimeBonusPercent',10],
    ['RoundMinutes',30],['Timezone','Asia/Bangkok']].forEach(([k,v]) => { if (!haveCfg.has(k)) appendObj_(pct,{Key:k,Value:v}); });
  make(SH.PAYRATES, ['EmployeeID', 'HourlyRate', 'UpdatedBy', 'UpdatedAt'], []);

  const def = ss.getSheetByName('Sheet1') || ss.getSheetByName('ชีต1');
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(def);
}

/* ---------- ลงเวลาและสรุปค่าจ้างรายสัปดาห์ ---------- */

function payrollCfg_() {
  const out = {};
  table_(SH.PAYCFG).rows.forEach(r => out[String(r.Key)] = r.Value);
  return out;
}

function payrollPunchDate_(ymd, hm) {
  const d = String(ymd || ''), t = String(hm || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(d) || !/^([01]\d|2[0-3]):[0-5]\d$/.test(t)) throw new Error('กรุณาระบุวันที่และเวลาให้ถูกต้อง');
  const tz = payrollCfg_().Timezone || 'Asia/Bangkok';
  const dt = Utilities.parseDate(d + ' ' + t, tz, 'yyyy-MM-dd HH:mm');
  if (Utilities.formatDate(dt, tz, 'yyyy-MM-dd HH:mm') !== d + ' ' + t) throw new Error('วันหรือเวลาไม่ถูกต้อง');
  return dt;
}

function payrollRoundHour_(dt) {
  const tz = payrollCfg_().Timezone || 'Asia/Bangkok';
  const date = Utilities.formatDate(dt, tz, 'yyyy-MM-dd');
  let hour = Number(Utilities.formatDate(dt, tz, 'HH'));
  if (Number(Utilities.formatDate(dt, tz, 'mm')) >= 30) hour++;
  if (hour >= 24) {
    const last = Utilities.parseDate(date + ' 23:00', tz, 'yyyy-MM-dd HH:mm');
    return new Date(last.getTime() + 3600000);
  }
  return Utilities.parseDate(date + ' ' + String(hour).padStart(2,'0') + ':00', tz, 'yyyy-MM-dd HH:mm');
}

function saveAttendancePhoto_(user, dataUrl, kind) {
  const m = String(dataUrl || '').match(/^data:(image\/(?:jpeg|png|webp));base64,([A-Za-z0-9+/=]+)$/);
  if (!m || m[2].length > 8 * 1024 * 1024) throw new Error('กรุณาแนบรูปภาพ JPEG, PNG หรือ WebP ขนาดไม่เกิน 6 MB');
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], norm_(user.EmployeeID) + '_' + kind + '_' + Date.now());
  const file = ensureAttendanceFolder_().createFile(blob);
  file.setDescription('หลักฐานลงเวลา ' + kind + ' · ' + safe_(user.Name) + ' (' + norm_(user.EmployeeID) + ')');
  return file.getId();
}

function attendanceStatus(token) {
  const tok = readToken_(token, ['auth']);
  if (!tok) return { ok:false, expired:true, message:'เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่' };
  const user = findUser_(tok.code); if (!user || mustChange_(user)) return { ok:false, expired:true, message:'กรุณาเข้าสู่ระบบใหม่' };
  const tz = payrollCfg_().Timezone || 'Asia/Bangkok', today = Utilities.formatDate(new Date(),tz,'yyyy-MM-dd');
  const rows = table_(SH.ATTENDANCE).rows.filter(r => norm_(r.EmployeeID)===tok.code).sort((a,b)=>String(a.Date).localeCompare(String(b.Date)));
  const open = rows.slice().reverse().find(r => r.ClockIn && !r.ClockOut);
  return {ok:true,name:safe_(user.Name),employeeId:tok.code,today,open:!!open,openDate:open?str_(open.Date):'',clockIn:open&&open.ClockIn instanceof Date?Utilities.formatDate(open.ClockIn,tz,'HH:mm'):'',
    todayRows:rows.filter(r=>str_(r.Date)===today).map(r=>({in:r.ClockIn instanceof Date?Utilities.formatDate(r.ClockIn,tz,'HH:mm'):'',out:r.ClockOut instanceof Date?Utilities.formatDate(r.ClockOut,tz,'HH:mm'):'',hours:Number(r.NetHours)||0}))};
}

function attendancePunch(token, kind, dataUrl) {
  const tok = readToken_(token, ['auth']);
  if (!tok) return { ok: false, expired: true, message: 'เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่' };
  const user = findUser_(tok.code);
  if (!user || mustChange_(user)) return { ok: false, expired: true, message: 'กรุณาเข้าสู่ระบบใหม่' };
  if (['in', 'out'].indexOf(kind) < 0) return { ok: false, message: 'ประเภทลงเวลาไม่ถูกต้อง' };
  try {
    const tz = payrollCfg_().Timezone || 'Asia/Bangkok';
    const dt = new Date(), date = Utilities.formatDate(dt,tz,'yyyy-MM-dd');
    const rounded = payrollRoundHour_(dt);
    const photoId = saveAttendancePhoto_(user,dataUrl,kind);
    const lock = LockService.getScriptLock(); lock.waitLock(10000);
    try {
      const t = table_(SH.ATTENDANCE);
      const row = t.rows.filter(r => norm_(r.EmployeeID)===tok.code).sort((a,b)=>String(b.Date).localeCompare(String(a.Date))).find(r=>r.ClockIn&&!r.ClockOut);
      if (kind === 'in') {
        if (row) return { ok: false, message: 'คุณมีเวรที่ยังไม่ได้ลงเวลาออก' };
        appendObj_(t, { Date: date, EmployeeID: tok.code, Name: safe_(user.Name), ClockIn: dt, RoundedIn: rounded, ClockInPhoto:photoId, UpdatedAt: new Date() });
      } else {
        if (!row || !row.ClockIn) return { ok: false, message: 'ยังไม่มีเวรที่เปิดอยู่ กรุณาลงเวลาเข้าก่อน' };
        if (dt.getTime() <= new Date(row.ClockIn).getTime()) return { ok: false, message: 'เวลาออกงานต้องอยู่หลังเวลาเข้างาน' };
        const rin = new Date(row.RoundedIn || payrollRoundHour_(new Date(row.ClockIn)));
        const rout = rounded;
        if (rout.getTime() <= rin.getTime()) return { ok: false, message: 'เมื่อปัดเป็นรายชั่วโมงแล้ว ระยะเวลาทำงานต้องมากกว่า 0 ชั่วโมง' };
        const cfg = payrollCfg_();
        let breakMs = 0;
        String(cfg.BreakIntervals || '').split(/[;,\n]/).map(x => x.trim()).filter(Boolean).forEach(x => {
          const m = x.match(/^([01]\d|2[0-3]):[0-5]\d\s*-\s*([01]\d|2[0-3]):[0-5]\d$/);
          if (!m) return;
          const a = payrollPunchDate_(date, m[0].split('-')[0].trim());
          const b = payrollPunchDate_(date, m[0].split('-')[1].trim());
          if (b <= a) b.setTime(a.getTime() + 24 * 3600000);
          breakMs += Math.max(0, Math.min(rout.getTime(), b.getTime()) - Math.max(rin.getTime(), a.getTime()));
        });
        const netHours = Math.max(0, (rout.getTime() - rin.getTime() - breakMs) / 3600000);
        setCell_(t, row, 'ClockOut', dt); setCell_(t, row, 'RoundedOut', rout);
        setCell_(t, row, 'ClockOutPhoto', photoId);
        setCell_(t, row, 'BreakHours', breakMs / 3600000); setCell_(t, row, 'NetHours', netHours);
        setCell_(t, row, 'UpdatedAt', new Date());
      }
      audit_(whoOf_(user), kind === 'in' ? 'ลงเวลาเข้า' : 'ลงเวลาออก', date, Utilities.formatDate(dt,tz,'HH:mm'));
      return { ok: true, message: kind === 'in' ? 'บันทึกเวลาเข้าแล้ว' : 'บันทึกเวลาออกแล้ว' };
    } finally { lock.releaseLock(); }
  } catch (e) { return { ok: false, message: e.message || String(e) }; }
}

function payrollPeriod_(date) {
  const tz = payrollCfg_().Timezone || 'Asia/Bangkok';
  const d = payrollPunchDate_(date, '12:00');
  const dow = Number(Utilities.formatDate(d, tz, 'u')); // Mon=1 ... Sun=7
  const sinceSat = dow % 7;
  const start = new Date(d.getTime()); start.setDate(start.getDate() - sinceSat);
  const end = new Date(start.getTime()); end.setDate(end.getDate() + 6);
  return { from: Utilities.formatDate(start, tz, 'yyyy-MM-dd'), to: Utilities.formatDate(end, tz, 'yyyy-MM-dd') };
}

function adminPayrollData(token, date) {
  const adm = requireAdmin_(token); if (adm.err) return adm.err;
  try {
    const period = payrollPeriod_(date || Utilities.formatDate(new Date(), payrollCfg_().Timezone || 'Asia/Bangkok', 'yyyy-MM-dd'));
    const cfg = payrollCfg_();
    const valOr = (v,d) => v == null || String(v).trim() === '' || !isFinite(Number(v)) ? d : Number(v);
    const minHours = Math.max(0, valOr(cfg.MinimumHoursPerDay,3));
    const minWeek = Math.max(0, valOr(cfg.MinimumHoursPerWeek,21));
    const dailyEnabled = cfg.DailyMinimumEnabled == null || String(cfg.DailyMinimumEnabled).trim() === '' ? true : truthy_(cfg.DailyMinimumEnabled);
    const weeklyEnabled = cfg.WeeklyMinimumEnabled == null || String(cfg.WeeklyMinimumEnabled).trim() === '' ? true : truthy_(cfg.WeeklyMinimumEnabled);
    const overtimeThreshold = Math.max(0, valOr(cfg.OvertimeThresholdHours,21));
    const overtimePercent = Math.max(0, valOr(cfg.OvertimeBonusPercent,10));
    const rates = {};
    table_(SH.PAYRATES).rows.forEach(r => rates[norm_(r.EmployeeID)] = Number(r.HourlyRate) || 0);
    const tz = cfg.Timezone || 'Asia/Bangkok';
    const items = table_(SH.ATTENDANCE).rows.filter(r => {
      const d = r.Date instanceof Date ? Utilities.formatDate(r.Date, tz, 'yyyy-MM-dd') : str_(r.Date);
      return d >= period.from && d <= period.to;
    }).map(r => {
      const dateKey = r.Date instanceof Date ? Utilities.formatDate(r.Date, tz, 'yyyy-MM-dd') : str_(r.Date);
      const hours = Number(r.NetHours) || 0, rate = Number(rates[norm_(r.EmployeeID)]) || 0;
      const complete = !!r.ClockOut;
      const eligible = complete && (!dailyEnabled || hours >= minHours);
      const fmt = v => v instanceof Date ? Utilities.formatDate(v, tz, 'HH:mm') : '';
      return { date: dateKey, employeeId: str_(r.EmployeeID), name: str_(r.Name), clockIn: fmt(r.ClockIn), clockOut: fmt(r.ClockOut),
        roundedIn: fmt(r.RoundedIn), roundedOut: fmt(r.RoundedOut), breakHours: Number(r.BreakHours) || 0,
        hours, hourlyRate: rate, complete, eligible, payableHours: eligible ? hours : 0, basePay: eligible ? hours * rate : 0, pay: 0 };
    }).sort((a,b) => a.date.localeCompare(b.date) || a.name.localeCompare(b.name));
    const byPerson = {};
    items.forEach(x => {
      const g = byPerson[x.employeeId] || (byPerson[x.employeeId] = { employeeId:x.employeeId, name:x.name, rate:x.hourlyRate, days:0, hours:0, basePay:0 });
      if (x.eligible) { g.days++; g.hours += x.payableHours; g.basePay += x.basePay; }
    });
    const people = Object.keys(byPerson).map(id => {
      const g = byPerson[id];
      g.weeklyEligible = !weeklyEnabled || g.hours >= minWeek;
      g.basePayBeforeWeeklyMinimum = g.basePay;
      g.basePay = g.weeklyEligible ? g.basePayBeforeWeeklyMinimum : 0;
      g.bonusPercent = g.weeklyEligible && g.hours > overtimeThreshold ? overtimePercent : 0;
      g.bonus = g.weeklyEligible ? g.basePay * g.bonusPercent / 100 : 0;
      g.totalPay = g.weeklyEligible ? g.basePay + g.bonus : 0;
      items.filter(x => x.employeeId === id).forEach(x => { x.weeklyEligible = g.weeklyEligible; x.pay = g.weeklyEligible ? x.basePay : 0; });
      return g;
    }).sort((a,b) => a.name.localeCompare(b.name));
    const users = table_(SH.USERS).rows.filter(u => norm_(u.EmployeeID) !== 'DEV001').map(u => ({
      employeeId: str_(u.EmployeeID), name: str_(u.Name), rank: str_(u.Rank), hourlyRate: Number(rates[norm_(u.EmployeeID)]) || 0
    }));
    return { ok: true, period, minHours, dailyEnabled, minWeek, weeklyEnabled, overtimeThreshold, overtimePercent,
      breaks: str_(cfg.BreakIntervals), items, people, users };
  } catch (e) { return { ok: false, message: e.message || String(e) }; }
}

function adminSavePayroll(token, config, rates) {
  const adm = requireAdmin_(token); if (adm.err) return adm.err;
  config = config || {}; rates = rates || [];
  const min = Number(config.minimumHours);
  if (!isFinite(min) || min < 0 || min > 24) return { ok: false, message: 'ชั่วโมงขั้นต่ำต้องอยู่ระหว่าง 0 ถึง 24' };
  const minWeek = Number(config.minimumHoursPerWeek), overtimeThreshold = Number(config.overtimeThreshold), overtimePercent = Number(config.overtimePercent);
  if (!isFinite(minWeek) || minWeek < 0 || minWeek > 168) return { ok:false, message:'ชั่วโมงขั้นต่ำต่อสัปดาห์ต้องอยู่ระหว่าง 0 ถึง 168' };
  if (!isFinite(overtimeThreshold) || overtimeThreshold < 0 || overtimeThreshold > 168) return { ok:false, message:'เกณฑ์โบนัสต้องอยู่ระหว่าง 0 ถึง 168 ชั่วโมง' };
  if (!isFinite(overtimePercent) || overtimePercent < 0 || overtimePercent > 1000) return { ok:false, message:'เปอร์เซ็นต์โบนัสต้องอยู่ระหว่าง 0 ถึง 1000' };
  const breaks = String(config.breaks || '').trim();
  const parts = breaks ? breaks.split(/[;,\n]/).map(x => x.trim()).filter(Boolean) : [];
  if (parts.some(x => !/^([01]\d|2[0-3]):[0-5]\d\s*-\s*([01]\d|2[0-3]):[0-5]\d$/.test(x))) return { ok: false, message: 'รูปแบบเวลาพักไม่ถูกต้อง ใช้ HH:mm-HH:mm คั่นแต่ละช่วงด้วย , เช่น 12:00-13:00,17:00-18:00' };
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const cfg = table_(SH.PAYCFG);
    [['BreakIntervals', parts.join(',')], ['MinimumHoursPerDay', min], ['DailyMinimumEnabled', !!config.dailyMinimumEnabled],
      ['MinimumHoursPerWeek', minWeek], ['WeeklyMinimumEnabled', !!config.weeklyMinimumEnabled],
      ['OvertimeThresholdHours', overtimeThreshold], ['OvertimeBonusPercent', overtimePercent]].forEach(([k,v]) => {
      const row = cfg.rows.find(r => str_(r.Key) === k);
      if (row) setCell_(cfg, row, 'Value', v); else appendObj_(cfg, { Key:k, Value:v });
    });
    const t = table_(SH.PAYRATES);
    (Array.isArray(rates) ? rates : []).forEach(x => {
      const id = norm_(x.employeeId), rate = Number(x.hourlyRate);
      if (!id || !isFinite(rate) || rate < 0) return;
      const row = t.rows.find(r => norm_(r.EmployeeID) === id);
      if (row) { setCell_(t,row,'HourlyRate',rate); setCell_(t,row,'UpdatedBy',whoOf_(adm.user)); setCell_(t,row,'UpdatedAt',new Date()); }
      else appendObj_(t,{ EmployeeID:id, HourlyRate:rate, UpdatedBy:whoOf_(adm.user), UpdatedAt:new Date() });
    });
    audit_(whoOf_(adm.user),'แก้ไขตั้งค่าเงินเดือน','', 'ขั้นต่ำ '+min+' ชม./วัน, '+minWeek+' ชม./สัปดาห์ · โบนัส >'+overtimeThreshold+' ชม. '+overtimePercent+'% · พัก '+parts.join(','));
    return { ok:true };
  } finally { lock.releaseLock(); }
}

