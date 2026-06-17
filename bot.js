require('dotenv').config();
const { Bot, InputFile, Keyboard, InlineKeyboard } = require('grammy');
const { PDFParse } = require('pdf-parse');
const OpenAI = require('openai');
const { createClient } = require('@supabase/supabase-js');
const ExcelJS = require('exceljs');
const { createCanvas } = require('@napi-rs/canvas');

const bot = new Bot(process.env.BOT_TOKEN);
const openai = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });
const db = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_KEY);

const pendingBudget = new Map();
const pendingGoal = new Map();
const pendingTxn = new Map();
const pendingAccName = new Map();

const CATEGORIES = [
  { name: 'Транспорт',   keywords: ['yandex.go','astana lrt','taksofon','taxofon','такси','indrive','avtobys'] },
  { name: 'Связь',       keywords: ['beeline','услуг связи','tele2','activ','kcell'] },
  { name: 'Еда',         keywords: ['magnum','small','galmart','abyroy','ярмарка','coffee','кофейн','cake','beans','mastercoffee','travelers','zebra','coffeeterra','coffee in time','plov','zeytun','zirvak','pampushka','yandex.eda','yandex.delivery','супермаркет','минимаркет','памир','oneshott','ризық','султан','цветная','stolovaya','food city','laim market','jaz duken','saudager'] },
  { name: 'Образование', keywords: ['astana it university','aitu'] },
  { name: 'Покупки',     keywords: ['lc waikiki','kari','ozon','jd group','elle silk','wildberries','salekz','сандык','kaspi magazin','serdtse stolitsy','shabyt'] },
  { name: 'Подписки',    keywords: ['anthropic','claude','cursor','github','google one','yandex.plus','tripo','hoster','openai','netflix','spotify'] },
  { name: 'Развлечения', keywords: ['steam','kinopark','кино','клуб','игр','riot'] },
  { name: 'Аренда',      keywords: ['аренда','квартир','кск','коммунал'] },
];
const BUDGET_CATS = ['Еда', 'Транспорт', 'Покупки', 'Подписки', 'Развлечения', 'Переводы', 'Прочее'];

const mainKb = new Keyboard().text('📊 Отчёт').text('🎯 Цели').row().text('📋 Бюджет').text('🤖 Советы').row().text('❓ Помощь').resized().persistent();

function categorize(desc) { const d = (desc || '').toLowerCase(); for (const c of CATEGORIES) if (c.keywords.some(k => d.includes(k))) return c.name; return 'Прочее'; }
function categorizeOp(desc) { const d = (desc || '').toLowerCase(); if (d.includes('со своего счет') || d.includes('на свой счет')) return 'Свои счета'; if (d.includes('комиссия')) return 'Комиссия'; if (d.startsWith('пополнение') || d.startsWith('поступление')) return 'Доход'; if (d.startsWith('снятие')) return 'Снятие'; if (d.startsWith('перевод')) return 'Переводы'; return categorize(desc); }
function fmt(n) { return Math.abs(n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ' '); }
function toISO(d) { let [dd, mm, yy] = d.split('.'); if (yy.length === 2) yy = '20' + yy; return yy + '-' + mm + '-' + dd; }
function progressBar(pct) { const f = Math.max(0, Math.min(10, Math.round(pct / 10))); return '▰'.repeat(f) + '▱'.repeat(10 - f); }
function dayWord(n) { n = Math.abs(n) % 100; const n1 = n % 10; if (n > 10 && n < 15) return 'дней'; if (n1 === 1) return 'день'; if (n1 >= 2 && n1 <= 4) return 'дня'; return 'дней'; }
function parseKaspi(text) { const re = /^(\d{2}\.\d{2}\.\d{2})\s+([+-])\s*([\d\s]+,\d{2})\s*(?:₸|T|тг)?\s*(.+)$/; const txns = []; for (const line of text.split('\n')) { const m = line.trim().match(re); if (!m) continue; const date = m[1], sign = m[2]; let amount = Math.round(parseFloat(m[3].replace(/\s/g, '').replace(',', '.'))); if (sign === '-') amount = -amount; const desc = m[4].trim(); const category = categorizeOp(desc); if (category === 'Свои счета') continue; txns.push({ date, desc, amount, category }); } return txns; }
function parseHalyk(text) { const lines = text.split('\n').map(l => l.trim()); const startRe = /^\d{2}\.\d{2}\.\d{4}\s+\d{2}\.\d{2}\.\d{4}\s/; const records = []; let cur = null; for (const line of lines) { if (startRe.test(line)) { if (cur) records.push(cur); cur = line; } else if (cur) cur += ' ' + line; } if (cur) records.push(cur); const txns = []; for (const rec of records) { const dm = rec.match(/^(\d{2}\.\d{2}\.\d{4})\s+\d{2}\.\d{2}\.\d{4}\s+(.*)$/); if (!dm) continue; const date = dm[1], rest = dm[2], low = rest.toLowerCase(); if (low.includes('конвертация') || low.includes('автоконверт')) continue; const numRe = /-?\d[\d\s]*,\d{2}/g; const nums = []; let firstIdx = -1, mm; while ((mm = numRe.exec(rest)) !== null) { if (firstIdx < 0) firstIdx = mm.index; nums.push(parseFloat(mm[0].replace(/\s/g, '').replace(',', '.'))); } if (nums.length < 3) continue; const desc = rest.slice(0, firstIdx).trim(); const net = Math.round(nums[nums.length - 3] + nums[nums.length - 2] + nums[nums.length - 1]); if (net === 0) continue; txns.push({ date, desc, amount: net, category: categorizeOp(desc) }); } return txns; }
function parseStatement(text) { if (/народный банк|halyk|hsbkkzkx/i.test(text)) return parseHalyk(text); return parseKaspi(text); }
function buildReport(txns) { const byCat = {}; let income = 0, expense = 0; for (const t of txns) { if (t.amount > 0) { income += t.amount; continue; } byCat[t.category] = (byCat[t.category] || 0) + (-t.amount); expense += (-t.amount); } const EMO = { 'Транспорт':'🚕','Еда':'🍔','Покупки':'🛍','Подписки':'💳','Развлечения':'🎮','Образование':'🎓','Связь':'📱','Аренда':'🏠','Переводы':'💸','Снятие':'🏧','Комиссия':'🏦','Прочее':'📦' }; let msg = '📊 Разбор выписки\nОпераций: ' + txns.length + '\n\nРасходы по категориям:\n'; Object.entries(byCat).sort((a, b) => b[1] - a[1]).forEach(([c, s]) => { msg += (EMO[c] || '•') + ' ' + c + ': -' + fmt(s) + ' ₸\n'; }); msg += '\n💰 Доход: +' + fmt(income) + ' ₸\n🧾 Всего расходов: -' + fmt(expense) + ' ₸\n⚖️ Баланс: ' + (income - expense >= 0 ? '+' : '-') + fmt(income - expense) + ' ₸'; return msg; }
async function askAI(userText) { const today = new Date().toISOString().slice(0, 10); const c = await openai.chat.completions.create({ model: 'gpt-4o-mini', response_format: { type: 'json_object' }, messages: [ { role: 'system', content: 'Разбери операцию. Сегодня ' + today + '. JSON: date,description,amount,category(Транспорт/Еда/Покупки/Подписки/Развлечения/Аренда/Переводы/Доход/Прочее),person. Поле person — имя человека, для кого или о ком трата, если оно явно названо в тексте (например «дочке», «маме», «Айгуль»). Если человек не указан — person пустая строка.' }, { role: 'user', content: userText } ] }); return JSON.parse(c.choices[0].message.content); }
async function saveTransactions(tgId, rows) { const toInsert = rows.map(r => ({ tg_id: tgId, op_date: r.op_date, description: r.description, amount: r.amount, category: r.category, person: r.person || null, account: r.account || null })); const { error } = await db.from('transactions').insert(toInsert); if (error) throw error; }
async function getAccounts(tgId) { const { data } = await db.from('accounts').select('*').eq('tg_id', tgId).order('name', { ascending: true }); return data || []; }
function txnAccKb(accs) { const ik = new InlineKeyboard(); accs.forEach(a => ik.text('💳 ' + a.name, 'txnacc:' + a.id).row()); ik.text('➕ Новый реквизит', 'txnaccnew').row(); ik.text('🗂 Без реквизита (в общий)', 'txnaccnone'); return ik; }
function pieChartPNG(labels, values) { const W = 420, H = 320 + 24 * labels.length + 20; const cv = createCanvas(W, H), ctx = cv.getContext('2d'); ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, W, H); const total = values.reduce((a, b) => a + b, 0) || 1; const colors = ['#4e79a7','#f28e2b','#e15759','#76b7b2','#59a14f','#edc948','#b07aa1']; let start = -Math.PI / 2; const cx = W / 2, cy = 160, r = 130; values.forEach((v, i) => { const a = (v / total) * Math.PI * 2; ctx.beginPath(); ctx.moveTo(cx, cy); ctx.arc(cx, cy, r, start, start + a); ctx.closePath(); ctx.fillStyle = colors[i % colors.length]; ctx.fill(); start += a; }); ctx.font = '15px Arial'; labels.forEach((l, i) => { const y = 320 + i * 24; ctx.fillStyle = colors[i % colors.length]; ctx.fillRect(16, y, 16, 16); ctx.fillStyle = '#222222'; ctx.fillText(l + ' — ' + Math.round(values[i] / total * 100) + '%', 40, y + 13); }); return cv.toBuffer('image/png'); }
function barChartPNG(labels, plan, actual) { const rowH = 46, padTop = 46, padLeft = 140, padRight = 70, W = 580, H = padTop + labels.length * rowH + 34; const cv = createCanvas(W, H), x = cv.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0, 0, W, H); x.fillStyle = '#222'; x.font = 'bold 16px Arial'; x.fillText('План vs Факт', padLeft, 26); const maxV = Math.max(1, ...plan, ...actual), barW = W - padLeft - padRight; x.font = '13px Arial'; labels.forEach((l, i) => { const y = padTop + i * rowH; x.fillStyle = '#222'; x.textAlign = 'right'; x.fillText(l, padLeft - 8, y + 15); x.textAlign = 'left'; x.fillStyle = '#9dc3e6'; x.fillRect(padLeft, y, (plan[i] / maxV) * barW, 14); x.fillStyle = actual[i] > plan[i] ? '#e15759' : '#59a14f'; x.fillRect(padLeft, y + 18, (actual[i] / maxV) * barW, 14); }); const ly = H - 22; x.fillStyle = '#9dc3e6'; x.fillRect(padLeft, ly, 14, 14); x.fillStyle = '#222'; x.fillText('План', padLeft + 20, ly + 12); x.fillStyle = '#59a14f'; x.fillRect(padLeft + 80, ly, 14, 14); x.fillStyle = '#222'; x.fillText('Факт <= плана', padLeft + 100, ly + 12); x.fillStyle = '#e15759'; x.fillRect(padLeft + 230, ly, 14, 14); x.fillStyle = '#222'; x.fillText('Факт > плана', padLeft + 250, ly + 12); return cv.toBuffer('image/png'); }

async function doHelp(ctx) { return ctx.reply('Что я умею:\n📄 PDF-выписка — разбор по категориям\n✍️ трата текстом («кофе 1500»)\n📋 Бюджет — план по категориям\n🎯 Цели — копить к дате\n📊 Отчёт — Excel\n🤖 Советы', { reply_markup: mainKb }); }
async function doAdvice(ctx) { try { const { data } = await db.from('transactions').select('*').eq('tg_id', ctx.from.id); if (!data || !data.length) return ctx.reply('Пока нет данных.', { reply_markup: mainKb }); let income = 0, expense = 0; const byCat = {}; for (const r of data) { const a = Number(r.amount); if (a > 0) income += a; else { byCat[r.category] = (byCat[r.category] || 0) + (-a); expense += -a; } } const summary = Object.entries(byCat).map(([c, s]) => c + ': ' + s).join(', '); await ctx.reply('🤖 Анализирую...'); const cc = await openai.chat.completions.create({ model: 'gpt-4o-mini', messages: [ { role: 'system', content: 'Ты заботливый финпомощник. Дай 3-4 коротких добрых конкретных совета по экономии на русском. Без осуждения. Каждый с эмодзи и новой строки.' }, { role: 'user', content: 'Доход: ' + income + ' тг. Расходы: ' + summary + '. Всего: ' + expense + ' тг.' } ] }); await ctx.reply(cc.choices[0].message.content + '\n\n(общие советы, не финконсультация)', { reply_markup: mainKb }); } catch (err) { console.error(err); await ctx.reply('⚠️ Не смог собрать советы.'); } }
async function doBudgetMenu(ctx) { const { data } = await db.from('budgets').select('*').eq('tg_id', ctx.from.id); let msg = '📋 Бюджеты:\n'; if (!data || !data.length) msg += '(не заданы)\n'; else for (const x of data) msg += '• ' + x.category + ': ' + fmt(Number(x.amount)) + ' ₸/мес\n'; msg += '\nВыбери категорию:'; const ik = new InlineKeyboard(); BUDGET_CATS.forEach((c, i) => { ik.text(c, 'budcat:' + c); if (i % 2 === 1) ik.row(); }); return ctx.reply(msg, { reply_markup: ik }); }
async function doGoals(ctx) { const { data } = await db.from('goals').select('*').eq('tg_id', ctx.from.id).order('deadline', { ascending: true, nullsFirst: false }); let msg = '🎯 Твои цели:\n'; if (!data || !data.length) msg += '(пока нет)\n'; else for (const g of data) { const pct = g.target ? Math.round(Number(g.saved) / Number(g.target) * 100) : 0; msg += '\n• ' + g.name + (g.deadline ? ' (до ' + g.deadline + ')' : '') + '\n' + progressBar(pct) + ' ' + pct + '% — ' + fmt(Number(g.saved)) + '/' + fmt(Number(g.target)) + ' ₸\n'; } msg += '\nВыбери действие:'; const ik = new InlineKeyboard().text('➕ Новая цель', 'goalnew').row(); (data || []).forEach(g => ik.text('💰 Пополнить: ' + g.name, 'goaltop:' + g.id).row()); return ctx.reply(msg, { reply_markup: ik }); }

async function doReport(ctx, account) {
  try {
    let q = db.from('transactions').select('*').eq('tg_id', ctx.from.id);
    if (account) q = q.eq('account', account);
    const { data, error } = await q.order('op_date', { ascending: true });
    if (error) throw error;
    if (!data || data.length === 0) return ctx.reply('Пока нет данных по этому реквизиту.', { reply_markup: mainKb });
    const rows = data.map(r => ({ op_date: r.op_date, description: r.description, amount: Number(r.amount), category: r.category, person: r.person, account: r.account }));
    const { data: budData } = await db.from('budgets').select('*').eq('tg_id', ctx.from.id);
    const budgetByCat = {}; (budData || []).forEach(x => { budgetByCat[x.category] = Number(x.amount); });
    const { data: goalData } = await db.from('goals').select('*').eq('tg_id', ctx.from.id).order('deadline', { ascending: true, nullsFirst: false });
    const MONEY = '#,##0" ₸"'; const HEADFONT = { bold: true, color: { argb: 'FFFFFFFF' } }; const HEADBG = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2F5496' } };
    const income = rows.filter(r => r.amount > 0).reduce((a, r) => a + r.amount, 0);
    const expenseAbs = rows.filter(r => r.amount < 0).reduce((a, r) => a - r.amount, 0);
    const balance = income - expenseAbs; const savings = income > 0 ? balance / income : 0;
    const expByCat = {}; for (const r of rows) if (r.amount < 0) expByCat[r.category] = (expByCat[r.category] || 0) + (-r.amount);
    const byMonth = {}; for (const r of rows) { const mm = r.op_date.slice(0, 7); byMonth[mm] = byMonth[mm] || { inc: 0, exp: 0 }; if (r.amount > 0) byMonth[mm].inc += r.amount; else byMonth[mm].exp += -r.amount; }
    let biggestIdx = -1, biggestVal = 0; rows.forEach((r, i) => { if (r.amount < 0 && -r.amount > biggestVal) { biggestVal = -r.amount; biggestIdx = i; } });
    const months = [...new Set(rows.map(r => r.op_date.slice(0, 7)))].sort(); const lastMonth = months[months.length - 1] || '';
    const actualByCatM = {}; for (const r of rows) if (r.amount < 0 && r.op_date.slice(0, 7) === lastMonth) actualByCatM[r.category] = (actualByCatM[r.category] || 0) + (-r.amount);
    const wb = new ExcelJS.Workbook();
    const d = wb.addWorksheet('Дашборд');
    d.mergeCells('A1:B1'); d.getCell('A1').value = '💰 Финансовый дашборд'; d.getCell('A1').font = { bold: true, size: 16 }; d.getColumn(1).width = 24; d.getColumn(2).width = 20;
    [['Доход', income, MONEY], ['Расходы', -expenseAbs, MONEY], ['Баланс', balance, MONEY], ['% сбережений', savings, '0.0%']].forEach((rw, i) => { const r = d.getRow(i + 3); r.getCell(1).value = rw[0]; r.getCell(1).font = { bold: true }; const c = r.getCell(2); c.value = rw[1]; c.numFmt = rw[2]; });
    const b = wb.addWorksheet('Бюджет');
    [18, 14, 14, 14, 16].forEach((w, i) => { b.getColumn(i + 1).width = w; });
    b.mergeCells('A1:E1'); b.getCell('A1').value = 'Бюджет vs Факт за ' + (lastMonth || '—'); b.getCell('A1').font = { bold: true, size: 14 };
    const bh = b.getRow(3); bh.values = ['Категория', 'План', 'Факт', 'Разница', '% выполнения']; bh.font = HEADFONT; bh.fill = HEADBG;
    const bcats = [...new Set([...Object.keys(budgetByCat), ...Object.keys(actualByCatM)])];
    bcats.forEach((cat, i) => { const plan = budgetByCat[cat] || 0, fact = actualByCatM[cat] || 0, rn = i + 4; const r = b.getRow(rn); r.getCell(1).value = cat; r.getCell(2).value = plan; r.getCell(2).numFmt = MONEY; r.getCell(3).value = fact; r.getCell(3).numFmt = MONEY; r.getCell(4).value = { formula: 'B' + rn + '-C' + rn }; r.getCell(4).numFmt = MONEY; r.getCell(5).value = { formula: 'IF(B' + rn + '=0,0,C' + rn + '/B' + rn + ')' }; r.getCell(5).numFmt = '0.0%'; });
    if (bcats.length) { b.addConditionalFormatting({ ref: 'E4:E' + (bcats.length + 3), rules: [ { type: 'cellIs', operator: 'greaterThan', formulae: ['1'], priority: 1, style: { font: { color: { argb: 'FFC00000' }, bold: true } } }, { type: 'cellIs', operator: 'lessThanOrEqual', formulae: ['1'], priority: 2, style: { font: { color: { argb: 'FF008000' } } } } ] }); const pb = barChartPNG(bcats, bcats.map(c => budgetByCat[c] || 0), bcats.map(c => actualByCatM[c] || 0)); const idb = wb.addImage({ buffer: pb, extension: 'png' }); b.addImage(idb, { tl: { col: 6, row: 2 }, ext: { width: 580, height: 46 + bcats.length * 46 + 34 } }); }
    const ws = wb.addWorksheet('Операции');
    ws.columns = [{ header: 'Дата', key: 'op_date', width: 14 }, { header: 'Описание', key: 'description', width: 34 }, { header: 'Сумма, ₸', key: 'amount', width: 16 }, { header: 'Категория', key: 'category', width: 16 }, { header: 'Реквизит', key: 'account', width: 16 }];
    ws.getRow(1).font = HEADFONT; ws.getRow(1).fill = HEADBG; ws.views = [{ state: 'frozen', ySplit: 1 }];
    rows.forEach((r, i) => { const row = ws.addRow({ op_date: r.op_date, description: r.description, amount: r.amount, category: r.category, account: r.account || '' }); row.getCell(3).numFmt = MONEY; if (i === biggestIdx) { row.eachCell(c => { c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } }; }); row.getCell(2).note = 'Самая крупная трата'; } });
    const last = rows.length + 1; ws.autoFilter = { from: 'A1', to: 'E1' };
    const tr = ws.addRow({ description: 'ИТОГО', amount: { formula: 'SUM(C2:C' + last + ')' } }); tr.font = { bold: true }; tr.getCell(3).numFmt = MONEY;
    ws.addConditionalFormatting({ ref: 'C2:C' + last, rules: [{ type: 'cellIs', operator: 'lessThan', formulae: ['0'], priority: 1, style: { font: { color: { argb: 'FFC00000' } } } }, { type: 'cellIs', operator: 'greaterThan', formulae: ['0'], priority: 2, style: { font: { color: { argb: 'FF008000' } } } }] });
    const cats = Object.keys(expByCat);
    const s = wb.addWorksheet('Сводка');
    s.columns = [{ header: 'Категория', key: 'cat', width: 18 }, { header: 'Сумма, ₸', key: 'sum', width: 16 }, { header: '% от расходов', key: 'pct', width: 16 }];
    s.getRow(1).font = HEADFONT; s.getRow(1).fill = HEADBG;
    cats.forEach(cat => { const r = s.addRow({ cat, sum: expByCat[cat], pct: expenseAbs ? expByCat[cat] / expenseAbs : 0 }); r.getCell(2).numFmt = MONEY; r.getCell(3).numFmt = '0.0%'; });
    if (cats.length) s.addConditionalFormatting({ ref: 'B2:B' + (cats.length + 1), rules: [{ type: 'dataBar', cfvo: [{ type: 'min' }, { type: 'max' }], color: { argb: 'FF638EC6' } }] });
    if (cats.length) { const png = pieChartPNG(cats, cats.map(c => expByCat[c])); const imgId = wb.addImage({ buffer: png, extension: 'png' }); s.addImage(imgId, { tl: { col: 4, row: 1 }, ext: { width: 420, height: 320 + 24 * cats.length + 20 } }); }
    const m = wb.addWorksheet('По месяцам');
    m.columns = [{ header: 'Месяц', key: 'mon', width: 12 }, { header: 'Доход', key: 'inc', width: 16 }, { header: 'Расход', key: 'exp', width: 16 }, { header: 'Баланс', key: 'bal', width: 16 }];
    m.getRow(1).font = HEADFONT; m.getRow(1).fill = HEADBG;
    Object.keys(byMonth).sort().forEach(mon => { const r = m.addRow({ mon, inc: byMonth[mon].inc, exp: -byMonth[mon].exp, bal: byMonth[mon].inc - byMonth[mon].exp }); [2, 3, 4].forEach(ci => r.getCell(ci).numFmt = MONEY); });
    const byPerson = {}; for (const r of rows) if (r.amount < 0 && r.person) byPerson[r.person] = (byPerson[r.person] || 0) + (-r.amount);
    const persons = Object.keys(byPerson);
    if (persons.length) {
      const p = wb.addWorksheet('По людям');
      p.columns = [{ header: 'Человек', key: 'p', width: 20 }, { header: 'Потрачено, ₸', key: 'sum', width: 18 }, { header: '% от расходов', key: 'pct', width: 16 }];
      p.getRow(1).font = HEADFONT; p.getRow(1).fill = HEADBG;
      persons.sort((a, c) => byPerson[c] - byPerson[a]).forEach(pn => { const r = p.addRow({ p: pn, sum: byPerson[pn], pct: expenseAbs ? byPerson[pn] / expenseAbs : 0 }); r.getCell(2).numFmt = MONEY; r.getCell(3).numFmt = '0.0%'; });
      p.addConditionalFormatting({ ref: 'B2:B' + (persons.length + 1), rules: [{ type: 'dataBar', cfvo: [{ type: 'min' }, { type: 'max' }], color: { argb: 'FFE15759' } }] });
    }
    if (!account) {
      const byAcc = {}; for (const r of rows) if (r.amount < 0) byAcc[r.account || 'Без реквизита'] = (byAcc[r.account || 'Без реквизита'] || 0) + (-r.amount);
      const accs = Object.keys(byAcc);
      if (accs.length > 1) {
        const ap = wb.addWorksheet('По реквизитам');
        ap.columns = [{ header: 'Реквизит', key: 'a', width: 22 }, { header: 'Потрачено, ₸', key: 'sum', width: 18 }, { header: '% от расходов', key: 'pct', width: 16 }];
        ap.getRow(1).font = HEADFONT; ap.getRow(1).fill = HEADBG;
        accs.sort((a, c) => byAcc[c] - byAcc[a]).forEach(an => { const r = ap.addRow({ a: an, sum: byAcc[an], pct: expenseAbs ? byAcc[an] / expenseAbs : 0 }); r.getCell(2).numFmt = MONEY; r.getCell(3).numFmt = '0.0%'; });
        ap.addConditionalFormatting({ ref: 'B2:B' + (accs.length + 1), rules: [{ type: 'dataBar', cfvo: [{ type: 'min' }, { type: 'max' }], color: { argb: 'FF8C6BB1' } }] });
      }
    }
    if (goalData && goalData.length) {
      const g = wb.addWorksheet('Цели');
      [26, 16, 16, 12, 14, 14].forEach((w, i) => { g.getColumn(i + 1).width = w; });
      const gh = g.getRow(1); gh.values = ['Цель', 'Цель, ₸', 'Накоплено, ₸', '% готово', 'Дедлайн', 'Осталось дней']; gh.font = HEADFONT; gh.fill = HEADBG;
      const today = new Date();
      goalData.forEach((xg, i) => { const rn = i + 2; const r = g.getRow(rn); r.getCell(1).value = xg.name; r.getCell(2).value = Number(xg.target); r.getCell(2).numFmt = MONEY; r.getCell(3).value = Number(xg.saved); r.getCell(3).numFmt = MONEY; r.getCell(4).value = { formula: 'IF(B' + rn + '=0,0,C' + rn + '/B' + rn + ')' }; r.getCell(4).numFmt = '0%'; r.getCell(5).value = xg.deadline || ''; if (xg.deadline) r.getCell(6).value = Math.ceil((new Date(xg.deadline) - today) / 86400000); });
      g.addConditionalFormatting({ ref: 'D2:D' + (goalData.length + 1), rules: [{ type: 'dataBar', cfvo: [{ type: 'num', value: 0 }, { type: 'num', value: 1 }], color: { argb: 'FF59A14F' } }] });
    }
    const ab = await wb.xlsx.writeBuffer();
    await ctx.replyWithDocument(new InputFile(Buffer.from(ab), 'report.xlsx'), { caption: '📊 Отчёт — ' + (account || 'все реквизиты'), reply_markup: mainKb });
  } catch (err) { console.error(err); await ctx.reply('⚠️ Не смог собрать отчёт.'); }
}

async function doReportMenu(ctx) {
  const accs = await getAccounts(ctx.from.id);
  const { count } = await db.from('transactions').select('*', { count: 'exact', head: true }).eq('tg_id', ctx.from.id);
  if (!count) return ctx.reply('Пока нет данных. Пришли выписку или напиши трату.', { reply_markup: mainKb });
  if (!accs.length) return doReport(ctx, null);
  const ik = new InlineKeyboard().text('📊 Общий (все реквизиты)', 'report:all').row();
  accs.forEach(a => ik.text('💳 ' + a.name, 'report:' + a.id).row());
  return ctx.reply('📊 Выбери, по какому реквизиту собрать отчёт:', { reply_markup: ik });
}

async function checkReminders() {
  try {
    const today = new Date(); const todayStr = today.toISOString().slice(0, 10);
    const { data } = await db.from('goals').select('*').not('deadline', 'is', null);
    if (!data) return;
    for (const g of data) {
      if (Number(g.saved) >= Number(g.target)) continue;
      const daysLeft = Math.ceil((new Date(g.deadline) - today) / 86400000);
      if (![7, 3, 1, 0].includes(daysLeft)) continue;
      if (g.notified_on === todayStr) continue;
      const pct = g.target ? Math.round(Number(g.saved) / Number(g.target) * 100) : 0;
      const left = Number(g.target) - Number(g.saved);
      const whenTxt = daysLeft === 0 ? 'сегодня дедлайн' : 'осталось ' + daysLeft + ' ' + dayWord(daysLeft);
      const msg = '⏰ Напоминание по цели «' + g.name + '»\n' + whenTxt + '!\n' + progressBar(pct) + ' ' + pct + '%\nНакоплено ' + fmt(Number(g.saved)) + '/' + fmt(Number(g.target)) + ' ₸ (осталось ' + fmt(left) + ' ₸)';
      try { await bot.api.sendMessage(g.tg_id, msg); await db.from('goals').update({ notified_on: todayStr }).eq('id', g.id); } catch (e) { console.error('reminder send fail', e.message); }
    }
  } catch (e) { console.error('checkReminders', e); }
}

bot.command('start', (ctx) => ctx.reply('Привет! Помогу с деньгами 💰\n\nПришли PDF-выписку 📄 или напиши трату («кофе 1500»). Кнопки внизу 👇', { reply_markup: mainKb }));
bot.command('report', (ctx) => doReportMenu(ctx));
bot.command('advice', (ctx) => doAdvice(ctx));
bot.command('budget', (ctx) => doBudgetMenu(ctx));
bot.command('goals', (ctx) => doGoals(ctx));

bot.callbackQuery(/^budcat:(.+)$/, async (ctx) => { pendingBudget.set(ctx.from.id, ctx.match[1]); await ctx.answerCallbackQuery(); await ctx.reply('Напиши сумму для «' + ctx.match[1] + '» в тенге:'); });
bot.callbackQuery('goalnew', async (ctx) => { pendingGoal.set(ctx.from.id, { mode: 'new' }); await ctx.answerCallbackQuery(); await ctx.reply('Напиши цель: Название Сумма Дата\nНапример: День рождения мамы 50000 12.07.2026\n(дату можно не указывать)'); });
bot.callbackQuery(/^goaltop:(\d+)$/, async (ctx) => { pendingGoal.set(ctx.from.id, { mode: 'top', id: Number(ctx.match[1]) }); await ctx.answerCallbackQuery(); await ctx.reply('Сколько добавить к цели? Напиши сумму в тенге:'); });

bot.callbackQuery('report:all', async (ctx) => { await ctx.answerCallbackQuery(); await ctx.reply('⏳ Собираю отчёт...'); return doReport(ctx, null); });
bot.callbackQuery(/^report:(\d+)$/, async (ctx) => { const id = Number(ctx.match[1]); const accs = await getAccounts(ctx.from.id); const acc = accs.find(a => a.id === id); await ctx.answerCallbackQuery(); await ctx.reply('⏳ Собираю отчёт...'); return doReport(ctx, acc ? acc.name : null); });
bot.callbackQuery(/^txnacc:(\d+)$/, async (ctx) => { const id = Number(ctx.match[1]); const t = pendingTxn.get(ctx.from.id); if (!t) { await ctx.answerCallbackQuery('Операция уже записана'); return; } pendingTxn.delete(ctx.from.id); const accs = await getAccounts(ctx.from.id); const acc = accs.find(a => a.id === id); await ctx.answerCallbackQuery(); await saveTransactions(ctx.from.id, [{ ...t, account: acc ? acc.name : null }]); await ctx.reply('Записал ✅' + (acc ? ' → реквизит «' + acc.name + '»' : ''), { reply_markup: mainKb }); });
bot.callbackQuery('txnaccnone', async (ctx) => { const t = pendingTxn.get(ctx.from.id); if (!t) { await ctx.answerCallbackQuery('Операция уже записана'); return; } pendingTxn.delete(ctx.from.id); await ctx.answerCallbackQuery(); await saveTransactions(ctx.from.id, [{ ...t, account: null }]); await ctx.reply('Записал ✅ (в общий)', { reply_markup: mainKb }); });
bot.callbackQuery('txnaccnew', async (ctx) => { if (!pendingTxn.has(ctx.from.id)) { await ctx.answerCallbackQuery('Операция уже записана'); return; } pendingAccName.set(ctx.from.id, true); await ctx.answerCallbackQuery(); await ctx.reply('Напиши название нового реквизита (например: «Каспи Голд» или «Зарплата Halyk»):'); });

bot.on('message:document', async (ctx) => {
  const doc = ctx.message.document;
  if (!doc.file_name || !doc.file_name.toLowerCase().endsWith('.pdf')) return ctx.reply('⚠️ Пришли именно PDF-файл.');
  try {
    await ctx.reply('⏳ Разбираю выписку...');
    const file = await ctx.getFile(); const url = `https://api.telegram.org/file/bot${process.env.BOT_TOKEN}/${file.file_path}`; const res = await fetch(url); const buffer = Buffer.from(await res.arrayBuffer());
    const parser = new PDFParse({ data: buffer }); const result = await parser.getText(); await parser.destroy();
    const txns = parseStatement(result.text);
    if (txns.length === 0) return ctx.reply('Не нашёл операций.');
    const bank = /народный банк|halyk|hsbkkzkx/i.test(result.text) ? 'Halyk' : 'Kaspi';
    await db.from('accounts').upsert({ tg_id: ctx.from.id, name: bank }, { onConflict: 'tg_id,name' });
    const rows = txns.map(t => ({ op_date: toISO(t.date), description: t.desc, amount: t.amount, category: t.category, account: bank }));
    await saveTransactions(ctx.from.id, rows);
    await ctx.reply(buildReport(txns) + '\n\n💾 Сохранил в реквизит «' + bank + '». Жми 📊 Отчёт.', { reply_markup: mainKb });
  } catch (err) { console.error(err); await ctx.reply('⚠️ Не смог обработать файл.'); }
});

bot.on('message:text', async (ctx) => {
  const text = ctx.message.text;
  if (text.startsWith('/')) return;
  if (pendingAccName.has(ctx.from.id)) {
    pendingAccName.delete(ctx.from.id);
    const name = text.trim().slice(0, 40);
    const t = pendingTxn.get(ctx.from.id);
    if (!name) return ctx.reply('Пустое название. Введи трату ещё раз.', { reply_markup: mainKb });
    await db.from('accounts').upsert({ tg_id: ctx.from.id, name }, { onConflict: 'tg_id,name' });
    if (t) { pendingTxn.delete(ctx.from.id); await saveTransactions(ctx.from.id, [{ ...t, account: name }]); return ctx.reply('✅ Создал реквизит «' + name + '» и записал туда операцию.', { reply_markup: mainKb }); }
    return ctx.reply('✅ Реквизит «' + name + '» создан.', { reply_markup: mainKb });
  }
  if (pendingBudget.has(ctx.from.id)) {
    const cat = pendingBudget.get(ctx.from.id); pendingBudget.delete(ctx.from.id);
    const amount = parseInt(text.replace(/\D/g, ''), 10);
    if (!amount) return ctx.reply('Не понял сумму. Жми 📋 Бюджет.', { reply_markup: mainKb });
    const { error } = await db.from('budgets').upsert({ tg_id: ctx.from.id, category: cat, amount }, { onConflict: 'tg_id,category' });
    if (error) { console.error(error); return ctx.reply('Не смог сохранить.'); }
    return ctx.reply('✅ Бюджет «' + cat + '» = ' + fmt(amount) + ' ₸/мес', { reply_markup: mainKb });
  }
  if (pendingGoal.has(ctx.from.id)) {
    const p = pendingGoal.get(ctx.from.id); pendingGoal.delete(ctx.from.id);
    if (p.mode === 'new') {
      let str = text.trim(); let deadline = null;
      const dm = str.match(/(\d{2}\.\d{2}\.\d{4})\s*$/); if (dm) { deadline = toISO(dm[1]); str = str.slice(0, dm.index).trim(); }
      const am = str.match(/([\d\s]+)$/); const target = am ? parseInt(am[0].replace(/\D/g, ''), 10) : NaN; const name = am ? str.slice(0, am.index).trim() : str;
      if (!name || !target) return ctx.reply('Не понял. Формат: Название Сумма Дата. Жми 🎯 Цели.', { reply_markup: mainKb });
      const { error } = await db.from('goals').insert({ tg_id: ctx.from.id, name, target, saved: 0, deadline });
      if (error) { console.error(error); return ctx.reply('Не смог сохранить цель.'); }
      return ctx.reply('✅ Цель «' + name + '» на ' + fmt(target) + ' ₸' + (deadline ? ' до ' + deadline : '') + ' создана!', { reply_markup: mainKb });
    } else {
      const amount = parseInt(text.replace(/\D/g, ''), 10);
      if (!amount) return ctx.reply('Не понял сумму.', { reply_markup: mainKb });
      const { data: g } = await db.from('goals').select('saved,name,target').eq('id', p.id).single();
      const newSaved = Number(g.saved) + amount;
      await db.from('goals').update({ saved: newSaved }).eq('id', p.id);
      const pct = g.target ? Math.round(newSaved / Number(g.target) * 100) : 0;
      return ctx.reply('✅ +' + fmt(amount) + ' ₸ к «' + g.name + '»\n' + progressBar(pct) + ' ' + pct + '%', { reply_markup: mainKb });
    }
  }
  if (text === '📊 Отчёт') return doReportMenu(ctx);
  if (text === '🎯 Цели') return doGoals(ctx);
  if (text === '📋 Бюджет') return doBudgetMenu(ctx);
  if (text === '🤖 Советы') return doAdvice(ctx);
  if (text === '❓ Помощь') return doHelp(ctx);
  try {
    await ctx.reply('🤖 Разбираю...');
    const t = await askAI(text);
    pendingTxn.set(ctx.from.id, { op_date: t.date, description: t.description, amount: t.amount, category: t.category, person: t.person });
    const accs = await getAccounts(ctx.from.id);
    const sign = t.amount >= 0 ? '+' : '-';
    const preview = '📝 ' + t.description + '\n💵 ' + sign + fmt(t.amount) + ' ₸\n🏷 ' + t.category + (t.person ? '\n👤 для: ' + t.person : '');
    await ctx.reply('Разобрал:\n' + preview + '\n\nКуда записать?', { reply_markup: txnAccKb(accs) });
  } catch (err) { console.error(err); await ctx.reply('⚠️ Не понял операцию. Например: «кофе 1500».'); }
});

checkReminders();
setInterval(checkReminders, 6 * 60 * 60 * 1000);
bot.start();
console.log('Бот запущен ✅');