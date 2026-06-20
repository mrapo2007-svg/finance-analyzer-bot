require('dotenv').config();
const { Bot, InputFile, Keyboard, InlineKeyboard } = require('grammy');
const { PDFParse } = require('pdf-parse');
const OpenAI = require('openai');
const { toFile } = require('openai');
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
const pendingBulkTxns = new Map();
const pendingBulkAccName = new Map();
const pendingFeedback = new Map();
const pendingTxEdit = new Map();

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
const EDIT_CATS = ['Еда', 'Транспорт', 'Покупки', 'Подписки', 'Развлечения', 'Аренда', 'Связь', 'Образование', 'Переводы', 'Доход', 'Прочее'];

const mainKb = new Keyboard().text('📊 Отчёт').text('🎯 Цели').row().text('📋 Бюджет').text('🤖 Советы').row().text('❓ Помощь').text('💬 Фидбек').resized().persistent();

function categorize(desc) { const d = (desc || '').toLowerCase(); for (const c of CATEGORIES) if (c.keywords.some(k => d.includes(k))) return c.name; return 'Прочее'; }
function categorizeOp(desc) { const d = (desc || '').toLowerCase(); if (d.includes('со своего счет') || d.includes('на свой счет')) return 'Свои счета'; if (d.includes('комиссия')) return 'Комиссия'; if (d.startsWith('пополнение') || d.startsWith('поступление')) return 'Доход'; if (d.startsWith('снятие')) return 'Снятие'; if (d.startsWith('перевод')) return 'Переводы'; return categorize(desc); }
function fmt(n) { return Math.abs(n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ' '); }
function toISO(d) { let [dd, mm, yy] = d.split('.'); if (yy.length === 2) yy = '20' + yy; return yy + '-' + mm + '-' + dd; }
function progressBar(pct) { const f = Math.max(0, Math.min(10, Math.round(pct / 10))); return '▰'.repeat(f) + '▱'.repeat(10 - f); }
function dayWord(n) { n = Math.abs(n) % 100; const n1 = n % 10; if (n > 10 && n < 15) return 'дней'; if (n1 === 1) return 'день'; if (n1 >= 2 && n1 <= 4) return 'дня'; return 'дней'; }
let _rates = null, _ratesAt = 0;
async function getRates() {
  if (_rates && Date.now() - _ratesAt < 3600000) return _rates;
  try {
    const res = await fetch('https://open.er-api.com/v6/latest/USD');
    const data = await res.json();
    if (data && data.result === 'success' && data.rates && data.rates.KZT) { _rates = data.rates; _ratesAt = Date.now(); }
  } catch (e) { console.error('getRates failed:', e.message); }
  return _rates;
}
function convertAmount(amount, from, to, rates) {
  if (!rates || !rates[from] || !rates[to]) return null;
  return amount / rates[from] * rates[to];
}
const CUR_SIGN = { KZT: '₸', USD: '$', EUR: '€', RUB: '₽' };
function parseKaspi(text) { const re = /^(\d{2}\.\d{2}\.\d{2})\s+([+-])\s*([\d\s]+,\d{2})\s*(?:₸|T|тг)?\s*(.+)$/; const txns = []; for (const line of text.split('\n')) { const m = line.trim().match(re); if (!m) continue; const date = m[1], sign = m[2]; let amount = Math.round(parseFloat(m[3].replace(/\s/g, '').replace(',', '.'))); if (sign === '-') amount = -amount; const desc = m[4].trim(); const category = categorizeOp(desc); if (category === 'Свои счета') continue; txns.push({ date, desc, amount, category }); } return txns; }
function parseHalyk(text) { const lines = text.split('\n').map(l => l.trim()); const startRe = /^\d{2}\.\d{2}\.\d{4}\s+\d{2}\.\d{2}\.\d{4}\s/; const records = []; let cur = null; for (const line of lines) { if (startRe.test(line)) { if (cur) records.push(cur); cur = line; } else if (cur) cur += ' ' + line; } if (cur) records.push(cur); const txns = []; for (const rec of records) { const dm = rec.match(/^(\d{2}\.\d{2}\.\d{4})\s+\d{2}\.\d{2}\.\d{4}\s+(.*)$/); if (!dm) continue; const date = dm[1], rest = dm[2], low = rest.toLowerCase(); if (low.includes('конвертация') || low.includes('автоконверт')) continue; const numRe = /-?\d[\d\s]*,\d{2}/g; const nums = []; let firstIdx = -1, mm; while ((mm = numRe.exec(rest)) !== null) { if (firstIdx < 0) firstIdx = mm.index; nums.push(parseFloat(mm[0].replace(/\s/g, '').replace(',', '.'))); } if (nums.length < 3) continue; const desc = rest.slice(0, firstIdx).trim(); const net = Math.round(nums[nums.length - 3] + nums[nums.length - 2] + nums[nums.length - 1]); if (net === 0) continue; txns.push({ date, desc, amount: net, category: categorizeOp(desc) }); } return txns; }
function parseStatement(text) { if (/народный банк|halyk|hsbkkzkx/i.test(text)) return parseHalyk(text); return parseKaspi(text); }
function buildReport(txns) { const byCat = {}; let income = 0, expense = 0; for (const t of txns) { if (t.amount > 0) { income += t.amount; continue; } byCat[t.category] = (byCat[t.category] || 0) + (-t.amount); expense += (-t.amount); } const EMO = { 'Транспорт':'🚕','Еда':'🍔','Покупки':'🛍','Подписки':'💳','Развлечения':'🎮','Образование':'🎓','Связь':'📱','Аренда':'🏠','Переводы':'💸','Снятие':'🏧','Комиссия':'🏦','Прочее':'📦' }; let msg = '📊 Разбор выписки\nОпераций: ' + txns.length + '\n\nРасходы по категориям:\n'; Object.entries(byCat).sort((a, b) => b[1] - a[1]).forEach(([c, s]) => { msg += (EMO[c] || '•') + ' ' + c + ': -' + fmt(s) + ' ₸\n'; }); msg += '\n💰 Доход: +' + fmt(income) + ' ₸\n🧾 Всего расходов: -' + fmt(expense) + ' ₸\n⚖️ Баланс: ' + (income - expense >= 0 ? '+' : '-') + fmt(income - expense) + ' ₸'; return msg; }
async function buildContext(tgId) {
  const [{ data: txns }, { data: buds }, { data: goals }, accs] = await Promise.all([
    db.from('transactions').select('op_date,amount,category,account,person').eq('tg_id', tgId),
    db.from('budgets').select('category,amount').eq('tg_id', tgId),
    db.from('goals').select('name,target,saved,deadline').eq('tg_id', tgId),
    getAccounts(tgId),
  ]);
  const T = txns || [];
  if (!T.length && !(goals || []).length) return 'У пользователя пока нет данных (ни операций, ни целей).';
  let inc = 0, exp = 0; const byCat = {}, byAcc = {}, byMonth = {}, byPerson = {};
  for (const r of T) {
    const a = Number(r.amount);
    if (a > 0) inc += a; else exp += -a;
    if (a < 0) { byCat[r.category] = (byCat[r.category] || 0) + (-a); byAcc[r.account || 'Без реквизита'] = (byAcc[r.account || 'Без реквизита'] || 0) + (-a); if (r.person) byPerson[r.person] = (byPerson[r.person] || 0) + (-a); }
    const mo = (r.op_date || '').slice(0, 7); if (mo) { byMonth[mo] = byMonth[mo] || { inc: 0, exp: 0 }; if (a > 0) byMonth[mo].inc += a; else byMonth[mo].exp += -a; }
  }
  const top = o => Object.entries(o).sort((x, y) => y[1] - x[1]).map(([k, v]) => k + ': ' + Math.round(v) + '₸').join(', ');
  const months = Object.keys(byMonth).sort();
  const monthLines = months.slice(-6).map(mo => mo + ': доход ' + Math.round(byMonth[mo].inc) + '₸, расход ' + Math.round(byMonth[mo].exp) + '₸').join(' | ');
  let s = 'Реквизиты (счета): ' + ((accs || []).map(a => a.name).join(', ') || 'нет') + '.\n';
  s += 'Всего операций: ' + T.length + '. Общий доход: ' + Math.round(inc) + '₸. Общий расход: ' + Math.round(exp) + '₸. Баланс: ' + Math.round(inc - exp) + '₸.\n';
  s += 'Расходы по категориям: ' + (top(byCat) || 'нет') + '.\n';
  s += 'Расходы по реквизитам: ' + (top(byAcc) || 'нет') + '.\n';
  if (Object.keys(byPerson).length) s += 'Расходы по людям: ' + top(byPerson) + '.\n';
  s += 'По месяцам (последние): ' + (monthLines || 'нет') + '.\n';
  if ((buds || []).length) s += 'Бюджеты: ' + buds.map(b => b.category + ' ' + Math.round(b.amount) + '₸/мес').join(', ') + '.\n';
  if ((goals || []).length) s += 'Цели: ' + goals.map(g => g.name + ' ' + Math.round(g.saved) + '/' + Math.round(g.target) + '₸' + (g.deadline ? ' до ' + g.deadline : '')).join('; ') + '.\n';
  const rates = await getRates();
  if (rates && rates.KZT) {
    s += 'Курс сейчас: 1 USD = ' + Math.round(rates.KZT) + '₸';
    if (rates.EUR) s += ', 1 EUR = ' + Math.round(rates.KZT / rates.EUR) + '₸';
    if (rates.RUB) s += ', 1 RUB = ' + (rates.KZT / rates.RUB).toFixed(2) + '₸';
    s += '. Используй этот курс, если спросят про конвертацию или «сколько доллар/рубль/евро в тенге».\n';
  }
  return s;
}
async function assistant(userText, contextText) {
  const today = new Date().toISOString().slice(0, 10);
  const sys = 'Ты — личный финансовый ИИ-ассистент пользователя в Telegram, тёплый и умный, общаешься по-человечески на русском. Сегодня ' + today + '.\n' +
    'Твоя зона — личные финансы этого пользователя: траты, доходы, отчёты, бюджеты, цели, реквизиты (счета). Если просят не про финансы (нарисовать картинку, написать код, общие вопросы) — мягко откажись и предложи помочь с деньгами. На приветствия отвечай дружелюбно и коротко.\n' +
    'У тебя есть данные пользователя (ниже). Когда спрашивают про суммы, отчёт или «сколько потратил/заработал» — посчитай по данным и ответь живым текстом с конкретными цифрами в тенге, кратко и по делу.\n' +
    'Верни ТОЛЬКО JSON. Поле action — одно из: "reply","expense","budget_set","goal_create","account_create","report_file","advice","goals_show","budget_show","delete","tx_list","balance".\n' +
    '- reply: ответь текстом — приветствие, ответ на вопрос про финансы с цифрами, вежливый отказ. Текст положи в поле reply.\n' +
    '- expense: пользователь сообщил трату или доход → date (YYYY-MM-DD, по умолчанию сегодня), description, amount (число; расход отрицательный, доход положительный), currency (валюта суммы: KZT/USD/EUR/RUB; «доллар»→USD, «рубль»→RUB, «евро»→EUR, по умолчанию KZT), category (Транспорт/Еда/Покупки/Подписки/Развлечения/Аренда/Переводы/Доход/Прочее), person (имя или "").\n' +
    '- budget_set: category, amount.\n' +
    '- goal_create: создать цель накопления → target (ЧИСЛО — сумма цели в тенге, сколько накопить), name (короткое название на что копим, БЕЗ суммы и даты; если не сказано — "Цель"), deadline (YYYY-MM-DD или ""). Пример: «накопить 150000 на отпуск к 1 сентября» → name:"Отпуск", target:150000, deadline:"2026-09-01".\n' +
    '- account_create: name.\n' +
    '- delete: пользователь просит что-то удалить → what ("account" — реквизит/отчёт, "goal" — цель, "budget" — бюджет, "all_reports" — все реквизиты сразу), name (название реквизита/цели или категория бюджета; "все"/"" если все или не уточнил).\n' +
    '- report_file: пользователь просит именно ФАЙЛ/Excel-отчёт → account (название реквизита или ""), period (период: "all" — за всё время, "1m" — последний месяц, "3m" — 3 месяца, "6m" — полгода, "12m" — год; по умолчанию "all").\n' +
    '- advice: просит совет по экономии. goals_show: показать цели кнопками. budget_show: показать бюджеты. tx_list: показать последние операции / исправить или поправить трату («покажи последние траты», «исправь покупку», «измени сумму»). balance: «сколько у меня осталось», «мой баланс», «сколько денег», «я в плюсе или минусе».\n' +
    'Действуй уверенно: если это команда (запиши трату, покажи отчёт, удали, поставь бюджет, создай цель/реквизит) — сразу выбирай нужное действие, не переспрашивай. Отказывай (reply) только если просьба вообще не про финансы. Примеры: «запиши кофе 1500»→expense; «покажи отчёт по Kaspi за 3 месяца»→report_file; «удали цель отпуск»→delete; «поставь лимит на еду 80000»→budget_set; «создай реквизит Каспи»→account_create; «сколько потратил за месяц»→reply с цифрами.\n' +
    'Данные пользователя:\n' + contextText;
  const c = await openai.chat.completions.create({ model: 'gpt-4o-mini', response_format: { type: 'json_object' }, messages: [{ role: 'system', content: sys }, { role: 'user', content: userText }] });
  return JSON.parse(c.choices[0].message.content);
}
async function transcribeOnce(buffer, model, prompt) { const file = await toFile(buffer, 'voice.ogg'); const r = await openai.audio.transcriptions.create({ file, model, prompt }); return (r.text || '').trim(); }
async function transcribeVoice(buffer) {
  const prompt = 'Финансовый помощник. Возможные слова: тенге, доллар, рубль, евро, Kaspi, Halyk, реквизит, отчёт, бюджет, цель, потратил, заработал, доход, такси, продукты, кофе, зарплата, удали, покажи, создай, добавь.';
  try { return await transcribeOnce(buffer, 'gpt-4o-mini-transcribe', prompt); }
  catch (e) { console.error('transcribe fallback to whisper-1:', e.message); return await transcribeOnce(buffer, 'whisper-1', prompt); }
}
async function saveTransactions(tgId, rows) { const toInsert = rows.map(r => ({ tg_id: tgId, op_date: r.op_date, description: r.description, amount: r.amount, category: r.category, person: r.person || null, account: r.account || null })); const { error } = await db.from('transactions').insert(toInsert); if (error) throw error; }
async function getAccounts(tgId) { const { data } = await db.from('accounts').select('*').eq('tg_id', tgId).order('name', { ascending: true }); return data || []; }
async function isNewUser(tgId) { const { count } = await db.from('transactions').select('*', { count: 'exact', head: true }).eq('tg_id', tgId); return !count; }
async function upsertAccount(tgId, name) { const { error } = await db.from('accounts').upsert({ tg_id: tgId, name }, { onConflict: 'tg_id,name' }); if (error) { console.error('upsertAccount failed:', error); throw error; } }
function txnAccKb(accs) { const ik = new InlineKeyboard(); accs.forEach(a => ik.text('💳 ' + a.name, 'txnacc:' + a.id).row()); ik.text('➕ Новый реквизит', 'txnaccnew').row(); ik.text('🗂 Без реквизита (в общий)', 'txnaccnone'); return ik; }
function bulkAccKb(accs, bank) { const ik = new InlineKeyboard(); ik.text('📊 Общий «' + bank + '»', 'bulkacc:bank').row(); accs.filter(a => a.name !== bank).forEach(a => ik.text('💳 ' + a.name, 'bulkacc:' + a.id).row()); ik.text('➕ Создать новый отчёт', 'bulkaccnew'); return ik; }
async function saveBulk(ctx, txns, accountName) { try { await upsertAccount(ctx.from.id, accountName); const rows = txns.map(t => ({ op_date: toISO(t.date), description: t.desc, amount: t.amount, category: t.category, account: accountName })); await saveTransactions(ctx.from.id, rows); await ctx.reply(buildReport(txns) + '\n\n💾 Сохранил в отчёт «' + accountName + '». Жми 📊 Отчёт.', { reply_markup: mainKb }); } catch (err) { console.error(err); await ctx.reply('⚠️ Не смог сохранить операции.\n\n' + err.message); } }
function pieChartPNG(labels, values) { const W = 420, H = 320 + 24 * labels.length + 20; const cv = createCanvas(W, H), ctx = cv.getContext('2d'); ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, W, H); const total = values.reduce((a, b) => a + b, 0) || 1; const colors = ['#4e79a7','#f28e2b','#e15759','#76b7b2','#59a14f','#edc948','#b07aa1']; let start = -Math.PI / 2; const cx = W / 2, cy = 160, r = 130; values.forEach((v, i) => { const a = (v / total) * Math.PI * 2; ctx.beginPath(); ctx.moveTo(cx, cy); ctx.arc(cx, cy, r, start, start + a); ctx.closePath(); ctx.fillStyle = colors[i % colors.length]; ctx.fill(); start += a; }); ctx.font = '15px Arial'; labels.forEach((l, i) => { const y = 320 + i * 24; ctx.fillStyle = colors[i % colors.length]; ctx.fillRect(16, y, 16, 16); ctx.fillStyle = '#222222'; ctx.fillText(l + ' — ' + Math.round(values[i] / total * 100) + '%', 40, y + 13); }); return cv.toBuffer('image/png'); }
function barChartPNG(labels, plan, actual) { const rowH = 46, padTop = 46, padLeft = 140, padRight = 70, W = 580, H = padTop + labels.length * rowH + 34; const cv = createCanvas(W, H), x = cv.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0, 0, W, H); x.fillStyle = '#222'; x.font = 'bold 16px Arial'; x.fillText('План vs Факт', padLeft, 26); const maxV = Math.max(1, ...plan, ...actual), barW = W - padLeft - padRight; x.font = '13px Arial'; labels.forEach((l, i) => { const y = padTop + i * rowH; x.fillStyle = '#222'; x.textAlign = 'right'; x.fillText(l, padLeft - 8, y + 15); x.textAlign = 'left'; x.fillStyle = '#9dc3e6'; x.fillRect(padLeft, y, (plan[i] / maxV) * barW, 14); x.fillStyle = actual[i] > plan[i] ? '#e15759' : '#59a14f'; x.fillRect(padLeft, y + 18, (actual[i] / maxV) * barW, 14); }); const ly = H - 22; x.fillStyle = '#9dc3e6'; x.fillRect(padLeft, ly, 14, 14); x.fillStyle = '#222'; x.fillText('План', padLeft + 20, ly + 12); x.fillStyle = '#59a14f'; x.fillRect(padLeft + 80, ly, 14, 14); x.fillStyle = '#222'; x.fillText('Факт <= плана', padLeft + 100, ly + 12); x.fillStyle = '#e15759'; x.fillRect(padLeft + 230, ly, 14, 14); x.fillStyle = '#222'; x.fillText('Факт > плана', padLeft + 250, ly + 12); return cv.toBuffer('image/png'); }

const GUIDE = '📘 Как пользоваться — всё просто:\n\n' +
  '1️⃣  Добавляй траты\n' +
  'Пиши или говори как другу: «такси 800», «продукты 5000 вчера».\n' +
  'Или пришли PDF-выписку из банка 📄 — разберу всё сам.\n\n' +
  '2️⃣  Смотри, куда уходят деньги\n' +
  'Спроси словами: «сколько я потратил за месяц?» — отвечу.\n' +
  'Или жми 📊 Отчёт — пришлю наглядный файл.\n\n' +
  '3️⃣  Ставь цели и бюджеты\n' +
  '🎯 копить к дате  ·  📋 лимиты по категориям.\n\n' +
  '💡 Можно ничего не нажимать — просто скажи, что хочешь.\n' +
  '🔒 Твои данные видишь только ты.';
async function doHelp(ctx) { return ctx.reply(GUIDE, { reply_markup: mainKb }); }
async function doFeedbackPrompt(ctx) {
  pendingFeedback.set(ctx.from.id, true);
  return ctx.reply('💬 Что улучшить или какой фишки не хватает? Напиши одним сообщением — или пришли скриншот 📸 (можно с подписью). Я всё передам разработчику.\n\nНапример: «в отчёте не хватает фильтра по датам» или скрин с проблемой.', { reply_markup: mainKb });
}
async function saveFeedback(ctx, text) {
  const who = (ctx.from.first_name || '') + (ctx.from.username ? ' @' + ctx.from.username : '') + ' (id ' + ctx.from.id + ')';
  if (process.env.ADMIN_ID) {
    try { await bot.api.sendMessage(process.env.ADMIN_ID, '💬 Новый фидбек\nОт: ' + who + '\n\n' + text); }
    catch (e) { console.error('admin notify failed:', e.message); }
  } else {
    console.log('ADMIN_ID не задан — фидбек не отправлен в личку. Текст:', text);
  }
  try { await db.from('feedback').insert({ tg_id: ctx.from.id, username: ctx.from.username || null, name: ctx.from.first_name || null, text: text.slice(0, 2000) }); }
  catch (e) { console.error('feedback save failed:', e.message); }
}
async function doAdvice(ctx) { try { const { data } = await db.from('transactions').select('*').eq('tg_id', ctx.from.id); if (!data || !data.length) return ctx.reply('Пока нет данных.', { reply_markup: mainKb }); let income = 0, expense = 0; const byCat = {}; for (const r of data) { const a = Number(r.amount); if (a > 0) income += a; else { byCat[r.category] = (byCat[r.category] || 0) + (-a); expense += -a; } } const summary = Object.entries(byCat).map(([c, s]) => c + ': ' + s).join(', '); await ctx.reply('🤖 Анализирую...'); const cc = await openai.chat.completions.create({ model: 'gpt-4o-mini', messages: [ { role: 'system', content: 'Ты заботливый финпомощник. Дай 3-4 коротких добрых конкретных совета по экономии на русском. Без осуждения. Каждый с эмодзи и новой строки.' }, { role: 'user', content: 'Доход: ' + income + ' тг. Расходы: ' + summary + '. Всего: ' + expense + ' тг.' } ] }); await ctx.reply(cc.choices[0].message.content + '\n\n(общие советы, не финконсультация)', { reply_markup: mainKb }); } catch (err) { console.error(err); await ctx.reply('⚠️ Не смог собрать советы.'); } }
async function doBudgetMenu(ctx) { const { data } = await db.from('budgets').select('*').eq('tg_id', ctx.from.id); let msg = '📋 Бюджеты:\n'; if (!data || !data.length) msg += '\nПока не заданы. Бюджет — это лимит на категорию в месяц (например, «Еда — 80 000 ₸»), бот покажет, укладываешься ли ты.\n'; else for (const x of data) msg += '• ' + x.category + ': ' + fmt(Number(x.amount)) + ' ₸/мес\n'; msg += '\nВыбери категорию, чтобы задать/изменить:'; const ik = new InlineKeyboard(); BUDGET_CATS.forEach((c, i) => { ik.text(c, 'budcat:' + c); if (i % 2 === 1) ik.row(); }); if (data && data.length) { ik.row(); data.forEach(x => ik.text('🗑 Удалить: ' + x.category, 'buddel:' + x.category).row()); } return ctx.reply(msg, { reply_markup: ik }); }
async function doGoals(ctx) { const { data } = await db.from('goals').select('*').eq('tg_id', ctx.from.id).order('deadline', { ascending: true, nullsFirst: false }); let msg = '🎯 Твои цели:\n'; if (!data || !data.length) msg += '\nПока нет ни одной. Создай первую — например скажи: «накопить 300000 на отпуск к сентябрю», или жми ➕ Новая цель ниже.\n'; else for (const g of data) { const pct = g.target ? Math.round(Number(g.saved) / Number(g.target) * 100) : 0; msg += '\n• ' + g.name + (g.deadline ? ' (до ' + g.deadline + ')' : '') + '\n' + progressBar(pct) + ' ' + pct + '% — ' + fmt(Number(g.saved)) + '/' + fmt(Number(g.target)) + ' ₸\n'; } msg += '\nВыбери действие:'; const ik = new InlineKeyboard().text('➕ Новая цель', 'goalnew').row(); (data || []).forEach(g => { ik.text('💰 Пополнить: ' + g.name, 'goaltop:' + g.id).row(); ik.text('🗑 Удалить: ' + g.name, 'goaldel:' + g.id).row(); }); return ctx.reply(msg, { reply_markup: ik }); }

async function doReport(ctx, account, periodMonths) {
  try {
    let q = db.from('transactions').select('*').eq('tg_id', ctx.from.id);
    if (account) q = q.eq('account', account);
    let periodLabel = '';
    if (periodMonths && periodMonths > 0) {
      const cutoff = new Date(); cutoff.setMonth(cutoff.getMonth() - periodMonths);
      q = q.gte('op_date', cutoff.toISOString().slice(0, 10));
      periodLabel = ' за ' + periodMonths + ' мес.';
    }
    const { data, error } = await q.order('op_date', { ascending: true });
    if (error) throw error;
    if (!data || data.length === 0) return ctx.reply('Пока нет данных за этот период.', { reply_markup: mainKb });
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
    const baseName = (account || 'Общий') + (periodMonths ? ' за ' + periodMonths + ' мес' : '');
    const fileName = baseName.replace(/[\\/:*?"<>|]/g, '_').slice(0, 60) + '.xlsx';
    await ctx.replyWithDocument(new InputFile(Buffer.from(ab), fileName), { caption: '📊 Отчёт — ' + (account || 'все реквизиты') + periodLabel, reply_markup: mainKb });
  } catch (err) { console.error(err); await ctx.reply('⚠️ Не смог собрать отчёт.\n\n' + err.message); }
}

async function doReportMenu(ctx) {
  const accs = await getAccounts(ctx.from.id);
  const { count } = await db.from('transactions').select('*', { count: 'exact', head: true }).eq('tg_id', ctx.from.id);
  if (!count) return ctx.reply('Пока не из чего собрать отчёт 🙂\nПришли PDF-выписку 📄 или запиши пару трат («кофе 1500» / голосом) — и я всё посчитаю и покажу картину.', { reply_markup: mainKb });
  if (!accs.length) return doReport(ctx, null);
  const ik = new InlineKeyboard().text('📊 Общий (все реквизиты)', 'report:all').row();
  accs.forEach(a => ik.text('💳 ' + a.name, 'report:' + a.id).text('🗑', 'accdel:' + a.id).row());
  return ctx.reply('📊 Выбери, по какому реквизиту собрать отчёт (🗑 — удалить реквизит):', { reply_markup: ik });
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

async function buildDigest(tgId) {
  const now = new Date();
  const iso = d => d.toISOString().slice(0, 10);
  const d7 = new Date(now); d7.setDate(d7.getDate() - 7);
  const d14 = new Date(now); d14.setDate(d14.getDate() - 14);
  const { data } = await db.from('transactions').select('amount,category,op_date').eq('tg_id', tgId).gte('op_date', iso(d14));
  if (!data || !data.length) return null;
  let spentWeek = 0, incomeWeek = 0, spentPrev = 0; const byCat = {};
  const cut = iso(d7);
  for (const t of data) {
    const a = Number(t.amount);
    if (t.op_date >= cut) {
      if (a < 0) { spentWeek += -a; byCat[t.category] = (byCat[t.category] || 0) + (-a); }
      else incomeWeek += a;
    } else if (a < 0) spentPrev += -a;
  }
  if (spentWeek === 0 && incomeWeek === 0) return null;
  let msg = '📅 Итоги недели\n\n';
  msg += '💸 Потрачено: ' + fmt(Math.round(spentWeek)) + ' ₸\n';
  if (incomeWeek > 0) msg += '💰 Доход: ' + fmt(Math.round(incomeWeek)) + ' ₸\n';
  const top = Object.entries(byCat).sort((a, b) => b[1] - a[1])[0];
  if (top && spentWeek > 0) msg += '🏷 Больше всего на «' + top[0] + '» — ' + fmt(Math.round(top[1])) + ' ₸ (' + Math.round(top[1] / spentWeek * 100) + '%)\n';
  if (spentPrev > 0) {
    const diff = Math.round((spentWeek - spentPrev) / spentPrev * 100);
    if (diff > 5) msg += '📈 Это на ' + diff + '% больше прошлой недели.\n';
    else if (diff < -5) msg += '📉 Это на ' + Math.abs(diff) + '% меньше прошлой недели — молодец! 👏\n';
    else msg += '➡️ Примерно как на прошлой неделе.\n';
  }
  msg += '\nХочешь подробно — жми 📊 Отчёт.';
  return msg;
}

async function checkDigests() {
  try {
    const hourAstana = (new Date().getUTCHours() + 5) % 24;
    if (hourAstana < 9 || hourAstana >= 22) return; // не будим людей ночью
    const { data: rows } = await db.from('transactions').select('tg_id');
    if (!rows) return;
    const ids = [...new Set(rows.map(r => r.tg_id))];
    const todayStr = new Date().toISOString().slice(0, 10);
    for (const tgId of ids) {
      try {
        const { data: st } = await db.from('user_settings').select('last_digest').eq('tg_id', tgId);
        const last = st && st[0] && st[0].last_digest ? new Date(st[0].last_digest) : null;
        if (last && (Date.now() - last.getTime()) / 86400000 < 7) continue;
        const text = await buildDigest(tgId);
        if (text) { try { await bot.api.sendMessage(tgId, text, { reply_markup: mainKb }); } catch (e) { console.error('digest send fail', tgId, e.message); } }
        await db.from('user_settings').upsert({ tg_id: tgId, last_digest: todayStr }, { onConflict: 'tg_id' });
      } catch (e) { console.error('digest user', tgId, e.message); }
    }
  } catch (e) { console.error('checkDigests', e.message); }
}

async function doRecentTx(ctx) {
  const { data } = await db.from('transactions').select('*').eq('tg_id', ctx.from.id).order('op_date', { ascending: false }).order('id', { ascending: false }).limit(10);
  if (!data || !data.length) return ctx.reply('Пока нет ни одной операции. Запиши трату — скажи «кофе 1500» или пришли выписку 📄.', { reply_markup: mainKb });
  const ik = new InlineKeyboard();
  for (const t of data) {
    const sign = Number(t.amount) >= 0 ? '+' : '−';
    const d = (t.op_date || '').slice(5);
    const desc = (t.description || '').slice(0, 18);
    ik.text(d + ' · ' + desc + ' · ' + sign + fmt(Number(t.amount)) + '₸', 'txedit:' + t.id).row();
  }
  return ctx.reply('🧾 Последние операции — выбери, чтобы исправить:', { reply_markup: ik });
}
bot.command('edit', (ctx) => doRecentTx(ctx));
async function doBalance(ctx) {
  const { data } = await db.from('transactions').select('amount,op_date').eq('tg_id', ctx.from.id);
  if (!data || !data.length) return ctx.reply('Пока нет записей. Добавь траты или доходы — и я посчитаю баланс 💰', { reply_markup: mainKb });
  const ms = new Date(); ms.setDate(1); const msISO = ms.toISOString().slice(0, 10);
  let incM = 0, expM = 0, incA = 0, expA = 0;
  for (const t of data) {
    const a = Number(t.amount);
    if (a >= 0) { incA += a; if (t.op_date >= msISO) incM += a; }
    else { expA += -a; if (t.op_date >= msISO) expM += -a; }
  }
  const netM = Math.round(incM - expM), netA = Math.round(incA - expA);
  let msg = '💰 Баланс (по записям в боте)\n\n';
  msg += '📆 Этот месяц:\n';
  msg += '  ↗️ Доход: ' + fmt(Math.round(incM)) + ' ₸\n';
  msg += '  ↘️ Расход: ' + fmt(Math.round(expM)) + ' ₸\n';
  msg += '  ' + (netM >= 0 ? '✅ Осталось: +' : '⚠️ В минусе: −') + fmt(netM) + ' ₸\n\n';
  msg += '📊 За всё время:\n';
  msg += '  ↗️ Доход: ' + fmt(Math.round(incA)) + ' ₸\n';
  msg += '  ↘️ Расход: ' + fmt(Math.round(expA)) + ' ₸\n';
  msg += '  ' + (netA >= 0 ? '✅ Чистыми: +' : '⚠️ В минусе: −') + fmt(netA) + ' ₸\n\n';
  msg += 'ℹ️ Это по тому, что записано в боте, а не баланс банковской карты.';
  return ctx.reply(msg, { reply_markup: mainKb });
}
bot.command('balance', (ctx) => doBalance(ctx));
bot.callbackQuery(/^txedit:(\d+)$/, async (ctx) => {
  const id = Number(ctx.match[1]);
  const { data: t } = await db.from('transactions').select('*').eq('id', id).single();
  await ctx.answerCallbackQuery();
  if (!t) return ctx.reply('Операция не найдена.', { reply_markup: mainKb });
  const sign = Number(t.amount) >= 0 ? '+' : '−';
  const ik = new InlineKeyboard().text('✏️ Сумма', 'txf:amt:' + id).text('🏷 Категория', 'txf:cat:' + id).row().text('🗑 Удалить', 'txdel1:' + id);
  return ctx.reply('🧾 ' + (t.op_date || '') + ' · ' + (t.description || '') + '\n💵 ' + sign + fmt(Number(t.amount)) + ' ₸\n🏷 ' + (t.category || 'Прочее') + '\n\nЧто исправить?', { reply_markup: ik });
});
bot.callbackQuery(/^txf:amt:(\d+)$/, async (ctx) => { pendingTxEdit.set(ctx.from.id, { id: Number(ctx.match[1]), field: 'amount' }); await ctx.answerCallbackQuery(); await ctx.reply('Напиши новую сумму в тенге.\nПодсказка: «1500» сохранит знак как был, «+1500» — доход, «-1500» — расход.'); });
bot.callbackQuery(/^txf:cat:(\d+)$/, async (ctx) => {
  const id = Number(ctx.match[1]);
  await ctx.answerCallbackQuery();
  const ik = new InlineKeyboard();
  EDIT_CATS.forEach((c, i) => { ik.text(c, 'txsetc:' + id + ':' + c); if (i % 2 === 1) ik.row(); });
  await ctx.reply('Выбери новую категорию:', { reply_markup: ik });
});
bot.callbackQuery(/^txsetc:(\d+):(.+)$/, async (ctx) => {
  const id = Number(ctx.match[1]); const cat = ctx.match[2];
  await db.from('transactions').update({ category: cat }).eq('id', id).eq('tg_id', ctx.from.id);
  await ctx.answerCallbackQuery('Готово');
  await ctx.reply('✅ Категория изменена на «' + cat + '».', { reply_markup: mainKb });
});
bot.callbackQuery(/^txdel1:(\d+)$/, async (ctx) => { const id = Number(ctx.match[1]); await ctx.answerCallbackQuery(); const ik = new InlineKeyboard().text('🗑 Да, удалить', 'txdel1y:' + id).text('Отмена', 'celcancel'); await ctx.reply('Удалить эту операцию?', { reply_markup: ik }); });
bot.callbackQuery(/^txdel1y:(\d+)$/, async (ctx) => { const id = Number(ctx.match[1]); await db.from('transactions').delete().eq('id', id).eq('tg_id', ctx.from.id); await ctx.answerCallbackQuery('Удалено'); await ctx.reply('🗑 Операция удалена.', { reply_markup: mainKb }); });

bot.command('start', async (ctx) => {
  const name = ctx.from.first_name || 'друг';
  let fresh = true;
  try { fresh = await isNewUser(ctx.from.id); } catch (e) { console.error(e); }
  if (!fresh) return ctx.reply('С возвращением, ' + name + '! 👋\nЧем помочь? Скажи трату, пришли выписку или спроси про свои финансы.', { reply_markup: mainKb });
  const ik = new InlineKeyboard().text('✍️ Попробовать на примере', 'tryexample').row().text('❓ Как это работает', 'howto');
  return ctx.reply(
    'Привет, ' + name + '! 👋\n\n' +
    'Я помогаю видеть, куда уходят деньги — простым языком, без таблиц.\n\n' +
    'Со мной можно просто разговаривать. Чтобы добавить трату:\n' +
    '•  напиши:  кофе 1500\n' +
    '•  или скажи голосом 🎤\n' +
    '•  или пришли PDF-выписку 📄\n\n' +
    'Потом спроси «сколько я потратил?» — и я покажу.\n\n' +
    'С чего начнём? 👇',
    { reply_markup: ik });
});
bot.command('help', (ctx) => doHelp(ctx));
bot.command('report', (ctx) => doReportMenu(ctx));
bot.command('advice', (ctx) => doAdvice(ctx));
bot.command('budget', (ctx) => doBudgetMenu(ctx));
bot.command('goals', (ctx) => doGoals(ctx));
bot.command('feedbacks', async (ctx) => {
  const admin = process.env.ADMIN_ID;
  if (!admin) return ctx.reply('ℹ️ Просмотр фидбека пока не настроен.\nТвой Telegram ID: ' + ctx.from.id + '\n\nДобавь в файл .env строку:\nADMIN_ID=' + ctx.from.id + '\nи перезапусти бота — тогда здесь будут видны отзывы.');
  if (String(ctx.from.id) !== String(admin)) return;
  try {
    const { data, error } = await db.from('feedback').select('*').order('created_at', { ascending: false }).limit(20);
    if (error) throw error;
    if (!data || !data.length) return ctx.reply('Пока нет ни одного отзыва 🤷');
    let msg = '💬 Последние отзывы (' + data.length + '):\n';
    for (const f of data) {
      const who = ((f.name || '') + (f.username ? ' @' + f.username : '')).trim() || ('id' + f.tg_id);
      const date = (f.created_at || '').slice(0, 10);
      msg += '\n— ' + date + ' · ' + who + '\n' + (f.text || '').slice(0, 400) + '\n';
    }
    await ctx.reply(msg.slice(0, 4000));
  } catch (err) { console.error(err); await ctx.reply('⚠️ Не смог получить отзывы: ' + err.message); }
});

bot.callbackQuery(/^budcat:(.+)$/, async (ctx) => { pendingBudget.set(ctx.from.id, ctx.match[1]); await ctx.answerCallbackQuery(); await ctx.reply('Напиши сумму для «' + ctx.match[1] + '» в тенге:'); });
bot.callbackQuery('goalnew', async (ctx) => { pendingGoal.set(ctx.from.id, { mode: 'new' }); await ctx.answerCallbackQuery(); await ctx.reply('Напиши цель: Название Сумма Дата\nНапример: День рождения мамы 50000 12.07.2026\n(дату можно не указывать)'); });
bot.callbackQuery(/^goaltop:(\d+)$/, async (ctx) => {
  const id = Number(ctx.match[1]);
  await ctx.answerCallbackQuery();
  const ik = new InlineKeyboard()
    .text('+1 000', 'goaladd:' + id + ':1000').text('+5 000', 'goaladd:' + id + ':5000').row()
    .text('+10 000', 'goaladd:' + id + ':10000').text('+50 000', 'goaladd:' + id + ':50000').row()
    .text('✏️ Другая сумма', 'goalcustom:' + id);
  await ctx.reply('💰 Сколько добавить к цели?', { reply_markup: ik });
});
bot.callbackQuery(/^goaladd:(\d+):(\d+)$/, async (ctx) => {
  const id = Number(ctx.match[1]); const amount = Number(ctx.match[2]);
  await ctx.answerCallbackQuery();
  const { data: g } = await db.from('goals').select('saved,name,target').eq('id', id).single();
  if (!g) return ctx.reply('Цель не найдена.', { reply_markup: mainKb });
  const newSaved = Number(g.saved) + amount;
  await db.from('goals').update({ saved: newSaved }).eq('id', id);
  const pct = g.target ? Math.round(newSaved / Number(g.target) * 100) : 0;
  await ctx.reply('✅ +' + fmt(amount) + ' ₸ к «' + g.name + '»\n' + progressBar(pct) + ' ' + pct + '% — ' + fmt(newSaved) + '/' + fmt(Number(g.target)) + ' ₸', { reply_markup: mainKb });
});
bot.callbackQuery(/^goalcustom:(\d+)$/, async (ctx) => { pendingGoal.set(ctx.from.id, { mode: 'top', id: Number(ctx.match[1]) }); await ctx.answerCallbackQuery(); await ctx.reply('Напиши сумму в тенге, которую добавить к цели:'); });

bot.callbackQuery('celcancel', async (ctx) => { await ctx.answerCallbackQuery('Отменено'); await ctx.reply('Отменено 👌', { reply_markup: mainKb }); });
bot.callbackQuery('howto', async (ctx) => { await ctx.answerCallbackQuery(); await ctx.reply(GUIDE, { reply_markup: mainKb }); });
bot.callbackQuery('tryexample', async (ctx) => { await ctx.answerCallbackQuery(); await ctx.reply('Смотри 👇 я сам разберу фразу «кофе 1500» — тебе останется выбрать, куда записать:'); return handleText(ctx, 'кофе 1500'); });
bot.callbackQuery('delallyes', async (ctx) => { await ctx.answerCallbackQuery(); await db.from('transactions').delete().eq('tg_id', ctx.from.id); await db.from('accounts').delete().eq('tg_id', ctx.from.id); await ctx.reply('🗑 Все реквизиты и операции удалены. Чистый лист 🤍', { reply_markup: mainKb }); });
bot.callbackQuery(/^goaldel:(\d+)$/, async (ctx) => { const id = Number(ctx.match[1]); const { data: g } = await db.from('goals').select('name').eq('id', id).single(); await ctx.answerCallbackQuery(); const ik = new InlineKeyboard().text('🗑 Да, удалить', 'goaldelyes:' + id).text('Отмена', 'celcancel'); await ctx.reply('Удалить цель' + (g ? ' «' + g.name + '»' : '') + '? Это навсегда.', { reply_markup: ik }); });
bot.callbackQuery(/^goaldelyes:(\d+)$/, async (ctx) => { const id = Number(ctx.match[1]); await db.from('goals').delete().eq('id', id).eq('tg_id', ctx.from.id); await ctx.answerCallbackQuery('Удалено'); await ctx.reply('🗑 Цель удалена.', { reply_markup: mainKb }); });
bot.callbackQuery(/^buddel:(.+)$/, async (ctx) => { const cat = ctx.match[1]; await db.from('budgets').delete().eq('tg_id', ctx.from.id).eq('category', cat); await ctx.answerCallbackQuery('Бюджет удалён'); await ctx.reply('🗑 Бюджет «' + cat + '» удалён.', { reply_markup: mainKb }); });
bot.callbackQuery(/^accdel:(\d+)$/, async (ctx) => { const id = Number(ctx.match[1]); const accs = await getAccounts(ctx.from.id); const acc = accs.find(a => a.id === id); await ctx.answerCallbackQuery(); if (!acc) return ctx.reply('Реквизит не найден.', { reply_markup: mainKb }); const ik = new InlineKeyboard().text('🗑 Да, удалить всё', 'accdelyes:' + id).text('Отмена', 'celcancel'); await ctx.reply('Удалить реквизит «' + acc.name + '» и ВСЕ его операции? Это навсегда.', { reply_markup: ik }); });
bot.callbackQuery(/^accdelyes:(\d+)$/, async (ctx) => { const id = Number(ctx.match[1]); const accs = await getAccounts(ctx.from.id); const acc = accs.find(a => a.id === id); await ctx.answerCallbackQuery(); if (!acc) return ctx.reply('Реквизит не найден.', { reply_markup: mainKb }); await db.from('transactions').delete().eq('tg_id', ctx.from.id).eq('account', acc.name); await db.from('accounts').delete().eq('id', id).eq('tg_id', ctx.from.id); await ctx.reply('🗑 Реквизит «' + acc.name + '» и его операции удалены.', { reply_markup: mainKb }); });

bot.callbackQuery('report:all', async (ctx) => { await ctx.answerCallbackQuery(); await ctx.reply('⏳ Собираю отчёт...'); return doReport(ctx, null); });
bot.callbackQuery(/^report:(\d+)$/, async (ctx) => { const id = Number(ctx.match[1]); const accs = await getAccounts(ctx.from.id); const acc = accs.find(a => a.id === id); await ctx.answerCallbackQuery(); await ctx.reply('⏳ Собираю отчёт...'); return doReport(ctx, acc ? acc.name : null); });
bot.callbackQuery(/^txnacc:(\d+)$/, async (ctx) => { const id = Number(ctx.match[1]); const t = pendingTxn.get(ctx.from.id); if (!t) { await ctx.answerCallbackQuery('Операция уже записана'); return; } pendingTxn.delete(ctx.from.id); const accs = await getAccounts(ctx.from.id); const acc = accs.find(a => a.id === id); await ctx.answerCallbackQuery(); await saveTransactions(ctx.from.id, [{ ...t, account: acc ? acc.name : null }]); await ctx.reply('Записал ✅' + (acc ? ' → реквизит «' + acc.name + '»' : ''), { reply_markup: mainKb }); });
bot.callbackQuery('txnaccnone', async (ctx) => { const t = pendingTxn.get(ctx.from.id); if (!t) { await ctx.answerCallbackQuery('Операция уже записана'); return; } pendingTxn.delete(ctx.from.id); await ctx.answerCallbackQuery(); await saveTransactions(ctx.from.id, [{ ...t, account: null }]); await ctx.reply('Записал ✅ (в общий)', { reply_markup: mainKb }); });
bot.callbackQuery('txnaccnew', async (ctx) => { if (!pendingTxn.has(ctx.from.id)) { await ctx.answerCallbackQuery('Операция уже записана'); return; } pendingAccName.set(ctx.from.id, true); await ctx.answerCallbackQuery(); await ctx.reply('Напиши название нового реквизита (например: «Каспи Голд» или «Зарплата Halyk»):'); });

bot.callbackQuery('bulkacc:bank', async (ctx) => {
  const pending = pendingBulkTxns.get(ctx.from.id);
  if (!pending) { await ctx.answerCallbackQuery('Данные устарели, пришли выписку заново'); return; }
  pendingBulkTxns.delete(ctx.from.id);
  await ctx.answerCallbackQuery();
  await saveBulk(ctx, pending.txns, pending.bank);
});
bot.callbackQuery(/^bulkacc:(\d+)$/, async (ctx) => {
  const pending = pendingBulkTxns.get(ctx.from.id);
  if (!pending) { await ctx.answerCallbackQuery('Данные устарели, пришли выписку заново'); return; }
  const id = Number(ctx.match[1]);
  const accs = await getAccounts(ctx.from.id);
  const acc = accs.find(a => a.id === id);
  pendingBulkTxns.delete(ctx.from.id);
  await ctx.answerCallbackQuery();
  await saveBulk(ctx, pending.txns, acc ? acc.name : pending.bank);
});
bot.callbackQuery('bulkaccnew', async (ctx) => {
  if (!pendingBulkTxns.has(ctx.from.id)) { await ctx.answerCallbackQuery('Данные устарели, пришли выписку заново'); return; }
  pendingBulkAccName.set(ctx.from.id, true);
  await ctx.answerCallbackQuery();
  await ctx.reply('Напиши название для нового отчёта (например: «Kaspi — личная карта» или «Halyk зарплата»):');
});

bot.on('message:document', async (ctx) => {
  const doc = ctx.message.document;
  if (!doc.file_name || !doc.file_name.toLowerCase().endsWith('.pdf')) return ctx.reply('⚠️ Пришли именно PDF-файл.');
  try {
    await ctx.reply('⏳ Разбираю выписку...');
    const file = await ctx.getFile(); const url = `https://api.telegram.org/file/bot${process.env.BOT_TOKEN}/${file.file_path}`; const res = await fetch(url); const buffer = Buffer.from(await res.arrayBuffer());
    const parser = new PDFParse({ data: buffer }); const result = await parser.getText(); await parser.destroy();
    const isHalyk = /народный банк|halyk|hsbkkzkx/i.test(result.text);
    let bank = isHalyk ? 'Halyk' : 'Kaspi';
    let txns = isHalyk ? parseHalyk(result.text) : parseKaspi(result.text);
    if (!txns.length) { const alt = isHalyk ? parseKaspi(result.text) : parseHalyk(result.text); if (alt.length) { txns = alt; bank = isHalyk ? 'Kaspi' : 'Halyk'; } }
    if (!txns.length) return ctx.reply('Не нашёл операций в этом файле. Формат выписки может отличаться — попробуй прислать операции текстом или другой файл.', { reply_markup: mainKb });
    pendingBulkTxns.set(ctx.from.id, { txns, bank });
    const accs = await getAccounts(ctx.from.id);
    await ctx.reply('📄 Похоже на ' + bank + '. Нашёл ' + txns.length + ' операций.\n\n💾 Куда сохранить?', { reply_markup: bulkAccKb(accs, bank) });
  } catch (err) { console.error(err); await ctx.reply('⚠️ Не смог обработать файл.\n\n' + err.message); }
});

async function handleText(ctx, text) {
  text = (text || '').trim();
  if (!text || text.startsWith('/')) return;
  const MENU = ['📊 Отчёт', '🎯 Цели', '📋 Бюджет', '🤖 Советы', '❓ Помощь', '💬 Фидбек'];
  if (MENU.includes(text)) { const id = ctx.from.id; pendingBudget.delete(id); pendingGoal.delete(id); pendingTxn.delete(id); pendingAccName.delete(id); pendingBulkTxns.delete(id); pendingBulkAccName.delete(id); pendingFeedback.delete(id); pendingTxEdit.delete(id); }
  if (pendingTxEdit.has(ctx.from.id)) {
    const ed = pendingTxEdit.get(ctx.from.id); pendingTxEdit.delete(ctx.from.id);
    const raw = text.trim();
    const n = parseInt(raw.replace(/[^\d]/g, ''), 10);
    if (isNaN(n) || n <= 0) return ctx.reply('Не понял сумму. Напиши число, например 1500.', { reply_markup: mainKb });
    const { data: t } = await db.from('transactions').select('amount').eq('id', ed.id).single();
    if (!t) return ctx.reply('Операция не найдена.', { reply_markup: mainKb });
    let sign = Number(t.amount) < 0 ? -1 : 1;
    if (raw.startsWith('+')) sign = 1; else if (raw.startsWith('-')) sign = -1;
    const newAmount = sign * n;
    await db.from('transactions').update({ amount: newAmount }).eq('id', ed.id).eq('tg_id', ctx.from.id);
    return ctx.reply('✅ Сумма изменена на ' + (newAmount >= 0 ? '+' : '−') + fmt(newAmount) + ' ₸.', { reply_markup: mainKb });
  }
  if (pendingFeedback.has(ctx.from.id)) { pendingFeedback.delete(ctx.from.id); await saveFeedback(ctx, text); return ctx.reply('Спасибо, записал! 🙏 Твой фидбек поможет сделать бота лучше.', { reply_markup: mainKb }); }
  if (pendingBulkAccName.has(ctx.from.id)) {
    pendingBulkAccName.delete(ctx.from.id);
    const name = text.trim().slice(0, 40);
    const pending = pendingBulkTxns.get(ctx.from.id);
    pendingBulkTxns.delete(ctx.from.id);
    if (!name) return ctx.reply('Пустое название. Пришли выписку ещё раз.', { reply_markup: mainKb });
    if (!pending) return ctx.reply('Данные устарели, пришли выписку ещё раз.', { reply_markup: mainKb });
    return saveBulk(ctx, pending.txns, name);
  }
  if (pendingAccName.has(ctx.from.id)) {
    pendingAccName.delete(ctx.from.id);
    const name = text.trim().slice(0, 40);
    const t = pendingTxn.get(ctx.from.id);
    if (!name) return ctx.reply('Пустое название. Введи трату ещё раз.', { reply_markup: mainKb });
    try { await upsertAccount(ctx.from.id, name); } catch (err) { return ctx.reply('⚠️ Не смог создать реквизит.\n\n' + err.message, { reply_markup: mainKb }); }
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
      const dm = str.match(/(\d{1,2})[.\/](\d{1,2})(?:[.\/](\d{2,4}))?/);
      if (dm && Number(dm[1]) <= 31 && Number(dm[2]) <= 12) {
        const dd = dm[1].padStart(2, '0'), mm = dm[2].padStart(2, '0');
        const yyyy = dm[3] ? (dm[3].length === 2 ? '20' + dm[3] : dm[3]) : String(new Date().getFullYear());
        deadline = yyyy + '-' + mm + '-' + dd;
        str = (str.slice(0, dm.index) + ' ' + str.slice(dm.index + dm[0].length)).trim();
      }
      const nums = (str.match(/\d[\d\s]*\d|\d/g) || []).map(s => parseInt(s.replace(/\s/g, ''), 10)).filter(n => !isNaN(n) && n > 0);
      const target = nums.length ? Math.max.apply(null, nums) : NaN;
      let name = str.replace(/\d[\d\s]*/g, ' ').replace(/\s+/g, ' ').trim();
      if (!name) name = 'Цель';
      if (!target || isNaN(target)) return ctx.reply('Не понял сумму. Формат: Название Сумма Дата (например: Отпуск 300000 01.09). Жми 🎯 Цели.', { reply_markup: mainKb });
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
  if (text === '💬 Фидбек') return doFeedbackPrompt(ctx);
  try {
    await ctx.replyWithChatAction('typing');
    const ctxText = await buildContext(ctx.from.id);
    const r = await assistant(text, ctxText);
    if (r.action === 'reply') return ctx.reply(r.reply || 'Я рядом 🙂 Спрашивай про свои финансы или скажи трату.', { reply_markup: mainKb });
    if (r.action === 'report_file') {
      const monthsMap = { '1m': 1, '3m': 3, '6m': 6, '12m': 12 };
      const months = monthsMap[r.period] || 0;
      if (r.account) { const accs = await getAccounts(ctx.from.id); const acc = accs.find(a => a.name.toLowerCase() === String(r.account).toLowerCase()); if (acc) { await ctx.reply('⏳ Собираю отчёт...'); return doReport(ctx, acc.name, months); } }
      if (months) { await ctx.reply('⏳ Собираю отчёт...'); return doReport(ctx, null, months); }
      return doReportMenu(ctx);
    }
    if (r.action === 'advice') return doAdvice(ctx);
    if (r.action === 'goals_show') return doGoals(ctx);
    if (r.action === 'budget_show') return doBudgetMenu(ctx);
    if (r.action === 'tx_list') return doRecentTx(ctx);
    if (r.action === 'balance') return doBalance(ctx);
    if (r.action === 'delete') {
      const what = r.what || 'account';
      const nm = (r.name || '').toString().toLowerCase().trim();
      const all = !nm || nm === 'все' || nm === 'всё' || nm === 'all';
      if (what === 'goal') {
        const { data } = await db.from('goals').select('id,name').eq('tg_id', ctx.from.id);
        const g = (data || []).find(x => x.name.toLowerCase().includes(nm) || (nm && nm.includes(x.name.toLowerCase())));
        if (g) { const ik = new InlineKeyboard().text('🗑 Да, удалить', 'goaldelyes:' + g.id).text('Отмена', 'celcancel'); return ctx.reply('Удалить цель «' + g.name + '»? Это навсегда.', { reply_markup: ik }); }
        return doGoals(ctx);
      }
      if (what === 'budget') {
        const { data } = await db.from('budgets').select('category').eq('tg_id', ctx.from.id);
        const b = (data || []).find(x => nm.includes(x.category.toLowerCase()) || x.category.toLowerCase().includes(nm));
        if (b) { await db.from('budgets').delete().eq('tg_id', ctx.from.id).eq('category', b.category); return ctx.reply('🗑 Бюджет «' + b.category + '» удалён.', { reply_markup: mainKb }); }
        return doBudgetMenu(ctx);
      }
      const accs = await getAccounts(ctx.from.id);
      if (what === 'all_reports' || (all && what === 'account')) {
        if (!accs.length) return ctx.reply('Реквизитов пока нет.', { reply_markup: mainKb });
        const ik = new InlineKeyboard().text('🗑 Да, удалить всё', 'delallyes').text('Отмена', 'celcancel');
        return ctx.reply('Удалить ВСЕ реквизиты (' + accs.map(a => a.name).join(', ') + ') и все операции? Это навсегда.', { reply_markup: ik });
      }
      const acc = accs.find(a => a.name.toLowerCase().includes(nm) || (nm && nm.includes(a.name.toLowerCase())));
      if (acc) { const ik = new InlineKeyboard().text('🗑 Да, удалить всё', 'accdelyes:' + acc.id).text('Отмена', 'celcancel'); return ctx.reply('Удалить реквизит «' + acc.name + '» и все его операции? Это навсегда.', { reply_markup: ik }); }
      return doReportMenu(ctx);
    }
    if (r.action === 'budget_set') {
      if (BUDGET_CATS.includes(r.category) && r.amount) { await db.from('budgets').upsert({ tg_id: ctx.from.id, category: r.category, amount: Math.round(r.amount) }, { onConflict: 'tg_id,category' }); return ctx.reply('✅ Бюджет «' + r.category + '» = ' + fmt(Math.round(r.amount)) + ' ₸/мес', { reply_markup: mainKb }); }
      return doBudgetMenu(ctx);
    }
    if (r.action === 'goal_create') {
      const target = Math.round(Number(r.target));
      if (target > 0) { const name = (r.name && String(r.name).trim()) || 'Цель'; const deadline = r.deadline && /^\d{4}-\d{2}-\d{2}$/.test(r.deadline) ? r.deadline : null; const { error } = await db.from('goals').insert({ tg_id: ctx.from.id, name, target, saved: 0, deadline }); if (error) { console.error(error); return ctx.reply('Не смог создать цель.', { reply_markup: mainKb }); } return ctx.reply('✅ Цель «' + name + '» на ' + fmt(target) + ' ₸' + (deadline ? ' до ' + deadline : '') + ' создана!', { reply_markup: mainKb }); }
      return ctx.reply('Не понял сумму цели. Скажи, например: «накопить 150000 на отпуск к 1 сентября».', { reply_markup: mainKb });
    }
    if (r.action === 'account_create') {
      if (r.name) { try { await upsertAccount(ctx.from.id, String(r.name).slice(0, 40)); } catch (e) { return ctx.reply('Не смог создать реквизит.', { reply_markup: mainKb }); } return ctx.reply('✅ Реквизит «' + r.name + '» создан. Теперь можешь сохранять в него операции и строить по нему отчёт.', { reply_markup: mainKb }); }
      return ctx.reply('Как назвать реквизит?', { reply_markup: mainKb });
    }
    if (r.action === 'expense' && typeof r.amount === 'number' && r.amount !== 0 && r.description) {
      let amountKzt = Math.round(r.amount);
      let description = r.description;
      const cur = (r.currency || 'KZT').toUpperCase();
      if (cur !== 'KZT') {
        const rates = await getRates();
        const conv = convertAmount(r.amount, cur, 'KZT', rates);
        if (conv) { amountKzt = Math.round(conv); description = r.description + ' (' + Math.abs(Math.round(r.amount)) + (CUR_SIGN[cur] || cur) + ')'; }
        else { description = r.description + ' (' + Math.abs(Math.round(r.amount)) + (CUR_SIGN[cur] || cur) + ', курс недоступен)'; }
      }
      pendingTxn.set(ctx.from.id, { op_date: r.date || new Date().toISOString().slice(0, 10), description, amount: amountKzt, category: r.category || 'Прочее', person: r.person });
      const accs = await getAccounts(ctx.from.id);
      const sign = amountKzt >= 0 ? '+' : '-';
      const curLine = cur !== 'KZT' ? '\n💵 ' + (r.amount >= 0 ? '+' : '-') + fmt(Math.round(r.amount)) + ' ' + (CUR_SIGN[cur] || cur) + ' ≈ ' + sign + fmt(amountKzt) + ' ₸' : '\n💵 ' + sign + fmt(amountKzt) + ' ₸';
      const preview = '📝 ' + r.description + curLine + '\n🏷 ' + (r.category || 'Прочее') + (r.person ? '\n👤 для: ' + r.person : '');
      return ctx.reply('Разобрал:\n' + preview + '\n\nКуда записать?', { reply_markup: txnAccKb(accs) });
    }
    return ctx.reply(r.reply || 'Не совсем понял 🤔 Можешь сказать, например: «сколько я потратил за последний месяц», «покажи отчёт по Kaspi» или «потратил 3000 на такси».', { reply_markup: mainKb });
  } catch (err) { console.error(err); await ctx.reply('⚠️ Что-то пошло не так. Попробуй сказать иначе или жми кнопки.'); }
}

bot.on('message:photo', async (ctx) => {
  const id = ctx.from.id;
  if (!pendingFeedback.has(id)) {
    return ctx.reply('📸 Хочешь оставить отзыв со скрином? Нажми 💬 Фидбек и пришли картинку. А трату можно записать словами или голосом 🙂', { reply_markup: mainKb });
  }
  pendingFeedback.delete(id);
  const caption = (ctx.message.caption || '').trim();
  const who = (ctx.from.first_name || '') + (ctx.from.username ? ' @' + ctx.from.username : '') + ' (id ' + id + ')';
  const photo = ctx.message.photo[ctx.message.photo.length - 1];
  if (process.env.ADMIN_ID) {
    try { await bot.api.sendPhoto(process.env.ADMIN_ID, photo.file_id, { caption: '💬 Фидбек со скрином\nОт: ' + who + (caption ? '\n\n' + caption : '') }); }
    catch (e) { console.error('admin photo notify failed:', e.message); }
  }
  try { await db.from('feedback').insert({ tg_id: id, username: ctx.from.username || null, name: ctx.from.first_name || null, text: (caption || '[скриншот]') + ' 📸' }); }
  catch (e) { console.error('feedback save failed:', e.message); }
  return ctx.reply('Спасибо, скрин получил! 🙏 Очень помогает.', { reply_markup: mainKb });
});
bot.on('message:text', (ctx) => handleText(ctx, ctx.message.text));

bot.on(['message:voice', 'message:audio'], async (ctx) => {
  try {
    await ctx.reply('🎤 Слушаю...');
    const file = await ctx.getFile();
    const url = `https://api.telegram.org/file/bot${process.env.BOT_TOKEN}/${file.file_path}`;
    const res = await fetch(url); const buffer = Buffer.from(await res.arrayBuffer());
    const text = await transcribeVoice(buffer);
    if (!text) return ctx.reply('Не расслышал 🙈 Попробуй сказать ещё раз, например: «потратил 3000 на такси».', { reply_markup: mainKb });
    await ctx.reply('📝 Распознал: «' + text + '»');
    return handleText(ctx, text);
  } catch (err) { console.error(err); await ctx.reply('⚠️ Не смог распознать голос. Попробуй ещё раз или напиши текстом.'); }
});

bot.api.setMyCommands([
  { command: 'start', description: 'Начать / главное меню' },
  { command: 'help', description: 'Как пользоваться' },
  { command: 'report', description: 'Отчёт' },
  { command: 'goals', description: 'Цели' },
  { command: 'budget', description: 'Бюджет' },
  { command: 'advice', description: 'Совет по экономии' },
]).catch(e => console.error('setMyCommands', e.message));

require('http').createServer((req, res) => { res.writeHead(200, { 'Content-Type': 'text/plain' }); res.end('Bot is running'); }).listen(process.env.PORT || 3000, () => console.log('HTTP keep-alive на порту ' + (process.env.PORT || 3000)));

checkReminders();
setInterval(checkReminders, 6 * 60 * 60 * 1000);
setTimeout(checkDigests, 30000);
setInterval(checkDigests, 3 * 60 * 60 * 1000);
bot.start();
console.log('Бот запущен ✅');