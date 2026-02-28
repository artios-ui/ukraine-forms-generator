const RESULTS_SHEET_NAME = 'Results';
const ADAPTIVE_LINKS_SHEET_NAME = 'AdaptiveLinks';

const WEAK_TOPIC_THRESHOLD = 70;
const ADAPTIVE_TOTAL_QUESTIONS = 6;
const ADAPTIVE_FORM_TITLE_PREFIX = 'Адаптивний тренажер — ';

let SETTINGS_CACHE = null;

const DEFAULT_FREE_MODELS = [
  'nvidia/llama-nemotron-embed-vl-1b-v2:free',
  'qwen/qwen3-vl-30b-a3b-thinking'
];

function reloadSettings_() { SETTINGS_CACHE = null; }

function loadSettings_() {
  SETTINGS_CACHE = {};
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Settings');
  if (!sheet) return;

  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][0] || '').trim();
    if (key) SETTINGS_CACHE[key] = values[i][1];
  }
}

function getSetting_(key, defaultValue) {
  const scriptVal = PropertiesService.getScriptProperties().getProperty(key);
  if (scriptVal !== null && scriptVal !== undefined && String(scriptVal) !== '') {
    if (typeof defaultValue === 'number') {
      const num = parseFloat(String(scriptVal).replace(',', '.'));
      return isNaN(num) ? defaultValue : num;
    }
    if (typeof defaultValue === 'boolean') {
      const str = String(scriptVal).toLowerCase();
      return str === 'true' || str === '1' || str === 'yes' || str === 'так';
    }
    return scriptVal;
  }

  if (!SETTINGS_CACHE) loadSettings_();

  if (SETTINGS_CACHE && (key in SETTINGS_CACHE) && SETTINGS_CACHE[key] !== '') {
    const val = SETTINGS_CACHE[key];

    if (typeof defaultValue === 'number') {
      const num = parseFloat(String(val).replace(',', '.'));
      return isNaN(num) ? defaultValue : num;
    }

    if (typeof defaultValue === 'boolean') {
      const str = String(val).toLowerCase();
      return str === 'true' || str === '1' || str === 'yes' || str === 'так';
    }

    return val;
  }

  return defaultValue;
}

function parseScore_(v) {
  const s = String(v).trim().replace('%', '').replace(',', '.');
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function normalizeIndex_(idx, len) {
  if (idx === undefined || idx === null) return 0;
  const n = Number(idx);
  if (isNaN(n)) return 0;
  if (n >= 0 && n < len) return n;
  if (n >= 1 && n <= len) return n - 1;
  return 0;
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Інформатика AI')
    .addItem('Створити тест (клас/тема)…', 'createQuizViaPrompts_')
    .addSeparator()
    .addItem('Імпорт результатів (CSV)…', 'importResultsFromDriveCsv_')
    .addItem('Експорт результатів (Аркуш -> Results)', 'exportFromActiveResponseSheet_')
    .addSeparator()
    .addItem('Згенерувати адаптивні тести (всім)', 'generateAdaptiveQuizzesForAll_')
    .addItem('Адаптивний тест (одному)…', 'generateAdaptiveQuizForOnePrompt_')
    .addSeparator()
    .addItem('Розіслати листи', 'sendAdaptiveLinksByEmail_')
    .addToUi();
}

function callOpenRouterJson_(prompt) {
  const keysString = getSetting_('OPENROUTER_API_KEYS', '');
  if (!keysString) throw new Error('Додайте OPENROUTER_API_KEYS в аркуш Settings або Script Properties');

  const apiKeys = keysString.split(',').map(k => k.trim()).filter(Boolean);
  const modelsString = getSetting_('OPENROUTER_MODELS', '');
  const models = modelsString ? modelsString.split(',').map(m => m.trim()).filter(Boolean) : DEFAULT_FREE_MODELS;

  for (let i = apiKeys.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [apiKeys[i], apiKeys[j]] = [apiKeys[j], apiKeys[i]];
  }

  let lastError = null;

  for (const apiKey of apiKeys) {
    for (const model of models) {
      try {
        const json = makeJsonRequest_(apiKey, model, prompt);
        return json;
      } catch (e) {
        lastError = e;
      }
    }
  }

  throw new Error(`Всі моделі відмовили. Остання помилка: ${lastError ? lastError.message : 'Unknown'}`);
}

function makeJsonRequest_(apiKey, model, prompt) {
  const url = 'https://openrouter.ai/api/v1/chat/completions';

  const payload = {
    model: model,
    messages: [
      {
        role: 'system',
        content:
          'You are a professional educational content creator for Ukrainian schools. ' +
          'Strictly follow Ukrainian language rules. Avoid any Chinese, Russian, or English artifacts in the text. ' +
          'Ensure all Multiple Choice options are unique. ' +
          'Output ONLY valid JSON array.'
      },
      { role: 'user', content: prompt }
    ]
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    headers: {
      'Authorization': 'Bearer ' + apiKey,
      'HTTP-Referer': 'https://script.google.com/',
      'X-Title': 'School Quiz JSON Pro'
    },
    payload: JSON.stringify(payload)
  };

  const res = UrlFetchApp.fetch(url, options);
  const code = res.getResponseCode();
  const body = res.getContentText();

  if (code >= 300) throw new Error(`API ${code}: ${body}`);

  const data = JSON.parse(body);
  const content = data && data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content;
  if (!content) throw new Error('Empty content from AI');

  const cleanJson = String(content)
    .replace(/^```json/i, '')
    .replace(/^```/i, '')
    .replace(/```$/i, '')
    .trim();

  try {
    const parsed = JSON.parse(cleanJson);
    if (Array.isArray(parsed)) return parsed;
    if (parsed && typeof parsed === 'object' && Array.isArray(parsed.questions)) return parsed.questions;
    if (parsed && typeof parsed === 'object') return [parsed];
    throw new Error('Response is not an array');
  } catch (e) {
    throw new Error('JSON Parse Error: ' + e.message);
  }
}

function createQuizViaPrompts_() {
  const ui = SpreadsheetApp.getUi();

  const respGrade = ui.prompt('Клас', 'Наприклад: 9 клас', ui.ButtonSet.OK_CANCEL);
  if (respGrade.getSelectedButton() !== ui.Button.OK) return;
  const grade = respGrade.getResponseText();

  const respTopic = ui.prompt('Тема', 'Наприклад: Цикли в Python', ui.ButtonSet.OK_CANCEL);
  if (respTopic.getSelectedButton() !== ui.Button.OK) return;
  const topic = respTopic.getResponseText();

  const respCount = ui.prompt('Кількість питань', 'Рекомендовано: 10-12', ui.ButtonSet.OK_CANCEL);
  if (respCount.getSelectedButton() !== ui.Button.OK) return;
  const count = parseInt(respCount.getResponseText(), 10) || 10;

  ui.alert('Генерація тесту...');

  try {
    const questions = generateJsonQuestions_(grade, topic, count);
    const title = `${grade} — Тест: ${topic}`;
    const form = createFormFromJson_(questions, title, `Тема: ${topic}. Згенеровано автоматично.`);

    ui.alert(
      'Тест успішно створено!',
      `Редагувати:\n${form.getEditUrl()}\n\nПроходити:\n${form.getPublishedUrl()}`,
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('Помилка генерації: ' + e.message);
  }
}

function generateJsonQuestions_(grade, topic, count) {
  const prompt = `
Ти — професійний методист з інформатики в Україні. Створи якісний тест українською мовою.
Клас: ${grade}. Тема: "${topic}". Кількість питань: ${count}.

ВИМОГИ:
1. Мова: ЛІТЕРАТУРНА УКРАЇНСЬКА.
2. Варіанти відповідей мають бути УНІКАЛЬНИМИ.
3. Правильна відповідь має бути ОДНОЗНАЧНОЮ.
4. Структура:
   - 70% MC (одна правильна).
   - 20% CHECKBOX (кілька правильних).
   - 10% SHORT (коротка текстова).

ФОРМАТ (JSON Array):
[
  {
    "type": "MC",
    "text": "Питання?",
    "options": ["Варіант 1", "Варіант 2", "Варіант 3", "Варіант 4"],
    "correctIndex": 2,
    "points": 1
  },
  {
    "type": "CHECKBOX",
    "text": "Питання з кількома відповідями...",
    "options": ["Варіант 1", "Варіант 2", "Варіант 3", "Варіант 4"],
    "correctIndices": [2, 4],
    "points": 2
  },
  {
    "type": "SHORT",
    "text": "Коротка відповідь...",
    "points": 1
  }
]
`;
  return callOpenRouterJson_(prompt);
}

function generateAdaptiveQuizForStudent_(email) {
  const weakThreshold = getSetting_('WEAK_TOPIC_THRESHOLD', WEAK_TOPIC_THRESHOLD);
  const totalQuestions = getSetting_('ADAPTIVE_TOTAL_QUESTIONS', ADAPTIVE_TOTAL_QUESTIONS);

  const stats = getStudentsStats_(weakThreshold);
  const st = stats[email];

  if (!st) throw new Error(`Учня ${email} немає в Results`);
  if (!st.weakTopics.length) throw new Error('Немає слабких тем (результат високий).');

  const weakStr = st.weakTopics.map(t => `${t.topic} (${Math.round(t.avg)}%)`).join(', ');

  const prompt = `
Створи адаптивний тест з інформатики (JSON). Мова: Українська.
Слабкі теми учня: ${weakStr}.
Кількість питань: ${totalQuestions}.
Варіанти відповідей НЕ ПОВИННІ ДУБЛЮВАТИСЯ.

JSON формат:
[
  {"type":"MC","text":"...","options":["Варіант 1","Варіант 2","Варіант 3","Варіант 4"],"correctIndex":2,"points":1},
  {"type":"CHECKBOX","text":"...","options":["Варіант 1","Варіант 2","Варіант 3","Варіант 4"],"correctIndices":[2,3],"points":2}
]
`;
  const questions = callOpenRouterJson_(prompt);

  const prefix = getSetting_('ADAPTIVE_FORM_TITLE_PREFIX', ADAPTIVE_FORM_TITLE_PREFIX);
  const title = prefix + (st.name || email);
  const form = createFormFromJson_(questions, title, 'Індивідуальний адаптивний тест.');

  const ss = SpreadsheetApp.getActive();
  const linksSheetName = getSetting_('ADAPTIVE_LINKS_SHEET_NAME', ADAPTIVE_LINKS_SHEET_NAME);
  let sheet = ss.getSheetByName(linksSheetName);
  if (!sheet) {
    sheet = ss.insertSheet(linksSheetName);
    sheet.appendRow(['email', 'name', 'weak_topics', 'form_edit_url', 'form_url', 'sent']);
  }

  sheet.appendRow([st.email, st.name, weakStr, form.getEditUrl(), form.getPublishedUrl(), '']);

  if (getSetting_('SEND_EMAILS_AUTOMATICALLY', false)) {
    GmailApp.sendEmail(st.email, 'Тренувальний тест', `Ваш персональний тест: ${form.getPublishedUrl()}`);
    sheet.getRange(sheet.getLastRow(), 6).setValue(new Date());
  }

  return form;
}

function createFormFromJson_(questions, title, desc) {
  if (!questions || !questions.length) throw new Error('Отримано порожній список питань від AI');

  const form = FormApp.create(title);
  form.setIsQuiz(true);
  form.setCollectEmail(false);
  form.setDescription(desc);

  const uniquifyOptions = (opts) => {
    const seen = new Set();
    return opts.map(opt => {
      let val = String(opt).trim();
      if (!val) val = 'Варіант';
      while (seen.has(val)) val += ' ';
      seen.add(val);
      return val;
    });
  };

  questions.forEach(q => {
    let type = String(q.type || 'MC').toUpperCase();
    const hasMulti = Array.isArray(q.correctIndices) && q.correctIndices.length > 0;
    if (type.includes('CHECKBOX') || type.includes('MULTI') || hasMulti) type = 'CHECKBOX';

    if (type.includes('MC')) {
      const item = form.addMultipleChoiceItem();
      const titleText = q.text || 'Питання';
      item.setTitle(titleText);
      item.setPoints(q.points || 1);

      const rawOpts = (q.options || []).filter(Boolean).map(x => String(x));
      if (rawOpts.length < 2) {
        form.addTextItem().setTitle(titleText + ' (Введіть відповідь)').setPoints(q.points || 1);
        return;
      }

      const opts = uniquifyOptions(rawOpts);
      const correctIdx = normalizeIndex_(q.correctIndex, opts.length);

      const choices = opts.map((optText, i) => item.createChoice(optText, i === correctIdx));
      item.setChoices(choices);
    } else if (type.includes('CHECKBOX')) {
      const item = form.addCheckboxItem();
      const titleText = q.text || 'Оберіть правильні варіанти';
      item.setTitle(titleText);
      item.setPoints(q.points || 2);

      const rawOpts = (q.options || []).filter(Boolean).map(x => String(x));
      if (rawOpts.length < 2) {
        form.addParagraphTextItem().setTitle(titleText).setPoints(q.points || 1);
        return;
      }

      const opts = uniquifyOptions(rawOpts);
      const correctIndices = (q.correctIndices || [])
        .map(x => normalizeIndex_(x, opts.length))
        .filter((v, i, a) => a.indexOf(v) === i);

      const choices = opts.map((optText, i) => item.createChoice(optText, correctIndices.includes(i)));
      item.setChoices(choices);
    } else if (type.includes('SHORT')) {
      form.addTextItem().setTitle(q.text || 'Питання').setPoints(q.points || 1);
    } else {
      form.addParagraphTextItem().setTitle(q.text || 'Питання').setPoints(q.points || 1);
    }
  });

  return form;
}

function generateAdaptiveQuizzesForAll_() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('Починаю генерацію...');

  const weakThreshold = getSetting_('WEAK_TOPIC_THRESHOLD', WEAK_TOPIC_THRESHOLD);
  const stats = getStudentsStats_(weakThreshold);

  let count = 0;

  for (const email in stats) {
    try {
      if (stats[email].weakTopics.length) {
        generateAdaptiveQuizForStudent_(email);
        count++;
      }
    } catch (e) {}
  }

  ui.alert(`Створено тестів: ${count}`);
}

function generateAdaptiveQuizForOnePrompt_() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Email учня', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const email = resp.getResponseText().trim();
  if (!email) return;

  try {
    const form = generateAdaptiveQuizForStudent_(email);
    ui.alert('Готово!', `Редагувати: ${form.getEditUrl()}\nУчню: ${form.getPublishedUrl()}`, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert(e.message);
  }
}

function importResultsFromDriveCsv_() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Посилання на CSV або ID', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const url = resp.getResponseText().trim();
  if (!url) return;

  try {
    let id = url;
    const match = url.match(/[-\w]{25,}/);
    if (match) id = match[0];

    const blob = DriveApp.getFileById(id).getBlob();
    const data = Utilities.parseCsv(blob.getDataAsString());

    const ss = SpreadsheetApp.getActive();
    const resSheetName = getSetting_('RESULTS_SHEET_NAME', RESULTS_SHEET_NAME);
    let sh = ss.getSheetByName(resSheetName);
    if (!sh) sh = ss.insertSheet(resSheetName);
    else sh.clearContents();

    sh.getRange(1, 1, data.length, data[0].length).setValues(data);
    ui.alert('Імпортовано!');
  } catch (e) {
    ui.alert('Помилка: ' + e.message);
  }
}

function getStudentsStats_(customThreshold) {
  const ss = SpreadsheetApp.getActive();
  const resSheetName = getSetting_('RESULTS_SHEET_NAME', RESULTS_SHEET_NAME);
  const sheet = ss.getSheetByName(resSheetName);

  if (!sheet) return {};
  const vals = sheet.getDataRange().getValues();
  if (vals.length < 2) return {};

  const h = vals[0].map(x => String(x).toLowerCase());

  const iEmail = h.findIndex(x => x.includes('email'));
  const iTopic = h.findIndex(x => x.includes('topic'));
  const iScore = h.findIndex(x => x.includes('score') || x.includes('percent'));
  const iName = h.findIndex(x => x.includes('name'));

  if (iEmail < 0) return {};

  const students = {};
  for (let r = 1; r < vals.length; r++) {
    const row = vals[r];
    const email = String(row[iEmail] || '').trim();
    if (!email) continue;

    if (!students[email]) {
      students[email] = { email, name: iName >= 0 ? row[iName] : '', topics: {}, avgTopics: [], weakTopics: [] };
    }

    const topic = iTopic >= 0 ? String(row[iTopic] || 'General') : 'General';
    const score = iScore >= 0 ? parseScore_(row[iScore]) : 0;

    if (!students[email].topics[topic]) students[email].topics[topic] = [];
    students[email].topics[topic].push(score);
  }

  const weakTh = (customThreshold !== undefined) ? customThreshold : getSetting_('WEAK_TOPIC_THRESHOLD', WEAK_TOPIC_THRESHOLD);

  Object.values(students).forEach(st => {
    for (const [t, arr] of Object.entries(st.topics)) {
      const avg = arr.reduce((a, b) => a + b, 0) / arr.length;
      st.avgTopics.push({ topic: t, avg });
      if (avg < weakTh) st.weakTopics.push({ topic: t, avg });
    }
  });

  return students;
}

function exportFromActiveResponseSheet_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  const rT = ui.prompt('Тема', ui.ButtonSet.OK_CANCEL);
  if (rT.getSelectedButton() !== ui.Button.OK) return;
  const topic = rT.getResponseText();

  const rM = ui.prompt('Макс. бал', ui.ButtonSet.OK_CANCEL);
  if (rM.getSelectedButton() !== ui.Button.OK) return;
  const max = parseScore_(rM.getResponseText());
  if (!max) { ui.alert('Некоректний макс. бал'); return; }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;

  const h = data[0].map(x => String(x).toLowerCase());
  const iEmail = h.findIndex(x => x.includes('email') || x.includes('адреса'));
  const iScore = h.findIndex(x => x.includes('score') || x.includes('бал') || x.includes('результат'));
  const iName = h.findIndex(x => x.includes("ім'я") || x.includes('name'));

  if (iEmail < 0 || iScore < 0) { ui.alert('Не знайдено Email або Бал'); return; }

  const resSheetName = getSetting_('RESULTS_SHEET_NAME', RESULTS_SHEET_NAME);
  let resSheet = ss.getSheetByName(resSheetName);
  if (!resSheet) {
    resSheet = ss.insertSheet(resSheetName);
    resSheet.appendRow(['email', 'name', 'topic', 'score_percent']);
  }

  let added = 0;
  data.slice(1).forEach(row => {
    const email = row[iEmail];
    if (!email) return;

    const sc = parseScore_(row[iScore]);
    if (!isNaN(sc)) {
      resSheet.appendRow([email, iName >= 0 ? row[iName] : '', topic, (sc / max) * 100]);
      added++;
    }
  });

  ui.alert(`Експортовано ${added} рядків.`);
}

function sendAdaptiveLinksByEmail_() {
  const ss = SpreadsheetApp.getActive();
  const linksSheetName = getSetting_('ADAPTIVE_LINKS_SHEET_NAME', ADAPTIVE_LINKS_SHEET_NAME);
  const sheet = ss.getSheetByName(linksSheetName);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();

  let sent = 0;
  for (let i = 1; i < data.length; i++) {
    const email = data[i][0];
    const url = data[i][4];
    const wasSent = data[i][5];

    if (!wasSent && email && url) {
      GmailApp.sendEmail(String(email), 'Тренувальний тест', `Ваш тест: ${url}`);
      sheet.getRange(i + 1, 6).setValue(new Date());
      sent++;
    }
  }

  SpreadsheetApp.getUi().alert(`Надіслано: ${sent}`);
}

function showStudentProfilePrompt_() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Email учня:', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() === ui.Button.OK) {
    try {
      showStudentProfile_(resp.getResponseText().trim());
    } catch (e) {
      ui.alert(e.message);
    }
  }
}

function showStudentProfile_(email) {
  const stats = getStudentsStats_();
  const st = stats[email];
  if (!st) throw new Error('Не знайдено.');

  let msg = `Учень: ${st.name || email}\nСлабкі теми:\n`;
  if (st.weakTopics.length) {
    st.weakTopics.forEach(t => { msg += `- ${t.topic} (${Math.round(t.avg)}%)\n`; });
  } else {
    msg += '(відсутні)';
  }

  SpreadsheetApp.getUi().alert(msg);
}