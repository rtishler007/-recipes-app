/**
 * אפליקציית מתכונים אישית — צד השרת (Google Apps Script).
 *
 * כאן נמצא כל מה שרץ אצל גוגל ולא בדפדפן:
 * - הצגת דף האפליקציה (doGet)
 * - הקריאה ל-Claude API (המפתח נקרא מ-Script Properties ולא יוצא מהשרת)
 * - קריאה, כתיבה ומחיקה של מתכונים בגיליון
 */

const SPREADSHEET_ID = '12k4phj3tNjjroynrdybJ-EDC6pig5KwPbZQbvqzzCqQ';
const SHEET_NAME = 'recipes';

const MODEL = 'claude-sonnet-5-5';
const API_URL = 'https://api.anthropic.com/v1/messages';
const API_VERSION = '2026-10-21';
// אם Claude מסרב לבקשה מסיבות בטיחות, ה-API מנסה אוטומטית מודל חלופי
const FALLBACK_BETA = 'server-side-fallback-2026-07-01';

// מספרי העמודות בגיליון (A=1 ... K=11)
const COL = {
  ID: 1, NAME: 2, CATEGORY: 3, SERVINGS: 4, TOTAL_TIME: 5, DESCRIPTION: 6,
  INGREDIENTS: 7, STEPS: 8, CREATED: 9, UPDATED: 10, JSON: 11,
};
const NUM_COLS = 11;

const SYSTEM_PROMPT = [
  'אתה עוזר בישול שבונה מתכונים בשיחה עם משתמש דובר עברית.',
  '',
  'בכל תשובה החזר JSON עם שני שדות:',
  '- message: תשובה קצרה וידידותית בעברית (2–4 משפטים). אל תעתיק את המתכון להודעה — הוא מוצג למשתמש בנפרד.',
  '- recipe: המתכון המלא והעדכני (לא רק השינויים), או null אם עדיין אין מתכון (למשל כששאלת שאלת הבהרה). אם כבר יש מתכון והבקשה לא משנה אותו, החזר אותו שוב כמו שהוא.',
  '',
  'העדף להציע טיוטה מהר. שאל שאלת הבהרה רק כשהבקשה באמת עמומה.',
  '',
  'שדות המתכון:',
  '- name; category — אחת מ: ארוחת בוקר, מנה ראשונה, מנה עיקרית, תוספת, סלט, מרק, מאפה, קינוח, משקה, רוטב (או קטגוריה קצרה אחרת אם אף אחת לא מתאימה); servings; totalMinutes — זמן כולל מההתחלה ועד ההגשה, כולל המתנות; description — משפט או שניים.',
  '- ingredients: לכל מצרך id ("i1", "i2", ...), amount (כמות ויחידה, למשל "2 כוסות" או "לפי הטעם") ו-name.',
  '- steps: לכל שלב id ("s1", "s2", ...), title — פועל קצר עם הפרט החשוב, עד כחמש מילים (למשל "לקצוץ דק", "לאפות 25 דק\' ב-180°"); text — ההוראה המלאה; inputs; activeMinutes — זמן עבודה; waitMinutes — זמן המתנה פסיבי (אפייה, התפחה, בישול על אש קטנה בלי השגחה, קירור), אחרת 0.',
  '',
  'כללי מבנה (חובה — מהם נבנית טבלת המתכון):',
  '1. inputs של שלב = המזהים של המצרכים ושל השלבים הקודמים שהשלב משלב או מעבד. שלב מפנה רק לשלבים שמופיעים לפניו.',
  '2. כל מצרך מופיע ב-inputs של שלב אחד בדיוק. מצרך שמשמש בשני מקומות מפוצל לשתי שורות, למשל "חמאה (לבצק)" ו"חמאה (לציפוי)".',
  '3. כל שלב, חוץ מהאחרון, מופיע ב-inputs של שלב מאוחר יותר אחד בדיוק, כך שהכל מתנקז לשלב האחרון.',
  '4. השלב האחרון הוא שלב הסיום וההגשה, ויש לו inputs.',
  '5. שלב הכנה שלא מעבד מצרכים (למשל "לחמם תנור ל-180°") מקבל inputs ריק, ואף שלב לא מפנה אליו.',
  '6. סדר את המצרכים לפי סדר השימוש בהם.',
].join('\n');

const RECIPE_SCHEMA = {
  type: 'object',
  properties: {
    name: { type: 'string' },
    category: { type: 'string' },
    servings: { type: 'integer' },
    totalMinutes: { type: 'integer' },
    description: { type: 'string' },
    ingredients: {
      type: 'array',
      items: {
        type: 'object',
        properties: {
          id: { type: 'string' },
          amount: { type: 'string' },
          name: { type: 'string' },
        },
        required: ['id', 'amount', 'name'],
        additionalProperties: false,
      },
    },
    steps: {
      type: 'array',
      items: {
        type: 'object',
        properties: {
          id: { type: 'string' },
          title: { type: 'string' },
          text: { type: 'string' },
          inputs: { type: 'array', items: { type: 'string' } },
          activeMinutes: { type: 'integer' },
          waitMinutes: { type: 'integer' },
        },
        required: ['id', 'title', 'text', 'inputs', 'activeMinutes', 'waitMinutes'],
        additionalProperties: false,
      },
    },
  },
  required: ['name', 'category', 'servings', 'totalMinutes', 'description', 'ingredients', 'steps'],
  additionalProperties: false,
};

const REPLY_SCHEMA = {
  type: 'object',
  properties: {
    message: { type: 'string' },
    recipe: { anyOf: [RECIPE_SCHEMA, { type: 'null' }] },
  },
  required: ['message', 'recipe'],
  additionalProperties: false,
};

// ---------- הדף ----------

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('המתכונים שלי')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/** מכניס קובץ HTML אחר לתוך index.html */
function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

// ---------- צ'אט ----------

/**
 * מקבל את כל השיחה עד עכשיו (בפורמט של Claude API) ומחזיר:
 * newMessages — ההודעות שיש להוסיף להיסטוריה (כולל תיקון אוטומטי אם היה)
 * message — הטקסט להציג למשתמש
 * recipe — המתכון העדכני, או null
 * problems — בעיות מבנה שנשארו גם אחרי ניסיון התיקון
 */
function chat(history) {
  if (!Array.isArray(history) || history.length === 0) {
    throw new Error('השיחה ריקה.');
  }
  const messages = history.slice();
  const first = callClaude_(messages);
  const newMessages = [first.assistant];
  let recipe = first.reply.recipe;
  let problems = recipe ? validateRecipe(recipe) : [];

  // אם המבנה לא תקין, מבקשים מ-Claude לתקן פעם אחת — בלי שהמשתמש יצטרך
  if (problems.length) {
    const fix = {
      role: 'user',
      content: '[בדיקה אוטומטית של המערכת, לא הודעה מהמשתמש] במתכון שהחזרת יש בעיות מבנה:\n- ' +
        problems.join('\n- ') +
        '\nהחזר את אותו מתכון, מתוקן לפי כללי המבנה. בשדה message כתוב משפט אחד קצר.',
    };
    messages.push(first.assistant, fix);
    try {
      const second = callClaude_(messages);
      newMessages.push(fix, second.assistant);
      if (second.reply.recipe) {
        recipe = second.reply.recipe;
        problems = validateRecipe(recipe);
      }
    } catch (e) {
      // התיקון נכשל — מחזירים את התשובה הראשונה, והבעיות יוצגו למשתמש
      console.warn('Repair call failed: ' + e.message);
    }
  }

  return { newMessages: newMessages, message: first.reply.message, recipe: recipe, problems: problems };
}

function callClaude_(messages) {
  const body = {
    model: MODEL,
    max_tokens: 16000,
    system: SYSTEM_PROMPT,
    messages: messages,
    output_config: {
      effort: 'low',
      format: { type: 'json_schema', schema: REPLY_SCHEMA },
    },
    fallbacks: 'default',
  };
  const data = postToClaude_(body);

  if (data.stop_reason === 'refusal') {
    throw new Error('Claude סירב לבקשה הזאת. נסה לנסח אותה אחרת.');
  }
  if (data.stop_reason === 'max_tokens') {
    throw new Error('התשובה של Claude נקטעה כי הייתה ארוכה מדי. נסה לבקש מתכון קצר יותר.');
  }

  // אם היה מעבר למודל חלופי, הטקסט הרלוונטי הוא מה שבא אחרי בלוק ה-fallback האחרון
  const content = data.content || [];
  let start = 0;
  content.forEach(function (block, i) { if (block.type === 'fallback') start = i + 1; });
  const text = content.slice(start)
    .filter(function (block) { return block.type === 'text'; })
    .map(function (block) { return block.text; })
    .join('');

  let reply;
  try {
    reply = JSON.parse(text);
  } catch (e) {
    throw new Error('התשובה של Claude לא הגיעה בפורמט הצפוי. נסה לשלוח שוב.');
  }
  // את תשובת Claude שומרים בדיוק כמו שהגיעה, כולל בלוקי thinking — ה-API דורש להחזיר אותם בלי שינוי
  return { assistant: { role: 'assistant', content: content }, reply: reply };
}

function postToClaude_(body) {
  const key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!key) {
    throw new Error('חסר מפתח API. הוסף ANTHROPIC_API_KEY ב-Project Settings ← Script Properties (ראה README).');
  }
  let res;
  try {
    res = UrlFetchApp.fetch(API_URL, {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'x-api-key': key,
        'anthropic-version': API_VERSION,
        'anthropic-beta': FALLBACK_BETA,
      },
      payload: JSON.stringify(body),
      muteHttpExceptions: true,
    });
  } catch (e) {
    throw new Error('לא הצלחתי להגיע ל-Claude (ייתכן שהבקשה לקחה יותר מדי זמן). נסה שוב.');
  }

  const code = res.getResponseCode();
  let data = {};
  try { data = JSON.parse(res.getContentText()); } catch (e) { /* תשובה שאינה JSON */ }
  if (code === 200) return data;

  const msg = (data.error && data.error.message) || '';
  console.error('Claude API error ' + code + ': ' + res.getContentText());
  if (code === 401) throw new Error('מפתח ה-API לא תקין. בדוק את ANTHROPIC_API_KEY ב-Script Properties.');
  if (/credit balance/i.test(msg)) throw new Error('אין מספיק קרדיט בחשבון Anthropic. הוסף קרדיט ב-Billing בקונסולה.');
  if (code === 429) throw new Error('יותר מדי בקשות ברגע זה. חכה דקה ונסה שוב.');
  if (code >= 500) throw new Error('השרתים של Claude עמוסים כרגע. נסה שוב בעוד דקה.');
  throw new Error('שגיאה מ-Claude (' + code + '): ' + msg);
}

/**
 * בודק שהמתכון בנוי כעץ: כל מצרך נכנס לשלב אחד, כל שלב (חוץ מהאחרון) נכנס לשלב אחד,
 * והכל מתנקז לשלב האחרון. מחזיר רשימת בעיות (ריקה = תקין).
 */
function validateRecipe(recipe) {
  const problems = [];
  const ingredients = recipe.ingredients || [];
  const steps = recipe.steps || [];
  if (!ingredients.length) problems.push('אין מצרכים.');
  if (!steps.length) {
    problems.push('אין שלבים.');
    return problems;
  }

  const kind = {};
  const position = {};
  ingredients.forEach(function (ing) {
    if (kind[ing.id]) problems.push('המזהה ' + ing.id + ' מופיע יותר מפעם אחת.');
    kind[ing.id] = 'ingredient';
  });
  steps.forEach(function (step, i) {
    if (kind[step.id]) problems.push('המזהה ' + step.id + ' מופיע יותר מפעם אחת.');
    kind[step.id] = 'step';
    position[step.id] = i;
  });

  const uses = {};
  steps.forEach(function (step, i) {
    (step.inputs || []).forEach(function (id) {
      if (!kind[id]) {
        problems.push('שלב ' + step.id + ' מפנה למזהה שלא קיים: ' + id + '.');
      } else if (kind[id] === 'step' && position[id] >= i) {
        problems.push('שלב ' + step.id + ' מפנה לשלב ' + id + ' שלא מופיע לפניו.');
      }
      uses[id] = (uses[id] || 0) + 1;
    });
  });

  ingredients.forEach(function (ing) {
    const n = uses[ing.id] || 0;
    if (n === 0) problems.push('המצרך "' + ing.name + '" (' + ing.id + ') לא משמש באף שלב.');
    if (n > 1) problems.push('המצרך "' + ing.name + '" (' + ing.id + ') משמש ב-' + n + ' שלבים — יש לפצל אותו לשורות נפרדות.');
  });

  const last = steps[steps.length - 1];
  if (!(last.inputs || []).length) problems.push('לשלב האחרון (' + last.id + ') חייבים להיות inputs.');
  steps.slice(0, -1).forEach(function (step) {
    const n = uses[step.id] || 0;
    if (!(step.inputs || []).length) {
      if (n > 0) problems.push('שלב ההכנה ' + step.id + ' (בלי inputs) לא יכול לשמש קלט של שלב אחר.');
    } else if (n !== 1) {
      problems.push('שלב ' + step.id + ' משמש כקלט ' + n + ' פעמים, ולא פעם אחת בדיוק.');
    }
  });
  return problems;
}

// ---------- גיליון ----------

function listRecipes() {
  const sheet = getSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, NUM_COLS).getValues()
    .filter(function (row) { return row[COL.ID - 1] !== ''; })
    .map(function (row) {
      return {
        id: String(row[COL.ID - 1]),
        name: String(row[COL.NAME - 1]),
        category: String(row[COL.CATEGORY - 1]),
        servings: row[COL.SERVINGS - 1],
        totalTime: String(row[COL.TOTAL_TIME - 1]),
        description: String(row[COL.DESCRIPTION - 1]),
        ingredients: String(row[COL.INGREDIENTS - 1]),
        updated: toIso_(row[COL.UPDATED - 1]),
      };
    })
    .sort(function (a, b) { return a.updated < b.updated ? 1 : -1; });
}

function getRecipe(id) {
  const sheet = getSheet_();
  const row = findRow_(sheet, id);
  if (row === -1) throw new Error('המתכון לא נמצא — אולי נמחק.');
  const json = sheet.getRange(row, COL.JSON).getValue();
  try {
    return { id: id, recipe: JSON.parse(json) };
  } catch (e) {
    throw new Error('הנתונים בעמודה K בשורה ' + row + ' פגומים.');
  }
}

/** שומר מתכון. בלי id (או אם השורה נמחקה) — שורה חדשה. מחזיר את המזהה. */
function saveRecipe(recipe, id) {
  if (!recipe || !recipe.name || !Array.isArray(recipe.ingredients) || !Array.isArray(recipe.steps)) {
    throw new Error('המתכון חסר — אין מה לשמור.');
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet_();
    const now = new Date();
    const row = id ? findRow_(sheet, id) : -1;
    if (row === -1) id = 'r' + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
    const created = row === -1 ? now : sheet.getRange(row, COL.CREATED).getValue() || now;

    const values = [
      id,
      safeText_(recipe.name),
      safeText_(recipe.category || ''),
      Number(recipe.servings) || '',
      formatMinutes_(recipe.totalMinutes),
      safeText_(recipe.description || ''),
      recipe.ingredients.map(function (i) { return '• ' + i.amount + ' ' + i.name; }).join('\n'),
      recipe.steps.map(function (s, n) {
        const wait = s.waitMinutes > 0 ? ' (המתנה ' + formatMinutes_(s.waitMinutes) + ')' : '';
        return (n + 1) + '. ' + s.title + ' — ' + s.text + wait;
      }).join('\n'),
      created,
      now,
      JSON.stringify(recipe),
    ];
    if (row === -1) {
      sheet.appendRow(values);
    } else {
      sheet.getRange(row, 1, 1, NUM_COLS).setValues([values]);
    }
    return id;
  } finally {
    lock.releaseLock();
  }
}

function deleteRecipe(id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet_();
    const row = findRow_(sheet, id);
    if (row !== -1) sheet.deleteRow(row);
  } finally {
    lock.releaseLock();
  }
}

function getSheet_() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('לא נמצאה לשונית בשם "' + SHEET_NAME + '" בגיליון.');
  return sheet;
}

function findRow_(sheet, id) {
  const last = sheet.getLastRow();
  if (last < 2 || !id) return -1;
  const ids = sheet.getRange(2, COL.ID, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

/** טקסט שמתחיל ב-= או + וכו' הגיליון מפרש כנוסחה; גרש בהתחלה מונע את זה */
function safeText_(value) {
  const s = String(value);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function formatMinutes_(minutes) {
  const m = Math.round(Number(minutes) || 0);
  if (m <= 0) return '';
  if (m < 60) return m + ' דק\'';
  const h = Math.floor(m / 60);
  const rest = m % 60;
  const hours = h === 1 ? 'שעה' : h === 2 ? 'שעתיים' : h + ' שעות';
  return rest ? hours + ' ו-' + rest + ' דק\'' : hours;
}

function toIso_(value) {
  if (value instanceof Date) return value.toISOString();
  return value ? String(value) : '';
}

// ---------- בדיקת התקנה (מריצים ידנית מהעורך) ----------

/**
 * בחר את הפונקציה הזאת בעורך ולחץ "הפעלה" (Run).
 * היא בודקת את הגיליון ואת המפתח, ומבקשת מ-Claude מתכון קטן בדיוק כמו הצ'אט (עלות של סנטים בודדים).
 * שום דבר לא נשמר בגיליון. התוצאה מופיעה ביומן הביצוע (Execution log).
 */
function testSetup() {
  const sheet = getSheet_();
  const header = sheet.getRange(1, 1, 1, NUM_COLS).getValues()[0];
  console.log('✔ הגיליון נמצא. כותרות: ' + header.join(' | '));

  const res = chat([{ role: 'user', content: 'מתכון לחביתה פשוטה למנה אחת' }]);
  console.log('✔ Claude ענה: ' + res.message);
  if (res.recipe) {
    console.log('✔ התקבל מתכון: ' + res.recipe.name + ' (' + res.recipe.ingredients.length + ' מצרכים, ' + res.recipe.steps.length + ' שלבים)');
  }
  if (res.problems.length) console.log('⚠ בעיות מבנה: ' + res.problems.join(' / '));
  console.log('הכל מוכן. אפשר לפרסם את האפליקציה (Deploy).');
}
