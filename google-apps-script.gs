const SPREADSHEET_ID = '1HnaOW2i17_Kat5BFsb7EZcbRdkaLpWKV1FuzMYMPUA0';
const SHEET_NAME = 'результаты';
const AI_BASE_URL = 'http://111.228.46.150:3000/v1';
const AI_API_KEY = 'sk-BtsZuSh7mCzjwPmke9ZzuCqnJPOucf6KXfHrdrVtglpNjYdm';
const AI_MODEL = 'wj-interpret';

function doGet(e) {
  const params = e && e.parameter || {};
  if (params.action === 'check') {
    const exists = findStudentRow_(getSheet_(), params.studentName || '') > 0;
    const result = JSON.stringify({ok: true, exists: exists});
    if (params.callback) return ContentService.createTextOutput(params.callback + '(' + result + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
    return json_({ok: true, exists: exists});
  }
  return json_({ok: true, service: 'octal-quiz', sheet: SHEET_NAME});
}

function doPost(e) {
  try {
    const data = parseRequest_(e);
    if (!data.studentName) throw new Error('Не указано имя участника');
    const sheet = getSheet_();
    ensureHeaders_(sheet);
    if (data.action === 'ai-review') return processAiReview_(sheet, data);
    const row = [
      data.studentName || '',
      formatDate_(data.completedAt),
      data.variant || '',
      data.score || 0,
      data.total || 19
    ];
    (data.answers || []).forEach(function(answer) { row.push(formatAnswer_(answer)); });
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      if (findStudentRow_(sheet, data.studentName) > 0) throw new Error('Ученик с таким именем и фамилией уже проходил тест');
      sheet.appendRow(row);
    } finally {
      lock.releaseLock();
    }
    SpreadsheetApp.flush();
    return json_({ok: true, saved: true});
  } catch (error) {
    return json_({ok: false, error: String(error)});
  }
}

function parseRequest_(e) {
  const formPayload = e && e.parameter && e.parameter.payload;
  const raw = formPayload || (e && e.postData && e.postData.contents) || '{}';
  if (formPayload) return JSON.parse(formPayload);
  try {
    return JSON.parse(raw);
  } catch (error) {
    const params = raw.split('&').reduce(function(result, item) {
      const parts = item.split('=');
      result[decodeURIComponent(parts[0])] = decodeURIComponent(parts.slice(1).join('='));
      return result;
    }, {});
    return params.payload ? JSON.parse(params.payload) : {};
  }
}

function processAiReview_(sheet, data) {
  const aiResults = checkOpenAnswers_(data.answers || []);
  const rowNumber = findStudentRow_(sheet, data.studentName || '');
  if (rowNumber > 0) {
    sheet.getRange(rowNumber, 4).setValue(aiResults.filter(function(answer) { return answer.isCorrect; }).length);
    aiResults.forEach(function(answer, index) { sheet.getRange(rowNumber, 6 + index).setValue(formatAnswer_(answer)); });
    SpreadsheetApp.flush();
  }
  return json_({ok: true, reviewed: true});
}

function formatResults_(answers) {
  return answers.map(formatAnswer_).join('\n\n');
}

function formatAnswer_(answer) {
  const status = answer.isCorrect ? 'ВЕРНО' : (answer.answerText || answer.answer ? 'ОШИБКА' : 'НЕТ ОТВЕТА');
  return [
    'Вопрос: ' + answer.question,
    'Ответ: ' + (answer.answerText || answer.answer || 'нет ответа'),
    'Правильный ответ: ' + (answer.correctText || answer.correct || 'не указан'),
    'Результат: ' + status
  ].join('\n');
}

function formatDate_(value) {
  const date = value ? new Date(value) : new Date();
  if (isNaN(date.getTime())) return String(value || '');
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm:ss');
}

function getSheet_() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Лист «' + SHEET_NAME + '» не найден');
  return sheet;
}

function ensureHeaders_(sheet) {
  const headers = ['Имя Фамилия', 'Время завершения', 'Вариант', 'Баллы', 'Всего заданий'];
  for (let i = 1; i <= 19; i++) headers.push('Задание ' + i);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
}

function findStudentRow_(sheet, name) {
  if (!name || sheet.getLastRow() < 2) return -1;
  const names = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < names.length; i++) {
    if (String(names[i][0]).trim().toLowerCase() === String(name).trim().toLowerCase()) return i + 2;
  }
  return -1;
}

function checkOpenAnswers_(answers) {
  return answers.map(function(answer) {
    if (!answer.open || !answer.answer) return answer;
    try {
      const response = UrlFetchApp.fetch(AI_BASE_URL + '/chat/completions', {
        method: 'post',
        contentType: 'application/json',
        headers: {Authorization: 'Bearer ' + AI_API_KEY},
        payload: JSON.stringify({
          model: AI_MODEL,
          temperature: 0,
          messages: [
            {role: 'system', content: 'Проверь учебный ответ. Верни только JSON без markdown в формате {"correct":true} или {"correct":false}. Считай ответ правильным, если смысл совпадает с эталоном.'},
            {role: 'user', content: 'Вопрос: ' + answer.question + '\nЭталон: ' + answer.correct + '\nОтвет ученика: ' + answer.answer}
          ]
        }),
        muteHttpExceptions: true
      });
      const body = JSON.parse(response.getContentText());
      const text = body.choices && body.choices[0] && body.choices[0].message.content || '{}';
      const match = text.match(/\{[\s\S]*\}/);
      if (!match) throw new Error('AI did not return JSON');
      const parsed = JSON.parse(match[0]);
      answer.isCorrect = parsed.correct === true;
    } catch (error) {
      answer.aiError = String(error);
    }
    return answer;
  });
}

function json_(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}
