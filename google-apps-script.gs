const SPREADSHEET_ID = '1HnaOW2i17_Kat5BFsb7EZcbRdkaLpWKV1FuzMYMPUA0';
const SHEET_NAME = 'результаты';
const AI_BASE_URL = 'http://111.228.46.150:3000/v1';
const AI_API_KEY = 'sk-BtsZuSh7mCzjwPmke9ZzuCqnJPOucf6KXfHrdrVtglpNjYdm';
const AI_MODEL = 'wj-interpret';

function doGet() {
  return json_({ok: true, service: 'octal-quiz', sheet: SHEET_NAME});
}

function doPost(e) {
  try {
    const data = parseRequest_(e);
    if (!data.studentName) throw new Error('Не указано имя участника');
    const sheet = getSheet_();
    ensureHeaders_(sheet);
    if (data.action === 'ai-review') return processAiReview_(sheet, data);
    const resultText = data.results || JSON.stringify(data.answers || []);
    const row = [
      data.studentName || '',
      data.completedAt || new Date().toISOString(),
      data.variant || '',
      data.score || 0,
      data.total || 19,
      resultText
    ];
    const existing = findStudentRow_(sheet, data.studentName || '');
    if (existing > 0) {
      sheet.getRange(existing, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
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
  const resultText = JSON.stringify(aiResults);
  const rowNumber = findStudentRow_(sheet, data.studentName || '');
  if (rowNumber > 0) {
    sheet.getRange(rowNumber, 4).setValue(aiResults.filter(function(answer) { return answer.isCorrect; }).length);
    sheet.getRange(rowNumber, 6).setValue(resultText);
  }
  return json_({ok: true, reviewed: true});
}

function getSheet_() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Лист «' + SHEET_NAME + '» не найден');
  return sheet;
}

function ensureHeaders_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Имя Фамилия', 'Время завершения', 'Вариант', 'Баллы', 'Всего заданий', 'Результаты']);
  }
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
