// ═══════════════════════════════════════════════════════
//  DriveMe Дніпро — Google Apps Script
//  Вставити в: script.google.com → Новий проєкт
//  Деплой: Розгорнути → Веб-застосунок → Всі мають доступ
// ═══════════════════════════════════════════════════════

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Ліди') || ss.insertSheet('Ліди');

    // Заголовки (лише при першому запуску)
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Дата', 'Ім\'я', 'Телефон', 'Джерело']);
      sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
    }

    sheet.appendRow([
      data.date   || new Date().toLocaleString('uk-UA'),
      data.name   || '',
      data.phone  || '',
      data.source || 'Сайт'
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Тест вручну (запустити в редакторі для перевірки)
function testWrite() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Ліди') || ss.insertSheet('Ліди');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Дата', 'Ім\'я', 'Телефон', 'Джерело']);
  }
  sheet.appendRow(['21.05.2026 тест', 'Тест Іван', '+380991234567', 'Тест скрипту']);
  Logger.log('OK — рядок додано');
}
