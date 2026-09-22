// Google Apps Script для обработки заказов из формы на сайте
// Инструкция по установке:
// 1. Откройте Google Таблицу: https://docs.google.com/spreadsheets/d/1ZNSjyMbDX6uhpPreP6522YEaKCQKpGm1j70_uP8d5w/edit
// 2. Расширения -> Apps Script
// 3. Удалите весь существующий код и вставьте этот скрипт
// 4. Нажмите "Развернуть" -> "Новое развертывание"
// 5. Выберите тип "Веб-приложение"
// 6. В поле "У кого есть доступ" выберите "Все пользователи" (Anyone)
// 7. Нажмите "Развернуть" и скопируйте URL веб-приложения
// 8. Вставьте этот URL в index.html вместо 'YOUR_GOOGLE_APPS_SCRIPT_URL'

const SHEET_ID = '1ZNSjyMbDX6uhpPreP6522YEaKCQKpGm1j70_uP8d5w';
const EMAIL_TO = 'bton.main@yandex.ru';

function doPost(e) {
  try {
    // Парсим данные из запроса (поддержка как JSON, так и form-urlencoded)
    let data;
    if (e.postData && e.postData.contents) {
      const contentType = e.postData.type || '';
      if (contentType.indexOf('application/json') !== -1) {
        data = JSON.parse(e.postData.contents);
      } else {
        // Обработка form-urlencoded данных
        data = JSON.parse(JSON.stringify(e.parameter));
      }
    } else {
      data = e.parameter;
    }
    
    // Получаем текущую дату и время
    const now = new Date();
    const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd.MM.yyyy HH:mm");
    
    // Открываем таблицу
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    const sheet = spreadsheet.getActiveSheet();
    
    // Проверяем заголовки, если таблица пустая - создаем их
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Дата', 'Имя', 'Контактные данные', 'Комментарии', 'Заказ']);
      // Форматируем заголовки
      const headerRange = sheet.getRange(1, 1, 1, 5);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#f3f3f3');
    }
    
    // Добавляем новую строку с данными заказа
    sheet.appendRow([
      data.date || dateStr,
      data.name || '',
      data.contact || '',
      data.comment || '',
      data.order || ''
    ]);
    
    // Формируем текст письма для уведомления
    const subject = '🛒 Новый заказ B\'TON - ' + (data.date || dateStr);
    const htmlBody = `
      <h2>Новый заказ B'TON</h2>
      <p><strong>Дата:</strong> ${data.date || dateStr}</p>
      <p><strong>Имя:</strong> ${data.name || ''}</p>
      <p><strong>Контактные данные:</strong> ${data.contact || ''}</p>
      <p><strong>Комментарии:</strong> ${data.comment || ''}</p>
      <p><strong>Заказ:</strong> ${data.order || ''}</p>
      <hr>
      <p><em>Это автоматическое уведомление от системы заказов B'TON.</em></p>
    `;
    
    // Отправляем email уведомление
    MailApp.sendEmail({
      to: EMAIL_TO,
      subject: subject,
      htmlBody: htmlBody
    });
    
    // Возвращаем успешный ответ
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success', message: 'Order saved successfully' }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    // Логируем ошибку и возвращаем её
    Logger.log('Error processing order: ' + error.toString());
    
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ message: 'B\'TON Order Processing API is running' }))
    .setMimeType(ContentService.MimeType.JSON);
}
