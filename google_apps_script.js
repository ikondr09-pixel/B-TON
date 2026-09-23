// Google Apps Script для обработки заказов из формы на сайте
// Инструкция по установке:
// 1. Откройте Google Таблицу: https://docs.google.com/spreadsheets/d/1ZNSjyMbDX6uhpPreP6522YEaKCQKpGm1j70_uP8d5w/edit
// 2. Расширения -> Apps Script
// 3. Удалите весь существующий код и вставьте этот скрипт
// 4. Нажмите "Развернуть" -> "Новое развертывание"
// 5. Выберите тип "Веб-приложение"
// 6. В поле "Выполнять как" выберите "От имени владельца" (Me)
// 7. В поле "Кто имеет доступ" выберите "Все пользователи" (Anyone) - НЕ "Все пользователи с аккаунтом Google"
// 8. Нажмите "Развернуть" и скопируйте URL веб-приложения
// 9. Вставьте этот URL в index.html вместо 'YOUR_GOOGLE_APPS_SCRIPT_URL'
// 
// ВАЖНО: Если вы уже создавали развертывание, нужно создать НОВОЕ (Manage deployments -> Edit -> New deployment)
// или удалить старое и создать заново. Просто изменить настройки существующего недостаточно!
// 
// ПРИЧИНА ОШИБКИ 404: Google Apps Script делает редирект при запросе. Браузер в режиме no-cors не может 
// отследить редирект на script.googleusercontent.com. Это нормально - заказ всё равно обрабатывается.

// ВНИМАНИЕ (результат диагностики, сентябрь 2026): этот SHEET_ID НЕ СОВПАДАЕТ с ссылкой в инструкции выше!
//   ссылка:     .../d/1ZNSjyMbDX6uhpPreP6522YEaKCQKpGm1j70_uP8d5w/edit   (43 символа)
//   в коде:     1ZNSjyMbDX6uhpPreP6522YEaKCQKpGm1j70_uP8d5w              (42 символа — пропущен 1 символ)
// Google отвечает «does not exist» на оба варианта без авторизации, поэтому точный ID нужно взять из
// реальной таблицы (он всегда в URL строки браузера: docs.google.com/spreadsheets/d/<ВОТ ЭТОТ ID>/edit).
// Именно из-за неверного ID SpreadsheetApp.openById(SHEET_ID) бросает исключение внутри try/catch,
// скрипт его «проглатывает», письмо уходит, а строка в таблицу НЕ добавляется.
// ПОДСТАВЬТЕ СЮДА реальный ID вашей таблицы, а затем обновите существующее развертывание
// до новой версии кода: Развернуть -> Управление развертываниями -> ✏️ -> Версия: Новая -> Обновить.
// При таком обновлении URL /exec НЕ меняется, править index.html не нужно.
const SHEET_ID = '1ZNSjyMbDX6uhpPreP6522YEaKCQKpGm1j70_uP8d5w';
const EMAIL_TO = 'bton.main@yandex.ru';

function doPost(e) {
  try {
    Logger.log('========== START ORDER PROCESSING ==========');
    Logger.log('Received POST request');
    
    // Парсим данные из запроса
    let data;
    if (e && e.postData && e.postData.contents) {
      const contentType = e.postData.type || '';
      Logger.log('Content-Type: ' + contentType);
      
      if (contentType.indexOf('application/json') !== -1) {
        data = JSON.parse(e.postData.contents);
        Logger.log('Parsed JSON data');
      } else if (contentType.indexOf('application/x-www-form-urlencoded') !== -1) {
        // Для form-urlencoded данные уже распарсены в e.parameter
        data = e.parameter;
        Logger.log('Parsed form-urlencoded data');
      } else {
        // Пытаемся распарсить как JSON, если не получится - используем parameter
        try {
          data = JSON.parse(e.postData.contents);
          Logger.log('Parsed as JSON (fallback)');
        } catch (jsonError) {
          data = e.parameter;
          Logger.log('Using parameter data (fallback)');
        }
      }
    } else if (e && e.parameter) {
      data = e.parameter;
      Logger.log('Using parameter data');
    } else {
      throw new Error('No data received in request');
    }
    
    Logger.log('Final parsed data: ' + JSON.stringify(data));
    
    // Получаем текущую дату и время
    const now = new Date();
    const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd.MM.yyyy HH:mm");
    
    // ============================================
    // ШАГ 1: ОТПРАВКА EMAIL УВЕДОМЛЕНИЯ (ПЕРВЫМ!)
    // ============================================
    let emailSent = false;
    try {
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
      
      Logger.log('Sending email to: ' + EMAIL_TO);
      MailApp.sendEmail({
        to: EMAIL_TO,
        subject: subject,
        htmlBody: htmlBody
      });
      emailSent = true;
      Logger.log('✅ Email notification sent successfully');
    } catch (emailError) {
      Logger.log('⚠️ Error sending email: ' + emailError.toString());
      // Не прерываем выполнение, продолжаем запись в таблицу
    }
    
    // ============================================
    // ШАГ 2: ЗАПИСЬ В ТАБЛИЦУ
    // ============================================
    let sheetSaved = false;
    let resultSheetError = null;
    try {
      // Проверяем, что SHEET_ID вообще похож на валидный ID таблицы.
      // В этом файле ID на 1 символ короче, чем ссылка в инструкции выше (см. комментарий у SHEET_ID):
      // скорее всего это опечатка, и openById() падает здесь -> письмо уходит, строки в таблице нет.
      if (!SHEET_ID || SHEET_ID.length < 33 || SHEET_ID.length > 44 || /^YOUR_/.test(SHEET_ID)) {
        throw new Error('SHEET_ID некорректен: "' + SHEET_ID + '". Возьмите ID из URL таблицы: docs.google.com/spreadsheets/d/<ID>/edit');
      }

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
      
      sheetSaved = true;
      Logger.log('✅ Order saved to sheet successfully');
    } catch (sheetError) {
      Logger.log('⚠️ Error saving to sheet: ' + sheetError.toString());

      // Резервная копия заказа в свойствах скрипта (PropertiesService), чтобы заказы не терялись,
      // пока проблема с таблицей не исправлена. Дальше их можно выгрузить функцией dumpPendingOrders().
      try {
        const props = PropertiesService.getScriptProperties();
        const pending = JSON.parse(props.getProperty('PENDING_ORDERS') || '[]');
        pending.push({ ts: now.toISOString(), data: data, error: sheetError.toString() });
        props.setProperty('PENDING_ORDERS', JSON.stringify(pending.slice(-200)));
        Logger.log('Order stored in PENDING_ORDERS (total: ' + pending.length + ')');
      } catch (backupError) {
        Logger.log('Backup failed: ' + backupError.toString());
      }

      // Пробрасываем текст ошибки наружу — иначе по ответу API нельзя понять, почему таблица не пишется.
      resultSheetError = sheetError.toString();
      // Email уже отправлен, так что заказ не потерян
    }
    
    // ============================================
    // ВОЗВРАЩАЕМ ОТВЕТ
    // ============================================
    const result = {
      result: 'success',
      emailSent: emailSent,
      sheetSaved: sheetSaved,
      sheetError: resultSheetError,
      message: 'Order processed (email: ' + (emailSent ? 'yes' : 'no') + ', sheet: ' + (sheetSaved ? 'yes' : 'no') + ')' +
               (resultSheetError ? ' | SHEET ERROR: ' + resultSheetError : '')
    };
    
    Logger.log('Result: ' + JSON.stringify(result));
    Logger.log('========== END ORDER PROCESSING ==========');
    
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    // Логируем ошибку и возвращаем её
    Logger.log('❌ CRITICAL ERROR processing order: ' + error.toString());
    Logger.log('Stack trace: ' + error.stack);
    
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: error.toString(), stack: error.stack }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ message: 'B\'TON Order Processing API is running', status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Выгрузка заказов, которые не удалось записать в таблицу (резервная копия из PropertiesService).
// Запустите из редактора Apps Script (Run -> dumpPendingOrders), результат появится в Log.
// После исправления SHEET_ID и создания НОВОГО развертывания эти заказы можно перенести в таблицу вручную.
function dumpPendingOrders() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty('PENDING_ORDERS');
  if (!raw) {
    Logger.log('No pending orders stored.');
    return [];
  }
  const pending = JSON.parse(raw);
  Logger.log('Pending orders: ' + pending.length);
  pending.forEach(function (p, i) {
    Logger.log('#' + (i + 1) + ' [' + p.ts + '] ' + JSON.stringify(p.data) + ' | error: ' + p.error);
  });
  return pending;
}

// Полная очистка резервной копии (после переноса заказов в таблицу).
function clearPendingOrders() {
  PropertiesService.getScriptProperties().deleteProperty('PENDING_ORDERS');
  Logger.log('PENDING_ORDERS cleared.');
}
