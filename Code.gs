var employeeCache = null;
const CACHE_EXPIRATION = 5 * 60; // 5 минут

function GENERATE_UUID() {
  return Utilities.getUuid();
}

// ==== ОЧИСТКА КЭША ====
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Протоколы')
    .addItem('Создать протокол встречи', 'showMeetingDialog')
    .addItem('Окно записей', 'showRecordDialog')
    .addSeparator()
    .addItem('Обновить кэш сотрудников', 'clearEmployeeCache') // Новая кнопка
    .addToUi();
}

function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() === 'Сотрудники') {
    clearEmployeeCache();
    SpreadsheetApp.getUi().alert('Кэш сотрудников обновлен!');
  }
}

function clearEmployeeCache() {
  const cache = CacheService.getScriptCache();
  const allKeys = cache.getAll([]); // Получаем все ключи
  const ourKeys = Object.keys(allKeys).filter(k => k.startsWith(CACHE_KEY_PREFIX));
  
  if (ourKeys.length > 0) {
    cache.removeAll(ourKeys);
  }
}

// ==== 
function showMeetingDialog() {
  var html = HtmlService.createHtmlOutputFromFile('MeetingForm')
    .setWidth(600)
    .setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, 'Новая встреча');
}

function showRecordDialog(meetingId, meetingNumber) {
  var html = HtmlService.createHtmlOutputFromFile('RecordForm')
    .setWidth(800)
    .setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, 'Протокол №' + (meetingNumber || ''));
  PropertiesService.getScriptProperties().setProperty('currentMeetingId', meetingId);
}

function getNextMeetingNumber() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('Встречи');
  return sheet.getLastRow() === 1 ? 1 : sheet.getRange(sheet.getLastRow(), 2).getValue() + 1;
}

function getColumnIndexes(headers) {
  return {
    name: headers.findIndex(h => h.trim() === 'Имя'),
    surname: headers.findIndex(h => h.trim() === 'Фамилия'),
    email: headers.findIndex(h => h.trim() === 'Почта'),
    id: headers.findIndex(h => h.trim() === 'Row ID'),
    org: headers.findIndex(h => h.trim() === 'Организация'),
    dept: headers.findIndex(h => h.trim() === 'Подразделение'),
    unit: headers.findIndex(h => h.trim() === 'Отдел')
  };
}

function getEmployees() {
  const cache = CacheService.getScriptCache();
  const CACHE_KEY_PREFIX = 'employees_part_'; // Префикс для ключей
  const CHUNK_SIZE = 90 * 1024; // Добавьте эту строку

  // Проверяем in-memory кэш
  if (employeeCache) return employeeCache;

  // Получаем ВСЕ ключи из кэша
  const allKeys = cache.getAll([]); // Пустой массив = все ключи
  const ourKeys = Object.keys(allKeys).filter(k => k.startsWith(CACHE_KEY_PREFIX));

  if (ourKeys.length > 0) {
    // Получаем только нужные части
    const cachedParts = cache.getAll(ourKeys);
    employeeCache = [];
    Object.keys(cachedParts)
      .sort((a, b) => 
        parseInt(a.replace(CACHE_KEY_PREFIX, '')) - 
        parseInt(b.replace(CACHE_KEY_PREFIX, ''))
      )
      .forEach(key => {
        employeeCache.push(...JSON.parse(cachedParts[key]));
      });
    return employeeCache;
  }

  // Загрузка данных из таблицы
  const sheet = SpreadsheetApp.getActive().getSheetByName('Сотрудники');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const columns = getColumnIndexes(headers);

  employeeCache = data.slice(1).map(row => {
    const firstName = row[columns.name] || '';
    const lastName = row[columns.surname] || '';
    
    return {
      id: row[columns.id],
      firstName: firstName,
      lastName: lastName,
      email: row[columns.email],
      organization: row[columns.org] || '',
      department: row[columns.dept] || '',
      unit: row[columns.unit] || '',
      displayName: `${firstName} ${lastName}`.trim() // Автоматическая генерация
    };
  }).filter(e => e.email);

  // Разбиваем на части и сохраняем
  const jsonData = JSON.stringify(employeeCache);
  const chunks = [];
  
  for (let i = 0; i < jsonData.length; i += CHUNK_SIZE) {
    const chunk = jsonData.substring(i, i + CHUNK_SIZE);
    chunks.push({
      key: CACHE_KEY_PREFIX + (i / CHUNK_SIZE),
      value: chunk,
      expiration: CACHE_EXPIRATION
    });
  }

  const chunkData = {};
  chunks.forEach((chunk, index) => {
    chunkData[CACHE_KEY_PREFIX + index] = chunk.value;
  });
  cache.putAll(chunkData, CACHE_EXPIRATION);

  return employeeCache;
}

// function getColumn(sheetName, columnName) {
//   var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
//   var data = sheet.getDataRange().getValues();
//   var headers = data[0]
//   return columnNumber = nameColumnIndex1 = headers.findIndex(h => h.trim() === columnName.trim());
// }

function getRecordTypes() {
  return getTableData('Данные', 'Типы записей');
}

function getImportanceLevels() {
  return getTableData('Данные', 'Значимость');
}

function getPriorityLevels() {
  return getTableData('Данные', 'Приоритет');
}

function getTableData(sheetName, columnName) {
  console.log('Загрузка данных из листа "%s", столбец "%s"', sheetName, columnName);
  try {
    var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!sheet) throw new Error('Лист "' + sheetName + '" не найден');
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var columnIndex = headers.findIndex(h => h.trim() === columnName.trim());
    
    if (columnIndex === -1) throw new Error('Столбец "' + columnName + '" не найден');
    
    return data.slice(1)
      .map(row => row[columnIndex])
      .filter(value => value !== '' && value !== null && value !== undefined);
  } catch (e) {
    console.error('Ошибка в getTableData: ', e);
    throw e; // Перебрасываем ошибку для обработки в вызывающем коде
  }
}

function getMeetingAttendees(meetingId) {
  try {
    var sheet = SpreadsheetApp.getActive().getSheetByName('Встречи');
    if (!sheet) throw new Error('Лист "Встречи" не найден');
    
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return []; // Если нет данных кроме заголовков
    
    var headers = data[0];
    var idCol = headers.indexOf('ID встречи');
    var attendeeIdsCol = headers.indexOf('ID участников');
    
    if (idCol === -1 || attendeeIdsCol === -1) {
      throw new Error('Не найдены необходимые колонки');
    }
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][idCol] === meetingId) {
        var attendeeIds = data[i][attendeeIdsCol];
        return attendeeIds ? attendeeIds.toString().split(',').map(id => id.trim()).filter(id => id) : [];
      }
    }
    return [];
  } catch (e) {
    console.error('Ошибка в getMeetingAttendees: ', e);
    return [];
  }
}

function createMeeting(meetingData) {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Встречи');
  const meetingId = Utilities.getUuid();
  const employees = getEmployees();

  // Подготовка данных
  const formattedNames = [];
  const attendeeIds = [];
  
  try {
    // 1. Валидация даты
    const meetingDate = new Date(meetingData.date);
    if (isNaN(meetingDate.getTime())) {
      throw new Error("Некорректный формат даты: " + meetingData.date);
    }

    // 2. Обработка участников
    meetingData.attendees.forEach(email => {
      const employee = employees.find(e => e.email === email);
      
      if (!employee) {
        throw new Error(`Сотрудник с email ${email} не найден`);
      }

      // 3. Форматирование имени
      const firstNameChar = employee.firstName 
        ? employee.firstName[0].toUpperCase() + "." 
        : "";
      formattedNames.push(`${firstNameChar} ${employee.lastName}`);
      
      attendeeIds.push(employee.id);
    });

    // 4. Запись в таблицу
    sheet.appendRow([
      meetingId,                         // A: ID встречи
      meetingData.meetingNumber,         // B: Номер встречи
      meetingDate,                      // C: Дата (корректный Date объект)
      meetingData.topic,                 // D: Тема
      formattedNames.join(', '),         // E: Имена
      attendeeIds.join(', ')             // F: ID участников
    ]);

    return { 
      id: meetingId, 
      number: meetingData.meetingNumber 
    };

  } catch (e) {
    console.error("Ошибка создания встречи:", e);
    throw new Error("Не удалось сохранить встречу. " + e.message);
  }
}

function createRecords(recordsData) {
  var sheet = SpreadsheetApp.getActive().getSheetByName('Записи');
  var meetingId = PropertiesService.getScriptProperties().getProperty('currentMeetingId');
  
  var rows = recordsData.map(function(record) {
    return [
      Utilities.getUuid(),
      meetingId,
      record.type,
      record.text,
      record.dueDate,
      record.responsible.join(', '),
      record.importance,
      record.priority,
      record.recordNumber
    ];
  });
  
  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow()+1, 1, rows.length, rows[0].length).setValues(rows);
  }
  return `Сохранено ${rows.length} записей`;
}
