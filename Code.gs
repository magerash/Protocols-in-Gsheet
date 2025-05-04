var employeeCache = null;
const CACHE_EXPIRATION = 5 * 60; // 5 минут

function GENERATE_UUID() {
  return Utilities.getUuid();
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Протоколы')
    .addItem('Создать протокол встречи', 'showMeetingDialog')
    .addItem('Окно записей', 'showRecordDialog')
    .addToUi();
}

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
  
  // 1. Всегда сначала проверяем in-memory кэш
  if (employeeCache) {
    console.log('Returning from memory cache');
    return employeeCache;
  }

  // 2. Проверяем persistent кэш
  const cached = cache.get('employees');
  if (cached) {
    console.log('Loading from script cache');
    employeeCache = JSON.parse(cached);
    return employeeCache;
  }
  
  // 3. Загрузка из таблицы если кэши пустые  
  console.log('Loading fresh data');
  const sheet = SpreadsheetApp.getActive().getSheetByName('Сотрудники');
  const data = sheet.getDataRange().getValues(); // Единственный запрос данных
  const headers = data[0]; // Получаем заголовки из первого элемента данных
  const columns = getColumnIndexes(headers);

  // Валидация обязательных колонок
  if (columns.name === -1) throw new Error('Колонка "Имя" не найдена');
  if (columns.surname === -1) throw new Error('Колонка "Фамилия" не найдена');
  if (columns.email === -1) throw new Error('Колонка "Почта" не найдена');

  employeeCache = data.slice(1).map(row => ({
    id: row[columns.id],
    firstName: row[columns.name],
    lastName: row[columns.surname],
    email: row[columns.email],
    organization: row[columns.org] || '',
    department: row[columns.dept] || '',
    unit: row[columns.unit] || '',
    displayName: `${row[columns.name]} ${row[columns.surname]}`
  })).filter(e => e.email);
  
  // 4. Обновляем оба кэша
  cache.put('employees', JSON.stringify(employeeCache), CACHE_EXPIRATION);
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
  var sheet = SpreadsheetApp.getActive().getSheetByName('Встречи');
  var meetingId = Utilities.getUuid();

  // Получаем данные сотрудников
  var employees = getEmployees();

  // Формируем данные для сохранения
  var formattedNames = [];
  var attendeeIds = [];

  
  // Обрабатываем каждого участника
  meetingData.attendees.forEach(email => {
    var employee = employees.find(e => e.email === email);
    if (employee) {
      // Форматируем имя как "И. Иванов"
      var formattedName = employee.firstName[0].toUpperCase() + '. ' + 
                         employee.lastName[0].toUpperCase() + 
                         employee.lastName.slice(1).toLowerCase();
      
      formattedNames.push(formattedName);
      attendeeIds.push(employee.id);
    }
  });

  sheet.appendRow([
    meetingId,                         // A: ID встречи
    meetingData.meetingNumber,         // B: Номер встречи
    meetingData.date,                  // C: Дата создания
    meetingData.topic,                 // D: Тема
    formattedNames.join(', '),         // E: Форматированные имена ("И. Иванов, П. Петров")
    attendeeIds.join(', ')             // F: ID участников через запятую ("1,2,3")
  ]);
  
  return {
    id: meetingId,
    number: meetingData.meetingNumber
  };
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
