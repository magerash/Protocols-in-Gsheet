//файл: code.gs

const DEBUG_MODE = true;

// Пример улучшенного логирования
function customLog(funcName, message, ...args) {
  if (DEBUG_MODE) {
    const logMessage = `[${funcName}] ${message}`;
    if (args.length > 0) {
      Logger.log(logMessage, ...args);
    } else {
      Logger.log(logMessage);
    }
  }
}

function GENERATE_UUID() {
  uuid = Utilities.getUuid()
  Logger.log('[GENERATE_UUID] cгенерированный UUID: ' + uuid);
  
  return Utilities.getUuid();
}

var employeeCache = null;
const CACHE_EXPIRATION = 5 * 60; // 5 минут
const CACHE_KEY_PREFIX = 'employees_part_';
const CHUNK_SIZE = 90 * 1024; 

function onEdit(e) {
  Logger.log('[onEdit] Событие редактирования. Лист: %s', e.source.getActiveSheet().getName());
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() === 'Сотрудники') {
    Logger.log('[onEdit] Обнаружено редактирование листа Сотрудники');
    clearEmployeeCache();
    SpreadsheetApp.getUi().alert('Кэш сотрудников обновлен!');
  }
}

function clearEmployeeCache() {
  Logger.log('[clearEmployeeCache] Очистка кэша сотрудников');
  const cache = CacheService.getScriptCache();
  const allKeys = cache.getAll([]); // Получаем все ключи
  const ourKeys = Object.keys(allKeys).filter(k => k.startsWith(CACHE_KEY_PREFIX));
  Logger.log('[clearEmployeeCache] Найдено ключей: %s', ourKeys.length);
  if (ourKeys.length > 0) {
    cache.removeAll(ourKeys);
    Logger.log('[clearEmployeeCache] Кэш успешно очищен');
  }
}

function getNextMeetingNumber() {
  Logger.log('[getNextMeetingNumber] Получение следующего номера встречи');  
  var sheet = SpreadsheetApp.getActive().getSheetByName('Встречи');
  const lastRow = sheet.getLastRow();
  const result = lastRow === 1 ? 1 : sheet.getRange(lastRow, 2).getValue() + 1;
  Logger.log('[getNextMeetingNumber] Результат: %s', result);
  return result;
}

function getColumnIndexes(headers) {
  Logger.log('[getColumnIndexes] Поиск индексов колонок');
  const indexes = {
    name: headers.findIndex(h => h.trim() === 'Имя'),
    surname: headers.findIndex(h => h.trim() === 'Фамилия'),
    email: headers.findIndex(h => h.trim() === 'Почта'),
    id: headers.findIndex(h => h.trim() === 'Row ID'),
    org: headers.findIndex(h => h.trim() === 'Организация'),
    dept: headers.findIndex(h => h.trim() === 'Подразделение'),
    unit: headers.findIndex(h => h.trim() === 'Отдел')
  };
  Logger.log('[getColumnIndexes] Результат: %s', JSON.stringify(indexes));
  return indexes;
}

function getEmployees() {
  Logger.log('[getEmployees] Начало загрузки сотрудников');  
  const cache = CacheService.getScriptCache();

  if (employeeCache) {
    Logger.log('[getEmployees] Используем кэш из памяти');
    return employeeCache;
  }

  const allKeys = cache.getAll([]); // Пустой массив = все ключи
  const ourKeys = Object.keys(allKeys).filter(k => k.startsWith(CACHE_KEY_PREFIX));
  Logger.log('[getEmployees] Найдено частей в кэше: %s', ourKeys.length);

  if (ourKeys.length > 0) {
    // Получаем только нужные части
    Logger.log('[getEmployees] Восстанавливаем из кэша');
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
    Logger.log('[getEmployees] Загружено записей: %s', employeeCache.length);      
    return employeeCache;
  }

  Logger.log('[getEmployees] Загрузка из таблицы');
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
  Logger.log('[getEmployees] Отфильтровано записей: %s', employeeCache.length);

  Logger.log('[getEmployees] Разбиваем на части и сохраняем записи');
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
  Logger.log('[getEmployees] Сохранение в кэш. Частей: %s', chunks.length);

  const chunkData = {};
  chunks.forEach((chunk, index) => {
    chunkData[CACHE_KEY_PREFIX + index] = chunk.value;
  });
  cache.putAll(chunkData, CACHE_EXPIRATION);

  console.log('Loaded employees:', employeeCache.map(e => e.email)); // Логирование email
  return employeeCache;
}

function getRecordTypes() {
  Logger.log('[getRecordTypes] Получение типов записей');
  const result = getTableData('Данные', 'Типы записей');
  Logger.log('[getRecordTypes] Найдено типов: %s', result.length);
  return result;
}

function getImportanceLevels() {
  Logger.log('[getImportanceLevels] Получение уровней значимости');
  const result = getTableData('Данные', 'Значимость');
  Logger.log('[getImportanceLevels] Найдено уровней: %s', result.length);
  return result;
}

function getPriorityLevels() {
  Logger.log('[getPriorityLevels] Получение уровней приоритета');
  const result = getTableData('Данные', 'Приоритет');
  Logger.log('[getPriorityLevels] Найдено уровней: %s', result.length);
  return result;
}

function getTableData(sheetName, columnName) {
  Logger.log('[getTableData] Начало обработки. Лист: "%s", Колонка: "%s"', sheetName, columnName);
  
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!sheet) {
      Logger.error('[getTableData] Лист "%s" не найден', sheetName);
      throw new Error('Лист "' + sheetName + '" не найден');
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const columnIndex = headers.findIndex(h => h.trim() === columnName.trim());
    
    if (columnIndex === -1) {
      Logger.error('[getTableData] Колонка "%s" не найдена', columnName);
      throw new Error('Столбец "' + columnName + '" не найден');
    }
    
    const result = data.slice(1)
      .map(row => row[columnIndex])
      .filter(value => value !== '' && value !== null && value !== undefined);
    
    Logger.log('[getTableData] Успешно. Найдено записей: %s', result.length);
    return result;
    
  } catch (e) {
    Logger.error('[getTableData] Ошибка: %s', e.toString());
    throw e;
  }
}

function createMeeting(meetingData) {
  Logger.log('[createMeeting] Начало создания встречи. Данные: %s', JSON.stringify(meetingData));
  
  const sheet = SpreadsheetApp.getActive().getSheetByName('Встречи');
  const meetingId = Utilities.getUuid();
  const employees = getEmployees();
  Logger.log('[createMeeting] Получено сотрудников: %s', employees.length);

  let formattedNames = [];
  let attendeeIds = [];
  let invalidEmails = [];
  
  try {
    const calendarData = PropertiesService.getScriptProperties()
      .getProperty('CALENDAR_EVENT_DATA');
    if (calendarData) { // Добавить проверку на наличие данных
      const parsedData = JSON.parse(calendarData);
      if (parsedData.startTime) {
        meetingData.date = parsedData.startTime;
        meetingData.attendees = parsedData.attendees || [];
      }
    }
    // // Проверяем данные из календаря
    // const calendarData = JSON.parse(
    //   PropertiesService.getScriptProperties()
    //     .getProperty('CALENDAR_EVENT_DATA') || '{}'
    // );
    
    // // Если есть данные из календаря, дополняем meetingData
    // if(calendarData.startTime) {
    //   meetingData.date = calendarData.startTime;
    //   meetingData.attendees = calendarData.attendees;
    // }
        
    // Валидация даты
    Logger.log('[createMeeting] Валидация даты: %s', meetingData.date);
    const meetingDate = new Date(meetingData.date);
    if (isNaN(meetingDate.getTime())) {
      throw new Error("Некорректный формат даты: " + meetingData.date);
    }

    // Обработка участников
    Logger.log('[createMeeting] Обработка %s участников', meetingData.attendees.length);
    meetingData.attendees.forEach((email, index) => {
      const lowerEmail = email.toLowerCase();
      const employee = employees.find(e => e.email.toLowerCase() === lowerEmail);
      
      if (employee) {
        const firstNameInitial = employee.firstName ? employee.firstName[0].toUpperCase() + '.' : '';
        formattedNames.push(`${firstNameInitial} ${employee.lastName}`);
        attendeeIds.push(employee.id);
        Logger.log('[createMeeting] Участник %s: %s добавлен', index + 1, email);
      } else {
        invalidEmails.push(email);
        Logger.log('[createMeeting] Невалидный email: %s', email);
      }
    });

    if (attendeeIds.length === 0) {
      Logger.error('[createMeeting] Нет валидных участников');
      throw new Error("Нет ни одного корректного участника");
    }

    // Запись в таблицу
    Logger.log('[createMeeting] Запись в лист "Встречи"');
    sheet.appendRow([
      meetingId,
      meetingData.meetingNumber,
      meetingDate,
      meetingData.topic,
      formattedNames.join(', '),
      attendeeIds.join(', '),
      meetingData.location
    ]);

    const result = { 
      id: meetingId, 
      number: meetingData.meetingNumber,
      success: true,
      invalidEmails: invalidEmails,
      message: invalidEmails.length > 0 
        ? `Встреча сохранена, но не найдены: ${invalidEmails.join(', ')}`
        : 'Встреча успешно сохранена'      
    };
    
    Logger.log('[createMeeting] Успешно создано. Результат: %s', JSON.stringify(result));
    // Очищаем кэш календарных данных
    PropertiesService.getScriptProperties()
      .deleteProperty('CALENDAR_EVENT_DATA');

    return result;

  } catch (e) {
    Logger.error("Ошибка создания встречи:", e);
    throw new Error("Не удалось сохранить встречу. " + e.message);    
  }
}

function createRecords(recordsData) {
  Logger.log('[createRecords] Начало создания записей. Количество: %s', recordsData.length);
  
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName('Записи');
    const meetingId = PropertiesService.getScriptProperties().getProperty('currentMeetingId');
    Logger.log('[createRecords] ID встречи: %s', meetingId);

    const rows = recordsData.map((record, index) => {
      Logger.log('[createRecords] Обработка записи %s: %s', index + 1, JSON.stringify(record));
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
      const range = sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length);
      Logger.log('[createRecords] Запись диапазона: %s', range.getA1Notation());
      range.setValues(rows);
    }

    const result = `Сохранено ${rows.length} записей`;
    Logger.log('[createRecords] Успешно. %s', result);
    return result;

  } catch (e) {
    Logger.error('[createRecords] Ошибка: %s', e.toString());
    throw e;
  }
}

function getMeetingAttendees(meetingId) {
  Logger.log('[getMeetingAttendees] Начало обработки для meetingId: %s', meetingId);
  
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName('Встречи');
    const data = sheet.getDataRange().getValues();
    const header = data[0];
    
    const ID_COL = header.indexOf('ID встречи');
    const ATTENDEES_COL = header.indexOf('ID участников');
    
    Logger.log('[getMeetingAttendees] Индексы колонок: ID_COL=%s, ATTENDEES_COL=%s', ID_COL, ATTENDEES_COL);
    
    const meeting = data.find(row => row[ID_COL] === meetingId); // Исправить 79 на meetingId
    if (!meeting) {
      Logger.log('[getMeetingAttendees] Встреча не найдена');
      return [];
    }
    
    const attendeeIds = meeting[ATTENDEES_COL].split(', ').map(Number);
    Logger.log('[getMeetingAttendees] Найдено ID участников: %s', attendeeIds.join(', '));
    
    const employees = getEmployees();
    Logger.log('[getMeetingAttendees] Получено сотрудников: %s', employees.length);
    
    const result = employees
      .filter(e => attendeeIds.includes(Number(e.id)))
      .map(e => ({
        email: e.email,
        name: `${e.firstName} ${e.lastName}`.trim(),
        id: e.id
      }));
    
    Logger.log('[getMeetingAttendees] Результат: %s записей', result.length);
    return result;
    
  } catch (e) {
    Logger.log('[getMeetingAttendees] Ошибка: %s', e.toString());
    throw e;
  }
}