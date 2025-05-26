//файл: meeting.gs

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
const DEBUG_MODE = true;

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
      // id: row[columns.id],
      id: String(row[columns.id]), // Преобразуем ID в строку
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
      attendees: attendeeIds, // Добавляем ID участников в ответ
      success: true,
      invalidEmails: invalidEmails,
      message: invalidEmails.length > 0 
        ? `Встреча сохранена, но не найдены: ${invalidEmails.join(', ')}`
        : 'Встреча успешно сохранена'      
    };
    // Кэшируем данные участников
    PropertiesService.getScriptProperties()
      .setProperty('currentMeetingAttendees', JSON.stringify(attendeeIds));

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
    
    const attendeeIds = meeting[ATTENDEES_COL].split(', ');
    Logger.log('[getMeetingAttendees] Найдено ID участников: %s', attendeeIds.join(', '));
    
    const employees = getEmployees();
    Logger.log('[getMeetingAttendees] Получено сотрудников: %s', employees.length);
    
    const result = employees
      .filter(e => attendeeIds.includes(Number(e.id)))
      .map(e => ({
        email: e.email,
        name: `${e.firstName} ${e.lastName}`.trim(),
        id: String(e.id), // Гарантируем строковый формат
        firstName: e.firstName || 'Имя не указано', // Запасные значения
        lastName: e.lastName || 'Фамилия не указана'        
      }));
    
    Logger.log('[getMeetingAttendees] Результат: %s записей', result.length);
    return result;
    
  } catch (e) {
    Logger.log('[getMeetingAttendees] Ошибка: %s', e.toString());
    throw e;
  }
}

function getCurrentMeetingData() {
  const props = PropertiesService.getScriptProperties();
  return {
    id: props.getProperty('currentMeetingId'),
    number: props.getProperty('currentMeetingNumber')
  };
}

/**
 * Получает полные данные встречи по ID
 * @param {string} meetingId - ID встречи
 * @returns {Object} - Объект с данными встречи
 */
function getMeetingById(meetingId) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `meeting_${meetingId}`;
 
  // Проверка кэша
  const cachedData = cache.get(cacheKey);
  if (cachedData) {
    console.log('Данные из кэша:', cachedData);
    return JSON.parse(cachedData);
  }

  const sheet = SpreadsheetApp.getActive().getSheetByName('Встречи');
  const data = sheet.getDataRange().getValues();
  // Явная проверка доступа
  if (!sheet.getSheetId()) {
    throw new Error('Нет доступа к листу "Встречи"');
  }
  // Нормализуем заголовки
  const headers = data[0].map(h => h.trim().toLowerCase()); // Все в нижний регистр

  // Получаем индексы с правильными названиями
  const getHeaderIndex = (name) => {
    const normalized = name.trim().toLowerCase();
    return headers.findIndex(h => h === normalized);
  };

  const idIndex = getHeaderIndex('ID встречи');
  const numberIndex = getHeaderIndex('Номер встречи');
  const dateIndex = getHeaderIndex('Дата');
  const topicIndex = getHeaderIndex('Тема встречи');
  const attendeesIndex = getHeaderIndex('Участники');
  const attendeesIdIndex = getHeaderIndex('ID участников');
  const locationIndex = getHeaderIndex('Место встречи');

  const meeting = data.find(row => {
    const rowId = row[idIndex]?.toString().trim();
    // console.log('Проверка ID:', rowId);
    return rowId === meetingId;
  });

  // Преобразование даты с обработкой ошибок
  let meetingDate;
  try {
    meetingDate = new Date(meeting[dateIndex]);
    if (isNaN(meetingDate.getTime())) {
      throw new Error('Некорректный формат даты');
    }
  } catch(e) {
    console.error('Ошибка парсинга даты:', e.message);
    meetingDate = new Date(); // Значение по умолчанию
  }

  // Проверяем индексы
  if (idIndex === -1) throw new Error('Не найдена колонка "ID встречи"');

  if (!meeting) {
    console.error('Встреча не найдена. Полученные данные:', data);
    throw new Error('Встреча не найдена');
  }

  // Логируем сырые данные
  console.log('Найдена запись:', meeting);

  // Форматируем данные
  const result = {
    id: meetingId,
    number: meeting[numberIndex]?.toString(),
    date: Utilities.formatDate(
      meetingDate, 
      Session.getScriptTimeZone(), 
      "dd.MM.yyyy HH:mm"
    ),
    topic: meeting[topicIndex],
    location: meeting[locationIndex],
    attendees: {
      ids: meeting[attendeesIdIndex]?.split(',').map(id => id.trim()) || [],
      emails: meeting[attendeesIndex]?.split(',').map(e => e.trim()) || []
    },
    meta: {
      created: Utilities.formatDate(
        new Date(), 
        Session.getScriptTimeZone(), 
        "yyyy-MM-dd'T'HH:mm:ss'Z'"
      ),
      lastModified: Utilities.formatDate(
        new Date(), 
        Session.getScriptTimeZone(), 
        "yyyy-MM-dd'T'HH:mm:ss'Z'"
      )
    }
  };
  console.log('Результат:', JSON.stringify(result, null, 2));
  
  // Кэшируем на 5 минут
  cache.put(cacheKey, JSON.stringify(result), 300);
  return result;
}

function testGetMeeting() {
  getMeetingById('ae76e9b4-368d-4c85-91a6-30dd665ded4f');
}

/**
 * Получает все записи для встречи
 * @param {string} meetingId - ID встречи
 * @returns {Array} - Массив объектов записей
 */
function getRecordsByMeetingId(meetingId) {
  console.log('Начало обработки для meetingId:', meetingId);
  const cache = CacheService.getScriptCache();
  const cacheKey = `records_${meetingId}`;
  
  // Проверка кэша
  const cachedData = cache.get(cacheKey);
  if (cachedData) {
    return JSON.parse(cachedData);
  }

  const sheet = SpreadsheetApp.getActive().getSheetByName('Записи');
  if (!sheet) {
    console.error('Лист "Записи" не найден');
    throw new Error('Лист записей отсутствует');
  }
  const data = sheet.getDataRange().getValues();
  console.log('Всего записей в листе:', data.length);
  const headers = data[0].map(h => h.toLowerCase().trim());
  console.log('Заголовки:', headers);

  const meetingIdIndex = headers.indexOf('id встречи');
  console.log('Индекс ID встречи:', meetingIdIndex);

  const filtered = data.filter(row => 
    row[meetingIdIndex]?.toString().trim() === meetingId
  );
  console.log('Найдено записей:', filtered.length);

  const parseDate = (dateValue) => {
    try {
      // Если значение уже является объектом Date
      if (dateValue instanceof Date) {
        return dateValue;
      }
      
      // Если это строка формата DD-MM-YYYY
      if (typeof dateValue === 'string') {
        const [day, month, year] = dateValue.split('-');
        return new Date(year, month-1, day);
      }
      
      // Для других случаев (timestamp, etc)
      return new Date(dateValue);
    } catch(e) {
      console.error('Ошибка парсинга даты:', e);
      return new Date(); // Возвращаем текущую дату как fallback
    }
  };

  // Проверка индексов колонок
  const getHeaderIndex = (name) => {
    const index = headers.indexOf(name.toLowerCase().trim());
    if (index === -1) throw new Error(`Колонка "${name}" не найдена`);
    return index;
  };

  const recordIdIndex = getHeaderIndex('ID записи');
  const typeIndex = getHeaderIndex('Запись');
  const textIndex = getHeaderIndex('Текст записи');
  const dateIndex = getHeaderIndex('Срок');
  const responsibleIndex = getHeaderIndex('Ответственные ID'); 
  const importanceIndex = getHeaderIndex('Значимость');
  const priorityIndex = getHeaderIndex('Приоритет');
  const completedIndex = getHeaderIndex('Выполнено');

  const records = filtered.map((row, idx) => {
    console.log(`Обработка строки ${idx + 1}:`, JSON.stringify(row));
    // Логирование сырых данных даты
    const rawDate = row[dateIndex];
    console.log('Тип данных даты:', typeof rawDate);
    console.log('Значение даты:', rawDate);

    return {
      id: row[recordIdIndex]?.toString(),
      type: row[typeIndex],
      text: row[textIndex],
      dueDate: (() => {
        try {
          const date = parseDate(rawDate);
          console.log('Успешно распарсено:', date);
          return Utilities.formatDate(
            date, 
            Session.getScriptTimeZone(), 
            "dd.MM.yyyy"
          );
        } catch(e) {
          console.error('Финальная обработка даты:', e);
          return 'Некорректная дата';
        }
      })(),  
      responsible: (() => {
        const raw = row[responsibleIndex]?.toString().trim();
        if (!raw) return [];
        
        return raw.split(';').map(r => {
          const match = r.match(/(.*?)\s\((.*?)\)/);
          return match ? { 
            name: match[1].trim(), 
            email: match[2].trim() 
          } : {name: r.trim(), email: ''};
        });
      })(),      
      status: {
        importance: row[importanceIndex] || 'Не указано',
        priority: row[priorityIndex] || 'Не указано',
        completed: row[completedIndex]?.toString().toLowerCase()
      }
    }
  });

  // Кэшируем на 5 минут
  // cache.put(cacheKey, JSON.stringify(records), 300);
  console.log('Итоговые записи:', JSON.stringify(records, null, 2));
  return records;
}

function testGetRecords() {
  getRecordsByMeetingId('e845a1dc-1c75-474f-b495-d930cd60f5d8');
}
