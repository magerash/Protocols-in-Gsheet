// файл: dialogs.gs

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Протоколы')
    .addItem('Создать протокол встречи', 'showMeetingTypeSelector')
    .addItem('Окно записей', 'showRecordDialog')
    .addSeparator()
    .addItem('Обновить кэш сотрудников', 'clearEmployeeCache') // Новая кнопка
    .addToUi();
}

function showMeetingDialog() {
  // Явная очистка данных при открытии не из календаря
  PropertiesService.getScriptProperties().deleteProperty('CALENDAR_EVENT_DATA');

  var html = HtmlService.createHtmlOutputFromFile('meetingForm')
    .setWidth(600)
    .setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, 'Новая встреча');
}

function showMeetingDialogWithData(data) {
  console.log('Received calendar data:', JSON.stringify(data));
  // Принудительная очистка предыдущих данных
  PropertiesService.getScriptProperties().deleteProperty('CALENDAR_EVENT_DATA');

  // Явно преобразуем даты в строки
  const preparedData = {
    title: data.title,
    startTime: new Date(data.startTime).toISOString(),
    attendees: data.attendees || [],
    location: data.location || ""
  };
  
  PropertiesService.getScriptProperties()
    .setProperty('CALENDAR_EVENT_DATA', JSON.stringify(preparedData));

  console.log('Saved to properties:', 
  PropertiesService.getScriptProperties().getProperty('CALENDAR_EVENT_DATA'));
  // Задержка для гарантии сохранения данных
  Utilities.sleep(1000);
  
  const html = HtmlService.createHtmlOutputFromFile('meetingForm')
    .setWidth(600)
    .setHeight(650);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Новая встреча');
}

function getCachedAttendees() {
  const props = PropertiesService.getScriptProperties();
  return JSON.parse(props.getProperty('currentMeetingAttendees') || []);
}

function saveMeetingAttendees(attendees) {
  PropertiesService.getScriptProperties()
    .setProperty('currentMeetingAttendees', JSON.stringify(attendees));
}

function showRecordDialog(meetingData = '') {
  const props = PropertiesService.getScriptProperties();
  
  try {
    if (meetingData && typeof meetingData === 'object') {
      props.setProperty('currentMeetingId', meetingData.id || '');
      props.setProperty('currentMeetingNumber', meetingData.number?.toString() || '');
      props.setProperty('currentMeetingAttendees', JSON.stringify(meetingData.attendees?.ids || []));
    } else {
      // Получаем свойства по одному
      const currentMeetingId = props.getProperty('currentMeetingId');
      
      if (!currentMeetingId) {
        throw new Error('Данные встречи не найдены');
      }
    }

    // Открываем диалог
    const html = HtmlService.createHtmlOutputFromFile('recordForm')
      .setWidth(800)
      .setHeight(650);
    SpreadsheetApp.getUi().showModalDialog(html, 'Записи встречи');

  } catch(e) {
    console.error('Dialog open error:', e);
    throw new Error('Ошибка открытия окна записей: ' + e.message);
  }
}

function getCurrentMeetingData() {
  const props = PropertiesService.getScriptProperties();
  return {
    id: props.getProperty('currentMeetingId'),
    number: props.getProperty('currentMeetingNumber'),
    attendees: JSON.parse(props.getProperty('currentMeetingAttendees') || [])
  };
}

function showCalendarEventsModal() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('calendarEventsModal')
      .setWidth(500)
      .setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(html, 'Выбор события из календаря');
  } catch(e) {
    console.error('Ошибка открытия окна:', e);
    throw new Error('Не удалось открыть окно календаря');
  }
}

function showMeetingTypeSelector() {
  var html = HtmlService.createHtmlOutputFromFile('meetingTypeSelector')
    .setWidth(650)
    .setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(html, 'Создание встречи');
}