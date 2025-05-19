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

function showRecordDialog(meetingId, meetingNumber) {
  var html = HtmlService.createHtmlOutputFromFile('recordForm')
    .setWidth(800)
    .setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, 'Записи встречи');
  PropertiesService.getScriptProperties().setProperty('currentMeetingId', meetingId);
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