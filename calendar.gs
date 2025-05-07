// ==== Google meet ====

// function getMeetingAttendees(meetingId) {
//   try {
//     var sheet = SpreadsheetApp.getActive().getSheetByName('Встречи');
//     if (!sheet) throw new Error('Лист "Встречи" не найден');
    
//     var data = sheet.getDataRange().getValues();
//     if (data.length < 2) return []; // Если нет данных кроме заголовков
    
//     var headers = data[0];
//     var idCol = headers.indexOf('ID встречи');
//     var attendeeIdsCol = headers.indexOf('ID участников');
    
//     if (idCol === -1 || attendeeIdsCol === -1) {
//       throw new Error('Не найдены необходимые колонки');
//     }
    
//     for (var i = 1; i < data.length; i++) {
//       if (data[i][idCol] === meetingId) {
//         var attendeeIds = data[i][attendeeIdsCol];
//         return attendeeIds ? attendeeIds.toString().split(',').map(id => id.trim()).filter(id => id) : [];
//       }
//     }
//     return [];
//   } catch (e) {
//     console.error('Ошибка в getMeetingAttendees: ', e);
//     return [];
//   }
// }

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

function getUpcomingMeetings(startDate, endDate) {
  try {
    const employees = getEmployees();
    const emailMap = new Map();
    
    // Создаем карту email → "Имя Фамилия"
    employees.forEach(emp => {
      if (emp.email && emp.firstName && emp.lastName) {
        const fullName = `${emp.firstName} ${emp.lastName}`.trim();
        emailMap.set(emp.email.toLowerCase(), fullName);
      }
    });

    const calendar = CalendarApp.getDefaultCalendar();
    const events = calendar.getEvents(new Date(startDate), new Date(endDate));
    
    return events.map(e => {
      const creatorEmail = e.getCreators()[0].toLowerCase();
      const organizer = emailMap.get(creatorEmail) || creatorEmail.split('@')[0];

      const participants = e.getGuestList()
        .filter(g => g.getGuestStatus() !== CalendarApp.GuestStatus.NO)
        .map(g => {
          const email = g.getEmail().toLowerCase();
          return emailMap.get(email) || email.split('@')[0];
        });

      return {
        id: e.getId().split('@')[0],
        title: e.getTitle(),
        startTime: e.getStartTime()?.toISOString(),
        endTime: e.getEndTime()?.toISOString(),
        location: e.getLocation(),
        timeZone: Session.getScriptTimeZone(),
        organizer: organizer,
        participants: participants
      };
    });
  } catch(e) {
    console.error('Calendar API Error:', e);
    throw new Error('Ошибка загрузки событий календаря');
  }
}


// Вспомогательная функция
function getEventById(eventId) {
  const calendar = CalendarApp.getDefaultCalendar();
  const events = calendar.getEvents(
    new Date(Date.now() - 30*24*60*60*1000),
    new Date(Date.now() + 30*24*60*60*1000)
  );
  
  const foundEvent = events.find(e => 
    e.getId().split('@')[0] === eventId
  );
  
  if (!foundEvent) throw new Error('Событие не найдено');
  return foundEvent;
}

function formatDateTime(date) {
  return new Date(date)
    .toISOString()
    .slice(0, 16)
    .replace('T', ' ');
}

function saveSelectedEventData(data) {
  PropertiesService.getScriptProperties()
    .setProperty('SELECTED_EVENT', JSON.stringify(data));
}

function getSelectedEventData() {
  const data = PropertiesService.getScriptProperties()
    .getProperty('SELECTED_EVENT');
  return data ? JSON.parse(data) : null;
}

function clearSelectedEventData() {
  PropertiesService.getScriptProperties()
    .deleteProperty('SELECTED_EVENT');
}
