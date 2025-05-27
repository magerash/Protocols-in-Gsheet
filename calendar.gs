//файл: calendar.gs

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
    })
    .filter(e => e.participants.length > 0);
    
  } catch(e) {
    console.error('Calendar API Error:', e);
    throw new Error('Ошибка загрузки событий календаря');
  }
}

function selectEvent(eventId) {
  try {
    const event = getEventById(eventId); // Используйте существующую функцию поиска
    const creatorEmail = event.getCreators()[0];
    console.log('Calendar Event Data:', {
      title: event.getTitle(),
      start: event.getStartTime(),
      attendees: event.getGuestList().map(g => g.getEmail())
    });
    const data = {
      title: event.getTitle(),
      startTime: event.getStartTime().toISOString(),
      endTime: event.getEndTime().toISOString(),
      location: event.getLocation(),
      attendees: [creatorEmail, ...event.getGuestList().map(g => g.getEmail())] // Все email напрямую
    };
    return data;
  } catch(e) {
    throw new Error("Не удалось загрузить данные встречи: " + e.message);
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
  try {
    PropertiesService.getScriptProperties()
      .setProperty('SELECTED_EVENT', JSON.stringify(data));
    return true; // Явно возвращаем результат
  } catch (e) {
    console.error('Save error:', e);
    return false;
  }
}

function getSelectedEventData() {
  try {
    const data = PropertiesService.getScriptProperties()
      .getProperty('CALENDAR_EVENT_DATA');
    
    // Возвращаем структуру по умолчанию если данных нет
    return data || JSON.stringify({
      title: '',
      startTime: new Date().toISOString(),
      attendees: [],
      location: ''
    });
    
  } catch(e) {
    console.error('Error in getSelectedEventData:', e);
    // Всегда возвращаем валидный JSON
    return JSON.stringify({});
  }
}

function clearSelectedEventData() {
  PropertiesService.getScriptProperties()
    .deleteProperty('SELECTED_EVENT');
}
