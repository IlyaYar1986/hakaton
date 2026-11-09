// Добавить отдельным файлом в тот же проект Apps Script, что и Code.gs.
// После добавления выбрать функцию setupHackathonCalendar и запустить один раз.
const CALENDAR_NAME = 'Хакатон';
const CALENDAR_YEAR = 2026;
const CALENDAR_TIMEZONE = 'Europe/Moscow';
const EVENT_MINUTES = 60;
const CALENDAR_PROPERTY = 'hackathon_calendar_id_v1';
const EVENT_PREFIX = 'hackathon_event_v1_';

function setupHackathonCalendar() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty(CALENDAR_PROPERTY);
  let calendar = id && CalendarApp.getCalendarById(id);
  if (id && !calendar) throw new Error('Не найден ранее настроенный календарь. Проверьте доступ.');
  if (!calendar) {
    const found = CalendarApp.getOwnedCalendarsByName(CALENDAR_NAME);
    if (found.length) throw new Error('Календарь «Хакатон» уже существует. Для защиты существующих событий настройка остановлена.');
    calendar = CalendarApp.createCalendar(CALENDAR_NAME, {timeZone: CALENDAR_TIMEZONE});
    props.setProperty(CALENDAR_PROPERTY, calendar.getId());
  }
  calendar.setColor(CalendarApp.Color.ORANGE);
  calendar.setTimeZone(CALENDAR_TIMEZONE);
  syncHackathonCalendar();
  const triggers = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'syncHackathonCalendar');
  if (!triggers.length) ScriptApp.newTrigger('syncHackathonCalendar').timeBased().everyMinutes(5).create();
  Logger.log('Календарь настроен: ' + calendar.getId());
}

function calendarSlotDate_(day, time) {
  const months = {'января':1,'февраля':2,'марта':3,'апреля':4,'мая':5,'июня':6,'июля':7,'августа':8,'сентября':9,'октября':10,'ноября':11,'декабря':12};
  const dateMatch = String(day).toLowerCase().match(/,\s*(\d{1,2})\s+([а-яё]+)/);
  const timeMatch = String(time).match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
  if (!dateMatch || !timeMatch || !months[dateMatch[2]]) throw new Error('Не удалось разобрать дату или время: ' + day + ' ' + time);
  const month=months[dateMatch[2]], date=Number(dateMatch[1]), hour=Number(timeMatch[1]), minute=Number(timeMatch[2]);
  if (date < 1 || date > 31 || hour > 23 || minute > 59) throw new Error('Некорректная дата или время.');
  // Москва UTC+3 без перехода на летнее время.
  const start = new Date(Date.UTC(CALENDAR_YEAR, month-1, date, hour-3, minute));
  const parts=Utilities.formatDate(start,CALENDAR_TIMEZONE,'yyyy-MM-dd HH:mm');
  const expected=CALENDAR_YEAR+'-'+String(month).padStart(2,'0')+'-'+String(date).padStart(2,'0')+' '+String(hour).padStart(2,'0')+':'+String(minute).padStart(2,'0');
  if (parts!==expected) throw new Error('Несуществующая дата: '+day);
  return start;
}

function syncHackathonCalendar() {
  const lock=LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const props=PropertiesService.getScriptProperties();
    const id=props.getProperty(CALENDAR_PROPERTY);
    const calendar=id && CalendarApp.getCalendarById(id);
    if (!calendar) throw new Error('Сначала запустите setupHackathonCalendar.');
    const slots=readSlots_();
    if (!slots.length) throw new Error('В таблице не найдено ни одного слота; синхронизация остановлена.');
    const desired={};
    slots.filter(s=>s.status==='Подтверждён').forEach(s=>{
      if (!s.team) throw new Error('В подтверждённой строке '+s.id+' не указан номер команды.');
      const start=calendarSlotDate_(s.day,s.time);
      const key=EVENT_PREFIX+digest_(s.week+'|'+s.day+'|'+s.time);
      if (desired[key]) throw new Error('Обнаружены совпадающие слоты: '+s.day+' '+s.time);
      desired[key]={title:'Хакатон · № '+s.team+(s.name?' · '+s.name:''),start:start,end:new Date(start.getTime()+EVENT_MINUTES*60000),description:'Встреча с командой. Источник: лист «Илья Чурилов», строка '+s.id+'.\n'+(s.comment||'')};
    });
    const all=props.getProperties();
    // Ошибки чтения и проверки останавливают синхронизацию до удаления событий.
    Object.keys(desired).forEach(key=>{
      const item=desired[key], existing=all[key] && calendar.getEventById(all[key]);
      if (existing) { existing.setTitle(item.title); existing.setTime(item.start,item.end); existing.setDescription(item.description); }
      else { const event=calendar.createEvent(item.title,item.start,item.end,{description:item.description}); props.setProperty(key,event.getId()); }
    });
    Object.keys(all).filter(key=>key.startsWith(EVENT_PREFIX) && !desired[key]).forEach(key=>{
      const event=calendar.getEventById(all[key]);
      if (event) event.deleteEvent();
      props.deleteProperty(key);
    });
  } finally { lock.releaseLock(); }
}
