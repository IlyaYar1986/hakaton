// Развернуть как веб-приложение от имени владельца таблицы.
const SPREADSHEET_ID = '1r5-hsnzBsU0tZzM8JGgqcF2dYw4hkSx2_hwACWsTZjs';
const SHEET_NAME = 'Илья Чурилов';

function result_(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}
function doGet() {
  try { return result_({ok:true, slots:readSlots_().map(s=>({id:s.id,week:s.week,day:s.day,time:s.time,occupied:s.occupied,confirmed:s.status==='Подтверждён'}))}); }
  catch (error) { return result_({ok:false, error:'Не удалось загрузить расписание.'}); }
}
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const data = JSON.parse(e.postData.contents);
    lock.waitLock(20000);
    return result_(change_(data));
  } catch (error) {
    return result_({ok:false, error:'Не удалось сохранить запись. Обновите страницу и попробуйте снова.'});
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}
function sheet_() { return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME); }
function readSlots_() {
  const rows = sheet_().getDataRange().getDisplayValues();
  let week='', day='', inSlots=false;
  const result=[];
  rows.forEach((cells, index) => {
    const a=String(cells[0]||'').trim();
    if (/^НЕДЕЛЯ\s+\d/i.test(a)) { week=a; inSlots=false; }
    else if (/^(ПОНЕДЕЛЬНИК|ВТОРНИК|СРЕДА|ЧЕТВЕРГ|ПЯТНИЦА|СУББОТА|ВОСКРЕСЕНЬЕ),/i.test(a)) { day=a; inSlots=false; }
    else if (a==='№ слота') { inSlots=true; }
    else if (inSlots && /^\d+$/.test(a)) {
      const time=String(cells[1]||'').trim();
      if (time) result.push({id:index+1,week:week,day:day,time:time,team:String(cells[2]||''),name:String(cells[3]||''),status:String(cells[4]||''),comment:String(cells[5]||''), occupied:Boolean(cells[2]||cells[3]||cells[4]||cells[5])});
    } else if (a) { inSlots=false; }
  });
  return result;
}
function digest_(token) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,token,Utilities.Charset.UTF_8).map(n => ('0'+((n+256)%256).toString(16)).slice(-2)).join('');
}
function change_(data) {
  if (!data || !['book','cancel'].includes(data.action) || !Number.isInteger(data.id) || !/^[a-f0-9]{64}$/.test(data.token||'')) return {ok:false,error:'Некорректный запрос.'};
  const slot=readSlots_().find(s=>s.id===data.id);
  if (!slot) return {ok:false,error:'Слот больше недоступен.'};
  const props=PropertiesService.getScriptProperties();
  const key='slot_'+SPREADSHEET_ID+'_'+slot.id, previous=props.getProperty(key);
  const cell=sheet_().getRange(slot.id,3,1,4);
  if (data.action==='book') {
    if (slot.occupied) return {ok:false,error:'Этот слот уже занят. Обновите расписание.'};
    const team=String(data.team||'').trim(), name=String(data.name||'').trim(), comment=String(data.comment||'').trim();
    if (!/^\d{1,3}$/.test(team) || !name || name.length>120 || comment.length>300) return {ok:false,error:'Укажите номер и название команды.'};
    // Предотвращаем внедрение формул через поля формы.
    const safe=s=>/^[=+\-@]/.test(s)?"'"+s:s;
    props.setProperty(key,JSON.stringify({hash:digest_(data.token),week:slot.week,day:slot.day,time:slot.time,team:team,name:name,comment:comment}));
    try { cell.setValues([[team,safe(name),'Забронирован',safe(comment)]]); SpreadsheetApp.flush(); }
    catch(error) { props.deleteProperty(key); throw error; }
    return {ok:true};
  }
  let saved; try { saved=JSON.parse(previous); } catch (_) { saved=null; }
  if (!slot.occupied || !saved || saved.hash!==digest_(data.token) || saved.week!==slot.week || saved.day!==slot.day || saved.time!==slot.time || saved.team!==slot.team || saved.name!==slot.name || saved.comment!==slot.comment) return {ok:false,error:'Запись изменилась или принадлежит другому браузеру. Обратитесь к эксперту.'};
  if (slot.status && slot.status!=='Забронирован') return {ok:false,error:'Запись уже подтверждена, для изменения обратитесь к трекеру в чате.'};
  cell.clearContent(); SpreadsheetApp.flush(); props.deleteProperty(key);
  return {ok:true};
}
