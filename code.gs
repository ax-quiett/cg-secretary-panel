/** Code.gs — Capital Group · Supervisor & Secretary Panel

 * (final+fix · 2025-11-12 · PERF: faster echo ~150ms; SHIFTS cleaned; ARRIVE/LUNCH/BREAK aggregates; weekend 10-21; holidays header; DASHBOARD sheet)

 * Base: final+fix · 2025-11-11 · HARDENED+flush+echo+fresh-read

 */



const SS = SpreadsheetApp.getActive();

const TZ = Session.getScriptTimeZone() || "Europe/Moscow";

const S  = { CFG:'Config', STAFF:'Staff', TOK:'Tokens', STAT:'Statuses', ABS:'Absences',

             HOL:'Holidays', PLAN:'Planner', PLANW:'WeekDraft', ROSTER:'Roster', QC:'QC', ERR:'Errors', DASH:'Dashboard' };



const VALID_STATUSES = new Set(['OFF','На линии','Обед','Перерыв']);



// ---------- utils ----------

function sh_(n){ var s=SS.getSheetByName(n); if(!s){ s=SS.insertSheet(n); } return s; }

/** НЕ очищаем лист — только заголовки. */

function header_(n, heads){

  var sh=sh_(n);

  var w=heads.length;

  var cur=sh.getRange(1,1,1,Math.max(w, sh.getLastColumn()||w)).getValues()[0] || [];

  var same=true;

  for(var i=0;i<w;i++){ if(String(cur[i]||'')!==String(heads[i])){ same=false; break; } }

  if(!same){ sh.getRange(1,1,1,w).setValues([heads]); }

}

function data_(n,c){

  var sh=sh_(n); var lr=sh.getLastRow();

  if(lr<2) return [];

  return sh.getRange(2,1,lr-1,c).getValues();

}

function putRows_(n, rows, heads){

  var sh=sh_(n);

  sh.clear();

  sh.getRange(1,1,1,heads.length).setValues([heads]);

  if(rows.length) sh.getRange(2,1,rows.length,heads.length).setValues(rows);

}

function iso(d){ return Utilities.formatDate(new Date(d), TZ, "yyyy-MM-dd"); }

function hhmm(d){ return Utilities.formatDate(new Date(d), TZ, "HH:mm"); }

function nowIso(){ return Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ss"); }

function toHHmm_(v){

  if(!v) return '';

  var t = (v instanceof Date) ? v : (new Date(v));

  return Utilities.formatDate(t, TZ, "HH:mm");

}

function minDiffHHmm_(a,b){

  function p(x){ var m=/^(\d{2}):(\d{2})$/.exec(String(x||"")); return m? (+m[1]*60+ +m[2]) : null; }

  var pa=p(a), pb=p(b); if(pa==null||pb==null) return null; return pa-pb;

}

function normHHmm_(v){

  if(!v) return '';

  if (v instanceof Date) return Utilities.formatDate(v, TZ, "HH:mm");

  var s=String(v).trim();

  var m = /^(\d{1,2}):(\d{2})/.exec(s);

  if(m){ var h=('0'+(+m[1])).slice(-2), mm=('0'+(+m[2])).slice(-2); return h+':'+mm; }

  if(/^\d+$/.test(s)){ var mins=+s; var h2=('0'+Math.floor(mins/60)).slice(-2), mm2=('0'+(mins%60)).slice(-2); return h2+':'+mm2; }

  return s;

}



// ---------- locking ----------

function withLock_(ms, fn){

  var lock = LockService.getScriptLock();

  lock.waitLock(ms||10000);

  try{ return fn(); }

  finally{ try{ lock.releaseLock(); }catch(_e){} }

}

function withUserLock_(userKey, ms, fn){ return withLock_(ms, fn); }



// ---------- names ----------

function normName_(s){ return String(s||'').replace(/\s+/g,' ').trim().toLowerCase(); }

function staff_(){

  return data_(S.STAFF,3).map(function(r){ return ({

    name:String(r[0]||'').trim(),

    norm:normName_(r[0]),

    isSenior:r[1]===true,

    isWeekend:r[2]===true

  }); }).filter(function(x){ return x.name; });

}

function staffMap_(){ var map=new Map(); staff_().forEach(function(s){ map.set(s.norm, s.name); }); return map; }

function canonicalName_(input){

  var m = staffMap_(); var nn = normName_(input);

  return m.get(nn) || String(input||'').trim();

}

function senior_(){ return staff_().find(function(x){return x.isSenior;}); }

function weekend_(){ return staff_().find(function(x){return x.isWeekend;}); }



// ---------- bootstrap ----------

function ensure_(){

  header_(S.CFG,  ["KEY","VALUE"]);

  header_(S.STAFF,["Name","IsSenior","IsWeekendSecretary"]);

  header_(S.TOK,  ["Token","Role","Name","Active","AccessURL"]);

  header_(S.STAT, ["Name","Status","UpdatedAt"]);

  header_(S.ABS,  ["Date","Name","Type"]);

  header_(S.HOL,  ["Date","Title"]);

  header_(S.PLAN, ["Date","Name","Start","End","Role","Source"]);

  header_(S.PLANW,["Date","Name","Shift"]);

  header_(S.ROSTER,["Date","Name","Start","End","Role"]);

  header_(S.QC,   ["Employee","Reviewer","Date","CallID","Tone","Script","Parasites","Comment","CreatedAt"]);

  header_(S.ERR,  ["Ts","Level","Source","Action","User","Payload","Message","Stack","ErrId"]);

  header_(S.DASH, ["PeriodFrom","PeriodTo","Name","DaysPresent","LateCount","LateFromLunch","LongBreaks","VacationDays","SickDays","QC_Avg","FirstSeen","LastSeen"]);



  var st=sh_(S.STAFF);

  if(st.getLastRow()<2){

    st.getRange(2,1,5,3).setValues([

      ["Старший Секретарь", true,  false],

      ["Анна",              false, false],

      ["Мария",             false, false],

      ["Ольга",             false, false],

      ["Выходной Секретарь",false, true ],

    ]);

  }

}



// ---------- config ----------

function cfgGet(k){ var m={}; data_(S.CFG,2).forEach(function(r){ m[String(r[0])]=r[1]; }); return k?m[k]:m; }

function cfgSet(k,v){

  var sh=sh_(S.CFG); var rows=data_(S.CFG,2);

  for(var i=0;i<rows.length;i++){ if(String(rows[i][0])===k){ sh.getRange(i+2,2).setValue(v); return; } }

  sh.appendRow([k,v]);

}



// ---------- logging ----------

function log_(level, src, act, payload, msg){

  try{

    var id=(Utilities.getUuid()||"").slice(0,8).replace(/-/g,"");

    sh_(S.ERR).appendRow([nowIso(), level, src, act, '', JSON.stringify(payload||{}), String(msg||''), '', id]);

    return id;

  }catch(e){ return ""; }

}

function logErr_(src,act,payload,err,user){

  try{

    var id=(Utilities.getUuid()||"").slice(0,8).replace(/-/g,"");

    sh_(S.ERR).appendRow([nowIso(),"ERROR",src,act,user||"",JSON.stringify(payload||{}),(err && (err.message||String(err)))||"", (err && err.stack)||"", id]);

    return id;

  }catch(e){}

  return "";

}

function logWarn_(src,act,payload,msg){ return log_("WARN", src, act, payload, msg); }

function safe_(src,act,payload,fn,user){

  try{ ensure_(); log_("INFO", src, act+":start", payload, "start"); var res = fn(); log_("INFO", src, act+":ok", payload, "ok"); return res; }

  catch(e){ return {ok:false,errId:logErr_(src,act,payload,e,user),message:String(e)}; }

}

function api_frontLog(source, action, message, stack, payload){

  return safe_("front","clientLog",{source,action,message,stack,payload},function(){

    var id=logErr_(source||'CLIENT', action||'', payload||{}, {message:message,stack:stack}, '');

    return {ok:true,errId:id};

  });

}



// ---------- tokens / auth ----------

function validateToken_(t){

  var rows=data_(S.TOK,5);

  if(!rows.length) return {ok:true,role:"supervisor",name:"Demo",token:""}; // демо вход

  var r=rows.find(function(x){ return String(x[0])===String(t) && x[3]===true; });

  if(!r) return {ok:false,error:"token"};

  return {ok:true,role:String(r[1]),name:String(r[2]),token:String(t)};

}

function api_tokens(){

  return safe_("data","tokens",{},function(){

    var base=cfgGet('WEBAPP_URL')||'';

    return data_(S.TOK,5).map(function(r){ return ({

      token:String(r[0]), role:String(r[1]), name:String(r[2]),

      active:r[3]===true,

      url:String(r[4]||'')|| (base?(base+(base.includes('?')?'&':'?')+'token='+String(r[0])):'')

    }); });

  });

}

function api_who(token){ return safe_("auth","who",{hasToken: !!token}, function (){ return validateToken_(token||''); }); }



// ---------- STATUSES ----------

function readStatusesFresh_(){

  var reopened = SpreadsheetApp.openById(SS.getId());

  var sh = reopened.getSheetByName(S.STAT) || reopened.insertSheet(S.STAT);

  var lr = sh.getLastRow();

  if(lr<2) return [];

  return sh.getRange(2,1,lr-1,3).getValues();

}

/** вспомогательно: обычное чтение последнего статуса */

function getStatusByName_(canonName){

  var key = normName_(canonName);

  var last = null;

  data_(S.STAT,3).forEach(function(r){

    if(normName_(r[0])!==key) return;

    var st = VALID_STATUSES.has(String(r[1])) ? String(r[1]) : 'OFF';

    last = {name: String(r[0]), status: st, updated: String(r[2]||'')};

  });

  return last;

}

/** «свежее» чтение последнего статуса (самая поздняя запись) */

function getStatusByNameFresh_(canonName){

  var key = normName_(canonName);

  var rows = readStatusesFresh_();

  for(var i=rows.length-1;i>=0;i--){

    var r = rows[i];

    if(normName_(String(r[0]||'')) === key){

      var status = String(r[1]||'');

      if(!VALID_STATUSES.has(status)){ status = 'OFF'; }

      return { name: String(r[0]||''), status: status, updated: String(r[2]||'') };

    }

  }

  return { name: canonicalName_(canonName), status: 'OFF', updated: '' };

}



function setStatus_(name,status){

  var canon = canonicalName_(name);

  var nkey  = normName_(canon);

  var stVal = VALID_STATUSES.has(status) ? status : 'OFF';

  log_("INFO","status","set_attempt",{name:canon,input:status,validated:stVal},"setting status");



  return withUserLock_(nkey, 10000, function () {

    header_(S.STAT, ["Name","Status","UpdatedAt"]);

    var sh=sh_(S.STAT);

    var rows=data_(S.STAT,3);

    var now=nowIso();



    // если уже такой же статус — ничего не пишем

    for(var i=rows.length-1;i>=0;i--){

      if(normName_(rows[i][0])===nkey){

        var cur = String(rows[i][1]||'');



        // PATCH: при переходе в «На линии» логируем «IN» метку (если реально смена статуса)

        try{

          if(stVal==="На линии" && cur!=="На линии"){

            var dnow = new Date(); var dIso = iso(dnow); var hh = hhmm(dnow);

            sh_(S.PLAN).appendRow([dIso, canon, hh, hh, "IN", "auto"]);

          }

        }catch(_e){}



        if(cur===stVal){

          SpreadsheetApp.flush();

          Utilities.sleep(150);

          var echo = getStatusByNameFresh_(canon) || {};

          log_("INFO","status","echo",{name:canon, asked:stVal, echoStatus:echo.status||"", echoUpdated:echo.updated||""},"ok");

          return {ok:true, name:canon, status:stVal, rowWritten:i+2, sheet:S.STAT, lastRow:sh.getLastRow(), skipped:true};

        }

        break;

      }

    }



    // дисциплина: опоздание (мягкий лог)

    try{

      if(stVal==="На линии"){

        var d = iso(new Date());

        var roster = rosterShiftsFor_(d).find(function(x){ return normName_(x.name)===nkey; });

        if(roster && roster.start){

          var nowHH = hhmm(new Date());

          var late = minDiffHHmm_(nowHH, roster.start);

          if(late!=null && late>5 && late<=240){

            logWarn_("discipline","late",{name:canon, date:d, start:roster.start, at:nowHH, lateMin:late}, "Опоздание: "+late+" мин");

          }

        }

      }

    }catch(_e){}



    // запись/перезапись

    var fresh = data_(S.STAT,3);

    for(var j=fresh.length-1;j>=0;j--){

      if(normName_(fresh[j][0])===nkey){

        sh.getRange(j+2,1,1,3).setValues([[canon,stVal,now]]);

        SpreadsheetApp.flush();



        // PATCH: если новый статус «На линии», поставим IN-метку (когда перезаписываем строку)

        try{

          if(stVal==="На линии"){

            var dnow1 = new Date(); var dIso1 = iso(dnow1); var hh1 = hhmm(dnow1);

            sh_(S.PLAN).appendRow([dIso1, canon, hh1, hh1, "IN", "auto"]);

          }

        }catch(_e){}



        // короткий эхо-пинг (3 попытки с быстрыми задержками)

        var okEcho=false, echoObj=null;

        for(var a=0;a<3;a++){

          Utilities.sleep(150);

          echoObj = getStatusByNameFresh_(canon) || {};

          if(echoObj && echoObj.status===stVal){ okEcho=true; break; }

        }

        if(!okEcho){

          logWarn_("status","mismatch_after_set",{name:canon, want:stVal, echoStatus:echoObj && echoObj.status, row:j+2, lastRow:sh.getLastRow()}, "echo mismatch");

        } else {

          log_("INFO","status","echo",{name:canon, echoStatus:echoObj.status, echoUpdated:echoObj.updated},"ok");

        }

        return {ok:true, name:canon, status:stVal, rowWritten:j+2, sheet:S.STAT, lastRow:sh.getLastRow()};

      }

    }



    sh.appendRow([canon,stVal,now]);

    SpreadsheetApp.flush();



    // PATCH: если это первая запись и статус «На линии», тоже ставим IN-метку

    try{

      if(stVal==="На линии"){

        var dnow2 = new Date(); var dIso2 = iso(dnow2); var hh2 = hhmm(dnow2);

        sh_(S.PLAN).appendRow([dIso2, canon, hh2, hh2, "IN", "auto"]);

      }

    }catch(_e){}



    var echo=null, okEcho2=false;

    for(var b=0;b<3;b++){

      Utilities.sleep(150);

      echo = getStatusByNameFresh_(canon) || {};

      if(echo && echo.status===stVal){ okEcho2=true; break; }

    }

    if(!okEcho2){

      logWarn_("status","mismatch_after_append",{name:canon, want:stVal, echoStatus:echo && echo.status, row:sh.getLastRow()}, "echo mismatch");

    } else {

      log_("INFO","status","echo",{name:canon, echoStatus:echo.status, echoUpdated:echo.updated},"ok");

    }

    return {ok:true, name:canon, status:stVal, rowWritten:sh.getLastRow(), sheet:S.STAT, lastRow:sh.getLastRow()};

  });

}



/** интервал «сейчас»: now..now+1m */

function whoOnShiftNow_(dateIso){

  var d=iso(dateIso||new Date());

  var now=new Date(); var from=hhmm(now); var to=hhmm(new Date(now.getTime()+60000));

  var act=activeOnSlot_(from,to,d);



  // PATCH: исключаем старшего из распределения ролей «сейчас»

  var seniors = staff_().filter(function(s){ return s.isSenior; }).map(function(s){ return s.name; });

  act = act.filter(function(n){ return seniors.indexOf(n)===-1; });



  var roleByName={};

  if(act.length===1){ roleByName[act[0]]="CALLS+MEET"; }

  else if(act.length===2){ roleByName[act[0]]="CALLS"; roleByName[act[1]]="MEET"; }

  else if(act.length>=3){

    var h=new Date().getHours();

    roleByName[act[(h)%act.length]]="CALLS";

    roleByName[act[(h+1)%act.length]]="CALLS";

    roleByName[act[(h+2)%act.length]]="MEET";

  }

  return {roleByName:roleByName};

}



function statuses_(){

  // читаем только «свежо» и берём самую свежую запись по каждому имени

  var st=new Map();

  var rows = readStatusesFresh_();

  rows.forEach(function(r){

    var nm = String(r[0]||'').trim();

    if(!nm) return;

    var status = String(r[1]||'');

    if(!VALID_STATUSES.has(status)){ status = 'OFF'; }

    var updated = String(r[2]||'');

    var key = normName_(nm);

    var ex = st.get(key);

    if(!ex || (updated && String(updated) > String(ex.updated||""))){

      st.set(key, {name:nm, status:status, updated:updated});

    }

  });



  var today=iso(new Date());

  var nowRoles = whoOnShiftNow_(today).roleByName;



  var allStaff = staff_();

  return allStaff.map(function(s){

    var rec = st.get(s.norm);



    // агрегаты на сегодня

    var arrive = firstInTimeFor_(today, s.name);

    var lunchMin = sumIntervalMinFor_(today, s.name, "LUNCH");

    var breakMin = sumIntervalMinFor_(today, s.name, "BREAK");



    // маркер опоздания относительно Roster

    var r = rosterShiftsFor_(today).find(function(x){ return normName_(x.name)===s.norm; });

    var late = (r && r.start && arrive) ? latenessMark_(arrive, r.start) : {mark:'', hhmm:arrive||''};



    return {

      name: s.name,

      status: rec ? rec.status : 'OFF',

      updated: rec ? rec.updated : '',

      roleNow: s.isSenior ? '' : (nowRoles[s.name]||''), // старшему роль не ставим

      arriveHHmm: late.hhmm || '',

      arriveMark: late.mark || '',

      lunchMin: lunchMin || 0,

      breakMin: breakMin || 0

    };

  });

}



// ---------- absences & holidays ----------

function absencesMapFor_(dateIso){

  var d=iso(dateIso); var m={};

  data_(S.ABS,3).forEach(function(r){ if(iso(r[0])===d){ m[normName_(r[1])]=String(r[2]); }});

  return m;

}

function isHoliday_(dateIso){

  var d=iso(dateIso); var dt=new Date(d);

  var wknd=[0,6].includes(dt.getDay());

  var hol = data_(S.HOL,2).some(function(r){ return iso(r[0])===d; });

  return wknd || hol;

}

function syncRuHolidays_(){

  return withLock_(15000, function(){

    var now=new Date(); var to=new Date(now.getTime()+365*864e5);

    var calId = 'ru.russian#holiday@group.v.calendar.google.com';

    var ev=CalendarApp.getCalendarById(calId).getEvents(now,to);

    var rows=ev.map(function(e){ return [iso(e.getAllDayStartDate()), e.getTitle()]; });

    putRows_(S.HOL, rows, ["Date","Title"]);

    return {ok:true,count:rows.length};

  });

}

function api_holidaysSync(){ return safe_("abs","holidaysSync",{},function(){ return syncRuHolidays_(); }); }

function api_absList(fromIso,toIso){

  return safe_("abs","list",{fromIso:fromIso,toIso:toIso},function(){

    var from = fromIso?new Date(fromIso):new Date("2000-01-01");

    var to   = toIso?new Date(toIso):new Date("2100-01-01");

    return data_(S.ABS,3).map(function(r){ return ({date:iso(r[0]),name:String(r[1]),type:String(r[2])}); })

      .filter(function(x){ return new Date(x.date)>=from && new Date(x.date)<=to; });

  });

}

function api_absAddRange(fromIso,toIso,name,type){

  return safe_("abs","addRange",{fromIso:fromIso,toIso:toIso,name:name,type:type},function(){

    return withUserLock_(normName_(name), 10000, function(){

      var sh=sh_(S.ABS);

      var from=new Date(fromIso), to=new Date(toIso);

      var canon=canonicalName_(name);

      for(var d=new Date(from); d<=to; d.setDate(d.getDate()+1)){

        sh.appendRow([iso(d), canon, type||'VACATION']);

      }

      return {ok:true};

    });

  });

}



// ---------- week planner ----------

var SHIFTS = ["OFF","9-18","10-19","12-21", "10-21"]; // очищено: только валидные смены

function api_shifts(){ return safe_("week","shifts",{},function(){ return SHIFTS.slice(); }); }

function shiftTimes_(code){

  if(code==="9-18") return ["09:00","18:00"];

  if(code==="10-19")return ["10:00","19:00"];

  if(code==="12-21")return ["12:00","21:00"];

  if(code==="10-21")return ["10:00","21:00"];

  return [null,null]; // OFF/прочее

}

function weekDays_(anchor){

  var a=new Date(anchor||new Date()); var wd=(a.getDay()+6)%7; a.setDate(a.getDate()-wd);

  return Array.from({length:7}).map(function(_v,i){ return iso(new Date(a.getFullYear(),a.getMonth(),a.getDate()+i)); });

}

function draftMap_(days){

  var m={}; days.forEach(function(d){ m[d]={}; });

  data_(S.PLANW,3).forEach(function(r){ var d=iso(r[0]); if(!m[d]) return; m[d][normName_(r[1])]=String(r[2]); });

  return m;

}

function rosterMap_(days){

  var m={}; days.forEach(function(d){ m[d]={}; });

  data_(S.ROSTER,5).forEach(function(r){

    var d=iso(r[0]); if(!m[d]) return;

    if(String(r[4])!=="SHIFT")return;

    m[d][normName_(r[1])]={start:normHHmm_(r[2]),end:normHHmm_(r[3])};

  });

  return m;

}

function weekGetDraft_(anchor){

  var days=weekDays_(anchor);

  var all=staff_();

  var dm=draftMap_(days);

  var rm=rosterMap_(days);

  var absByDay={}; days.forEach(function(d){ absByDay[d]=absencesMapFor_(d); });



  function toCode(st,en){

    var s=normHHmm_(st), e=normHHmm_(en);

    if(s==="09:00"&&e==="18:00") return "9-18";

    if(s==="10:00"&&e==="19:00") return "10-19";

    if(s==="12:00"&&e==="21:00") return "12-21";

    if(s==="10:00"&&e==="21:00") return "10-21";

    return "OFF";

  }



  // PATCH: заголовочные метки праздников/выходных

  function holidayTitle_(d){

    var dt = new Date(d);

    var isWknd = [0,6].includes(dt.getDay());

    var title = '';

    data_(S.HOL,2).forEach(function(r){ if(iso(r[0])===d) title = String(r[1]||''); });

    if(title) return title;

    return isWknd ? 'Выходной' : '';

  }

  var holTitles = days.map(function(d){ return holidayTitle_(d); });



  var rows = all.map(function(s){ return ({

    name:s.name,

    days: days.map(function(d){

      var draft = (dm[d]||{})[s.norm];

      if(draft && draft!=="OFF") return draft;

      var r = (rm[d]||{})[s.norm];

      if(r && r.start && r.end){ return toCode(r.start,r.end); }

      return "OFF";

    }),

    abs:  days.map(function(d){ return (absByDay[d]||{})[s.norm] || ""; })

  }); });



  return {days:days, rows:rows, holidays: holTitles};

}

function weekSetDraft_(dateIso, name, code){

  return withUserLock_(normName_(name), 10000, function(){

    var all = data_(S.PLANW,3);

    var key = normName_(name);

    var keep = all.filter(function(r){ return !(iso(r[0])===iso(dateIso) && normName_(r[1])===key); });

    putRows_(S.PLANW, keep, ["Date","Name","Shift"]);

    sh_(S.PLANW).appendRow([iso(dateIso), canonicalName_(name), String(code)]);

    return {ok:true};

  });

}

function weekClearDraft_(anchor){

  return withLock_(10000, function(){

    var days=weekDays_(anchor); var set=new Set(days);

    var rows=data_(S.PLANW,3).filter(function(r){ return !set.has(iso(r[0])); });

    putRows_(S.PLANW, rows, ["Date","Name","Shift"]);

    return {ok:true};

  });

}

function weekPublishDraft_(anchor){

  return withLock_(15000, function(){

    var days=weekDays_(anchor);

    var set=new Set(days);

    var roster=data_(S.ROSTER,5).filter(function(r){ return !(set.has(iso(r[0])) && String(r[4])==="SHIFT"); });

    putRows_(S.ROSTER, roster, ["Date","Name","Start","End","Role"]);

    var dm=draftMap_(days); var toAdd=[];

    days.forEach(function(d){

      var rec=dm[d]||{};

      Object.keys(rec).forEach(function(nkey){

        var code = rec[nkey];

        if(!code || code==="OFF") return;

        var se=shiftTimes_(code); var st=se[0], en=se[1];

        if(st && en) toAdd.push([d, canonicalName_(nkey), st,en,"SHIFT"]);

      });

    });

    if(toAdd.length){

      var sh=sh_(S.ROSTER);

      sh.getRange(sh.getLastRow()+1,1,toAdd.length,5).setValues(toAdd);

    }

    return {ok:true, rows:toAdd.length};

  });

}



// авто-черновик недели (учитывает отсутствия; + выходной секретарь 10-21)

function weekAutoDraft_(anchor){

  return withLock_(15000, function(){

    var days=weekDays_(anchor);

    var all=staff_();

    var sec=all.filter(function(s){ return !s.isSenior && !s.isWeekend; });

    var senior=all.find(function(s){ return s.isSenior; });

    var wknd=weekend_();

    var out=[];

    var rot3=0; var rotSenior=0;



    function pushUnique(d,name,code){

      var key=normName_(name);

      for(var i=out.length-1;i>=0;i--){ if(out[i][0]===d && normName_(out[i][1])===key) out.splice(i,1); }

      out.push([d,canonicalName_(name),code]);

    }



    days.forEach(function(d){

      if(isHoliday_(d)){

        if(wknd) pushUnique(d, wknd.name, "10-21"); // PATCH: выходной/праздник = 10-21

        return;

      }

      var abs = absencesMapFor_(d);

      var present = sec.filter(function(s){ return !abs[s.norm]; });



      if(senior) pushUnique(d, senior.name, "10-19");



      if(present.length===3){

        var order=["9-18","10-19","12-21","10-21"];

        present.forEach(function(p,i){ pushUnique(d,p.name,order[(i+rot3)%3]); });

        rot3=(rot3+1)%3;

      }else if(present.length===2){

        pushUnique(d, present[0].name, "9-18");

        pushUnique(d, present[1].name, "12-21");

      }else if(present.length===1){

        var p=present[0];

        var seniorShift = (rotSenior%2===0) ? "9-18" : "12-21";

        var secShift     = (seniorShift==="9-18") ? "12-21" : "9-18";

        pushUnique(d,p.name,secShift);

        if(senior){ pushUnique(d, senior.name, seniorShift); }

        rotSenior++;

      }

    });



    var keep = data_(S.PLANW,3).filter(function(r){ return !days.includes(iso(r[0])); });

    putRows_(S.PLANW, keep, ["Date","Name","Shift"]);

    if(out.length){

      var sh=sh_(S.PLANW);

      sh.getRange(sh.getLastRow()+1,1,out.length,3).setValues(out);

    }

    return {ok:true,rows:out.length};

  });

}



// ---------- day roles ----------

function rosterShiftsFor_(dateIso){

  var d=iso(dateIso);

  return data_(S.ROSTER,5)

    .filter(function(r){ return iso(r[0])===d && String(r[4])==="SHIFT"; })

    .map(function(r){ return ({name:canonicalName_(r[1]),start:normHHmm_(r[2]),end:normHHmm_(r[3])}); });

}

function blocksForDate_(dateIso){

  var d=iso(dateIso);

  return data_(S.PLAN,6).filter(function(r){ return iso(r[0])===d; })

    .map(function(r){ return ({name:canonicalName_(r[1]),start:normHHmm_(r[2]),end:normHHmm_(r[3]),role:String(r[4])}); });

}

function intersects_(A1,A2,B1,B2){

  function p(x){ var m=/^(\d{2}):(\d{2})$/.exec(String(x||"")); return m? (+m[1]*60+ +m[2]) : null; }

  var a1=p(A1), a2=p(A2), b1=p(B1), b2=p(B2);

  if([a1,a2,b1,b2].some(function(v){return v==null;})) return false;

  return !(a2<=b1 || b2<=a1);

}

function inShift_(tFrom,tTo,shift){ return intersects_(tFrom,tTo,shift.start,shift.end); }

function blocked_(tFrom,tTo,blocks,name){

  var key=normName_(name);

  return blocks.some(function(b){ return normName_(b.name)===key && (b.role==="LUNCH"||b.role==="BREAK") && intersects_(tFrom,tTo,b.start,b.end); });

}

function activeOnSlot_(from,to,dateIso){

  var d=iso(dateIso);

  var blocks = blocksForDate_(d);

  var shifts = rosterShiftsFor_(d);

  return shifts

    .filter(function(s){ return inShift_(from,to,s); })

    .filter(function(s){ return !blocked_(from,to,blocks,s.name); })

    .map(function(s){ return s.name; })

    .sort();

}

function dayMenu_(dateIso){

  var d=iso(dateIso||new Date());

  var hours=[];

  var t=new Date(d+"T09:00:00");

  var end=new Date(d+"T21:00:00");



  var shifts = rosterShiftsFor_(d);

  var blocks = blocksForDate_(d);

  var supObj = senior_();

  var supName = supObj ? String(supObj.name||'').trim() : '';



  function inShiftLoc(tFrom,tTo,shift){ return intersects_(tFrom,tTo,shift.start,shift.end); }

  function blockedLoc(tFrom,tTo,name){

    var key=normName_(name);

    return blocks.some(function(b){ return normName_(b.name)===key && (b.role==="LUNCH"||b.role==="BREAK") && intersects_(tFrom,tTo,b.start,b.end); });

  }



  var rr=0;

  while(t<end){

    var from=hhmm(t), to=hhmm(new Date(t.getTime()+3600000));

    var availAll = shifts

      .filter(function(s){ return inShiftLoc(from,to,s); })

      .filter(function(s){ return !blockedLoc(from,to,String(s.name||'').trim()); })

      .map(function(s){ return String(s.name||'').trim(); })

      .sort();



    // исключаем старшего из пула ролей

    var seniors = staff_().filter(function(s){ return s.isSenior; }).map(function(s){ return s.name; });

    var basePool = availAll.filter(function(n){ return n && seniors.indexOf(n)===-1; });



    var supShift = supName ? shifts.find(function(s){ return normName_(s.name)===normName_(supName); }) : null;

    var supFree  = !!(supShift && inShiftLoc(from,to,supShift) && !blockedLoc(from,to,supName));



    var calls=[], meet="";



    if(basePool.length===0){

      if(supFree){ calls=[supName]; meet=supName; }

      else { calls=[]; meet=""; }

    } else if(basePool.length===1){

      calls=[basePool[0]];

      meet = supFree ? supName : basePool[0];

    } else if(basePool.length===2){

      calls=[basePool[0]]; meet=basePool[1];

    } else {

      calls=[basePool[(rr)%basePool.length], basePool[(rr+1)%basePool.length]];

      meet=basePool[(rr+2)%basePool.length];

      rr=(rr+1)%basePool.length;

    }

    if(!meet && calls.length>0 && supFree){ meet=supName; }



    hours.push({from:from,to:to,CALLS:calls,MEET:meet||""});

    t=new Date(t.getTime()+3600000);

  }

  return {date:d, hours:hours};

}



// ---------- intervals (lunch/break) ----------

function markInterval_(name, role, minutes){

  return withUserLock_(normName_(name), 8000, function(){

    var canon=canonicalName_(name);

    var d=new Date(); var date=iso(d); var st=hhmm(d); var en=hhmm(new Date(d.getTime()+ (minutes||60)*60000));

    sh_(S.PLAN).appendRow([date,canon,st,en,role,"front"]);

    return {ok:true, wrote:String(role)+" "+String(minutes||0)+"m", name:canon};

  });

}

function endIntervalIfAny_(name, role){

  return withUserLock_(normName_(name), 8000, function(){

    var canon=canonicalName_(name);

    var d=new Date(); var date=iso(d); var now=hhmm(d);

    var sh=sh_(S.PLAN); var rows=data_(S.PLAN,6);

    var nkey=normName_(canon);

    for(var i=rows.length-1;i>=0;i--){

      var r=rows[i]; if(iso(r[0])!==date || normName_(r[1])!==nkey || r[4]!==role) continue;

      if(String(r[2])<now && now<String(r[3])){

        sh.getRange(i+2,4).setValue(now);

        SpreadsheetApp.flush();

        Utilities.sleep(120);

        try{

          var dur = minDiffHHmm_(now, String(r[2]));

          var limit = role==="LUNCH"? 60 : 15;

          if(dur!=null && dur>limit){

            logWarn_("discipline", role.toLowerCase()+"_long", {name:canon, date:date, start:String(r[2]), end:now, minutes:dur, limit:limit}, "Длительный "+role+": "+dur+" мин");

          }

        }catch(_e){}

        break;

      }

    }

    return {ok:true};

  });

}

/** Единая точка: интервал + статус */

function setStatusAndInterval_(name, kind, minutes, isStart){

  var canon=canonicalName_(name);

  var human = (kind==='LUNCH'?'Обед':'Перерыв');

  if(isStart){

    markInterval_(canon, kind, minutes);

    return setStatus_(canon, human);

  }else{

    endIntervalIfAny_(canon, kind);

    return setStatus_(canon, "На линии");

  }

}



// ---------- PATCH helpers: дневные агрегаты ----------

function firstInTimeFor_(dateIso, name){

  var d = iso(dateIso); var key = normName_(name); var first = null;

  data_(S.PLAN,6).forEach(function(r){

    if(iso(r[0])!==d) return;

    if(normName_(r[1])!==key) return;

    if(String(r[4])!=="IN") return;

    var t = normHHmm_(r[2]);

    if(!first || (t && t < first)) first = t;

  });

  return first; // 'HH:MM' | null

}

function sumIntervalMinFor_(dateIso, name, role){

  var d = iso(dateIso); var key = normName_(name);

  var total = 0;

  data_(S.PLAN,6).forEach(function(r){

    if(iso(r[0])!==d) return;

    if(normName_(r[1])!==key) return;

    if(String(r[4])!==role) return;

    var st = normHHmm_(r[2]); var en = normHHmm_(r[3]);

    if(!st || !en) return;

    var m = minDiffHHmm_(en, st);

    if(m!=null && m>0) total += m;

  });

  return total || 0;

}

function latenessMark_(arriveHHmm, rosterStartHHmm){

  var df = minDiffHHmm_(arriveHHmm, rosterStartHHmm);

  if(df==null) return {mark:'', hhmm:''};

  if(df<=0) return {mark:'ok', hhmm:arriveHHmm};          // вовремя/раньше

  if(df<=10) return {mark:'warn', hhmm:arriveHHmm};       // до 10 минут

  return {mark:'bad', hhmm:arriveHHmm};                   // >10 минут

}



// ---------- QC ----------

function qcSubmit_(f){

  return withUserLock_(normName_(f.employee||''), 5000, function(){

    var canonEmp=canonicalName_(f.employee);

    var canonRev=canonicalName_(f.reviewer);

    sh_(S.QC).appendRow([canonEmp,canonRev,iso(f.date),f.callId,f.tone,f.script,f.parasites,f.comment,nowIso()]);

    return {ok:true};

  });

}

function qcList_(fromIso,toIso){

  var from=fromIso?new Date(fromIso):new Date("2000-01-01");

  var to=toIso?new Date(toIso):new Date("2100-01-01");

  return data_(S.QC,9).map(function(r){ return ({

    employee:String(r[0]),reviewer:String(r[1]),date:iso(r[2]),

    callId:String(r[3]||""),tone:+r[4]||0,script:+r[5]||0,parasites:+r[6]||0,comment:String(r[7]||"")

  }); }).filter(function(x){ return new Date(x.date)>=from && new Date(x.date)<=to; });

}

function qcAvg_(days){

  var since=new Date(); since.setDate(since.getDate()-days);

  var m={};

  data_(S.QC,9).forEach(function(r){

    var dt=new Date(r[8]||new Date()); if(dt<since) return;

    var e=String(r[0]); var v=((+r[4]||0)+(+r[5]||0)+(+r[6]||0))/3;

    (m[e]||(m[e]={sum:0,cnt:0})).sum+=v; m[e].cnt++;

  });

  var out={}; Object.keys(m).forEach(function(k){ out[k]= +(m[k].sum/m[k].cnt).toFixed(2); });

  return out;

}



// ---------- my week ----------

function myWeek_(name, anchor){

  var canon=canonicalName_(name);

  var days=weekDays_(anchor); var rm=rosterMap_(days);

  function toCode(st,en){

    var s=normHHmm_(st), e=normHHmm_(en);

    if(s==="09:00"&&e==="18:00") return "9-18";

    if(s==="10:00"&&e==="19:00") return "10-19";

    if(s==="12:00"&&e==="21:00") return "12-21";

    if(s==="10:00"&&e==="21:00") return "10-21"

    return "OFF";

  }

  return {days:days, shifts: days.map(function(d){

    var r=(rm[d]||{})[normName_(canon)]; if(!r) return "OFF"; return toCode(r.start,r.end);

  })};

}



// ---------- HTTP ----------

function doGet(e){

  ensure_();

  var token = e && e.parameter && e.parameter.token;

  var ctx = validateToken_(token);

  var tplName = ctx.ok && ctx.role==='secretary' ? 'Secretary' : 'Supervisor';

  var t;

  try{ t=HtmlService.createTemplateFromFile(tplName); }

  catch(err){

    return HtmlService.createHtmlOutput("<!doctype html><meta charset='utf-8'><h3>Создайте файлы <b>Supervisor</b> и <b>Secretary</b></h3>");

  }

  t.ctx=ctx;

  return t.evaluate().setTitle("Capital Group · Панель").setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

}



// ---------- APIs ----------

function api_bootstrap(){ return safe_("core","bootstrap",{},function(){ return ({ok:true}); }); }

function api_staff(){ return safe_("data","staff",{},function(){ return staff_().map(function(s){ return ({name:s.name,isSenior:s.isSenior,isWeekend:s.isWeekend}); }); }); }

function api_statuses(){ return safe_("data","statuses",{},function(){ return statuses_(); }); }

function api_myStatus(token){

  return safe_("status","my",{hasToken: !!token}, function(){

    var c=validateToken_(token);

    if(!c.ok) return {ok:false,error:'auth'};

    var echo = getStatusByNameFresh_(c.name);

    return {ok:true, name:c.name, status:(echo && echo.status) || 'OFF', updated:(echo && echo.updated)||''};

  });

}

function api_setStatusSecure(name,status,token){

  return safe_("status","set",{name:name,status:status},function(){

    var c=validateToken_(token);

    if(!c.ok) return {ok:false,error:'auth'};

    var targetRaw = (name && String(name).trim()) ? String(name).trim() : String(c.name||'').trim();

    var target = canonicalName_(targetRaw);

    var names = staff_().map(function(s){ return s.norm; });

    if(names.indexOf(normName_(target))===-1) return {ok:false,error:'invalid_name'};



    if(status==="Обед"){ return setStatusAndInterval_(target,"LUNCH",60,true); }

    if(status==="Перерыв"){ return setStatusAndInterval_(target,"BREAK",15,true); }

    if(status==="На линии"){ endIntervalIfAny_(target,"LUNCH"); endIntervalIfAny_(target,"BREAK"); return setStatus_(target,"На линии"); }

    if(status==="OFF"){ return setStatus_(target,"OFF"); }



    return setStatus_(target,status);

  });

}

function api_quickEventSecure(event, token){

  return safe_("status","quick",{event:event},function(){

    var c=validateToken_(token); if(!c.ok) return {ok:false,error:'auth'};

    var map={IN:"На линии", OUT:"OFF", LUNCH_START:"Обед", LUNCH_END:"На линии", BREAK_START:"Перерыв", BREAK_END:"На линии"};

    return api_setStatusSecure('', map[event]||"На линии", token);

  });

}

/** интервалы */

function api_lunchStart(mins,token){

  return safe_("plan","lunchStart",{mins:mins},function(){

    var c=validateToken_(token); if(!c.ok) return {ok:false,error:'auth'};

    return setStatusAndInterval_(c.name,"LUNCH",mins||60,true);

  });

}

function api_lunchEnd(token){

  return safe_("plan","lunchEnd",{},function(){

    var c=validateToken_(token); if(!c.ok) return {ok:false,error:'auth'};

    return setStatusAndInterval_(c.name,"LUNCH",0,false);

  });

}

function api_breakStart(mins,token){

  return safe_("plan","breakStart",{mins:mins},function(){

    var c=validateToken_(token); if(!c.ok) return {ok:false,error:'auth'};

    return setStatusAndInterval_(c.name,"BREAK",mins||15,true);

  });

}

function api_breakEnd(token){

  return safe_("plan","breakEnd",{},function(){

    var c=validateToken_(token); if(!c.ok) return {ok:false,error:'auth'};

    return setStatusAndInterval_(c.name,"BREAK",0,false);

  });

}



function api_dayMenu(dateIso){ return safe_("roles","dayMenu",{dateIso:dateIso},function(){ return dayMenu_(dateIso); }); }

function api_weekGet(anchor){ return safe_("week","getDraft",{anchor:anchor},function(){ return weekGetDraft_(anchor); }); }

function api_weekSetDraft(dateIso,name,code){

  return safe_("week","setDraft",{dateIso:dateIso,name:name,code:code}, function(){

    return weekSetDraft_(dateIso,name,code);

  });

}

function api_weekClearDraft(anchor){ return safe_("week","clearDraft",{anchor:anchor},function(){ return weekClearDraft_(anchor); }); }

function api_weekAuto(anchor){ return safe_("week","autoDraft",{anchor:anchor}, function(){ return weekAutoDraft_(anchor); }); }

function api_weekPublish(anchor){ return safe_("week","publish",{anchor:anchor},function(){ return weekPublishDraft_(anchor); }); }



function weekReplaceRosterCell_(dateIso,name,code){

  return withUserLock_(normName_(name), 8000, function(){

    var d=iso(dateIso); var se=shiftTimes_(code); var st=se[0], en=se[1];

    if(!st||!en){ return {ok:false,error:'bad-shift'}; }

    var key=normName_(name);

    var rows=data_(S.ROSTER,5).filter(function(r){ return !(iso(r[0])===d && normName_(r[1])===key && String(r[4])==="SHIFT"); });

    putRows_(S.ROSTER, rows, ["Date","Name","Start","End","Role"]);

    sh_(S.ROSTER).appendRow([d,canonicalName_(name),st,en,"SHIFT"]);

    return {ok:true};

  });

}

function api_weekReplaceRoster(dateIso,name,code){

  return safe_("week","replaceRoster",{dateIso:dateIso,name:name,code:code},function(){ return weekReplaceRosterCell_(dateIso,name,code); });

}



// -------- QC APIs --------

function api_qcSubmit(f){ return safe_("qc","submit",f,function(){ return qcSubmit_(f); }); }

function api_qcList(fromIso,toIso){ return safe_("qc","list",{fromIso:fromIso,toIso:toIso},function(){ return qcList_(fromIso,toIso); }); }

function api_qcAvg(days){ return safe_("qc","avg",{days:days},function(){ return qcAvg_(+days||14); }); }



// -------- My week API --------

function api_myWeek(name, anchor){ return safe_("week","my",{name:name,anchor:anchor},function(){ return myWeek_(name, anchor); }); }



// -------- Errors API --------

function api_errorsList(fromIso,toIso,q,limit){

  return safe_("log","list",{fromIso:fromIso,toIso:toIso,q:q,limit:limit},function(){

    var from = fromIso ? new Date(fromIso) : new Date("2000-01-01");

    var to   = toIso   ? new Date(toIso)   : new Date("2100-01-01");

    var needle = (q||"").toLowerCase();

    var rows = data_(S.ERR,9).map(function(r){ return ({

      ts:String(r[0]||""), source:String(r[2]||""), action:String(r[3]||""), user:String(r[4]||""),

      payload:String(r[5]||""), message:String(r[6]||""), stack:String(r[7]||""), errId:String(r[8]||"")

    }); }).filter(function(e){

      var dt=new Date(e.ts||"2000-01-01");

      if(dt<from || dt>to) return false;

      if(!needle) return true;

      var hay=(e.source+" "+e.action+" "+e.message).toLowerCase();

      return hay.indexOf(needle)!==-1;

    });

    var lim = Math.max(1, +limit||100);

    return rows.slice(-lim);

  });

}

function api_errorsClear(){

  return safe_("log","clear",{},function(){

    return withLock_(8000, function(){

      putRows_(S.ERR, [], ["Ts","Level","Source","Action","User","Payload","Message","Stack","ErrId"]);

      return {ok:true};

    });

  });

}



// -------- Диагностика --------

function api_debug_snapshot(dateIso){

  return safe_("diag","snapshot",{dateIso:dateIso},function(){

    var d = iso(dateIso || new Date());

    var statuses = data_(S.STAT,3).map(function(r){ return ({name:String(r[0]),status:String(VALID_STATUSES.has(String(r[1]))?r[1]:'OFF'),updated:r[2]||''}); });

    var planToday = data_(S.PLAN,6).filter(function(r){ return iso(r[0])===d; })

      .map(function(r){ return ({name:String(r[1]),start:toHHmm_(r[2]),end:toHHmm_(r[3]),role:String(r[4])}); });

    var rosterToday = rosterShiftsFor_(d);

    var now=new Date(); var from=hhmm(now); var to=hhmm(new Date(now.getTime()+60000));

    var activeNow = activeOnSlot_(from, to, d);

    return {

      ok:true,

      spreadsheetId: SS.getId(),

      d:d,

      counts: { statuses: statuses.length, planToday: planToday.length, rosterToday: rosterToday.length },

      sample: { statuses: statuses.slice(0,10), planToday: planToday.slice(0,10), rosterToday: rosterToday.slice(0,10), activeNow: activeNow }

    };

  });

}



// ---------- DASHBOARD ----------

/**

 * Собирает свод за период по каждому сотруднику:

 * - DaysPresent: дни, где есть отметка IN в Planner

 * - LateCount: опоздания (first IN > Roster.start)

 * - LateFromLunch: дней с обедом > 60 мин

 * - LongBreaks: дней с перерывами > 15 мин суммарно

 * - VacationDays / SickDays: по Absences

 * - QC_Avg: среднее (Tone+Script+Parasites)/3 по датам в период (QC.Date)

 * - FirstSeen / LastSeen: первая/последняя отметка IN в периоде (по времени)

 */

function dashboardBuild_(fromIso, toIso){

  ensure_();

  var from = new Date(iso(fromIso||new Date())); var to = new Date(iso(toIso||new Date()));

  // расширим to до включительно

  to.setDate(to.getDate()+0);



  var staff = staff_().filter(function(s){ return !s.isWeekend; }); // исключим выходного секретаря из свода

  var names = staff.map(function(s){ return s.name; });



  // предзагрузка данных

  var planRows = data_(S.PLAN,6).filter(function(r){ var d=iso(r[0]); var dt=new Date(d); return dt>=from && dt<=to; });

  var rosterRows = data_(S.ROSTER,5).filter(function(r){ var d=iso(r[0]); var dt=new Date(d); return dt>=from && dt<=to && String(r[4])==="SHIFT"; });

  var absRows = data_(S.ABS,3).filter(function(r){ var d=iso(r[0]); var dt=new Date(d); return dt>=from && dt<=to; });

  var qcRows = data_(S.QC,9).filter(function(r){ var d=iso(r[2]); var dt=new Date(d); return dt>=from && dt<=to; });



  // индексы по дате и имени

  function key(d,n){ return iso(d)+'|'+normName_(n); }



  var firstIN = {}; var lastIN = {};

  var daysPresent = {}; var lateCount = {}; var lunchOver = {}; var breakOver = {};

  var vacDays = {}; var sickDays = {};

  var qcMap = {}; // name -> {sum,cnt}



  // roster start lookup

  var rosterStart = {}; // key(date|name) -> "HH:MM"

  rosterRows.forEach(function(r){

    rosterStart[key(r[0], r[1])] = normHHmm_(r[2]);

  });



  // presence + intervals + late

  var byDayByName = {}; // key -> {firstIn, lunchMin, breakMin}

  planRows.forEach(function(r){

    var d = iso(r[0]); var nm = canonicalName_(r[1]); var k = key(d,nm);

    var role = String(r[4]||"");

    var st = normHHmm_(r[2]); var en = normHHmm_(r[3]);

    var o = byDayByName[k] || (byDayByName[k]={firstIn:null, lunch:0, br:0});

    if(role==="IN"){

      if(st && (!o.firstIn || st<o.firstIn)) o.firstIn = st;

      // глобальные first/last по периоду

      var dts = new Date(d+"T"+(st||"00:00")+":00");

      if(!firstIN[nm] || dts < firstIN[nm]) firstIN[nm] = dts;

      if(!lastIN[nm]  || dts > lastIN[nm])  lastIN[nm]  = dts;

    }else if(role==="LUNCH" && st && en){

      var m = minDiffHHmm_(en,st); if(m!=null && m>0) o.lunch += m;

    }else if(role==="BREAK" && st && en){

      var m2 = minDiffHHmm_(en,st); if(m2!=null && m2>0) o.br += m2;

    }

    byDayByName[k]=o;

  });



  // агрегирование по дням

  names.forEach(function(nm){

    daysPresent[nm]=0; lateCount[nm]=0; lunchOver[nm]=0; breakOver[nm]=0;

  });



  Object.keys(byDayByName).forEach(function(k){

    var parts = k.split('|'); var d=parts[0]; var nn=parts[1];

    var canon=canonicalName_(nn);

    if(names.indexOf(canon)===-1) return;

    var o = byDayByName[k];

    if(o.firstIn){ daysPresent[canon] = (daysPresent[canon]||0)+1; }

    var rs = rosterStart[k];

    if(o.firstIn && rs){

      var df = minDiffHHmm_(o.firstIn, rs);

      if(df!=null && df>0){ lateCount[canon] = (lateCount[canon]||0)+1; }

    }

    if(o.lunch>60){ lunchOver[canon] = (lunchOver[canon]||0)+1; }

    if(o.br>15){ breakOver[canon] = (breakOver[canon]||0)+1; }

  });



  // abs

  absRows.forEach(function(r){

    var d=iso(r[0]); var nm=canonicalName_(r[1]); var type=String(r[2]||"");

    if(names.indexOf(nm)===-1) return;

    if(type==="VACATION"){ vacDays[nm]=(vacDays[nm]||0)+1; }

    if(type==="SICK"){ sickDays[nm]=(sickDays[nm]||0)+1; }

  });



  // qc

  qcRows.forEach(function(r){

    var nm=canonicalName_(r[0]);

    if(names.indexOf(nm)===-1) return;

    var v=((+r[4]||0)+(+r[5]||0)+(+r[6]||0))/3;

    (qcMap[nm]||(qcMap[nm]={sum:0,cnt:0})).sum+=v; qcMap[nm].cnt++;

  });



  // write

  var heads=["PeriodFrom","PeriodTo","Name","DaysPresent","LateCount","LateFromLunch","LongBreaks","VacationDays","SickDays","QC_Avg","FirstSeen","LastSeen"];

  var rows = names.map(function(nm){

    var avg = (qcMap[nm] && qcMap[nm].cnt>0) ? +(qcMap[nm].sum/qcMap[nm].cnt).toFixed(2) : "";

    var fs = firstIN[nm] ? Utilities.formatDate(firstIN[nm], TZ, "yyyy-MM-dd HH:mm") : "";

    var ls = lastIN[nm]  ? Utilities.formatDate(lastIN[nm], TZ, "yyyy-MM-dd HH:mm") : "";

    return [

      iso(from), iso(to),

      nm,

      daysPresent[nm]||0,

      lateCount[nm]||0,

      lunchOver[nm]||0,

      breakOver[nm]||0,

      vacDays[nm]||0,

      sickDays[nm]||0,

      avg,

      fs, ls

    ];

  });



  putRows_(S.DASH, rows, heads);

  return {ok:true, from:iso(from), to:iso(to), people:rows.length};

}



function api_dashboardRecalc(fromIso,toIso){

  return safe_("dash","recalc",{fromIso:fromIso,toIso:toIso}, function(){ return dashboardBuild_(fromIso,toIso); });

}



// ---------- menu ----------

function onOpen(){

  ensure_();

  SpreadsheetApp.getUi().createMenu('🏢 CG Menu')

    .addItem('Создать и заполнить листы', 'menu_bootstrapSheets')

    .addItem('Связать с веб-приложением', 'menu_linkManual')

    .addItem('Сгенерировать токены', 'menu_tokens')

    .addItem('Показать лог ошибок', 'menu_showErrors')

    .addSeparator()

    .addSubMenu(SpreadsheetApp.getUi().createMenu('📈 Дашборд')

      .addItem('Пересчитать (последние 30 дней)', 'menu_dashLast30')

      .addItem('Пересчитать (задать период)', 'menu_dashCustom'))

    .addToUi();

}

function bootstrapAll_(){

  ensure_();

  if (!cfgGet('WEBAPP_URL')) cfgSet('WEBAPP_URL', '');

  return { ok:true, sheets:[S.CFG,S.STAFF,S.TOK,S.STAT,S.ABS,S.HOL,S.PLAN,S.PLANW,S.ROSTER,S.QC,S.ERR,S.DASH] };

}

function menu_bootstrapSheets(){

  var r=bootstrapAll_();

  SpreadsheetApp.getUi().alert('Готово. Листы созданы/проверены: ' + (r.sheets||[]).join(', '));

}

function menu_linkManual(){

  var ui=SpreadsheetApp.getUi();

  var r=ui.prompt('Вставьте URL Web App (/exec)','',ui.ButtonSet.OK_CANCEL);

  if(r.getSelectedButton()!==ui.Button.OK) return;

  var url=(r.getResponseText()||'').trim();

  if(!/^https:\/\/script\.google\.com\/.*\/exec(\?.*)?$/i.test(url)){

    ui.alert('URL должен заканчиваться на /exec'); return;

  }

  cfgSet('WEBAPP_URL',url);

  ui.alert('Готово: URL сохранён (Config.WEBAPP_URL)');

}

function menu_tokens(){

  var base=cfgGet('WEBAPP_URL');

  if(!base){ SpreadsheetApp.getUi().alert('Сначала свяжите URL через меню.'); return; }

  return withLock_(8000, function(){

    var sh=sh_(S.TOK);

    putRows_(S.TOK, [], ["Token","Role","Name","Active","AccessURL"]);

    function add(role,name){

      var t=Utilities.getUuid().replace(/-/g,'');

      var url=base+(base.includes('?')?'&':'?')+'token='+t;

      sh.appendRow([t,role,canonicalName_(name),true,url]);

    }

    var st=staff_(); var sup=st.find(function(s){return s.isSenior;}); if(sup) add('supervisor',sup.name);

    st.filter(function(s){return !s.isSenior;}).forEach(function(s){ add('secretary',s.name); });

    SpreadsheetApp.getUi().alert('Токены сгенерированы.');

  });

}

function menu_showErrors(){

  var rows=data_(S.ERR,9).slice(-100);

  SpreadsheetApp.getUi().alert(rows.map(function(r){ return '['+r[0]+'] '+r[2]+': '+r[3]; }).join('\n')||'Лог пуст');

}



// ---- Dashboard menu handlers ----

function menu_dashLast30(){

  var to=new Date(); var from=new Date(); from.setDate(to.getDate()-29);

  var r = dashboardBuild_(iso(from), iso(to));

  SpreadsheetApp.getUi().alert('Дашборд посчитан за '+r.from+' — '+r.to+' ('+r.people+' сотрудников).');

}

function menu_dashCustom(){

  var ui=SpreadsheetApp.getUi();

  var r1=ui.prompt('Период дашборда','Дата С (YYYY-MM-DD)',ui.ButtonSet.OK_CANCEL);

  if(r1.getSelectedButton()!==ui.Button.OK) return;

  var r2=ui.prompt('Период дашборда','Дата По включительно (YYYY-MM-DD)',ui.ButtonSet.OK_CANCEL);

  if(r2.getSelectedButton()!==ui.Button.OK) return;

  var from=(r1.getResponseText()||'').trim(), to=(r2.getResponseText()||'').trim();

  if(!/^\d{4}-\d{2}-\d{2}$/.test(from) || !/^\d{4}-\d{2}-\d{2}$/.test(to)){ ui.alert('Неверный формат даты'); return; }

  var r = dashboardBuild_(from,to);

  ui.alert('Дашборд посчитан за '+r.from+' — '+r.to+' ('+r.people+' сотрудников).');

}
