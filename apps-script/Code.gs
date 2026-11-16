/**
 * Centrum Přítomnost — booking backend
 * Deploy as a Web App owned by centrumpritomnost@gmail.com.
 */
const CONFIG = {
  SPREADSHEET_ID: '1h1uEVdEO_8-4c6YwMMuHA3t9SvqY3Nqofc8nSPiPY8k',
  ADMIN_EMAIL: 'centrumpritomnost@gmail.com',
  CALENDAR_ID: 'centrumpritomnost@gmail.com',
  TIME_ZONE: 'Europe/Prague',
  REQUEST_SHEET: 'Requests',
  SETTINGS_SHEET: 'Settings',
  PROFESSIONALS_SHEET: 'Professionals',
  HOLD_PREFIX: '[PENDING] '
};
const REQUEST_HEADERS = [
  'Request ID','Created At','Status','Category','Service','Date','Start Time',
  'Duration (minutes)','Customer Name','Customer Email','Customer Phone',
  'Message','Calendar Event ID','End Date','End Time','Nights','Professional'
];

function setupCentrum() {
  const ss=SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let requests=ss.getSheetByName(CONFIG.REQUEST_SHEET);
  if(!requests)requests=ss.insertSheet(CONFIG.REQUEST_SHEET);
  ensureRequestHeaders_(requests);
  let settings=ss.getSheetByName(CONFIG.SETTINGS_SHEET);
  if(!settings)settings=ss.insertSheet(CONFIG.SETTINGS_SHEET);
  if(settings.getLastRow()===0){
    settings.getRange(1,1,1,2).setValues([['Setting','Value']]);
    settings.getRange(2,1,4,2).setValues([
      ['Admin email',CONFIG.ADMIN_EMAIL],['Calendar ID',CONFIG.CALENDAR_ID],
      ['Time zone',CONFIG.TIME_ZONE],
      ['Booking policy','Requests require manual approval. Pending calendar events temporarily block the requested time.']
    ]);
    settings.setFrozenRows(1);settings.getRange(1,1,1,2).setFontWeight('bold');
  }
  ensureProfessionalsSheet_(ss);
  const professionals=getProfessionals_(ss);
  professionals.forEach(p=>ensureProfessionalSheet_(ss,p));
  const calendar=CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  if(!calendar)throw new Error('Primary Google Calendar is not accessible.');
  return 'Setup completed. Requests, Zuzana sheet and primary calendar are accessible.';
}
function ensureRequestHeaders_(sheet){
  const current=sheet.getLastColumn()?sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0]:[];
  if(!current.length||!current[0]){
    sheet.getRange(1,1,1,REQUEST_HEADERS.length).setValues([REQUEST_HEADERS]);
    sheet.setFrozenRows(1);sheet.getRange(1,1,1,REQUEST_HEADERS.length).setFontWeight('bold');
  }else{
    REQUEST_HEADERS.forEach((header,i)=>{
      if(current[i]!==header)sheet.getRange(1,i+1).setValue(header);
    });
  }
}
function doGet(e){
  try{
    if(getParam_(e,'action')==='availability')return availabilityOutput_(e,getAvailability_(getParam_(e,'date'),getParam_(e,'professional')));
    if(getParam_(e,'action')==='professionals')return availabilityOutput_(e,{ok:true,professionals:getProfessionals_(SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)).map(p=>({name:p.name}))});
    return jsonOutput_({ok:true,service:'Centrum Přítomnost booking backend',status:'ready'});
  }catch(err){
    const value={ok:false,error:err.message||'Unable to process request.'};
    if(getParam_(e,'action')==='availability'&&getParam_(e,'callback'))return availabilityOutput_(e,value);
    return jsonOutput_(value);
  }
}
function getParam_(e,name){
  if(e&&e.parameter&&e.parameter[name]!=null)return String(e.parameter[name]).trim();
  if(e&&e.parameters&&e.parameters[name]&&e.parameters[name].length)return String(e.parameters[name][0]).trim();
  return '';
}
function availabilityOutput_(e,value){
  const callback=getParam_(e,'callback');
  if(callback){
    if(!/^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback))return jsonOutput_({ok:false,error:'Invalid callback.'});
    return ContentService.createTextOutput(callback+'('+JSON.stringify(value)+');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return jsonOutput_(value);
}
function doPost(e){
  let response;
  try{response=createRequest_((e&&e.parameter)||{});}
  catch(err){response={ok:false,error:err.message||'Unable to process request.'};}
  return postMessageOutput_(getParam_(e,'responseToken'),response);
}
function postMessageOutput_(token,result){
  const safe=JSON.stringify({type:'centrum-booking-result',token:token,ok:!!result.ok,error:result.error||'',requestId:result.requestId||''}).replace(/</g,String.fromCharCode(92)+'u003c');
  const page='<!doctype html><html><head><meta charset="utf-8"></head><body><script>window.parent.postMessage('+safe+', "*");</script><p>Request processed.</p></body></html>';
  return HtmlService.createHtmlOutput(page).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function ensureProfessionalsSheet_(ss){
  let sheet=ss.getSheetByName(CONFIG.PROFESSIONALS_SHEET);
  if(!sheet)sheet=ss.insertSheet(CONFIG.PROFESSIONALS_SHEET);
  const headers=['Professional Name','Email','Active','Sheet Name'];
  if(sheet.getLastRow()===0){
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
    sheet.getRange(2,1,1,4).setValues([['Zuzana Pithardová, DiS.','zuzanapithardova@gmail.com',true,'Zuzana']]);
    sheet.setFrozenRows(1);sheet.getRange(1,1,1,headers.length).setFontWeight('bold');
  }else{
    const current=sheet.getRange(1,1,1,Math.max(sheet.getLastColumn(),headers.length)).getValues()[0];
    headers.forEach((h,i)=>{if(current[i]!==h)sheet.getRange(1,i+1).setValue(h);});
  }
  if(sheet.getLastRow()<2){
    sheet.getRange(2,1,1,4).setValues([['Zuzana Pithardová, DiS.','zuzanapithardova@gmail.com',true,'Zuzana']]);
  }
  return sheet;
}
function getProfessionals_(ss){
  const sheet=ensureProfessionalsSheet_(ss);
  if(sheet.getLastRow()<2)return [];
  return sheet.getRange(2,1,sheet.getLastRow()-1,4).getValues()
    .filter(r=>r[0]&&r[1]&&r[2]!==false&&String(r[2]).toUpperCase()!=='FALSE')
    .map(r=>({name:String(r[0]).trim(),email:String(r[1]).trim(),sheetName:String(r[3]||safeSheetName_(r[0])).trim()}));
}
function safeSheetName_(name){
  const cleaned=String(name||'Professional').replace(/[\\/?*[\]:]/g,' ').trim().slice(0,90);
  return cleaned||'Professional';
}
function professionalSheetHeaders_(){
  return ['Request ID','Status','Date','Start Time','End Date','End Time','Duration (minutes)','Service','Customer Name','Customer Email','Customer Phone','Message'];
}
function ensureProfessionalSheet_(ss,professional){
  const name=professional.sheetName||safeSheetName_(professional.name);
  let sheet=ss.getSheetByName(name);
  if(!sheet)sheet=ss.insertSheet(name);
  const headers=professionalSheetHeaders_();
  if(sheet.getLastRow()===0){
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
    sheet.setFrozenRows(1);sheet.getRange(1,1,1,headers.length).setFontWeight('bold');
  }else{
    const current=sheet.getRange(1,1,1,Math.max(sheet.getLastColumn(),headers.length)).getValues()[0];
    headers.forEach((h,i)=>{if(current[i]!==h)sheet.getRange(1,i+1).setValue(h);});
  }
  return sheet;
}
function getProfessional_(ss,name){
  return getProfessionals_(ss).find(p=>p.name===String(name||'').trim())||null;
}
function dateCellString_(value){
  if(value instanceof Date)return Utilities.formatDate(value,CONFIG.TIME_ZONE,'yyyy-MM-dd');
  return String(value||'').slice(0,10);
}
function timeCellString_(value,fallback){
  if(value instanceof Date)return Utilities.formatDate(value,CONFIG.TIME_ZONE,'HH:mm');
  const text=String(value||'');
  return text?text.slice(0,5):fallback;
}
function syncProfessionalStatuses_(ss,professional){
  const requests=ss.getSheetByName(CONFIG.REQUEST_SHEET);
  const mirror=ensureProfessionalSheet_(ss,professional);
  if(!requests||requests.getLastRow()<2||mirror.getLastRow()<2)return;
  const requestRows=requests.getRange(2,1,requests.getLastRow()-1,REQUEST_HEADERS.length).getValues();
  const statusById={};
  requestRows.forEach(r=>{if(r[0])statusById[String(r[0])]=String(r[2]||'PENDING').toUpperCase();});
  const mirrorRows=mirror.getRange(2,1,mirror.getLastRow()-1,12).getValues();
  mirrorRows.forEach((r,i)=>{
    const status=statusById[String(r[0])];
    if(status&&String(r[1]).toUpperCase()!==status)mirror.getRange(i+2,2).setValue(status);
  });
}
function getProfessionalBusy_(ss,dateString,professional){
  const requests=ss.getSheetByName(CONFIG.REQUEST_SHEET);
  if(!requests||requests.getLastRow()<2)return [];
  syncProfessionalStatuses_(ss,professional);
  const rows=requests.getRange(2,1,requests.getLastRow()-1,REQUEST_HEADERS.length).getValues(),busy=[];
  rows.forEach(r=>{
    if(String(r[16]||'').trim()!==professional.name)return;
    const status=String(r[2]||'PENDING').toUpperCase();
    if(status==='REJECTED'||status==='CANCELLED')return;
    const startDate=dateCellString_(r[5]),endDate=dateCellString_(r[13]||r[5]);
    if(startDate!==dateString&&endDate!==dateString)return;
    const start=localDateTime_(startDate,timeCellString_(r[6],'00:00'));
    const end=localDateTime_(endDate,timeCellString_(r[14],'23:59'));
    busy.push({start:Utilities.formatDate(start,CONFIG.TIME_ZONE,"yyyy-MM-dd'T'HH:mm:ss"),end:Utilities.formatDate(end,CONFIG.TIME_ZONE,"yyyy-MM-dd'T'HH:mm:ss")});
  });
  return busy;
}
function getAvailability_(dateString,professionalName){
  if(!/^\d{4}-\d{2}-\d{2}$/.test(dateString||''))return {ok:false,error:'Invalid date. Use YYYY-MM-DD.'};
  const start=localDateTime_(dateString,'00:00'),end=localDateTime_(dateString,'23:59');
  const calendar=CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  if(!calendar)throw new Error('Calendar is not accessible.');
  let busy=calendar.getEvents(start,new Date(end.getTime()+60000)).map(ev=>({
    start:Utilities.formatDate(ev.getStartTime(),CONFIG.TIME_ZONE,"yyyy-MM-dd'T'HH:mm:ss"),
    end:Utilities.formatDate(ev.getEndTime(),CONFIG.TIME_ZONE,"yyyy-MM-dd'T'HH:mm:ss")
  }));
  if(professionalName){
    const ss=SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const professional=getProfessional_(ss,professionalName);
    if(!professional)throw new Error('Selected professional is not available.');
    busy=busy.concat(getProfessionalBusy_(ss,dateString,professional));
  }
  return {ok:true,date:dateString,busy:busy,timeZone:CONFIG.TIME_ZONE};
}
function localDateTime_(dateString,timeString){
  const p=dateString.split('-').map(Number),t=timeString.split(':').map(Number);
  const utc=Date.UTC(p[0],p[1]-1,p[2],t[0],t[1],0);
  const offsetText=Utilities.formatDate(new Date(utc),CONFIG.TIME_ZONE,'Z');
  const sign=offsetText.charAt(0)==='-'?-1:1;
  const offset=(Number(offsetText.slice(1,3))*60+Number(offsetText.slice(3,5)))*sign;
  return new Date(utc-offset*60000);
}
function createRequest_(data){
  const required=['category','service','date','startTime','duration','name','email','phone'];
  required.forEach(k=>{if(!String(data[k]||'').trim())throw new Error('Please complete all required fields.');});
  if(!/^\d{4}-\d{2}-\d{2}$/.test(data.date))throw new Error('Invalid date.');
  if(!/^([01]\d|2[0-3]):[0-5]\d$/.test(data.startTime))throw new Error('Invalid start time.');
  const duration=Number(data.duration);
  if(!Number.isInteger(duration)||duration<30||duration>43200)throw new Error('Invalid reservation duration.');
  const professional=String(data.professional||'').trim();
  if(data.category==='Masáž / individuální péče'&&!professional)throw new Error('Please select a professional for your treatment.');
  const email=String(data.email).trim();
  if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email))throw new Error('Invalid email address.');
  const start=localDateTime_(data.date,data.startTime);
  let end;
  if(data.endDate&&data.endTime){
    if(!/^\d{4}-\d{2}-\d{2}$/.test(data.endDate)||!/^([01]\d|2[0-3]):[0-5]\d$/.test(data.endTime))throw new Error('Invalid checkout date or time.');
    end=localDateTime_(data.endDate,data.endTime);
    if(Math.abs(Math.round((end-start)/60000)-duration)>1)throw new Error('Reservation duration does not match the selected end time.');
  }else end=new Date(start.getTime()+duration*60000);
  if(start<=new Date()||end<=start)throw new Error('Please choose a valid future time.');
  const nights=data.nights?Number(data.nights):0;
  if(data.nights&&(!Number.isInteger(nights)||nights<1||nights>30))throw new Error('Invalid number of nights.');
  const lock=LockService.getScriptLock();lock.waitLock(20000);
  try{
    const calendar=CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
    if(!calendar)throw new Error('Calendar is not accessible.');
    if(calendar.getEvents(start,end).length)throw new Error('This time is no longer available. Please choose another time.');
    const ss=SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const professionalRecord=professional?getProfessional_(ss,professional):null;
    if(professional&&!professionalRecord)throw new Error('Selected professional is not active or is not configured.');
    if(professionalRecord){
      const pSheet=ensureProfessionalSheet_(ss,professionalRecord);
      syncProfessionalStatuses_(ss,professionalRecord);
      const rows=pSheet.getLastRow()>1?pSheet.getRange(2,1,pSheet.getLastRow()-1,12).getValues():[];
      const conflict=rows.some(r=>{const st=String(r[1]||'PENDING').toUpperCase();if(st==='REJECTED'||st==='CANCELLED')return false;const sd=dateCellString_(r[2]),ed=dateCellString_(r[4]||r[2]);if(!sd||!ed)return false;const a=localDateTime_(sd,timeCellString_(r[3],'00:00')),b=localDateTime_(ed,timeCellString_(r[5],'23:59'));return start<b&&end>a;});
      if(conflict)throw new Error(professionalRecord.name+' is already booked at this time. Please choose another time.');
    }
    const sheet=ss.getSheetByName(CONFIG.REQUEST_SHEET);
    if(!sheet)throw new Error('Requests sheet is missing. Contact the administrator.');
    ensureRequestHeaders_(sheet);
    const id=Utilities.getUuid(),name=String(data.name).trim(),phone=String(data.phone).trim(),message=String(data.message||'').trim();
    const title=CONFIG.HOLD_PREFIX+data.category+' — '+data.service+(professional?' — '+professional:'')+' — '+name;
    const description=[
      'Status: PENDING — manual approval required','Request ID: '+id,
      'Category: '+data.category,'Service: '+data.service,'Professional: '+professional,'Customer: '+name,
      'Email: '+email,'Phone: '+phone,'Message: '+message,
      'End: '+Utilities.formatDate(end,CONFIG.TIME_ZONE,'dd.MM.yyyy HH:mm')
    ].join('\n');
    const event=calendar.createEvent(title,start,end,{description:description});
    let requestRow=0;
    try{
      sheet.appendRow([id,new Date(),'PENDING',data.category,data.service,data.date,data.startTime,duration,name,email,phone,message,event.getId(),data.endDate||Utilities.formatDate(end,CONFIG.TIME_ZONE,'yyyy-MM-dd'),data.endTime||Utilities.formatDate(end,CONFIG.TIME_ZONE,'HH:mm'),nights,professional]);
      requestRow=sheet.getLastRow();
      if(professionalRecord){
        const pSheet=ensureProfessionalSheet_(ss,professionalRecord);
        pSheet.appendRow([id,'PENDING',data.date,data.startTime,data.endDate||Utilities.formatDate(end,CONFIG.TIME_ZONE,'yyyy-MM-dd'),data.endTime||Utilities.formatDate(end,CONFIG.TIME_ZONE,'HH:mm'),duration,data.service,name,email,phone,message]);
      }
    }catch(sheetError){
      if(requestRow>0)sheet.deleteRow(requestRow);
      calendar.deleteEvent(event.getId());
      throw new Error('Could not save the request in the spreadsheet. '+sheetError.message);
    }
    const formatted=Utilities.formatDate(start,CONFIG.TIME_ZONE,'dd.MM.yyyy HH:mm')+' – '+Utilities.formatDate(end,CONFIG.TIME_ZONE,'dd.MM.yyyy HH:mm');
    try{
      MailApp.sendEmail({to:CONFIG.ADMIN_EMAIL,subject:'New booking request — Centrum Přítomnost',body:'A new booking request is waiting for manual approval.\n\nCategory: '+data.category+'\nService: '+data.service+'\nProfessional: '+professional+'\nDate and time: '+formatted+'\nDuration: '+duration+' minutes\nCustomer: '+name+'\nEmail: '+email+'\nPhone: '+phone+'\nMessage: '+message+'\n\nThe calendar slot is temporarily blocked as PENDING.'});
      if(professionalRecord)MailApp.sendEmail({to:professionalRecord.email,name:'Centrum Přítomnost',replyTo:CONFIG.ADMIN_EMAIL,subject:'New booking request — '+professionalRecord.name,body:'A new booking request for you is waiting for review.\\n\\nProfessional: '+professionalRecord.name+'\\nService: '+data.service+'\\nDate and time: '+formatted+'\\nDuration: '+duration+' minutes\\nCustomer: '+name+'\\nEmail: '+email+'\\nPhone: '+phone+'\\nMessage: '+message+'\\n\\nStatus: PENDING — not yet confirmed.'});
      MailApp.sendEmail({to:email,subject:'We received your request — Centrum Přítomnost',body:'Hello '+name+',\n\nWe received your request for '+data.service+(professional?' with '+professional:'')+' ('+formatted+'). This is not a confirmed reservation. We will contact you after reviewing it.\n\nCentrum Přítomnost'});
    }catch(mailError){console.error('Request saved, but notification email failed: '+mailError.message);}
    return {ok:true,requestId:id,status:'PENDING'};
  }finally{lock.releaseLock();}
}
function jsonOutput_(value){return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);}
