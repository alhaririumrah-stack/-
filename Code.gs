const DB = {
  Users:['userId','username','password','role','name','pilgrimCode','status','createdAt'],
  Trips:['tripId','tripName','tripType','startDate','endDate','capacity','baseCurrency','status','createdAt'],
  Packages:['packageId','tripId','packageName','priceUSD','includes','status','createdAt'],
  Pilgrims:['pilgrimCode','tripId','familyId','relation','fullName','phone','nationality','passportNo','packageName','roomType','driveLink','passportStatus','visaStatus','priceUSD','paidUSD','balanceUSD','walletBalanceUSD','roomId','busId','guideId','createdAt'],
  Families:['familyCode','tripId','headPilgrimCode','familyName','membersCount','note','createdAt'],
  Payments:['paymentId','tripId','pilgrimCode','date','currency','amount','rateToUSD','amountUSD','method','note','createdAt'],
  Expenses:['expenseId','tripId','date','category','currency','amount','rateToUSD','amountUSD','paidBy','note','createdAt'],
  ExchangeRates:['rateId','date','currency','rateToUSD','source','note','createdAt'],
  Accounts:['accountId','accountName','accountType','currency','parentAccount','status','createdAt'],
  JournalEntries:['entryId','date','tripId','debitAccount','creditAccount','currency','amount','rateToUSD','amountUSD','description','createdAt'],
  Visas:['visaId','tripId','pilgrimCode','passportNo','status','requestDate','issueDate','note','createdAt'],
  Hotels:['hotelId','tripId','city','hotelName','checkIn','checkOut','nights','note','createdAt'],
  Rooms:['roomId','tripId','hotelId','roomNo','roomType','capacity','assignedCount','status','createdAt'],
  Buses:['busId','tripId','busNo','company','driverName','driverPhone','guideName','capacity','assignedCount','createdAt'],
  BusStops:['stopId','tripId','busId','stopName','time','status','outCount','backCount','note','createdAt'],
  Guides:['guideId','tripId','guideName','phone','language','maxPilgrims','advanceBalanceUSD','note','createdAt'],
  GuideAdvances:['advanceId','tripId','guideId','date','type','currency','amount','rateToUSD','amountUSD','category','note','createdAt'],
  Attendance:['attId','tripId','date','pilgrimCode','place','status','createdAt'],
  Luggage:['luggageId','tripId','pilgrimCode','tagNo','status','note','createdAt'],
  Notifications:['noticeId','tripId','title','target','body','createdAt'],
  Requirements:['reqId','tripId','itemName','qtyPerPilgrim','totalQty','status','createdAt'],
  Complaints:['complaintId','tripId','pilgrimCode','category','status','description','createdAt'],
  Logs:['logId','user','action','details','createdAt']
};
const PREFIX = {Trips:'T',Packages:'PKG',Pilgrims:'P',Families:'F',Payments:'PAY',Expenses:'EXP',ExchangeRates:'FX',Accounts:'ACC',JournalEntries:'JE',Visas:'V',Hotels:'H',Rooms:'R',Buses:'B',BusStops:'STOP',Guides:'G',GuideAdvances:'ADV',Attendance:'ATT',Luggage:'LUG',Notifications:'N',Requirements:'REQ',Complaints:'C',Users:'U',Logs:'LOG'};
function setup(){const ss=SpreadsheetApp.getActive(); Object.keys(DB).forEach(n=>{let sh=ss.getSheetByName(n)||ss.insertSheet(n); if(sh.getLastRow()===0) sh.appendRow(DB[n]);}); seed(); return out({ok:true,message:'تم إنشاء قاعدة البيانات بنجاح'});}
function seed(){const now=new Date(); const u=sheet('Users'); if(u.getLastRow()<2){u.appendRow(['U1','admin','123456','Admin','المدير العام','','Active',now]);u.appendRow(['U2','accountant','123456','Accountant','المحاسب','','Active',now]);u.appendRow(['U3','guide','123456','Guide','مرشد تجريبي','','Active',now]);}
 const fx=sheet('ExchangeRates'); if(fx.getLastRow()<2){[['USD',1],['EUR',1.08],['SAR',0.266],['TRY',0.031],['SYP',0.000077]].forEach((x,i)=>fx.appendRow(['FX'+(i+1),today(),x[0],x[1],'افتراضي','',now]));}
 const acc=sheet('Accounts'); if(acc.getLastRow()<2){[['1000','الصندوق','Asset','USD','','Active'],['1100','ذمم المعتمرين','Asset','USD','','Active'],['4000','إيرادات الرحلات','Revenue','USD','','Active'],['5000','مصروفات الرحلات','Expense','USD','','Active']].forEach(r=>acc.appendRow([...r,now]));}}
function doGet(e){return out({ok:true,message:'Umrah ERP API يعمل',version:'complete-1.0'});}
function doPost(e){try{return out(handle(JSON.parse(e.postData.contents||'{}')))}catch(err){return out({ok:false,error:String(err),stack:err.stack})}}
function handle(p){if(p.action==='login')return login(p); if(p.action==='list')return listRows(p.sheet); if(p.action==='upsert')return upsert(p.sheet,p.record,p.user); if(p.action==='delete')return deleteRow(p.sheet,p.key,p.user); if(p.action==='dashboard')return dashboard(); if(p.action==='reports')return reports(); if(p.action==='portalLookup')return portalLookup(p.key); return {ok:false,error:'Unknown action '+p.action};}
function login(p){const rows=getObjects('Users'); const u=rows.find(r=>String(r.username)===String(p.username)&&String(r.password)===String(p.password)&&String(r.status).toLowerCase()==='active'); if(!u)return {ok:false,error:'بيانات الدخول غير صحيحة أو المستخدم غير فعال'}; log(u.username,'login','دخول للنظام'); return {ok:true,user:{username:u.username,role:u.role,name:u.name,pilgrimCode:u.pilgrimCode}};}
function listRows(n){check(n); return {ok:true,rows:getObjects(n)};}
function upsert(n,rec,user){check(n); const sh=sheet(n), headers=DB[n], key=headers[0]; if(!rec[key]) rec[key]=nextId(n); if(!rec.createdAt && headers.includes('createdAt')) rec.createdAt=new Date(); autoCalc(n,rec); const rows=sh.getDataRange().getValues(); let found=-1; for(let i=1;i<rows.length;i++) if(String(rows[i][0])===String(rec[key])) found=i+1; const vals=headers.map(h=>rec[h]!==undefined?rec[h]:''); if(found>0) sh.getRange(found,1,1,headers.length).setValues([vals]); else sh.appendRow(vals); if(n==='Payments'||n==='Pilgrims') recalcPilgrim(rec.pilgrimCode); if(n==='GuideAdvances') recalcGuide(rec.guideId); log(user?.username||'system','upsert '+n,JSON.stringify(rec)); return {ok:true,id:rec[key]};}
function deleteRow(n,key,user){check(n); const sh=sheet(n), rows=sh.getDataRange().getValues(); for(let i=1;i<rows.length;i++){if(String(rows[i][0])===String(key)){sh.deleteRow(i+1); log(user?.username||'system','delete '+n,String(key)); return {ok:true};}} return {ok:false,error:'السجل غير موجود'};}
function portalLookup(key){const p=getObjects('Pilgrims').find(r=>String(r.pilgrimCode)===String(key)||String(r.phone)===String(key)); return p?{ok:true,pilgrim:p}:{ok:false};}
function dashboard(){const stats={}; ['Trips','Pilgrims','Payments','Expenses','Visas','Rooms','Buses','Guides','Complaints'].forEach(n=>stats[n]=Math.max(0,sheet(n).getLastRow()-1)); const pay=sum('Payments','amountUSD'), exp=sum('Expenses','amountUSD'); const finance=`إجمالي المقبوضات: ${pay.toFixed(2)} USD<br>إجمالي المصروفات: ${exp.toFixed(2)} USD<br>صافي تقديري: ${(pay-exp).toFixed(2)} USD`; const alerts=`جوازات غير مسلمة: ${countWhere('Pilgrims','passportStatus','!=','مسلم')}<br>تأشيرات غير صادرة: ${countWhere('Pilgrims','visaStatus','!=','صادرة')}<br>شكاوى مفتوحة: ${countWhere('Complaints','status','!=','مغلقة')}`; return {ok:true,stats,finance,alerts};}
function reports(){const pay=sum('Payments','amountUSD'), exp=sum('Expenses','amountUSD'); let txt='تقرير مختصر للنظام\n'; txt+='====================\n'; txt+='عدد الرحلات: '+(sheet('Trips').getLastRow()-1)+'\n'; txt+='عدد المعتمرين: '+(sheet('Pilgrims').getLastRow()-1)+'\n'; txt+='المقبوضات USD: '+pay.toFixed(2)+'\n'; txt+='المصروفات USD: '+exp.toFixed(2)+'\n'; txt+='النتيجة USD: '+(pay-exp).toFixed(2)+'\n\n'; txt+='ملاحظة: هذه تقارير تشغيلية داخل Google Sheets وليست بديلاً عن برنامج محاسبة قانوني عند التوسع الكبير.'; return {ok:true,text:txt};}
function autoCalc(n,r){['Payments','Expenses','JournalEntries','GuideAdvances'].forEach(s=>{if(n===s){const a=num(r.amount), rate=num(r.rateToUSD)||1; r.amountUSD=(a*rate).toFixed(2);}}); if(n==='Pilgrims'){r.priceUSD=num(r.priceUSD); r.paidUSD=num(r.paidUSD); r.balanceUSD=(num(r.priceUSD)-num(r.paidUSD)).toFixed(2);}}
function recalcPilgrim(code){if(!code)return; const ps=getObjects('Payments').filter(x=>String(x.pilgrimCode)===String(code)); const paid=ps.reduce((s,x)=>s+num(x.amountUSD),0); const sh=sheet('Pilgrims'), rows=sh.getDataRange().getValues(), h=DB.Pilgrims; const codeCol=h.indexOf('pilgrimCode'), paidCol=h.indexOf('paidUSD')+1, balCol=h.indexOf('balanceUSD')+1, priceCol=h.indexOf('priceUSD'); for(let i=1;i<rows.length;i++){if(String(rows[i][codeCol])===String(code)){const price=num(rows[i][priceCol]); sh.getRange(i+1,paidCol).setValue(paid); sh.getRange(i+1,balCol).setValue((price-paid).toFixed(2));}}}
function recalcGuide(id){if(!id)return; const rows=getObjects('GuideAdvances').filter(x=>String(x.guideId)===String(id)); const bal=rows.reduce((s,x)=>s+(String(x.type).includes('تسوية')?-num(x.amountUSD):num(x.amountUSD)),0); const sh=sheet('Guides'), vals=sh.getDataRange().getValues(), h=DB.Guides; const idc=h.indexOf('guideId'), bc=h.indexOf('advanceBalanceUSD')+1; for(let i=1;i<vals.length;i++) if(String(vals[i][idc])===String(id)) sh.getRange(i+1,bc).setValue(bal.toFixed(2));}
function sheet(n){return SpreadsheetApp.getActive().getSheetByName(n)||SpreadsheetApp.getActive().insertSheet(n)}
function check(n){if(!DB[n])throw new Error('جدول غير معروف: '+n);}
function getObjects(n){check(n); const sh=sheet(n), v=sh.getDataRange().getValues(); if(v.length<2)return []; const h=v[0]; return v.slice(1).filter(r=>r.join('')!=='').map(r=>Object.fromEntries(h.map((x,i)=>[x,r[i] instanceof Date?Utilities.formatDate(r[i],Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm:ss'):r[i]])));}
function nextId(n){return (PREFIX[n]||'ID')+'-'+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyyMMddHHmmss')+'-'+Math.floor(Math.random()*999);}
function sum(n,f){return getObjects(n).reduce((s,r)=>s+num(r[f]),0)}
function countWhere(n,f,op,val){return getObjects(n).filter(r=>op==='!='?String(r[f])!==String(val):String(r[f])===String(val)).length}
function log(user,action,details){const sh=sheet('Logs'); if(sh.getLastRow()===0) sh.appendRow(DB.Logs); sh.appendRow([nextId('Logs'),user,action,details,new Date()]);}
function out(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON)}
function num(x){return Number(x||0)||0}
function today(){return Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd')}
