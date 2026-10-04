const BASE="https://graph.microsoft.com/v1.0";
const escapeHtml=value=>String(value||"").replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll(">","&gt;").replaceAll('"',"&quot;");
export function syncHash(value){let hash=2166136261;for(const char of value){hash^=char.charCodeAt(0);hash=Math.imul(hash,16777619);}return(hash>>>0).toString(16);}
const utc=value=>({dateTime:new Date(`${value.length===16?value+":00":value}+09:00`).toISOString().slice(0,-1),timeZone:"UTC"});
const tokyo=value=>({dateTime:`${value.slice(0,16)}:00`,timeZone:"Tokyo Standard Time"});
export function todoPayload(record){const marks={防衛:"🔴",収益:"🔵",能力:"🟢",構造:"🟣",探索:"🟡",ルーティン:"⚪"},payload={title:record.title,importance:record.priority==="🔥HOT"?"high":"normal",status:"notStarted",categories:[marks[record.category]?`${marks[record.category]} ${record.category}`:record.category],body:{contentType:"html",content:`TASK CONTROL TaskID: ${escapeHtml(record.id)}<br>分類: ${escapeHtml(record.category)}<br>このタスクの構造・進捗・依存関係はTASK CONTROLで管理します。`},dueDateTime:record.dueDate?utc(`${record.dueDate}T23:59:00`):null,isReminderOn:Boolean(record.reminderAt)};if(record.reminderAt)payload.reminderDateTime=utc(record.reminderAt);return payload;}
export function calendarPayload(record){return{subject:`[TC] ${record.title}`,start:tokyo(record.calendarStart),end:tokyo(record.calendarEnd),showAs:"busy",isReminderOn:false,body:{contentType:"HTML",content:`TASK CONTROLから登録された予定です。<br>TaskID: ${escapeHtml(record.id)}<br>タスクの階層・進捗・期日はTASK CONTROLで管理してください。`},transactionId:record.calendarTransactionId};}
async function graph(path,access,method="GET",body){const response=await fetch(path.startsWith("https:")?path:BASE+path,{method,headers:{Authorization:`Bearer ${access}`,Accept:"application/json",Prefer:'outlook.timezone="Tokyo Standard Time", IdType="ImmutableId"',...(body?{"Content-Type":"application/json"}:{})},body:body?JSON.stringify(body):undefined});if(!response.ok){let payload={};try{payload=await response.json();}catch{}throw new Error(`${path.includes("/todo/")?"Microsoft To Do":"Outlook"}：${payload.error?.message||`HTTP ${response.status}`}`);}return response.status===204?null:response.json();}
async function collection(path,access){const items=[];while(path){const page=await graph(path,access);items.push(...(page.value||[]));path=page["@odata.nextLink"]||"";}return items;}
export async function syncMicrosoft(record,getToken,checkpoint,status){
 if(record.category==="外部メモ")return true;
 if(!record.todoDone){status("Microsoft To Do登録中…");const access=await getToken(["Tasks.ReadWrite"]);if(!access)return false;
  if(!record.todoListId){const lists=await collection("/me/todo/lists",access);let list=lists.find(x=>x.displayName==="TASK CONTROL リマインダー")||lists.find(x=>x.displayName==="TASK CONTROL");if(!list)list=await graph("/me/todo/lists",access,"POST",{displayName:"TASK CONTROL"});record.todoListId=list.id;checkpoint();}
  const path=`/me/todo/lists/${encodeURIComponent(record.todoListId)}/tasks`;
  if(!record.todoTaskId){const remotes=await collection(path+"?$top=100",access),remote=remotes.find(x=>String(x.body?.content||"").match(/TASK CONTROL TaskID:\s*([A-Za-z0-9-]+)/i)?.[1]===record.id);const created=remote||await graph(path,access,"POST",{title:record.title,body:todoPayload(record).body});record.todoTaskId=created.id;checkpoint();}
  const remote=await graph(`${path}/${encodeURIComponent(record.todoTaskId)}`,access,"PATCH",todoPayload(record));record.todoDone=true;record.todoModifiedAt=remote.lastModifiedDateTime||"";record.todoWebUrl=remote.webUrl||"";checkpoint();
 }
 if(record.calendarEnabled&&!record.calendarDone){status("Outlook予定登録中…");const access=await getToken(["Calendars.ReadWrite"]);if(!access)return false;
  if(!record.calendarEventId){const remote=await graph("/me/calendar/events",access,"POST",calendarPayload(record));record.calendarEventId=remote.id;record.calendarWebLink=remote.webLink||"";checkpoint();}
  record.calendarDone=true;checkpoint();
 }
 return true;
}
export function linkedRecords(record){if(record.category==="外部メモ")return[];const system=(prefix,title,fields)=>({id:`ZZZ-TC-${prefix}-${record.id}`,rootId:`ZZZ-TC-${prefix}-${record.id}`,title:`${title}｜${record.id}`,category:"その他",status:"完了",priority:"🧊COLD",owner:"",startDate:"",dueDate:"",progress:100,nextAction:"",blocker:"",notes:"",linkUrl:"",now:record.now,device:record.device,sortOrder:0,...fields}),out=[];
 if(record.reminderAt)out.push(system("REMINDER","リマインダー",{startDate:record.reminderAt.slice(0,10),dueDate:record.reminderAt.slice(0,10),nextAction:record.reminderAt}));
 if(record.todoTaskId)out.push(system("TODO","To Do連携",{blocker:record.todoListId,notes:record.todoTaskId,linkUrl:record.todoWebUrl||"",owner:syncHash(JSON.stringify(todoPayload(record))),nextAction:"notStarted",followNotes:JSON.stringify({type:"todo-reminder-v1",value:record.reminderAt||"",remoteModifiedAt:record.todoModifiedAt||"",syncedAt:new Date().toISOString()})}));
 if(record.calendarEventId)out.push(system("CALENDAR","Outlook予定",{rootId:record.calendarTransactionId,nextAction:record.calendarStart,blocker:record.calendarEnd,startDate:record.calendarStart.slice(0,10),dueDate:record.calendarEnd.slice(0,10),notes:record.calendarEventId,linkUrl:record.calendarWebLink||"",owner:syncHash(`${record.calendarStart}|${record.calendarEnd}`)}));
 return out;
}
