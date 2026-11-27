import {PublicClientApplication} from '@azure/msal-browser';
import {parseMemoWorkbook,buildMemoWorkbook,WORKBOOK_FILE_NAME,WORKBOOK_MIME,WORKBOOK_FORMAT_VERSION} from './nolty-memo-workbook.js';

const scopes=['Files.ReadWrite.AppFolder'];
const base='https://graph.microsoft.com/v1.0';
export function memoTarget(mode,date){
  if(!['weekly','daily'].includes(mode)||!/^\d{4}-\d{2}-\d{2}$/.test(date))throw new Error('Weekly / Dailyと日付を選択してください');
  const day=new Date(date+'T00:00:00Z');
  if(!Number.isFinite(day.getTime())||day.toISOString().slice(0,10)!==date)throw new Error('日付が正しくありません');
  day.setUTCDate(day.getUTCDate()-((day.getUTCDay()+6)%7));
  const memoType=mode==='weekly'?'DAY':'DAILY_VIEW';
  return {memoId:`${memoType}-${date}`,memoType,weekStart:day.toISOString().slice(0,10),targetDate:date};
}

// Use the NOLTY app registration, not the TASK CONTROL app folder.
export function createNoltyAuth(config,getAccount,onChange){
  let client,ready,job,verified=false,connectedId='',message='';
  function notify(){onChange?.();}
  function init(){if(!ready){client=new PublicClientApplication({auth:{clientId:config.NOLTY_CLIENT_ID,authority:config.AUTHORITY,redirectUri:config.REDIRECT_URI,navigateToLoginRequestUrl:false},cache:{cacheLocation:'localStorage'}});ready=client.initialize().catch(e=>{ready=null;throw e;});}return ready;}
  function verify(result,primary){if(!result?.account||result.account.homeAccountId!==primary.homeAccountId||getAccount()?.homeAccountId!==primary.homeAccountId)throw new Error('TASK CONTROLと同じMicrosoftアカウントでNOLTYに接続してください');client.setActiveAccount(result.account);verified=true;connectedId=primary.homeAccountId;return result.accessToken;}
  function run(action){if(job)return job;verified=false;message='';job=Promise.resolve().then(action).catch(e=>{message=String(e.errorCode||e.message||'認証失敗');return null;}).finally(()=>{job=null;notify();});notify();return job;}
  async function silent(){const primary=getAccount();if(!navigator.onLine||!primary)return null;return run(async()=>{await init();const account=client.getAllAccounts().find(a=>a.homeAccountId===primary.homeAccountId);let result;if(account){try{result=await client.acquireTokenSilent({account,scopes});}catch{result=await client.ssoSilent({account,scopes});}}else result=await client.ssoSilent({loginHint:primary.username,scopes});return verify(result,primary);});}
  async function connect(){if(!navigator.onLine||!getAccount()||job)return null;const primary=getAccount();await init();return run(async()=>verify(await client.loginPopup({scopes,loginHint:primary.username,redirectUri:config.REDIRECT_URI}),primary));}
  return {init,silent,connect,get busy(){return Boolean(job)},get connected(){return verified&&connectedId===getAccount()?.homeAccountId},get message(){return message}};
}

// Retain a checkpoint before every write. A lost response cannot silently
// append the same capture twice; ambiguous external edits remain pending.
export async function appendNoltyMemo(record,access,checkpoint,{fetcher=fetch,parse=parseMemoWorkbook,build=buildMemoWorkbook}={}){
  const target=memoTarget(record.noltyMode,record.noltyDate);
  const text=String(record.notes||'').trim();
  if(!text||/[\x00-\x08\x0b\x0c\x0e-\x1f]/.test(text))throw new Error('メモ内容が空か、Excelで使用できない文字が含まれています');
  const auth={Authorization:`Bearer ${access}`};
  for(let n=0;n<4;n++){
    const metaResponse=await fetcher(`${base}/me/drive/special/approot:/${encodeURIComponent(WORKBOOK_FILE_NAME)}?$select=id,eTag`,{headers:auth});
    if(metaResponse.status===404)throw new Error('Digital NOLTYのメモMasterがありません。先にDigital NOLTYでOneDrive同期してください');
    if(!metaResponse.ok)throw new Error(`NOLTY Master取得失敗（HTTP ${metaResponse.status}）`);
    const meta=await metaResponse.json();if(!meta.id||!meta.eTag)throw new Error('NOLTY Masterの更新確認情報を取得できません');
    const url=`${base}/me/drive/items/${encodeURIComponent(meta.id)}/content`;
    const content=await fetcher(url,{headers:auth});if(!content.ok)throw new Error(`NOLTY Master読込失敗（HTTP ${content.status}）`);
    const records=await parse(await content.arrayBuffer());
    if(records.workbookFormatVersion!==WORKBOOK_FORMAT_VERSION||records.some(r=>r.memoType==='WEEK')||new Set(records.map(r=>r.memoId)).size!==records.length)throw new Error('NOLTY Masterの形式を確認できません。先にDigital NOLTYで同期してください');
    const previous=records.find(r=>r.memoId===target.memoId),old=previous?.memoText||'';
    const receipt=`QC:${record.id}`;
    if(previous?.updatedBy===receipt)return true;
    if(record.noltyAttempt){const {before,after}=record.noltyAttempt;if(old===after||old.startsWith(after+'\n'))return true;if(old!==before)throw new Error('前回の送信結果と現在のNOLTYメモが異なります。Digital NOLTYで内容を確認してください（重複防止のため未同期保持）');}
    const nextText=old?`${old}${old.endsWith('\n')?'\n':'\n\n'}${text}`:text;
    if(nextText.length>32767)throw new Error('Excelのメモ文字数上限を超えます。Digital NOLTYでメモを整理してください');
    const replacement={...target,memoText:nextText,updatedAt:new Date().toISOString(),updatedBy:receipt,revision:(previous?.revision||0)+1};
    const next=await build([...records.filter(r=>r.memoId!==target.memoId),replacement]);
    record.noltyAttempt={before:old,after:nextText};await checkpoint(record);
    const response=await fetcher(url,{method:'PUT',headers:{...auth,'Content-Type':WORKBOOK_MIME,'If-Match':meta.eTag},body:next});
    if(response.ok)return true;
    if(response.status===409||response.status===412){delete record.noltyAttempt;await checkpoint(record);continue;}
    if(response.status>=400&&response.status<500){delete record.noltyAttempt;await checkpoint(record);}
    throw new Error(`NOLTYメモ保存失敗（HTTP ${response.status}）`);
  }
  throw new Error('同時更新が続いています。オンライン保存で再試行してください');
}
