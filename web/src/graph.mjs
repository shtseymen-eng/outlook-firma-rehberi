import {messageContacts,mergeContact} from './contacts.mjs';
export const GRAPH='https://graph.microsoft.com/v1.0';
export function trustedGraphUrl(value){
 const url=new URL(value.startsWith('/')?GRAPH+value:value);
 if(url.origin!=='https://graph.microsoft.com'||!url.pathname.startsWith('/v1.0/me')||!/^\/v1\.0\/me(?:\/|$)/.test(url.pathname)||url.username||url.password||url.hash)throw new Error('Geçersiz Microsoft Graph bağlantısı.');
 return url.href;
}
const abortIf=signal=>{if(signal?.aborted)throw new DOMException('Tarama durduruldu.','AbortError');};
export function sleep(ms,signal){return new Promise((resolve,reject)=>{abortIf(signal);const finish=()=>{signal?.removeEventListener('abort',cancel);resolve();};const timer=setTimeout(finish,ms);const cancel=()=>{clearTimeout(timer);reject(new DOMException('Tarama durduruldu.','AbortError'));};signal?.addEventListener('abort',cancel,{once:true});});}
export class GraphClient{
 constructor(tokenProvider,options={}){this.tokenProvider=tokenProvider;this.fetch=options.fetch||globalThis.fetch.bind(globalThis);this.sleep=options.sleep||sleep;}
 async request(account,path,signal){
  const url=trustedGraphUrl(path);
  for(let attempt=0;attempt<4;attempt++){
   abortIf(signal);const token=await this.tokenProvider(account);let response;
   try{response=await this.fetch(url,{method:'GET',headers:{Authorization:'Bearer '+token,Accept:'application/json',Prefer:'outlook.body-content-type="text", IdType="ImmutableId", odata.maxpagesize=100'},signal:signal?AbortSignal.any([signal,AbortSignal.timeout(45000)]):AbortSignal.timeout(45000)});}
   catch(error){abortIf(signal);if(attempt<3){await this.sleep(1000*2**attempt,signal);continue;}throw new Error('Microsoft bağlantısı kurulamadı. İnternet bağlantınızı kontrol edin.');}
   if([401,403].includes(response.status))throw new Error('Microsoft erişimi reddetti. Hesabı yeniden bağlayın; gerekirse kurum yöneticinizden Mail.Read izni alın.');
   if([429,500,502,503,504].includes(response.status)){
    if(attempt===3)throw new Error('Microsoft geçici olarak istekleri sınırladı. Daha sonra yeniden deneyin.');
    const raw=response.headers.get('Retry-After');let delay=raw?(Number.isFinite(Number(raw))?Number(raw)*1000:Date.parse(raw)-Date.now()):1000*2**attempt;
    delay=Math.max(0,Number.isFinite(delay)?delay:1000*2**attempt);if(delay>300000)throw new Error('Microsoft daha uzun bekleme istiyor. Taramayı daha sonra yeniden başlatın.');
    await this.sleep(delay,signal);continue;
   }
   if(response.status===410){const error=new Error('Klasörün ilerleme kaydı yenilenmeli. Baştan tara seçeneğini kullanın.');error.code='checkpoint_expired';throw error;}
   if(!response.ok)throw new Error('Microsoft Graph yanıtı: HTTP '+response.status);
   let payload;try{payload=await response.json();}catch{throw new Error('Microsoft geçersiz bir yanıt döndürdü.');}
   if(!payload||typeof payload!=='object'||Array.isArray(payload))throw new Error('Microsoft yanıtı okunamadı.');return payload;
  }
 }
 async profile(account,signal){return this.request(account,'/me?$select=id,displayName,mail,userPrincipalName',signal);}
 async folders(account,signal){
  const queue=[{url:'/me/mailFolders?includeHiddenFolders=true&$top=200&$select=id,displayName,parentFolderId,childFolderCount,isHidden,totalItemCount',parent:''}];
  const result=[],seen=new Set(),pages=new Set();
  while(queue.length){const {url,parent}=queue.shift();const full=trustedGraphUrl(url);if(pages.has(full))throw new Error('Klasör sayfalaması döngüye girdi.');pages.add(full);
   const payload=await this.request(account,url,signal);if(!Array.isArray(payload.value))throw new Error('Microsoft klasör listesi okunamadı.');
   for(const item of payload.value){if(!item.id||seen.has(item.id))continue;seen.add(item.id);const path=parent?parent+' › '+item.displayName:item.displayName;result.push({...item,path});
    if(item.childFolderCount>0)queue.push({url:'/me/mailFolders/'+encodeURIComponent(item.id)+'/childFolders?includeHiddenFolders=true&$top=200&$select=id,displayName,parentFolderId,childFolderCount,isHidden,totalItemCount',parent:path});
   }
   if(payload['@odata.nextLink'])queue.unshift({url:payload['@odata.nextLink'],parent});
  }
  return result;
 }
 async *messages(account,folder,checkpoint,signal){
  const fields='id,from,sender,toRecipients,ccRecipients,bccRecipients,body,receivedDateTime,sentDateTime';
  let url=checkpoint||'/me/mailFolders/'+encodeURIComponent(folder.id)+'/messages/delta?$select='+fields;const seen=new Set();
  while(url){const full=trustedGraphUrl(url);if(seen.has(full))throw new Error('Posta sayfalaması döngüye girdi.');seen.add(full);
   const payload=await this.request(account,url,signal);if(!Array.isArray(payload.value))throw new Error('Microsoft posta listesi okunamadı.');
   yield {value:payload.value,next:payload['@odata.nextLink']||'',delta:payload['@odata.deltaLink']||''};url=payload['@odata.nextLink'];
  }
 }
}
export async function scanFolders({graph,work,directory,checkpoints,own,signal,full=false,onProgress=()=>{},validateCheckpoint=trustedGraphUrl,persistCheckpoints=true}){
 const startedAt=new Date().toISOString();let processed=0,committedMessages=0,completed=0,cancelled=false;const warnings=[];
 const uniqueWork=[...new Map(work.map(item=>[JSON.stringify([item.account.id,item.folder.id]),item])).values()];
 for(const {account,folder} of uniqueWork){
  const key=JSON.stringify([account.id,folder.id]),staged=new Map();let delta='',folderMessages=0;
  try{
   abortIf(signal);const start=full?null:checkpoints.get(key);
   const source={account:account.email,accountId:account.id,folderId:folder.id,folder:folder.path};
   for await(const page of graph.messages(account,folder,start,signal)){
    abortIf(signal);if(page.delta)delta=validateCheckpoint(page.delta);
    for(const message of page.value){abortIf(signal);for(const contact of messageContacts(message,source,own))staged.set(contact.email,mergeContact(staged.get(contact.email),contact));processed++;folderMessages++;}
    onProgress({processed,completed,total:uniqueWork.length,contacts:directory.size,folder:folder.path,account:account.email});
   }
   abortIf(signal);if(!delta)throw new Error('Klasör taraması tamamlanma kaydı alınamadı; bu klasör yeniden denenebilir.');
   for(const contact of staged.values())directory.set(contact.email,mergeContact(directory.get(contact.email),contact));
   if(persistCheckpoints)checkpoints.set(key,delta);completed++;committedMessages+=folderMessages;
   onProgress({processed,completed,total:uniqueWork.length,contacts:directory.size,folder:folder.path,account:account.email});
  }catch(error){warnings.push({account:account.email,folder:folder.path,message:error.message});if(signal?.aborted||error.name==='AbortError'){cancelled=true;break;}}
 }
 return {startedAt,completedAt:new Date().toISOString(),complete:!cancelled&&completed===uniqueWork.length&&warnings.length===0,processed,committedMessages,completed,total:uniqueWork.length,accounts:new Set(uniqueWork.map(x=>x.account.id)).size,cancelled,warnings};
}
