import {scanFolders} from './graph.mjs';
export function localSession(value){
 const url=new URL(value),match=url.hash.match(/^#mail=([a-f0-9]{64})$/);
 return url.protocol==='http:'&&url.hostname==='127.0.0.1'&&url.port&&match?{origin:url.origin,token:match[1]}:null;
}
export class LocalMailClient{
 constructor(session,options={}){if(!localSession(session.origin+'/#mail='+session.token))throw new Error('Geçersiz yerel bağlantı.');this.origin=session.origin;this.token=session.token;this.fetch=options.fetch||globalThis.fetch.bind(globalThis);this.signatures=false;}
 async request(command,payload={},signal){
  if(!['accounts','folders','snapshot','messages'].includes(command))throw new Error('Geçersiz Mail işlemi.');
  let response;
  try{response=await this.fetch(this.origin+'/api/'+command,{method:'POST',headers:{Authorization:'Bearer '+this.token,'Content-Type':'application/json'},body:JSON.stringify(payload),credentials:'omit',redirect:'error',signal:signal?AbortSignal.any([signal,AbortSignal.timeout(190000)]):AbortSignal.timeout(190000)});}
  catch(error){if(signal?.aborted)throw new DOMException('Tarama durduruldu.','AbortError');throw new Error('Mac Mail yardımcısına ulaşılamadı. Yardımcıyı açık tutup rehber ekranını yeniden açın.');}
  const data=await response.json();if(!response.ok)throw new Error(data.error||'Mail okunamadı.');return data;
 }
 async accounts(signal){const data=await this.request('accounts',{},signal);if(!Array.isArray(data.accounts))throw new Error('Hesap listesi okunamadı.');return data.accounts.map(a=>({id:a.id,email:a.addresses?.[0]||a.name,name:a.name,own:a.addresses||[],source:'local'}));}
 async folders(account,signal){const data=await this.request('folders',{account:account.id},signal);if(!Array.isArray(data.folders))throw new Error('Klasörler okunamadı.');return data.folders;}
 async *messages(account,folder,_checkpoint,signal){
  const payload={account:account.id,path:folder.parts};
  const snapshot=await this.request('snapshot',payload,signal);
  if(!Array.isArray(snapshot.ids)||snapshot.ids.some(id=>!Number.isSafeInteger(id)||id<0)||new Set(snapshot.ids).size!==snapshot.ids.length)throw new Error('Klasör mesaj listesi okunamadı.');
  for(let offset=0;offset<snapshot.ids.length;offset+=25){
   const ids=snapshot.ids.slice(offset,offset+25),data=await this.request('messages',{...payload,ids,signatures:this.signatures},signal);
   if(!Array.isArray(data.messages)||data.messages.length!==ids.length||data.messages.some((m,i)=>String(m.id)!==String(ids[i])))throw new Error('Klasörde bazı mesajlar okunamadı. Bu klasörün sonuçları kaydedilmedi; yeniden deneyin.');
   yield {value:data.messages};
  }
  yield {value:[],delta:'local-complete'};
 }
}
export function scanLocalFolders(options){
 return scanFolders({...options,graph:options.client,full:true,persistCheckpoints:false,validateCheckpoint:value=>{if(value!=='local-complete')throw new Error('Klasör tamamlanamadı.');return value;}});
}
