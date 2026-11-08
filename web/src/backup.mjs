import {normalizeEmail,importLegacy,mergeContact} from './contacts.mjs';
import {trustedGraphUrl} from './graph.mjs';
const text=value=>String(value??'').slice(0,5000);
const rank=value=>value&&typeof value==='object'?{value:text(value.value),quality:Math.max(0,Math.min(100,Number(value.quality)||0)),at:Number.isFinite(Date.parse(value.at))?new Date(value.at).toISOString():'1970-01-01T00:00:00Z'}:null;
function safeContact(value){
 if(!value||typeof value!=='object')throw new Error('Rehber kaydı geçersiz.');const email=normalizeEmail(value.email);if(!email)throw new Error('Rehberde geçersiz e-posta adresi var.');
 const list=x=>Array.isArray(x)?[...new Set(x.slice(0,100).map(text))]:[];
 const sources=Array.isArray(value.sources)?value.sources.slice(0,500).map(s=>({account:text(s.account),accountId:text(s.accountId),folderId:text(s.folderId),folder:text(s.folder)})):[];
 const overrides={};for(const field of ['name','title','company'])if(Object.hasOwn(value.overrides||{},field))overrides[field]=text(value.overrides[field]);
 return {email,domain:email.split('@')[1],name:rank(value.name),title:rank(value.title),company:rank(value.company),mobiles:list(value.mobiles),landlines:list(value.landlines),sources,signature:text(value.signature),signatureAt:text(value.signatureAt),overrides};
}
export function serializeWorkspace({directory,checkpoints,aliases,result}){
 return JSON.stringify({kind:'outlook-firma-rehberi-web',version:1,savedAt:new Date().toISOString(),contacts:[...directory.values()],checkpoints:[...checkpoints],aliases:[...aliases],result:result||null});
}
export function restoreWorkspace(payload){
 if(payload?.kind!=='outlook-firma-rehberi-web'){
  const legacy=importLegacy(payload);return {directory:new Map(legacy.contacts.map(c=>[c.email,c])),checkpoints:new Map(),aliases:new Map(),result:{complete:false,processed:0,completed:0,total:0,accounts:0,warnings:[{message:'Eski v4 rehberinden aktarıldı; yeni tarama henüz yapılmadı.'}]},skipped:legacy.skipped};
 }
 if(payload.version!==1||!Array.isArray(payload.contacts)||payload.contacts.length>250000)throw new Error('Çalışma yedeği biçimi veya sürümü desteklenmiyor.');
 const directory=new Map();for(const raw of payload.contacts){const contact=safeContact(raw);directory.set(contact.email,mergeContact(directory.get(contact.email),contact));}
 const checkpoints=new Map();for(const pair of payload.checkpoints||[]){if(!Array.isArray(pair)||pair.length!==2||typeof pair[0]!=='string'||typeof pair[1]!=='string')throw new Error('Yedek ilerleme kaydı geçersiz.');checkpoints.set(pair[0],trustedGraphUrl(pair[1]));}
 const aliases=new Map();for(const [key,values] of payload.aliases||[]){if(typeof key!=='string'||!Array.isArray(values))throw new Error('Yedek hesap adresleri geçersiz.');aliases.set(key,values.map(normalizeEmail).filter(Boolean));}
 return {directory,checkpoints,aliases,result:payload.result&&typeof payload.result==='object'?{...payload.result,complete:payload.result.complete===true}:null,skipped:0};
}
