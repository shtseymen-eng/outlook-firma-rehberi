export const CATEGORY_LABELS={polisan_poliport:'Polisan-Poliport',customs_ygm:'Gümrük-YGM',corporate:'Diğer Kurumsal',personal:'Kişisel Mailler'};
const PERSONAL=new Set(['gmail.com','googlemail.com','hotmail.com','hotmail.com.tr','outlook.com','live.com','live.com.tr','windowslive.com','msn.com','yahoo.com','yahoo.com.tr','ymail.com','yandex.com','yandex.com.tr','yandex.ru','icloud.com','me.com','proton.me','protonmail.com','aol.com','aim.com']);
const TITLES=['müdür','uzman','şef','direktör','koordinatör','asistan','yetkili','sorumlu','müşavir','operatör','lojistik','satın alma','muhasebe','finans','insan kaynakları','pazarlama','satış','yönetici','mühendis','tekniker','danışman','analist','planlama','depo','sevkiyat','ihracat','ithalat','operasyon','manager','director','chief','officer','specialist','coordinator','assistant','analyst','engineer','consultant','supervisor','executive','head of','logistics','procurement','operations','sales','marketing','finance','accounting','transport','shipping','export','import','supply chain','customs broker'];
const GREETING=/^\s*(?:--+|iyi çalışmalar|saygılarımla|regards|best regards|thanks|teşekkürler|kind regards|br,?|cheers|sincerely|iyi günler|good day)\s*[,.]?\s*$/i;
const QUOTED=/^\s*(?:from|kimden|sent|gönderilen|to|kime|subject|konu)\s*:|^\s*-{2,}\s*(?:original message|özgün ileti)\s*-{2,}\s*$|^\s*on\s+.+\s+wrote\s*:\s*$/i;
const SUFFIX=/\b(?:a\.?\s*ş\.?|ltd\.?|şti\.?|inc\.?|co\.?|llc\.?|gmbh)\b/i;
export const searchText=value=>String(value||'').toLocaleLowerCase('tr-TR').replace(/ı/g,'i').normalize('NFKD').replace(/\p{M}/gu,'').replace(/[^a-z0-9]+/g,' ').trim();
export function normalizeEmail(value){
 const raw=String(value||'').trim();if(raw.split('@').length!==2)return null;
 const [local,rawDomain]=raw.toLowerCase().split('@');const domain=rawDomain.replace(/\.+$/,'');
 if(!/^[a-z0-9.!#$%&'*+/=?^_`{|}~-]+$/.test(local)||!domain.includes('.'))return null;
 let ascii;try{ascii=new URL('https://'+domain).hostname;}catch{return null;}
 if(!/^[a-z0-9.-]+$/.test(ascii)||ascii.split('.').some(label=>!label||label.startsWith('-')||label.endsWith('-'))||/[\s/\\?#:@]/.test(domain))return null;
 return local+'@'+ascii;
}
export function classify(contact){
 const domain=String(contact.domain||contact.email?.split('@')[1]||'').toLowerCase();
 if(['polisan.com.tr','poliport.com'].includes(domain))return 'polisan_poliport';
 if(PERSONAL.has(domain))return 'personal';
 const haystack=searchText([contact.email,domain,contact.company,contact.title,contact.signature].join(' '));
 if(/(?:^|\s)ygm(?:$|\s)/.test(haystack)||['gumruk','gumrukleme','musavirlik','customs'].some(term=>haystack.includes(term)))return 'customs_ygm';
 return 'corporate';
}
function nameLine(line){const v=String(line||'').replace(/\([^)]*\)/g,'').replace(/"/g,'').trim();if(/[@\d]/.test(v)||GREETING.test(v)||SUFFIX.test(v)||TITLES.some(t=>v.toLowerCase().includes(t)))return '';return v.split(/\s+/).length>=2&&v.split(/\s+/).length<=5?v:'';}
function htmlText(html){
 // Text-only conversion: never attach message HTML or load its images into the page.
 return String(html).replace(/<(script|style)\b[^>]*>[\s\S]*?<\/\1>/gi,'').replace(/<blockquote\b[^>]*>[\s\S]*$/gi,'').replace(/<[^>]*(?:gmail_quote|yahoo_quoted|divRplyFwdMsg|reference-message)[^>]*>[\s\S]*$/gi,'').replace(/<\/?(?:p|div|tr|br|li|td|th)\b[^>]*>/gi,'\n').replace(/<[^>]*>/g,'').replace(/&(?:nbsp|amp|lt|gt|quot|apos);/g,m=>({'&nbsp;':' ','&amp;':'&','&lt;':'<','&gt;':'>','&quot;':'"','&apos;':"'"}[m])).replace(/&#(x[0-9a-f]+|\d+);/gi,(_,n)=>{const code=n[0].toLowerCase()==='x'?parseInt(n.slice(1),16):Number(n);return code<=0x10ffff?String.fromCodePoint(code):'';});
}
function signature(body){
 let lines=[];for(const line of body.split(/\r?\n/)){if(QUOTED.test(line))break;lines.push(line.trim());}
 const start=lines.findIndex(l=>GREETING.test(l));lines=start>=0?lines.slice(start):lines.slice(-20);
 const text=lines.join('\n').slice(0,5000);
 if(start<0&&(!lines.some(nameLine)||!/[\d@]|ltd|manager|müdür/i.test(text)))return '';
 return text;
}
const ranked=(value,quality,at)=>value?{value,quality,at}:null;
const pick=(a,b)=>!a?b:!b?a:b.quality>a.quality?b:b.quality<a.quality?a:(Date.parse(b.at)||0)>(Date.parse(a.at)||0)?b:a;
const unique=values=>[...new Set(values.filter(Boolean))].sort();
export function messageContacts(message,source,ownAddresses){
 if(message['@removed'])return [];
 const own=new Set([...ownAddresses].map(normalizeEmail).filter(Boolean));
 const at=message.receivedDateTime||message.sentDateTime||'1970-01-01T00:00:00Z';
 const contacts=new Map();
 function add(item,isSender=false){
  const addr=item?.emailAddress||{};const email=normalizeEmail(addr.address);if(!email||own.has(email))return;
  if(contacts.has(email))return;
  let sig='';if(isSender){const body=message.uniqueBody||message.body||{};sig=signature(body.contentType?.toLowerCase()==='html'?htmlText(body.content||''):String(body.content||''));}
  const lines=sig.split('\n');const sigName=lines.map(nameLine).find(Boolean)||'';
  const title=lines.find(l=>TITLES.some(t=>l.toLowerCase().includes(t)))||'';
  const company=lines.find(l=>l.length>2&&l.length<=100&&!/@|\d{4,}/.test(l)&&(SUFFIX.test(l)||(/[A-ZÇĞİÖŞÜ]/.test(l)&&l===l.toLocaleUpperCase('tr-TR')&&!TITLES.some(t=>l.toLowerCase().includes(t)))))||'';
  const mobileMatches=sig.match(/(?<!\d)(?:\+?90[\s.\-]?)?0?5\d{2}[\s.\-]?\d{3}[\s.\-]?\d{2}[\s.\-]?\d{2}(?!\d)/g)||[];
  const fixedMatches=sig.match(/(?<!\d)(?:0\s*[2-4]\d{2})[\s.\-]?\d{3}[\s.\-]?\d{2}[\s.\-]?\d{2}(?!\d)/g)||[];
  const display=nameLine(addr.name);
  contacts.set(email,{email,domain:email.split('@')[1],name:ranked(sigName||display,sigName?30:20,at),title:ranked(title,30,at),company:ranked(company,30,at),mobiles:unique(mobileMatches.map(v=>v.replace(/\D/g,''))),landlines:unique(fixedMatches.map(v=>v.replace(/\D/g,''))),sources:[source],signature:sig,signatureAt:at});
 }
 add(message.from||message.sender,true);
 for(const item of [...(message.toRecipients||[]),...(message.ccRecipients||[]),...(message.bccRecipients||[])])add(item);
 return [...contacts.values()];
}
export function mergeContact(a,b){
 if(!a)return structuredClone(b);if(a.email!==b.email)throw new Error('Farklı e-posta adresleri birleştirilemez.');
 const sources=new Map([...a.sources,...b.sources].map(s=>[JSON.stringify([s.accountId,s.folderId]),s]));
 const latest=Date.parse(b.signatureAt)>Date.parse(a.signatureAt)&&b.signature?b:a;
 return {...a,name:pick(a.name,b.name),title:pick(a.title,b.title),company:pick(a.company,b.company),mobiles:unique([...a.mobiles,...b.mobiles]),landlines:unique([...a.landlines,...b.landlines]),sources:[...sources.values()],signature:latest.signature||'',signatureAt:latest.signatureAt||'',overrides:{...b.overrides,...a.overrides}};
}
export function contactValues(contact){
 const overrides=contact.overrides||{};const value=field=>Object.hasOwn(overrides,field)?overrides[field]:contact[field]?.value||'';
 const flat={email:contact.email,domain:contact.domain,name:value('name'),title:value('title'),company:value('company')||contact.domain.split('.')[0].toUpperCase(),mobiles:contact.mobiles,landlines:contact.landlines,sources:contact.sources,signature:contact.signature||''};
 flat.category=classify(flat);return flat;
}
export function importLegacy(payload){
 if(!payload||typeof payload.contacts!=='object'||Array.isArray(payload.contacts))throw new Error('Geçerli bir eski rehber yedeği seçin.');
 const contacts=[];let skipped=0;const at=new Date().toISOString();
 for(const [key,item] of Object.entries(payload.contacts)){
  const email=item&&typeof item==='object'?normalizeEmail(item.email||key):null;if(!email){skipped++;continue;}
  const source={account:'Eski v4',accountId:'legacy-v4',folderId:'legacy-import',folder:'v4 içe aktarım'};
  contacts.push({email,domain:email.split('@')[1],name:ranked(String(item.name||'').slice(0,500),25,at),title:ranked(String(item.title||'').slice(0,500),25,at),company:ranked(String(item.firma_display||item.firma||'').slice(0,500),25,at),mobiles:unique(Array.isArray(item.gsm)?item.gsm.map(String):[]),landlines:unique(Array.isArray(item.sabit)?item.sabit.map(String):[]),sources:[source],signature:'',signatureAt:at});
 }
 return {contacts,skipped};
}
