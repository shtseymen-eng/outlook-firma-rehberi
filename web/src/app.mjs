import {MicrosoftAuth,validClientId,redirectUri,authError} from './auth.mjs';
import {GraphClient,scanFolders} from './graph.mjs';
import {LocalMailClient,localSession,scanLocalFolders} from './local-mail.mjs';
import {normalizeEmail,contactValues,CATEGORY_LABELS,searchText} from './contacts.mjs';
import {serializeWorkspace,restoreWorkspace} from './backup.mjs';
const $=id=>document.getElementById(id);
const pairing=localSession(location.href),local=pairing?new LocalMailClient(pairing):null;
if(location.hash.startsWith('#mail='))history.replaceState(null,'',location.pathname);
const isLocal=location.hostname==='127.0.0.1';
$('helperPanel').hidden=isLocal;
if(isLocal){$('setupButton').hidden=true;$('addAccount').textContent='Mac Mail hesaplarını göster';$('localOptions').hidden=false;}
const accounts=new Map(),trees=new Map();let directory=new Map(),checkpoints=new Map(),aliases=new Map(),selected=new Set(),lastResult=null,auth=null,busy=false,abort=null,page=0,selectedEmail='',editing='';
const PAGE_SIZE=75;let clientId='';
try{clientId=localStorage.getItem('outlook-rehberi-web-client-id')||import.meta.env.VITE_OUTLOOK_REHBERI_CLIENT_ID||'';}catch{}
if(validClientId(clientId))auth=new MicrosoftAuth(clientId);
$('clientId').value=clientId;$('callback').value=redirectUri();
function notice(message,error=false){$('notice').hidden=false;$('notice').textContent=message;$('notice').classList.toggle('error',error);}
function log(message,error=false){const p=document.createElement('p');p.textContent=new Date().toLocaleTimeString('tr-TR')+' · '+message;p.classList.toggle('error',error);$('log').append(p);while($('log').children.length>80)$('log').firstChild.remove();$('log').scrollTop=$('log').scrollHeight;}
function own(){return new Set([...accounts.values()].flatMap(a=>a.own).concat([...aliases.values()].flat()).map(normalizeEmail).filter(Boolean));}
function removeOwn(){const excluded=own();for(const email of excluded)directory.delete(email);}
function rawContacts(){const excluded=own();return [...directory.values()].filter(c=>!excluded.has(c.email));}
function controls(){
 const hasAccounts=accounts.size>0,hasContacts=rawContacts().length>0;
 for(const id of ['addAccount','setupButton','importBackup','clientId','saveSetup'])$(id).disabled=busy;
 for(const id of ['refreshFolders','activeAccount','editAliases','disconnect','selectFolders','clearFolders'])$(id).disabled=busy||!hasAccounts;
 $('fullScan').disabled=busy;$('readSignatures').disabled=busy;$('scan').disabled=busy||!selected.size||!hasAccounts;$('cancel').disabled=!busy||!abort;
 $('exportExcel').disabled=busy||!hasContacts||(!lastResult?.complete&&!$('allowPartial').checked);
 $('saveBackup').disabled=busy||!hasContacts;$('clearDirectory').disabled=busy||!hasContacts;$('editContact').disabled=busy||!selectedEmail;
 $('partialWarning').hidden=!hasContacts||lastResult?.complete===true;
 $('authStatus').textContent=isLocal?(local?'Mac Mail yardımcısı bağlı':'Yardımcıdan rehber ekranını yeniden açın'):(auth?'Microsoft girişi hazır':'Mac Mail yardımcısıyla tarayabilirsiniz');$('authStatus').classList.toggle('pending',isLocal?!local:!auth);
 $('accountCount').textContent=accounts.size;$('folderCount').textContent=selected.size;$('contactCount').textContent=rawContacts().length;
}
function renderAccounts(){
 const previous=$('activeAccount').value;$('activeAccount').replaceChildren();
 if(!accounts.size){const option=document.createElement('option');option.value='';option.textContent='Hesap seçin';$('activeAccount').append(option);}
 for(const account of accounts.values()){const option=document.createElement('option');option.value=account.id;option.textContent=account.email;$('activeAccount').append(option);}
 if(accounts.has(previous))$('activeAccount').value=previous;
 $('folderSummary').textContent=accounts.size?accounts.size+' hesap · '+selected.size+' klasör seçili':'Bağlı hesap yok';
 if(!accounts.size){controls();return;}
 $('folders').replaceChildren();
 for(const account of accounts.values()){
  const group=document.createElement('details');group.open=true;group.className='account-group';const summary=document.createElement('summary');const folders=trees.get(account.id)||[];const accountCheck=document.createElement('input');accountCheck.type='checkbox';accountCheck.setAttribute('aria-label',account.email+' tüm klasörler');const picked=folders.filter(f=>selected.has(JSON.stringify([account.id,f.id]))).length;accountCheck.checked=folders.length>0&&picked===folders.length;accountCheck.indeterminate=picked>0&&picked<folders.length;accountCheck.disabled=busy;accountCheck.addEventListener('click',event=>event.stopPropagation());accountCheck.addEventListener('change',()=>{for(const folder of folders){const key=JSON.stringify([account.id,folder.id]);if(accountCheck.checked)selected.add(key);else selected.delete(key);}renderAccounts();});summary.append(accountCheck,document.createTextNode(' '+account.email));summary.addEventListener('click',()=>{$('activeAccount').value=account.id;});group.append(summary);

  if(!folders.length){const p=document.createElement('p');p.className='hint';p.textContent='Klasör listesi alınamadı veya boş. Klasörleri yenile düğmesini kullanın.';group.append(p);}
  for(const folder of folders){const key=JSON.stringify([account.id,folder.id]),label=document.createElement('label'),input=document.createElement('input'),name=document.createElement('span'),count=document.createElement('small');
   label.className='folder-row';label.style.paddingLeft=16+folder.path.split(' › ').length*13+'px';input.type='checkbox';input.checked=selected.has(key);input.disabled=busy;name.textContent=folder.path.split(' › ').at(-1);label.title=folder.path;count.textContent=String(folder.totalItemCount||0);input.addEventListener('change',()=>{for(const item of folders.filter(f=>f.id===folder.id||f.path.startsWith(folder.path+' › '))){const itemKey=JSON.stringify([account.id,item.id]);if(input.checked)selected.add(itemKey);else selected.delete(itemKey);}renderAccounts();});label.append(input,name,count);group.append(label);
  }
  $('folders').append(group);
 }
 controls();
}
function renderContacts(){
 const query=searchText($('search').value),category=$('category').value;
 const items=rawContacts().map(contactValues).filter(c=>(!category||c.category===category)&&(!query||searchText([c.name,c.title,c.company,c.email,...c.mobiles,...c.landlines].join(' ')).includes(query))).sort((a,b)=>a.domain.localeCompare(b.domain)||a.email.localeCompare(b.email));
 const pages=Math.max(1,Math.ceil(items.length/PAGE_SIZE));page=Math.min(page,pages-1);$('contacts').replaceChildren();
 for(const c of items.slice(page*PAGE_SIZE,(page+1)*PAGE_SIZE)){
  const row=document.createElement('tr');row.tabIndex=0;row.classList.toggle('selected',c.email===selectedEmail);row.setAttribute('aria-label',c.email+' kaydı');
  for(const value of [c.name,c.title,c.company,c.email,CATEGORY_LABELS[c.category]]){const td=document.createElement('td');td.textContent=value||'—';row.append(td);}
  row.addEventListener('click',()=>{selectedEmail=c.email;renderContacts();});row.addEventListener('dblclick',()=>edit(c.email));row.addEventListener('keydown',event=>{if(event.key==='Enter'){selectedEmail=c.email;edit(c.email);}});$('contacts').append(row);
 }
 $('contactsEmpty').hidden=items.length>0;$('pageInfo').textContent=items.length?`${items.length} kayıt · ${page+1} / ${pages}`:'0 kayıt';$('previous').disabled=page===0;$('next').disabled=page===pages-1;controls();
}
function graph(){return new GraphClient(account=>auth.token(account));}
async function loadFolders(account){
 const items=await (account.source==='local'?local:graph()).folders(account);const oldKeys=new Set([...selected].filter(k=>JSON.parse(k)[0]===account.id));const hadTree=trees.has(account.id);
 trees.set(account.id,items);for(const key of oldKeys)selected.delete(key);for(const item of items){const key=JSON.stringify([account.id,item.id]);if(hadTree&&oldKeys.has(key))selected.add(key);}
}
$('setupButton').addEventListener('click',()=>{$('setupDialog').showModal();});
$('saveSetup').addEventListener('click',async()=>{
 const value=$('clientId').value.trim();if(!validClientId(value)){notice('Geçerli Microsoft uygulama kimliğini girin.',true);return;}
 if(accounts.size){notice('Giriş ayarını değiştirmeden önce bağlı hesapları ayırın.',true);return;}
 try{const candidate=new MicrosoftAuth(value);await candidate.ready;auth=candidate;clientId=value;try{localStorage.setItem('outlook-rehberi-web-client-id',value);}catch{}$('setupDialog').close();notice('Bağlantı ayarı kaydedildi. Microsoft tarafında dönüş adresinin de tanımlanmış olması gerekir.');controls();}catch(error){notice(authError(error),true);}
});
$('addAccount').addEventListener('click',async()=>{
 if(isLocal){await connectLocal();return;}
 if(!auth){$('setupDialog').showModal();notice('Hesap bağlamak için Microsoft giriş kurulumu bir kez tamamlanmalı.');return;}
 busy=true;controls();
 try{
  const login=await auth.connect();const ms=login.account;if(!ms)throw new Error('Hesap alınamadı.');
  const account={id:ms.homeAccountId,email:normalizeEmail(ms.username)||ms.username,name:ms.name||ms.username,auth:ms,own:[ms.username]};
  const profile=await graph().profile(account);account.email=normalizeEmail(profile.mail)||normalizeEmail(profile.userPrincipalName)||account.email;account.name=profile.displayName||account.name;account.own=[...new Set([account.email,profile.mail,profile.userPrincipalName,ms.username].map(normalizeEmail).filter(Boolean))];
  accounts.set(account.id,account);removeOwn();log(account.email+' hesabı bağlandı. Klasörler alınıyor…');await loadFolders(account);notice('Hesap bağlandı. Taranacak posta klasörlerini seçebilirsiniz.');
 }catch(error){notice(error.errorCode?authError(error):error.message,true);log('Hesap / klasör bağlantısı tamamlanamadı.',true);}
 finally{busy=false;renderAccounts();renderContacts();}
});
$('refreshFolders').addEventListener('click',async()=>{
 busy=true;renderAccounts();let failed=0;
 for(const account of accounts.values()){try{await loadFolders(account);log(account.email+' klasörleri yenilendi.');}catch(error){failed++;log(account.email+' · '+error.message,true);}}
 busy=false;renderAccounts();notice(failed?'Bazı hesapların klasörleri yenilenemedi. İşlem durumunu kontrol edin.':'Posta klasörleri yenilendi.',!!failed);
});
$('disconnect').addEventListener('click',async()=>{const account=accounts.get($('activeAccount').value);if(!account)return;try{if(account.source!=='local')await auth.disconnect(account);accounts.delete(account.id);trees.delete(account.id);for(const key of [...selected])if(JSON.parse(key)[0]===account.id)selected.delete(key);renderAccounts();renderContacts();notice('Hesap bu sekmeden ayrıldı. Oluşturulmuş rehber korunuyor.');}catch(error){notice('Hesap bağlantısı kaldırılamadı. Sayfayı yenileyerek bu sekmenin oturumunu kapatabilirsiniz.',true);}});
$('editAliases').addEventListener('click',()=>{const account=accounts.get($('activeAccount').value);if(!account)return;$('aliasesAccount').textContent=account.email;$('aliasesText').value=[...new Set([...account.own,...(aliases.get(account.id)||[])])].join('\n');$('aliasesDialog').showModal();});
$('saveAliases').addEventListener('click',()=>{const account=accounts.get($('activeAccount').value);if(!account)return;const raw=$('aliasesText').value.split(/[\n,;]+/).map(v=>v.trim()).filter(Boolean);if(raw.some(v=>!normalizeEmail(v))){notice('Kendi adresleri listesinde geçersiz e-posta var.',true);return;}aliases.set(account.id,[...new Set(raw.map(normalizeEmail))]);removeOwn();$('aliasesDialog').close();renderContacts();notice('Kendi adresleriniz rehberden çıkarıldı.');});
$('selectFolders').addEventListener('click',()=>{for(const [id,folders] of trees)for(const folder of folders)selected.add(JSON.stringify([id,folder.id]));renderAccounts();});
$('clearFolders').addEventListener('click',()=>{selected.clear();renderAccounts();});
$('scan').addEventListener('click',async()=>{
 const work=[];for(const key of selected){const [id,folderId]=JSON.parse(key);const account=accounts.get(id),folder=trees.get(id)?.find(f=>f.id===folderId);if(account&&folder)work.push({account,folder});}
 if(!work.length){notice('En az bir posta klasörü seçin.',true);return;}
 if($('fullScan').checked){directory=new Map();checkpoints=new Map();selectedEmail='';}
 busy=true;abort=new AbortController();$('allowPartial').checked=false;$('progress').value=0;$('messageCount').textContent='0';$('scanStatus').textContent='Taranıyor';renderAccounts();renderContacts();log('Tarama başladı. '+work.length+' klasör seçildi.');
 try{
  if(local)local.signatures=$('readSignatures').checked;
  const scan=local?scanLocalFolders:scanFolders;
  lastResult=await scan({client:local,graph:local||graph(),work,directory,checkpoints,own:own(),signal:abort.signal,full:$('fullScan').checked,onProgress:progress=>{$('messageCount').textContent=progress.processed;$('progress').value=progress.total?progress.completed/progress.total*100:0;$('progressText').textContent=`${progress.account} · ${progress.folder} · ${progress.processed} mesaj · ${progress.completed}/${progress.total} klasör`;renderContacts();}});
  for(const warning of lastResult.warnings)log(warning.account+' › '+warning.folder+' · '+warning.message,true);
  $('scanStatus').textContent=lastResult.complete?'Tamamlandı':lastResult.cancelled?'İptal edildi':'Eksik tarama';$('progressText').textContent=`${lastResult.completed}/${lastResult.total} klasör tamamlandı · ${rawContacts().length} tekil e-posta`;
  notice(lastResult.complete?'Tarama tamamlandı. Firma rehberini Excel olarak kaydedebilirsiniz.':'Tarama tam değil. Tamamlanan klasörler korundu; sonuçları kontrol ederek eksik çıktı alabilir veya yeniden tarayabilirsiniz.',!lastResult.complete);log($('scanStatus').textContent+'. '+rawContacts().length+' tekil e-posta.');
 }catch(error){lastResult={complete:false,warnings:[{message:'Beklenmeyen tarama hatası. Tamamlanan sonuçlar korundu.'}]};notice('Tarama tamamlanamadı. Tamamlanan klasörler korundu.',true);$('scanStatus').textContent='Eksik tarama';}
 finally{busy=false;abort=null;renderAccounts();renderContacts();}
});
$('cancel').addEventListener('click',()=>{abort?.abort();$('scanStatus').textContent='Durduruluyor…';});
function saveBlob(blob,name){const url=URL.createObjectURL(blob),link=document.createElement('a');link.href=url;link.download=name;document.body.append(link);link.click();link.remove();setTimeout(()=>URL.revokeObjectURL(url),60000);}
$('exportExcel').addEventListener('click',async()=>{
 if(!lastResult?.complete&&!$('allowPartial').checked)return;const raw=rawContacts();if(!raw.length)return;
 let handle;
 try{if(window.showSaveFilePicker)handle=await window.showSaveFilePicker({suggestedName:'Outlook_Firma_Rehberi.xlsx',types:[{description:'Excel çalışma kitabı',accept:{'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':['.xlsx']}}]});}
 catch(error){if(error.name!=='AbortError')notice('Excel kayıt konumu seçilemedi.',true);return;}
 busy=true;controls();
 try{const {buildWorkbook}=await import('./export.mjs');const buffer=await buildWorkbook(raw,lastResult||{complete:false});const blob=new Blob([buffer],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  if(handle){const writer=await handle.createWritable();try{await writer.write(blob);await writer.close();}catch(error){try{await writer.abort();}catch{}throw error;}notice('Excel rehberi kaydedildi.');}else{saveBlob(blob,'Outlook_Firma_Rehberi.xlsx');notice('Excel indirmesi başlatıldı.');}
 }catch(error){notice('Excel kaydedilemedi. Kayıt klasörünü ve boş disk alanını kontrol edin.',true);}finally{busy=false;controls();}
});
$('saveBackup').addEventListener('click',()=>{const json=serializeWorkspace({directory,checkpoints,aliases,result:lastResult});saveBlob(new Blob([json],{type:'application/json'}),'Outlook_Rehberi_Calisma_Yedegi.json');notice('Çalışma yedeğinin indirmesi başlatıldı. Hesap şifresi veya giriş belirteci içermez.');});
$('importBackup').addEventListener('click',()=>$('backupFile').click());
$('backupFile').addEventListener('change',async event=>{
 const file=event.target.files[0];if(!file)return;
 try{if(file.size>30*1024*1024)throw new Error('Yedek dosyası en fazla 30 MB olabilir.');const restored=restoreWorkspace(JSON.parse(await file.text()));directory=restored.directory;checkpoints=restored.checkpoints;aliases=restored.aliases;lastResult=restored.result;selectedEmail='';page=0;removeOwn();$('allowPartial').checked=false;renderContacts();notice(directory.size+' kayıt içe aktarıldı.'+(restored.skipped?' '+restored.skipped+' geçersiz kayıt atlandı.':'')+' Posta taramasına devam etmek için hesaplarınızı bağlayın.');log('Rehber yedeği içe aktarıldı.');}catch(error){notice(error instanceof SyntaxError?'Geçerli bir JSON rehber yedeği seçin.':error.message,true);}event.target.value='';
});
function edit(email){if(busy)return;const raw=directory.get(email);if(!raw)return;editing=email;const c=contactValues(raw);$('editingEmail').textContent=email;$('contactName').value=c.name;$('contactTitle').value=c.title;$('contactCompany').value=c.company;$('editDialog').showModal();}
$('editContact').addEventListener('click',()=>edit(selectedEmail));
$('saveContact').addEventListener('click',()=>{const contact=directory.get(editing);if(!contact)return;contact.overrides={name:$('contactName').value.trim().slice(0,500),title:$('contactTitle').value.trim().slice(0,500),company:$('contactCompany').value.trim().slice(0,500)};$('editDialog').close();renderContacts();});
$('clearDirectory').addEventListener('click',()=>{$('confirmText').textContent=directory.size+' kişi kaydı ve bu sekmenin tarama ilerlemesi temizlenecek. Bilgisayarınıza kaydettiğiniz Excel ve yedek dosyaları etkilenmez.';$('confirmDialog').showModal();});
$('confirmClear').addEventListener('click',()=>{directory.clear();checkpoints.clear();lastResult=null;selectedEmail='';$('confirmDialog').close();renderContacts();notice('Bu oturumdaki rehber temizlendi.');});
$('allowPartial').addEventListener('change',controls);
for(const id of ['search','category'])$(id).addEventListener(id==='search'?'input':'change',()=>{page=0;renderContacts();});
$('previous').addEventListener('click',()=>{page--;renderContacts();});$('next').addEventListener('click',()=>{page++;renderContacts();});
for(const button of document.querySelectorAll('[data-close]'))button.addEventListener('click',()=>$(button.dataset.close).close());
window.addEventListener('beforeunload',event=>{if(directory.size||busy){event.preventDefault();event.returnValue='';}});
renderAccounts();renderContacts();

async function connectLocal(){
 if(!local){notice('Mac Mail Rehberi yardımcısındaki Rehberi Aç düğmesiyle bu ekranı yeniden açın.',true);return;}
 busy=true;controls();notice('Mac Mail hesapları okunuyor. macOS izin sorarsa Mail erişimine izin verin.');
 try{
  const items=await local.accounts();
  for(const item of items){accounts.set(item.id,item);}
  removeOwn();
  for(const item of items){try{await loadFolders(item);}catch(error){log(item.email+' · '+error.message,true);}}
  notice(items.length?items.length+' Mail hesabı bulundu. Klasörleri seçip taramayı başlatın.':'Mail’de hesap bulunamadı. Önce Mail uygulamasına hesabınızı ekleyin.');
 }catch(error){notice(error.message,true);}
 finally{busy=false;renderAccounts();renderContacts();}
}
