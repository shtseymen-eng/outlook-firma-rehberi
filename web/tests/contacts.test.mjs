import test from 'node:test';
import assert from 'node:assert/strict';
import {normalizeEmail,classify,messageContacts,mergeContact,contactValues,importLegacy} from '../src/contacts.mjs';
const source={account:'self@example.com',accountId:'a',folderId:'inbox',folder:'Gelen Kutusu'};
const address=(email,name='')=>({emailAddress:{address:email,name}});
const message={id:'m1',from:address('person@acme.example','Ayşe Test'),toRecipients:[address('self@example.com'),address('other@acme.example','Ali Test')],ccRecipients:[address('PERSON@ACME.EXAMPLE')],bccRecipients:[address('third@other.example')],receivedDateTime:'2026-11-08T10:00:00Z',body:{contentType:'text',content:'Merhaba\nSaygılarımla\nAyşe Test\nOperations Manager\nACME LTD.\n+90 532 123 45 67\n0212 123 45 67\nFrom: Old Sender\nWrong Person\nWrong Company LTD.'}};
test('normalizes and validates email addresses',()=>{
 assert.equal(normalizeEmail(' Person@Acme.Example. '),'person@acme.example');
 assert.equal(normalizeEmail('not an email'),null);
 assert.equal(normalizeEmail('a@bad..example'),null);
});
test('deduplicates all participant fields and excludes own addresses',()=>{
 const result=messageContacts(message,source,new Set(['self@example.com']));
 assert.deepEqual(result.map(r=>r.email).sort(),['other@acme.example','person@acme.example','third@other.example']);
});
test('enriches only sender from new signature, never quoted message or another recipient',()=>{
 const rows=messageContacts(message,source,new Set());const sender=contactValues(rows.find(r=>r.email==='person@acme.example'));
 assert.equal(sender.name,'Ayşe Test');assert.equal(sender.title,'Operations Manager');assert.equal(sender.company,'ACME LTD.');
 assert.deepEqual(sender.mobiles,['905321234567']);assert.deepEqual(sender.landlines,['02121234567']);
 assert.equal(contactValues(rows.find(r=>r.email==='other@acme.example')).title,'');
 assert.ok(!JSON.stringify(rows).includes('Wrong Company'));
});
test('newer evidence of equal quality wins; names from metadata cannot overwrite signature names',()=>{
 const first=messageContacts(message,source,new Set())[0];
 const newer=messageContacts({...message,receivedDateTime:'2026-11-09T10:00:00Z',body:{contentType:'text',content:'Saygılarımla\nAyşe Yeni\nSales Director\nACME LTD.'}},source,new Set())[0];
 assert.equal(contactValues(mergeContact(first,newer)).name,'Ayşe Yeni');
 const metadata=messageContacts({...message,receivedDateTime:'2026-11-10T10:00:00Z',from:address('person@acme.example','Yanlış Ad'),body:{contentType:'text',content:''}},source,new Set())[0];
 assert.equal(contactValues(mergeContact(first,metadata)).name,'Ayşe Test');
});
test('classifies company, customs and personal addresses',()=>{
 assert.equal(classify({email:'x@poliport.com',domain:'poliport.com'}),'polisan_poliport');
 assert.equal(classify({email:'x@gmail.com',domain:'gmail.com',company:'ACME LTD.'}),'personal');
 assert.equal(classify({email:'x@acme.example',domain:'acme.example',company:'Örnek Gümrük Müşavirliği'}),'customs_ygm');
});
test('legacy v4 import preserves contacts without fabricating missing details',()=>{
 const result=importLegacy({contacts:{'x@acme.example':{name:'Örnek Kişi',firma_display:'Örnek Firma',gsm:['05321234567']},invalid:{name:'Invalid'}}});
 assert.equal(result.contacts.length,1);assert.equal(result.skipped,1);assert.equal(contactValues(result.contacts[0]).company,'Örnek Firma');
});
