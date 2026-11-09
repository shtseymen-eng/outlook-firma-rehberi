import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import {readFileSync} from 'node:fs';
const script=readFileSync(new URL('../../mac_helper/mail_reader.js',import.meta.url),'utf8');
const person={address:()=> 'other@example.org',name:()=> 'Other Person'};
const message={id:()=>7,sender:()=> 'Other Person <other@example.org>',toRecipients:()=>[person],ccRecipients:()=>[],bccRecipients:()=>[],dateReceived:()=>new Date('2026-11-08T10:00:00Z'),dateSent:()=>new Date('2026-11-08T10:00:00Z'),content:()=> 'Signature Company'};
const messages=()=>[message];messages.id=()=>[7];messages.byId=id=>{assert.equal(id,7);return message;};
const box={name:()=> 'Inbox',messages,mailboxes:()=>[]};
const account={id:()=> 'abc',name:()=> 'Example',emailAddresses:()=>['self@example.com'],mailboxes:()=>[box]};
function read(input){const context=vm.createContext({Application:id=>{assert.equal(id,'com.apple.mail');return {accounts:()=>[account],mailboxes:()=>[]};},ObjC:{import:()=>{},unwrap:x=>x},$:{NSFileHandle:{fileHandleWithStandardInput:{readDataToEndOfFile:JSON.stringify(input)}},NSString:{alloc:{initWithDataEncoding:x=>x}},NSUTF8StringEncoding:4}});vm.runInContext(script,context);return JSON.parse(context.run());}
test('read-only Mail adapter maps accounts, nested folder path and message participants',()=>{
 assert.equal(read({command:'accounts'}).accounts[0].id,'mail:abc');
 const folders=read({command:'folders',account:'mail:abc'}).folders;
 assert.deepEqual(folders[0].parts,['Inbox']);
 assert.deepEqual(read({command:'snapshot',account:'mail:abc',path:['Inbox']}).ids,[7]);
 const payload={command:'messages',account:'mail:abc',path:['Inbox'],ids:[7]};
 assert.equal(read(payload).messages[0].from.emailAddress.address,'other@example.org');
 assert.equal(read(payload).messages[0].body.content,'');
 assert.equal(read({...payload,signatures:true}).messages[0].body.content,'Signature Company');
 assert.throws(()=>read({...payload,path:['Missing']}),/Mailbox unavailable/);
});
