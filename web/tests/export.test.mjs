import test from 'node:test';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import {buildWorkbook} from '../src/export.mjs';
import {serializeWorkspace,restoreWorkspace} from '../src/backup.mjs';
import {importLegacy} from '../src/contacts.mjs';
const contacts=importLegacy({contacts:{'a@acme.example':{name:'=HYPERLINK("x")',firma:'ACME LTD.'},'b@acme.example':{name:'Örnek Kişi',firma:'ACME LTD.'}}}).contacts;
test('exports five sheets, merged company groups and literal text cells',async()=>{
 const buffer=await buildWorkbook(contacts,{complete:true,processed:20,total:1,completed:1,accounts:1,completedAt:'2026-11-08T10:00:00Z'});
 const workbook=new ExcelJS.Workbook();await workbook.xlsx.load(buffer);
 assert.deepEqual(workbook.worksheets.map(s=>s.name),['Özet','Polisan-Poliport','Gümrük-YGM','Diğer Kurumsal','Kişisel Mailler']);
 const sheet=workbook.getWorksheet('Diğer Kurumsal');
 assert.equal(sheet.getCell('A4').value,'=HYPERLINK("x")');assert.equal(sheet.getCell('A4').type,ExcelJS.ValueType.String);
 assert.equal(sheet.getCell('D5').isMerged,true);assert.equal(sheet.getCell('C5').isMerged,true);
});
test('incomplete exports are clearly marked in the summary',async()=>{
 const buffer=await buildWorkbook(contacts,{complete:false,processed:3,total:2,completed:1,accounts:1,warnings:[{message:'İptal'}]});
 const workbook=new ExcelJS.Workbook();await workbook.xlsx.load(buffer);
 assert.match(JSON.stringify(workbook.getWorksheet('Özet').model),/EKSİK TARAMA/);
});
test('backup contains directory and checkpoints but no authentication data',()=>{
 const key=JSON.stringify(['account','folder']);
 const json=serializeWorkspace({directory:new Map(contacts.map(c=>[c.email,c])),checkpoints:new Map([[key,'https://graph.microsoft.com/v1.0/me/mailFolders/f/messages/delta?$deltatoken=x']]),aliases:new Map([['account',['self@example.com']]]),result:{complete:true},accessToken:'MUST_NOT_EXPORT'});
 assert.ok(!json.includes('MUST_NOT_EXPORT'));const restored=restoreWorkspace(JSON.parse(json));
 assert.equal(restored.directory.size,2);assert.equal(restored.checkpoints.size,1);
});
test('backup rejects non-Microsoft pagination links and invalid email records',()=>{
 assert.throws(()=>restoreWorkspace({kind:'outlook-firma-rehberi-web',version:1,contacts,checkpoints:[['x','https://attacker.example/steal']],aliases:[]}),/Geçersiz Microsoft/);
});
test('per-contact company edits are preserved even for the same email domain',async()=>{
 const changed=structuredClone(contacts);changed[0].overrides={company:'Firma A'};changed[1].overrides={company:'Firma B'};
 const buffer=await buildWorkbook(changed,{complete:true});const workbook=new ExcelJS.Workbook();await workbook.xlsx.load(buffer);
 const sheet=workbook.getWorksheet('Diğer Kurumsal');const values=new Map();for(const row of [4,5])values.set(sheet.getCell(row,5).value,sheet.getCell(row,4).value);
 assert.equal(values.get('a@acme.example'),'Firma A');assert.equal(values.get('b@acme.example'),'Firma B');
});
