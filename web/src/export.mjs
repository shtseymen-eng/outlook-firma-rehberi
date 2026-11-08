import ExcelJS from 'exceljs';
import {contactValues,CATEGORY_LABELS} from './contacts.mjs';
const COLORS={polisan_poliport:'175B85',customs_ygm:'76589D',corporate:'167B79',personal:'B57732'};
const HEADERS=['Ad Soyad','Görev / Unvan','Domain','Firma','E-posta','Cep Telefonu','Sabit Telefon','Kaynak Hesap(lar)','Kaynak Klasör(ler)'];
const WIDTHS=[27,30,25,35,38,25,25,38,52];
function paint(cell,bg='FFFFFF',bold=false,color='26384A'){
 cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:bg}};cell.font={name:'Calibri',size:11,bold,color:{argb:color}};cell.alignment={vertical:'middle',wrapText:true};cell.border={bottom:{style:'thin',color:{argb:'DFE7EF'}}};
}
function title(sheet,text,accent){sheet.views=[{showGridLines:false,state:'frozen',ySplit:3}];sheet.mergeCells('A1:I1');sheet.getCell('A1').value=text;paint(sheet.getCell('A1'),accent,true,'FFFFFF');sheet.getCell('A1').alignment={horizontal:'center',vertical:'middle'};sheet.getRow(1).height=34;}
export async function buildWorkbook(rawContacts,result={}){
 const workbook=new ExcelJS.Workbook();workbook.creator='Outlook Firma Rehberi';workbook.created=new Date();
 const unique=new Map(rawContacts.map(c=>[c.email,contactValues(c)]));const contacts=[...unique.values()];
 const summary=workbook.addWorksheet('Özet');title(summary,'OUTLOOK FİRMA REHBERİ — ÖZET','153F62');
 summary.mergeCells('A2:I2');summary.getCell('A2').value=result.complete?'TARAMA TAMAMLANDI':'EKSİK TARAMA — Yalnızca tamamlanan klasörlerin sonuçlarıdır.';paint(summary.getCell('A2'),result.complete?'E3F3E9':'FFF0D6',true);
 const info=[['Oluşturma zamanı',new Date().toLocaleString('tr-TR')],['Tekil e-posta adresi',contacts.length],['Hesap sayısı',result.accounts||0],['Tamamlanan klasör',result.completed||0],['Seçili klasör',result.total||0],['İşlenen mesaj',result.processed||0],...Object.entries(CATEGORY_LABELS).map(([key,label])=>[label,contacts.filter(c=>c.category===key).length])];
 let row=4;for(const [label,value] of info){summary.getCell(row,1).value=label;summary.getCell(row,2).value=value;paint(summary.getCell(row,1),'F1F5F9',true);paint(summary.getCell(row,2));row++;}
 if(result.warnings?.length){row++;summary.getCell(row++,1).value='Eksik / Hata Alan Klasörler';for(const warning of result.warnings){summary.getCell(row,1).value=String(warning.account||'');summary.getCell(row,2).value=String(warning.folder||'');summary.mergeCells(row,3,row,9);summary.getCell(row,3).value=String(warning.message||'');row++;}}
 summary.getColumn(1).width=34;summary.getColumn(2).width=45;
 for(const [category,label] of Object.entries(CATEGORY_LABELS)){
  const sheet=workbook.addWorksheet(label),accent=COLORS[category];title(sheet,'OUTLOOK FİRMA REHBERİ — '+label.toLocaleUpperCase('tr-TR'),accent);
  const rows=contacts.filter(c=>c.category===category).sort((a,b)=>a.domain.localeCompare(b.domain)||a.company.localeCompare(b.company)||a.email.localeCompare(b.email));
  sheet.mergeCells('A2:I2');sheet.getCell('A2').value='Toplam tekil e-posta: '+rows.length;paint(sheet.getCell('A2'),'EAF0F6');
  HEADERS.forEach((header,i)=>{const cell=sheet.getCell(3,i+1);cell.value=header;paint(cell,accent,true,'FFFFFF');sheet.getColumn(i+1).width=WIDTHS[i];});sheet.getRow(3).height=27;
  let current=4;let previousDomain='',groupIndex=-1;
  for(const contact of rows){
    if(contact.domain!==previousDomain){groupIndex++;previousDomain=contact.domain;}
    const values=[contact.name,contact.title,contact.domain,contact.company,contact.email,(contact.mobiles||[]).join(' / '),(contact.landlines||[]).join(' / '),[...new Set(contact.sources.map(s=>s.account))].sort().join(' / '),[...new Set(contact.sources.map(s=>s.folder))].sort().join(' / ')];
    values.forEach((value,i)=>{const cell=sheet.getCell(current,i+1);cell.value=String(value??'');paint(cell,groupIndex%2?'FFFFFF':'F0F6FC');});sheet.getRow(current).height=32;current++;
  }
  // Merge only genuinely equal adjacent values; edited company names remain distinct.
  for(const column of [3,4]){
    let start=4;
    for(let end=5;end<=current;end++){
      const equal=end<current&&sheet.getCell(end,column).value===sheet.getCell(start,column).value&&(column===3||sheet.getCell(end,3).value===sheet.getCell(start,3).value);
      if(!equal){if(end-start>1)sheet.mergeCells(start,column,end-1,column);start=end;}
    }
  }
  sheet.autoFilter={from:'A3',to:'I'+Math.max(3,current-1)};sheet.pageSetup={orientation:'landscape',fitToPage:true,fitToWidth:1,fitToHeight:0};
 }
 return workbook.xlsx.writeBuffer();
}
