import test from 'node:test';
import assert from 'node:assert/strict';
import {GraphClient,scanFolders} from '../src/graph.mjs';
const account={id:'a',email:'self@example.com',own:['self@example.com']};
const folder={id:'f',path:'Gelen Kutusu'};
const message=id=>({id,from:{emailAddress:{address:id+'@acme.example',name:'Örnek Kişi'}},receivedDateTime:'2026-11-08T10:00:00Z',body:{contentType:'text',content:''}});
test('rejects untrusted pagination URLs before sending an access token',async()=>{
 let tokens=0,calls=0;
 const graph=new GraphClient(async()=>{tokens++;return 'test';},{fetch:async()=>{calls++;return new Response('{}');}});
 await assert.rejects(()=>graph.request(account,'https://attacker.example/steal'),/Geçersiz Microsoft/);
 assert.equal(tokens,0);assert.equal(calls,0);
});
test('recursively loads all folder pages and children',async()=>{
 const payloads=[{value:[{id:'f',displayName:'Gelen',childFolderCount:1}], '@odata.nextLink':'https://graph.microsoft.com/v1.0/me/mailFolders?$skiptoken=x'},{value:[{id:'s',displayName:'Gönderilen',childFolderCount:0}]},{value:[{id:'c',displayName:'Alt',childFolderCount:0}]}];
 const graph=new GraphClient(async()=>'test',{fetch:async()=>new Response(JSON.stringify(payloads.shift()))});
 const result=await graph.folders(account);
 assert.deepEqual(result.map(f=>f.path),['Gelen','Gönderilen','Gelen › Alt']);
});
test('commits contacts and delta checkpoint only after the entire folder succeeds',async()=>{
 const directory=new Map(),checkpoints=new Map();
 const graph={async *messages(){yield {value:[message('first')],next:'next'};throw new Error('network');}};
 const result=await scanFolders({graph,work:[{account,folder}],directory,checkpoints,own:new Set(account.own)});
 assert.equal(result.complete,false);assert.equal(directory.size,0);assert.equal(checkpoints.size,0);
});
test('successful folder creates unique contacts and persists checkpoint for next scan',async()=>{
 const directory=new Map(),checkpoints=new Map();let start;
 const graph={async *messages(a,f,checkpoint){start=checkpoint;yield {value:[message('first'),message('first')],delta:'https://graph.microsoft.com/v1.0/me/mailFolders/f/messages/delta?$deltatoken=x'};}};
 const config={graph,work:[{account,folder}],directory,checkpoints,own:new Set(account.own)};
 const result=await scanFolders(config);
 assert.equal(result.complete,true);assert.equal(directory.size,1);assert.equal(checkpoints.size,1);
 await scanFolders(config);assert.ok(start.includes('$deltatoken'));assert.equal(directory.size,1);
});
test('cancellation keeps completed folders but does not commit the interrupted folder',async()=>{
 const abort=new AbortController(),directory=new Map(),checkpoints=new Map();
 const graph={async *messages(a,f){yield {value:[message(f.id)]};if(f.id==='second')abort.abort();yield {value:[],delta:'https://graph.microsoft.com/v1.0/me/mailFolders/'+f.id+'/messages/delta?$deltatoken=x'};}};
 const result=await scanFolders({graph,work:[{account,folder},{account,folder:{id:'second',path:'İkinci'}}],directory,checkpoints,own:new Set(account.own),signal:abort.signal});
 assert.equal(result.complete,false);assert.equal(directory.size,1);assert.equal(checkpoints.size,1);
});
test('throttling honors Retry-After and retries a read-only request',async()=>{
 let count=0,waited=0;
 const graph=new GraphClient(async()=>'test',{fetch:async()=>++count===1?new Response('{}',{status:429,headers:{'Retry-After':'2'}}):new Response('{"value":[]}'),sleep:async ms=>{waited=ms;}});
 await graph.request(account,'/me/mailFolders');assert.equal(count,2);assert.equal(waited,2000);
});
