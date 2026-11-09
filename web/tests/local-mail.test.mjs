import test from 'node:test';
import assert from 'node:assert/strict';
import {LocalMailClient,localSession,scanLocalFolders} from '../src/local-mail.mjs';
const token='a'.repeat(64);
const account={id:'mail:a',email:'self@example.com'};
const folder={id:'["Inbox"]',parts:['Inbox'],path:'Inbox'};
test('pairing secrets only accepted on loopback and never from hosted URL',()=>{
 assert.equal(localSession('https://shtseymen-eng.github.io/#mail='+token),null);
 assert.equal(localSession('http://127.0.0.1:9876/#mail=bad'),null);
 assert.deepEqual(localSession('http://127.0.0.1:9876/#mail='+token),{origin:'http://127.0.0.1:9876',token});
});
test('local scan batches a stable snapshot and commits only fully read folders',async()=>{
 let calls=[];
 const client=new LocalMailClient({origin:'http://127.0.0.1:9876',token},{fetch:async(url,options)=>{
  calls.push({url,options});const request=JSON.parse(options.body);
  const value=url.endsWith('/snapshot')?{ids:[1,2]}:{messages:request.ids.map(id=>({id:String(id),from:{emailAddress:{address:'person@example.org',name:'Test'}}}))};
  return new Response(JSON.stringify(value));
 }});
 const directory=new Map();let result=await scanLocalFolders({client,work:[{account,folder}],directory,own:new Set(),checkpoints:new Map()});
 assert.equal(result.complete,true);assert.equal(result.processed,2);assert.equal(directory.size,1);
 assert.equal(calls[0].options.headers.Authorization,'Bearer '+token);
 assert.equal(calls[0].options.credentials,'omit');
 client.fetch=async url=>new Response(JSON.stringify(url.endsWith('/snapshot')?{ids:[1,2]}:{messages:[{id:'1'}]}));
 const fresh=new Map();result=await scanLocalFolders({client,work:[{account,folder}],directory:fresh,own:new Set(),checkpoints:new Map()});
 assert.equal(result.complete,false);assert.equal(fresh.size,0);
});
test('empty local folders complete and do not create Graph checkpoints',async()=>{
 const checkpoints=new Map(),client=new LocalMailClient({origin:'http://127.0.0.1:9876',token},{fetch:async()=>new Response('{"ids":[]}')});
 const result=await scanLocalFolders({client,work:[{account,folder}],directory:new Map(),own:new Set(),checkpoints});
 assert.equal(result.complete,true);assert.equal(checkpoints.size,0);
});
test('browser fetch is called with its Window receiver',async()=>{
 const original=globalThis.fetch;
 try{
  globalThis.fetch=async function(){assert.equal(this,globalThis);return new Response('{"accounts":[]}');};
  const client=new LocalMailClient({origin:'http://127.0.0.1:9876',token});
  assert.deepEqual(await client.accounts(),[]);
 }finally{globalThis.fetch=original;}
});
