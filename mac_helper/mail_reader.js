// Read-only Mail scripting adapter. Dictionary: /System/Applications/Mail.app/.../Mail.sdef
ObjC.import('Foundation');
function run() {
    const data = $.NSFileHandle.fileHandleWithStandardInput.readDataToEndOfFile;
    const input = JSON.parse(ObjC.unwrap($.NSString.alloc.initWithDataEncoding(data, $.NSUTF8StringEncoding)));
    const mail = Application('com.apple.mail');
    const accounts = mail.accounts();
    const localBoxes = mail.mailboxes().filter(b => {try {const owner=b.account(); return !owner || !owner.id();} catch (_) {return true;}});
    if (input.command === 'accounts') {
        const list=accounts.map(a => ({id:'mail:'+a.id(),name:a.name(),addresses:a.emailAddresses()}));
        if(localBoxes.length) list.push({id:'mail:local',name:'Mac’imdeki Postalar',addresses:[]});
        return JSON.stringify({accounts:list});
    }
    const account = input.account==='mail:local'?{mailboxes:()=>localBoxes}:accounts.find(a => 'mail:'+a.id() === input.account);
    if (!account) throw Error('Account unavailable');
    if (input.command === 'folders') {
        const result = [];
        function walk(boxes, path) {
            for (const box of boxes) {
                const parts = path.concat([box.name()]);
                result.push({id:JSON.stringify(parts),parts:parts,path:parts.join(' › '),totalItemCount:box.messages.length});
                walk(box.mailboxes(), parts);
            }
        }
        walk(account.mailboxes(), []);
        return JSON.stringify({folders:result});
    }
    let parent = account;
    for (const name of input.path) {
        const matches = parent.mailboxes().filter(b => b.name() === name);
        if (matches.length !== 1) throw Error('Mailbox unavailable or ambiguous');
        parent = matches[0];
    }
    if (input.command === 'snapshot') return JSON.stringify({ids:parent.messages.id()});
    if (input.command !== 'messages') throw Error('Unsupported operation');
    function recipient(r) { return {emailAddress:{address:r.address(),name:r.name()||''}}; }
    function sender(value) {
        const match = String(value).match(/^(.*?)\s*<([^>]+)>\s*$/);
        return {emailAddress:match?{name:match[1].replace(/^"|"$/g,''),address:match[2]}:{name:'',address:String(value)}};
    }
    const messages = input.ids.map(id => {
        const m = parent.messages.byId(id);
        // Missing/deleted messages fail the batch. Never report a partial folder as complete.
        const item = {id:String(m.id()),from:sender(m.sender()),
            toRecipients:m.toRecipients().map(recipient),ccRecipients:m.ccRecipients().map(recipient),bccRecipients:m.bccRecipients().map(recipient),
            receivedDateTime:m.dateReceived().toISOString(),sentDateTime:m.dateSent().toISOString(),
            body:{contentType:'text',content:input.signatures?String(m.content()).slice(0,20000):''}};
        return item;
    });
    return JSON.stringify({messages:messages});
}
