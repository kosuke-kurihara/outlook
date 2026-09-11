// Dev only: npm install --no-save playwright; node tests/test.cjs
const fs = require('node:fs');
const path = require('node:path');
const assert = require('node:assert/strict');
const { pathToFileURL } = require('node:url');
const { chromium } = require(process.env.ERH_PLAYWRIGHT || 'playwright');
const root = path.resolve(__dirname,'..');
(async()=>{
const browser = await chromium.launch({channel:'msedge',headless:true});
try {
const page = await browser.newPage();
await page.addScriptTag({path:path.join(root,'docs/runtime.js')});
const results = await page.evaluate(async()=>{
  const api=EnglishReplyHeaders, results=[];
  const check=(name,ok)=>{if(!ok)throw Error(name);results.push(name);};
  const jp='差出人: Alice <alice@example.com>\n送信: 2026 年 9 月 11 日 (金曜日) 14:05\n宛先: Bob\nCC: Carol\n件名: test';
  const html=(id='divRplyFwdMsg')=>'<div id="'+id+'"><b>差出人:</b> Alice &lt;alice@example.com&gt;<br><b>送信:</b> <span>2026年9月11日</span> （金曜日） 14:05<br><b>宛先:</b> Bob<br><b>CC:</b> Carol<br><b>件名:</b> test</div>';
  check('Japanese labels',api.labelKey(' 差出人： ')==='from');
  check('Date format',api.convertJapaneseDate('2026年9月11日 14:05')==='11 Sep 2026 14:05');
  check('Seconds',api.convertJapaneseDate('2026 年 9 月 11 日 （金曜日） 4:05:09')==='11 Sep 2026 04:05:09');
  check('Invalid date',api.convertJapaneseDate('2026年2月30日 14:05')==='2026年2月30日 14:05');
  check('Invalid time',api.convertJapaneseDate('2026年9月11日 25:05')==='2026年9月11日 25:05');
  let plain=api.replacePlainTextHeaders(jp);
  check('Plain conversion',plain.changed&&plain.text.includes('Sent: 11 Sep 2026 14:05'));
  check('Plain idempotence with older Japanese chain',!api.replacePlainTextHeaders(plain.text+'\n\n'+jp).changed);
  check('Plain older header preserved',api.replacePlainTextHeaders(jp+'\n\n'+jp).text.endsWith(jp));
  check('Casual subject untouched',!api.replacePlainTextHeaders('件名: 相談です').changed);
  check('Unrelated lines reject block',!api.replacePlainTextHeaders(jp.replace('送信:','unrelated prose\n送信:')).changed);
  check('CRLF preserved',api.replacePlainTextHeaders(jp.replaceAll('\n','\r\n')).text.includes('\r\n'));
  let converted=api.replaceHtmlHeaders(html());
  check('HTML conversion including span date',converted.changed&&converted.text.includes('11 Sep 2026 14:05'));
  check('HTML idempotence',!api.replaceHtmlHeaders(converted.text).changed);
  check('HTML English newest stops older conversion',!api.replaceHtmlHeaders(converted.text+html('divRplyFwdMsg_1')).changed);
  const prefix='<p id="signature">署名 <b>件名:</b> unchanged</p>';
  check('Signature and body unchanged',api.replaceHtmlHeaders(prefix+html()).text.startsWith(prefix));
  const ordered=api.replaceHtmlHeaders(html('x_divRplyFwdMsg')+html());
  check('Prefixed newest takes priority',ordered.text.endsWith(html()));
  const nested=html().replace('</div>',html('divRplyFwdMsg_1')+'</div>');
  check('Nested old header unchanged',api.replaceHtmlHeaders(nested).text.includes(html('divRplyFwdMsg_1')));
  check('No marker no conversion',!api.replaceHtmlHeaders('<p><b>件名:</b> test</p>').changed);
  check('Incomplete marker no conversion',!api.replaceHtmlHeaders('<div id="divRplyFwdMsg"><b>件名:</b> test</div>').changed);
  check('Separator not duplicated',(api.replaceHtmlHeaders('<div>----- Original Message -----</div>'+html()).text.match(/Original Message/g)||[]).length===1);
  api.config.addOriginalMessageSeparator=false;
  check('Separator configurable',!api.replacePlainTextHeaders(jp).text.includes('Original Message'));
  let writes=0,reads=0,changed=false;
  window.Office={AsyncResultStatus:{Succeeded:'ok'},CoercionType:{Html:'html'},MailboxEnums:{BodyMode:{HostConfig:'hostConfig'}},context:{mailbox:{item:{
    getComposeTypeAsync:cb=>cb({status:'ok',value:{composeType:'reply',coercionType:'text'}}),
    body:{getAsync:(type,options,cb)=>{check('HostConfig read '+reads,options.bodyMode==='hostConfig');reads++;cb({status:'ok',value:jp+(changed&&reads>1?' edited':'')});},setAsync:(value,options,cb)=>{writes++;check('HostConfig write',options.bodyMode==='hostConfig');cb({status:'ok'});}}
  }}}};
  await api.transformCurrentItem();check('Compose writes once',writes===1);
  changed=true;reads=0;await api.transformCurrentItem();check('Concurrent edit aborts write',writes===1);
  Office.context.mailbox.item.getComposeTypeAsync=cb=>cb({status:'ok',value:{composeType:'newMail'}});
  reads=0;await api.transformCurrentItem();check('New message skipped',reads===0);
  const notes=[];let completed=0;
  const item=Office.context.mailbox.item;
  item.notificationMessages={replaceAsync:(key,note,cb)=>{notes.push(note.message);cb({status:'ok'});}};
  item.getComposeTypeAsync=cb=>cb({status:'ok',value:{composeType:'reply',coercionType:'text'}});
  item.body.getTypeAsync=cb=>cb({status:'ok',value:'text'});
  Office.MailboxEnums.BodyMode.FullBody='fullBody';
  api.config.retryDelaysMs=[0];
  item.body.getAsync=(type,options,cb)=>cb({status:'ok',value:options.bodyMode==='fullBody'?jp:''});
  await api.manualConvert({completed:()=>completed++});
  check('Manual reports hidden quote scope',notes.at(-1).includes('[SCOPE]'));
  check('Read-only scope diagnosis never writes',writes===1);
  check('Manual completes once',completed===1);
  notes.length=0;
  item.body.getAsync=(type,options,cb)=>cb({status:'failed',error:{code:5001,message:'PRIVATE MAIL CONTENT'}});
  await api.manualConvert({completed:()=>completed++});
  check('Manual reports API error code',notes.at(-1).includes('5001'));
  check('Diagnostic does not expose mail content',!notes.join('').includes('PRIVATE MAIL'));
  check('Manual completes after API failure',completed===2);
  check('Unsupported markup not mistaken for English',api.replaceHtmlHeaders('<div id="divRplyFwdMsg"><span>差出人: Alice</span></div>').reason.includes('Unsupported'));
  return results;
});
results.forEach(x=>console.log('PASS '+x));
const errors=[];page.on('pageerror',e=>errors.push(e.message));
await page.goto(pathToFileURL(path.join(root,'configure.html')).href);
await page.locator('#url').fill('https://example.com/my-repo/');
for(const mode of ['user','admin-auto']){
 await page.locator('#mode').selectOption(mode);
 const [download]=await Promise.all([page.waitForEvent('download'),page.locator('#download').click()]);
 const xml=fs.readFileSync(await download.path(),'utf8');
 assert(xml.includes('<AppDomain>https://example.com</AppDomain>'));
 assert(!xml.includes('{{BASE_URL}}'));
 assert(xml.includes('<Permissions>ReadWriteItem</Permissions>'));
 const parsed=await page.evaluate(xml=>{const doc=new DOMParser().parseFromString(xml,'application/xml');return {error:!!doc.querySelector('parsererror'),events:doc.getElementsByTagName('LaunchEvent').length,actions:[...doc.getElementsByTagName('Action')].every(a=>a.getAttribute('xsi:type')==='ExecuteFunction')};},xml);
 assert(!parsed.error);assert(parsed.actions);assert.equal(parsed.events,mode==='user'?0:1);
 console.log('PASS '+mode+' generator and XML structure (not schema validation)');
}
await page.locator('#url').fill('http://example.com');await page.locator('#download').click();
assert((await page.locator('#status').textContent()).includes('HTTPS'));
assert.deepEqual(errors,[]);console.log('PASS invalid URL rejected, no script errors');
} finally {await browser.close();}
})().catch(e=>{console.error(e);process.exitCode=1;});
