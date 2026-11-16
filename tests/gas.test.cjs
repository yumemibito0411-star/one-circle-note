// npm依存なし: node --test tests/gas.test.cjs
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const {randomUUID} = require('node:crypto');
const root = path.resolve(__dirname, '..');

class Sheet {
  rows = [];
  maxRows = 1000;
  getLastRow() { return this.rows.length; }
  getMaxRows() { return this.maxRows; }
  insertRowsAfter(_after, count) { this.maxRows += count; }
  getRange(row, col, height, width) {
    return {
      getValues: () => Array.from({length: height}, (_, y) => Array.from({length: width}, (_, x) => this.rows[row - 1 + y]?.[col - 1 + x] ?? '')),
      setValues: values => {
        for (let y = 0; y < height; y++) {
          this.rows[row - 1 + y] ??= [];
          for (let x = 0; x < width; x++) this.rows[row - 1 + y][col - 1 + x] = values[y][x];
        }
      },
      clearContent: () => {
        for (let y = 0; y < height; y++) {
          for (let x = 0; x < width; x++) if (this.rows[row - 1 + y]) this.rows[row - 1 + y][col - 1 + x] = '';
        }
        while (this.rows.length && this.rows.at(-1).every(v => v === '')) this.rows.pop();
      }
    };
  }
}
function fixture({bound = true, setup = true} = {}) {
  const props = new Map();
  const sheets = new Map();
  let held = false, failLock = false, fetchCount = 0;
  const setupLogs = [];
  const db = {
    getId: () => 'mock-database',
    getSheetByName: name => sheets.get(name),
    insertSheet: name => { const sheet = new Sheet(); sheets.set(name, sheet); return sheet; }
  };
  const content = text => ({text, getContent: () => text, setTitle() { return this; }, setMimeType() { return this; }});
  const c = vm.createContext({
    console: {log: message => setupLogs.push(message)},
    PropertiesService: {getScriptProperties: () => ({getProperty: key => props.get(key), setProperty: (key, value) => props.set(key, value)})},
    LockService: {getScriptLock: () => ({
      tryLock() { if (failLock) return false; assert.equal(held, false); held = true; return true; },
      waitLock() { assert.equal(held, false); held = true; },
      releaseLock() { assert.equal(held, true); held = false; }
    })},
    SpreadsheetApp: {getActiveSpreadsheet: () => bound ? db : null, create: () => db,
      openById: id => { assert.equal(id, 'mock-database'); return db; }, flush() {}},
    Utilities: {getUuid: randomUUID, formatDate: (_date, zone, format) => {
      assert.equal(zone, 'Asia/Tokyo'); assert.equal(format, 'yyyy-MM-dd'); return '2026-11-16';
    }},
    ScriptApp: {getService: () => ({getUrl: () => 'https://example.invalid/exec'})},
    HtmlService: {createHtmlOutputFromFile: name => { assert.equal(name, 'index'); return content(fs.readFileSync(path.join(root, 'index.html'), 'utf8')); }, createHtmlOutput: content},
    ContentService: {createTextOutput: content, MimeType: {JSON:'application/json'}},
    UrlFetchApp: {fetch: () => { assert.equal(held, false, 'AI must release database lock'); fetchCount++; return {getResponseCode: () => 200, getContentText: () => JSON.stringify({stop_reason:'end_turn',content:[{type:'text',text:'{"summary":"ok"}'}]})}; }}
  });
  for (const file of ['Code.gs','ai.gs']) vm.runInContext(fs.readFileSync(path.join(root,'gas',file),'utf8'), c);
  let admin;
  if (setup) {
    c.setup();
    admin = JSON.parse(sheets.get('members').rows[1][1]);
  }
  const call = (key, action, payload = {}) => JSON.parse(c.api(key, action, JSON.stringify(payload)));
  const ok = (key, action, payload) => { const result = call(key, action, payload); assert.equal(result.ok, true, JSON.stringify(result)); return result.data; };
  return {c, props, sheets, admin, setupLogs, call, ok, isHeld: () => held, failLock: () => {failLock = true;}, fetchCount: () => fetchCount};
}
const person = (id = 'p1', name = 'テスト') => ({id,name,profile:{job:'営業'},actions:[]});

test('setup creates tables and only emits initial administrator key once', () => {
  const f = fixture();
  assert.equal(f.props.get('DB_ID'), 'mock-database');
  assert.deepEqual([...f.sheets.keys()], ['members','people','logs','handoffs','config']);
  assert.match(f.admin.key, /^[a-f0-9]{64}$/);
  assert.equal(f.admin.role, 'admin');
  assert.equal(f.setupLogs.length, 1);
  assert.ok(f.setupLogs[0].includes(f.admin.key));
  f.c.setup();
  assert.equal(f.setupLogs.length, 1);
  assert.equal(f.sheets.get('members').rows.length, 2);
  assert.equal(f.isHeld(), false);
});
test('setup also creates a spreadsheet when no spreadsheet is bound', () => {
  const f = fixture({bound:false});
  assert.equal(f.props.get('DB_ID'), 'mock-database');
});
test('not configured, invalid authentication, malformed payload and busy lock use envelopes', () => {
  const empty = fixture({setup:false});
  assert.equal(empty.call('x','bootstrap').error, 'not_setup');
  const f = fixture();
  for (const key of ['',null,{},'incorrect']) assert.deepEqual(f.call(key,'bootstrap'), {ok:false,error:'invalid_key'});
  assert.deepEqual(JSON.parse(f.c.api(f.admin.key,'bootstrap','not JSON')), {ok:false,error:'invalid_payload'});
  assert.equal(f.call(f.admin.key,'unknown').error, 'unknown_action');
  assert.equal(f.isHeld(), false);
  f.failLock();
  assert.equal(f.call(f.admin.key,'bootstrap').error, 'busy');
});
test('bootstrap returns frontend shape, savePerson replaces full record, owner is isolated', () => {
  const f = fixture();
  const a = f.admin;
  const member = f.ok(a.key,'addMember',{name:'メンバー',role:'member'});
  const boot = f.ok(member.key,'bootstrap');
  assert.deepEqual(Object.keys(boot).sort(), ['ai','global','handoffs','logs','me','members','mine','people','target'].sort());
  assert.equal(boot.me.id,member.id);
  assert.equal(boot.target.id,member.id);
  assert.equal(boot.ai,false);
  assert.equal('key' in boot.me,false);
  assert.deepEqual(boot.members,[]);
  f.ok(member.key,'savePerson',{person:person()});
  assert.equal(f.ok(member.key,'bootstrap').people[0].profile.job,'営業');
  assert.deepEqual(f.ok(a.key,'bootstrap').people,[]);
  f.ok(a.key,'savePerson',{person:person('p1','管理者のデータ')});
  f.ok(member.key,'savePerson',{person:person('p1','更新後')});
  assert.equal(f.ok(member.key,'bootstrap').people[0].name,'更新後');
  assert.equal(f.ok(a.key,'bootstrap').people[0].name,'管理者のデータ');
});
test('logs and handoffs preserve fields, reference own people, deletePerson cascades', () => {
  const f = fixture(); const key = f.admin.key;
  f.ok(key,'savePerson',{person:person()});
  const log = {id:'l1',personId:'p1',type:'meeting',date:'2026-11-16',text:'面談',createdAt:'2026-11-16T01:00:00Z'};
  const handoff = {id:'h1',personId:'p1',status:'sent',datetime:'2026-11-15T10:00',talked:''};
  f.ok(key,'addLog',{log}); f.ok(key,'saveHandoff',{handoff});
  assert.deepEqual(f.ok(key,'bootstrap').logs,[log]);
  assert.deepEqual(f.ok(key,'bootstrap').handoffs,[handoff]);
  assert.equal(f.call(key,'addLog',{log:{...log,id:'bad',personId:'other'}}).error,'not_found');
  f.ok(key,'deleteLog',{id:'l1'}); f.ok(key,'deleteHandoff',{id:'h1'});
  assert.deepEqual(f.ok(key,'bootstrap').logs,[]);
  assert.deepEqual(f.ok(key,'bootstrap').handoffs,[]);
  f.ok(key,'addLog',{log}); f.ok(key,'saveHandoff',{handoff});
  f.ok(key,'deletePerson',{id:'p1'});
  const boot = f.ok(key,'bootstrap');
  assert.deepEqual([boot.people,boot.logs,boot.handoffs],[[],[],[]]);
});
test('role restrictions, view, active, self_admin, and key rotation enforced on server', () => {
  const f = fixture(); const a = f.admin;
  const member = f.ok(a.key,'addMember',{name:'メンバー',role:'member'});
  const leader = f.ok(a.key,'addMember',{name:'リーダー',role:'leader'});
  assert.equal(f.call(member.key,'bootstrap',{view:a.id}).error,'forbidden');
  for (const [action,payload] of [['saveGlobal',{patch:{lineTemplate:'x'}}],['addMember',{name:'x'}],['updateMember',{id:member.id,role:'admin'}],['teamSummary',{}],['importData',{data:{people:[]}}]]) {
    assert.equal(f.call(member.key,action,payload).error,'forbidden');
  }
  assert.equal(f.call(leader.key,'addMember',{name:'x'}).error,'forbidden');
  const view = f.ok(leader.key,'bootstrap',{view:member.id});
  assert.equal(view.target.id,member.id); assert.equal('key' in view.target,false); assert.deepEqual(view.members,[]);
  assert.equal(f.call(leader.key,'savePerson',{view:member.id,person:person()}).error,'forbidden');
  for (const patch of [{active:false},{role:'member'}]) assert.equal(f.call(a.key,'updateMember',{id:a.id,...patch}).error,'self_admin');
  f.ok(a.key,'updateMember',{id:member.id,active:false});
  assert.equal(f.call(member.key,'bootstrap').error,'invalid_key');
  f.ok(a.key,'updateMember',{id:member.id,active:true});
  const updated = f.ok(a.key,'updateMember',{id:member.id,newKey:true});
  assert.notEqual(updated.key,member.key);
  assert.equal(f.call(member.key,'bootstrap').error,'invalid_key');
  assert.equal(f.call(updated.key,'bootstrap').ok,true);
  const rotatedAdmin = f.ok(a.key,'updateMember',{id:a.id,newKey:true});
  assert.equal(f.call(a.key,'bootstrap').error,'invalid_key');
  assert.equal(f.call(rotatedAdmin.key,'bootstrap').ok,true);
});
test('config deep merge, team summary metrics and no keys exposed', () => {
  const f = fixture(); const key = f.admin.key;
  f.ok(key,'saveGlobal',{patch:{lineTemplate:'template'}});
  f.ok(key,'saveMine',{patch:{goals:{'2026-11':{meeting:3,contact:4}},themes:{'2026-11':'テーマ'}}});
  f.ok(key,'saveMine',{patch:{goals:{'2026-11':{contact:6}}}});
  f.ok(key,'savePerson',{person:{...person(),actions:[{id:'a1',text:'連絡',due:'2026-11-15',done:false}]}});
  f.ok(key,'addLog',{log:{id:'l1',personId:'p1',type:'meeting',date:'2026-11-16'}});
  f.ok(key,'addLog',{log:{id:'l2',personId:'p1',type:'meeting',date:'2026-10-16'}});
  f.ok(key,'saveHandoff',{handoff:{id:'h1',personId:'p1',status:'sent',datetime:'2026-11-15T12:00',talked:''}});
  const summary = f.ok(key,'teamSummary')[0];
  assert.deepEqual(summary, {id:f.admin.id,name:f.admin.name,role:'admin',counts:{contact:0,meeting:1,propose:0,connect:0,won:0,new:0},goals:{meeting:3,contact:6},theme:'テーマ',list:1,late:1,back:1,last:'2026-11-16'});
  assert.equal(f.ok(key,'bootstrap').global.lineTemplate,'template');
  assert.equal(f.call(key,'saveMine',{patch:{businesses:[]}}).error,'invalid_payload');
});
test('import accepts raw bootstrap and envelope, overwrites IDs, targets explicit new member, ignores old credentials', () => {
  const f = fixture(); const key = f.admin.key;
  const member = f.ok(key,'addMember',{name:'取り込み先'});
  const source = {
    people:[person()],logs:[{id:'l1',personId:'p1',type:'meeting',date:'2026-11-16'}],
    handoffs:[{id:'h1',personId:'p1',talked:'完了'}],global:{lineTemplate:'旧書式'},mine:{me:'旧名',goals:{'2026-11':{meeting:3}}},
    me:{id:'old',key:'old-placeholder'},members:[{id:'old',key:'old-placeholder',role:'admin'}],ai:true
  };
  assert.deepEqual(f.ok(key,'importData',{memberId:member.id,data:source}).counts,{people:1,logs:1,handoffs:1});
  let boot = f.ok(member.key,'bootstrap');
  assert.equal(boot.people[0].name,'テスト'); assert.equal(boot.logs.length,1); assert.equal(boot.handoffs.length,1);
  assert.equal(boot.mine.me,'旧名'); assert.equal(boot.global.lineTemplate,'旧書式');
  assert.equal(boot.me.role,'member'); assert.equal(boot.ai,false);
  assert.equal(f.ok(key,'bootstrap').members.length,2);
  f.ok(key,'importData',{memberId:member.id,data:{ok:true,data:{people:[person('p1','上書き')]}}});
  boot = f.ok(member.key,'bootstrap'); assert.equal(boot.people.length,1); assert.equal(boot.people[0].name,'上書き');
  assert.deepEqual(f.ok(key,'bootstrap').people,[]);
  f.ok(key,'importData',{people:[person('self')]});
  assert.equal(f.ok(key,'bootstrap').people[0].id,'self');
});
test('invalid import does not partially persist and dangerous properties/IDs/oversize rejected', () => {
  const f = fixture(); const key = f.admin.key;
  const badImports = [
    {people:[person()],logs:[{id:'l1',personId:'missing'}]},
    {people:[person()],mine:{unknown:'field'}},
    {people:[person()],handoffs:{}},
    {people:[person('bad"id')]},
    {people:[{...person(),actions:'bad'}]},
    {people:[{...person(),memo:'x'.repeat(45001)}]},
    {people:[person()],global:null},
    {unrecognized:[]}
  ];
  for (const data of badImports) {
    assert.equal(f.call(key,'importData',{data}).error,'invalid_payload');
    assert.deepEqual(f.ok(key,'bootstrap').people,[]);
  }
  assert.equal(JSON.parse(f.c.api(key,'saveMine','{"patch":{"__proto__":{"polluted":true}}}')).error,'invalid_payload');
  assert.equal({}.polluted,undefined);
});
test('all API entry points share envelopes, doGet fills safe BOOT, no incidental key logs', () => {
  const f = fixture(); const key = f.admin.key;
  const post = request => JSON.parse(f.c.doPost({postData:{contents:JSON.stringify(request)}}).text);
  assert.deepEqual(post({key:'wrong',action:'bootstrap',payload:{}}), f.call('wrong','bootstrap'));
  assert.deepEqual(post({key,action:'bootstrap',payload:{}}), f.call(key,'bootstrap'));
  assert.equal(JSON.parse(f.c.doPost({postData:{contents:'not JSON'}}).text).error,'invalid_payload');
  const html = f.c.doGet({parameter:{k:key}}).text;
  const boot = JSON.parse(html.match(/<script type="application\/json" id="boot">(.*?)<\/script>/s)[1]);
  assert.deepEqual(boot,{key,url:'https://example.invalid/exec',external:false});
  const invalidHtml = f.c.doGet({parameter:{k:'</script><script>alert(1)</script>'}}).text;
  assert.equal(invalidHtml.includes('alert(1)'),false);
  assert.equal(invalidHtml.includes('__BOOT__'),false);
  assert.equal(f.setupLogs.length,1);
});
test('ai no_key uses shared envelope, model API works after configuration, releases lock', () => {
  const f = fixture(); const key = f.admin.key;
  assert.deepEqual(f.call(key,'ai',{prompt:'test'}),{ok:false,error:'no_key'});
  assert.equal(f.fetchCount(),0);
  f.props.set('ANTHROPIC_API_KEY','mock-placeholder');
  assert.equal(f.ok(key,'bootstrap').ai,true);
  assert.equal(f.ok(key,'ai',{prompt:'test'}),'{"summary":"ok"}');
  assert.equal(f.fetchCount(),1);
  assert.equal(f.isHeld(),false);
  assert.equal(f.setupLogs.length,1);
});

test('deleting an ID never deletes another member record with the same ID', () => {
  const f = fixture(); const key = f.admin.key;
  const member = f.ok(key,'addMember',{name:'別メンバー'});
  for (const ownKey of [key,member.key]) {
    f.ok(ownKey,'savePerson',{person:person()});
    f.ok(ownKey,'addLog',{log:{id:'l1',personId:'p1',type:'contact',date:'2026-11-16'}});
    f.ok(ownKey,'saveHandoff',{handoff:{id:'h1',personId:'p1'}});
    f.ok(ownKey,'saveMine',{patch:{me:ownKey === key ? '管理者' : 'メンバー'}});
  }
  f.ok(member.key,'deletePerson',{id:'p1'});
  const adminBoot = f.ok(key,'bootstrap');
  assert.equal(adminBoot.people.length,1); assert.equal(adminBoot.logs.length,1); assert.equal(adminBoot.handoffs.length,1);
  assert.equal(adminBoot.mine.me,'管理者');
  assert.equal(f.ok(member.key,'bootstrap').mine.me,'メンバー');
});

function frontendFixture() {
  const html = fs.readFileSync(path.join(root,'index.html'),'utf8');
  const actionSource = html.slice(html.indexOf('async function onAct('),html.indexOf('/* ============ 起動 ============ */'));
  const membersSource = html.slice(html.indexOf('function vMembers('),html.indexOf('/* ============ 目標・行動量 ============ */'));
  const notices = [], calls = [], histories = [], loads = [];
  const c = vm.createContext({
    S:{me:{id:'admin',role:'admin'},people:new Map(),members:[{id:'admin',name:'管理者',role:'admin',active:true,key:'old-placeholder'}],ui:{},d:{},canWrite:true},
    BOOT:{key:'old-placeholder',url:'https://example.invalid/exec',external:true},
    ROLE_LABEL:{member:'メンバー',leader:'リーダー',admin:'管理者'},
    isAdmin: () => c.S.me.role === 'admin',
    draft: (key, def = '') => key in c.S.d ? c.S.d[key] : def,
    memberUrl: () => '',
    esc: value => String(value ?? '').replace(/[&<>"']/g, char => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[char])),
    head: () => '', render() {}, toast: text => notices.push(text),
    loadData: async view => loads.push(view),
    call: async (action,payload) => { calls.push({action,payload}); return action === 'importData' ? {counts:{people:1,logs:2,handoffs:3}} : {...c.S.members[0],key:'new-placeholder'}; },
    URL, location:{href:'https://example.invalid/app#members',hash:'#members'},
    history:{replaceState: (_state,_title,url) => histories.push(String(url))},window:{}
  });
  vm.runInContext(membersSource + '\n' + actionSource,c);
  return {c,notices,calls,histories,loads};
}

test('frontend import is admin-only, rejects malformed JSON, sends selected owner and reloads', async () => {
  const f = frontendFixture();
  assert.ok(f.c.vMembers().includes('データを取り込む（JSONを貼り付け）'));
  f.c.S.me.role = 'member';
  assert.equal(f.c.vMembers().includes('data-act="importData"'),false);
  await f.c.onAct({}, {dataset:{}},'importData');
  assert.equal(f.calls.length,0);
  f.c.S.me.role = 'admin';
  f.c.S.d.import_json = 'bad JSON';
  await f.c.onAct({}, {dataset:{}},'importData');
  assert.equal(f.calls.length,0); assert.ok(f.notices.at(-1).includes('JSON'));
  f.c.S.d.import_json = JSON.stringify({people:[person()]});
  f.c.S.d.import_member = 'selected-member';
  await f.c.onAct({}, {dataset:{}},'importData');
  assert.equal(f.calls[0].action,'importData');
  assert.equal(f.calls[0].payload.memberId,'selected-member');
  assert.equal(f.calls[0].payload.data.people[0].id,'p1');
  assert.equal('import_json' in f.c.S.d,false);
  assert.deepEqual(f.loads,['']);
  assert.ok(f.notices.at(-1).includes('人物1件・記録2件・共有3件'));
});

test('self key rotation refreshes authentication and preserves page route', async () => {
  const f = frontendFixture();
  await f.c.onAct({}, {dataset:{mid:'admin',op:'rekey'}},'memberOp');
  assert.equal(f.c.BOOT.key,'new-placeholder');
  const url = new URL(f.histories[0]);
  assert.equal(url.searchParams.get('k'),'new-placeholder');
  assert.equal(url.hash,'#members');
  f.c.BOOT.external = false;
  let gasHistory;
  f.c.window.google = f.c.google = {script:{history:{replace: (...args) => {gasHistory = args;}}}};
  await f.c.onAct({}, {dataset:{mid:'admin',op:'rekey'}},'memberOp');
  assert.equal(gasHistory[1].k,'new-placeholder');
  assert.equal(gasHistory[2],'members');
});
