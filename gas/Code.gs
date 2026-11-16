/** ONE CIRCLE の独立サーバー。アプリ用の入口は doGet / doPost / api。setup は初期設定専用。
 * データ操作・認証は共通の排他区間で行う。末尾 _ の関数は
 * google.script.run から直接実行できない。キーはメンバー管理者だけに返す。
 */
var TABLES_ = ['members', 'people', 'logs', 'handoffs', 'config'];
var GLOBAL_FIELDS_ = ['businesses', 'kinds', 'lineTemplate'];
var MINE_FIELDS_ = ['me', 'goals', 'themes', 'shareAsk'];
var ERROR_CODES_ = ['invalid_key', 'self_admin', 'forbidden', 'not_found',
  'invalid_payload', 'not_setup', 'server_error', 'busy', 'unknown_action',
  'no_key', 'rate_limited', 'refused', 'too_long', 'ai_error'];

/** エディタから一度実行。最初のキーだけを実行ログに表示する。
 * 再実行はデータやキーを変更せず、既存キーも再表示しない。
 */
function setup() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var props = PropertiesService.getScriptProperties();
    var id = props.getProperty('DB_ID');
    var db = id ? SpreadsheetApp.openById(id) :
      (SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.create('ONE CIRCLE 人脈ノート'));
    TABLES_.forEach(function (name) {
      var sheet = db.getSheetByName(name) || db.insertSheet(name);
      if (sheet.getLastRow() === 0) sheet.getRange(1, 1, 1, 2).setValues([['id', 'json']]);
      var header = sheet.getRange(1, 1, 1, 2).getValues()[0];
      if (header[0] !== 'id' || header[1] !== 'json') throw new Error('既存シートの列名を確認してください: ' + name);
    });
    props.setProperty('DB_ID', db.getId());
    var state = readState_(db);
    if (state.members.length) return '設定済みです。既存データ・キーは変更していません。';
    var member = newMember_('最初の管理者', 'admin');
    put_(state, 'members', member.id, member);
    flush_(state);
    SpreadsheetApp.flush();
    // 初回setupだけの明示的な出力。api/doPost/doGetはキーをログ出力しない。
    var message = '初期管理者キー（この実行時だけ表示）: ' + member.key +
      '\nデプロイ後のウェブアプリURLの末尾に ?k= とこのキーを付けて開いてください。安全な場所に控えてください。';
    console.log(message);
    return message;
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  var key = e && e.parameter && e.parameter.k || '';
  // 他の入口と同じ認証・エラー変換。無効キーでも画面を返し、画面が案内を表示する。
  var result = JSON.parse(api(key, 'bootstrap', '{}'));
  var boot = { key: result.ok ? key : '', url: ScriptApp.getService().getUrl() || '', external: false };
  var json = JSON.stringify(boot).replace(/</g, '\\u003c').replace(/>/g, '\\u003e')
    .replace(/&/g, '\\u0026').replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029');
  var html = HtmlService.createHtmlOutputFromFile('index').getContent();
  html = html.replace('__BOOT__', function () { return json; });
  return HtmlService.createHtmlOutput(html).setTitle('ONE CIRCLE 人脈ノート');
}

function doPost(e) {
  var result;
  try {
    var request = parseObject_(e && e.postData && e.postData.contents);
    result = api(request.key, request.action, JSON.stringify(request.payload || {}));
  } catch (error) {
    result = JSON.stringify(errorResult_(error));
  }
  return ContentService.createTextOutput(result).setMimeType(ContentService.MimeType.JSON);
}

/** google.script.run もHTTPも {ok:true,data} / {ok:false,error} のJSON文字列を返す。 */
function api(key, action, payloadJson) {
  var lock = LockService.getScriptLock();
  var held = false;
  try {
    if (!lock.tryLock(30000)) throw new Error('busy');
    held = true;
    var id = PropertiesService.getScriptProperties().getProperty('DB_ID');
    if (!id) throw new Error('not_setup');
    var state = readState_(SpreadsheetApp.openById(id));
    var me = state.members.find(function (row) {
      return typeof key === 'string' && key.length > 0 && row.value.key === key && row.value.active === true;
    });
    if (!me) throw new Error('invalid_key');
    me = me.value;
    var payload = parseObject_(payloadJson);
    if (action === 'ai') {
      // 外部APIの待ち時間はスプレッドシートのロックを保持しない。
      lock.releaseLock(); held = false;
      return JSON.stringify({ ok: true, data: aiAction_(payload) });
    }
    var data = dispatch_(state, me, action, payload);
    flush_(state);
    SpreadsheetApp.flush();
    return JSON.stringify({ ok: true, data: data });
  } catch (error) {
    return JSON.stringify(errorResult_(error));
  } finally {
    if (held) lock.releaseLock();
  }
}

function errorResult_(error) {
  var code = error && error.message;
  return { ok: false, error: ERROR_CODES_.indexOf(code) !== -1 ? code : 'server_error' };
}
function object_(value) { return !!value && typeof value === 'object' && !Array.isArray(value); }
function parseObject_(text) {
  if (typeof text !== 'string' || text.length > 5000000) throw new Error('invalid_payload');
  try {
    var value = JSON.parse(text, function (key, item) {
      if (['__proto__', 'prototype', 'constructor'].indexOf(key) !== -1) throw new Error('invalid_payload');
      return item;
    });
    if (!object_(value)) throw new Error('invalid_payload');
    return value;
  } catch (error) { throw new Error('invalid_payload'); }
}
function id_(id) {
  // IDsは画面のdata属性やDOMのidにも入るため、引用符などを許可しない。
  if (typeof id !== 'string' || !/^[A-Za-z0-9_-]{1,160}$/.test(id)) throw new Error('invalid_payload');
  return id;
}
function admin_(me) { if (me.role !== 'admin') throw new Error('forbidden'); }
function leader_(me) { if (['admin', 'leader'].indexOf(me.role) === -1) throw new Error('forbidden'); }
function publicMember_(member) { return { id: member.id, name: member.name, role: member.role, active: member.active }; }
function newKey_() { return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, ''); }
function newMember_(name, role) {
  if (typeof name !== 'string' || !name.trim() || name.length > 200 ||
      ['admin', 'leader', 'member'].indexOf(role) === -1) throw new Error('invalid_payload');
  return { id: Utilities.getUuid(), name: name.trim(), role: role, active: true, key: newKey_() };
}
function readState_(db) {
  var state = { db: db, dirty: {} };
  TABLES_.forEach(function (name) {
    var sheet = db.getSheetByName(name);
    if (!sheet) throw new Error('not_setup');
    var rows = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues() : [];
    state[name] = rows.filter(function (row) { return row[0] !== ''; }).map(function (row) {
      return { id: String(row[0]), value: parseObject_(row[1]) };
    });
  });
  return state;
}
function put_(state, table, id, value) {
  // Sheetsの1セル上限を超えるレコードは、書き込み前に拒否する。
  if (JSON.stringify(value).length > 45000) throw new Error('invalid_payload');
  var row = state[table].find(function (item) { return item.id === id; });
  if (row) row.value = value; else state[table].push({ id: id, value: value });
  state.dirty[table] = true;
}
function remove_(state, table, predicate) {
  state[table] = state[table].filter(function (row) { return !predicate(row.value); });
  state.dirty[table] = true;
}
function flush_(state) {
  Object.keys(state.dirty).forEach(function (table) {
    var sheet = state.db.getSheetByName(table);
    var oldCount = sheet.getLastRow() - 1;
    var rows = state[table].map(function (row) { return [row.id, JSON.stringify(row.value)]; });
    if (rows.length + 1 > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), rows.length + 1 - sheet.getMaxRows());
    if (rows.length) sheet.getRange(2, 1, rows.length, 2).setValues(rows);
    if (oldCount > rows.length) sheet.getRange(rows.length + 2, 1, oldCount - rows.length, 2).clearContent();
  });
}
function records_(state, table, owner) {
  return state[table].filter(function (row) { return row.value.ownerId === owner; })
    .map(function (row) { return row.value.record; });
}
function config_(state, id) {
  var row = state.config.find(function (item) { return item.id === id; });
  return row ? row.value : {};
}
function merge_(target, patch) {
  Object.keys(patch).forEach(function (key) {
    if (object_(patch[key])) target[key] = merge_(object_(target[key]) ? target[key] : {}, patch[key]);
    else target[key] = patch[key];
  });
  return target;
}
function configPatch_(state, id, patch, fields) {
  if (!object_(patch) || Object.keys(patch).some(function (key) { return fields.indexOf(key) === -1; })) throw new Error('invalid_payload');
  Object.keys(patch).forEach(function (key) {
    var value = patch[key];
    if (['businesses', 'kinds'].indexOf(key) !== -1 && !Array.isArray(value)) throw new Error('invalid_payload');
    if (['lineTemplate', 'me', 'shareAsk'].indexOf(key) !== -1 && typeof value !== 'string') throw new Error('invalid_payload');
    if (['goals', 'themes'].indexOf(key) !== -1 && !object_(value)) throw new Error('invalid_payload');
  });
  var next = merge_(JSON.parse(JSON.stringify(config_(state, id))), patch);
  put_(state, 'config', id, next);
  return next;
}
function saveRecord_(state, table, owner, record) {
  if (!object_(record)) throw new Error('invalid_payload');
  id_(record.id);
  if (table === 'people') {
    ['actions', 'needs', 'frustrations', 'talkPoints', 'avoid'].forEach(function (key) {
      if (record[key] !== undefined && !Array.isArray(record[key])) throw new Error('invalid_payload');
    });
    ['profile', 'biz', 'fit'].forEach(function (key) {
      if (record[key] !== undefined && !object_(record[key])) throw new Error('invalid_payload');
    });
    (record.actions || []).forEach(function (action) {
      if (!object_(action) || (action.due !== undefined && typeof action.due !== 'string')) throw new Error('invalid_payload');
    });
  }
  ['date', 'datetime', 'talked', 'type', 'createdAt', 'updatedAt'].forEach(function (key) {
    if (record[key] !== undefined && typeof record[key] !== 'string') throw new Error('invalid_payload');
  });
  if ((table === 'logs' || table === 'handoffs')) id_(record.personId);
  // 行のIDは所有者＋レコードID。同一IDの他メンバーのデータを上書きしない。
  put_(state, table, owner + ':' + record.id, { ownerId: owner, record: record });
  return record;
}
function memberById_(state, id) {
  var row = state.members.find(function (item) { return item.id === id; });
  if (!row) throw new Error('not_found');
  return row.value;
}
function dispatch_(state, me, action, payload) {
  // 書き込み先は認証した本人に固定。閲覧先へ書き込むAPIは設けない。
  if (payload.view && payload.view !== me.id && action !== 'bootstrap') throw new Error('forbidden');
  switch (action) {
    case 'bootstrap': {
      var target = me;
      if (payload.view && payload.view !== me.id) {
        leader_(me);
        target = memberById_(state, id_(payload.view));
      }
      return { me: publicMember_(me), target: publicMember_(target),
        members: me.role === 'admin' ? state.members.map(function (row) { return row.value; }) : [],
        people: records_(state, 'people', target.id), logs: records_(state, 'logs', target.id),
        handoffs: records_(state, 'handoffs', target.id), global: config_(state, 'global'),
        mine: config_(state, 'mine:' + target.id), ai: aiEnabled_() };
    }
    case 'savePerson': return saveRecord_(state, 'people', me.id, payload.person);
    case 'deletePerson': {
      var personId = id_(payload.id);
      remove_(state, 'people', function (v) { return v.ownerId === me.id && v.record.id === personId; });
      ['logs', 'handoffs'].forEach(function (table) {
        remove_(state, table, function (v) { return v.ownerId === me.id && v.record.personId === personId; });
      });
      return null;
    }
    case 'addLog':
    case 'saveHandoff': {
      var table = action === 'addLog' ? 'logs' : 'handoffs';
      var record = action === 'addLog' ? payload.log : payload.handoff;
      if (!object_(record) || !records_(state, 'people', me.id).some(function (p) { return p.id === record.personId; })) throw new Error('not_found');
      return saveRecord_(state, table, me.id, record);
    }
    case 'deleteLog':
    case 'deleteHandoff': {
      var table = action === 'deleteLog' ? 'logs' : 'handoffs';
      var recordId = id_(payload.id);
      remove_(state, table, function (v) { return v.ownerId === me.id && v.record.id === recordId; });
      return null;
    }
    case 'saveGlobal': admin_(me); return configPatch_(state, 'global', payload.patch, GLOBAL_FIELDS_);
    case 'saveMine': return configPatch_(state, 'mine:' + me.id, payload.patch, MINE_FIELDS_);
    case 'addMember': {
      admin_(me);
      var member = newMember_(payload.name, payload.role || 'member');
      put_(state, 'members', member.id, member);
      return member;
    }
    case 'updateMember': {
      admin_(me);
      var member = JSON.parse(JSON.stringify(memberById_(state, id_(payload.id))));
      if (member.id === me.id && ((payload.role !== undefined && payload.role !== me.role) || payload.active === false)) throw new Error('self_admin');
      if (payload.role !== undefined) {
        if (['admin', 'leader', 'member'].indexOf(payload.role) === -1) throw new Error('invalid_payload');
        member.role = payload.role;
      }
      if (payload.active !== undefined) {
        if (typeof payload.active !== 'boolean') throw new Error('invalid_payload');
        member.active = payload.active;
      }
      if (payload.newKey === true) member.key = newKey_();
      put_(state, 'members', member.id, member);
      return member;
    }
    case 'teamSummary': leader_(me); return teamSummary_(state);
    case 'importData': admin_(me); return importData_(state, me, payload);
    default: throw new Error('unknown_action');
  }
}

function importData_(state, me, payload) {
  // dataには旧bootstrapのdata部分または {ok:true,data:...} をそのまま渡せる。
  var source = payload.data || payload;
  if (source.ok === true && object_(source.data)) source = source.data;
  if (!object_(source)) throw new Error('invalid_payload');
  var owner = payload.memberId ? memberById_(state, id_(payload.memberId)).id : me.id;
  var recognized = ['people', 'logs', 'handoffs', 'global', 'mine'];
  if (!recognized.some(function (key) { return Object.prototype.hasOwnProperty.call(source, key); })) throw new Error('invalid_payload');
  var counts = { people: 0, logs: 0, handoffs: 0 };
  ['people', 'logs', 'handoffs'].forEach(function (table) {
    if (source[table] === undefined) return;
    if (!Array.isArray(source[table])) throw new Error('invalid_payload');
    source[table].forEach(function (record) { saveRecord_(state, table, owner, record); counts[table]++; });
  });
  // 全検証が通った後にだけapi()がflushする。不正なJSONで部分反映しない。
  ['logs', 'handoffs'].forEach(function (table) {
    if (!source[table]) return;
    var ids = records_(state, 'people', owner).map(function (p) { return p.id; });
    source[table].forEach(function (record) { if (ids.indexOf(record.personId) === -1) throw new Error('invalid_payload'); });
  });
  if (source.global !== undefined) configPatch_(state, 'global', source.global, GLOBAL_FIELDS_);
  if (source.mine !== undefined) configPatch_(state, 'mine:' + owner, source.mine, MINE_FIELDS_);
  // me/target/members/aiなどは無視し、旧キー・権限・DB_IDを持ち込まない。
  return { memberId: owner, counts: counts };
}

function teamSummary_(state) {
  var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  var month = today.slice(0, 7);
  return state.members.filter(function (row) { return row.value.active; }).map(function (row) {
    var member = row.value;
    var people = records_(state, 'people', member.id);
    var logs = records_(state, 'logs', member.id);
    var handoffs = records_(state, 'handoffs', member.id);
    var mine = config_(state, 'mine:' + member.id);
    var counts = { contact: 0, meeting: 0, propose: 0, connect: 0, won: 0, new: 0 };
    logs.forEach(function (log) {
      if ((log.date || '').slice(0, 7) === month && Object.prototype.hasOwnProperty.call(counts, log.type)) counts[log.type]++;
    });
    var late = 0;
    people.forEach(function (person) {
      (person.actions || []).forEach(function (action) { if (!action.done && action.due && action.due < today) late++; });
    });
    var back = handoffs.filter(function (h) {
      return !(h.talked || '').trim() && h.status === 'sent' && h.datetime && h.datetime.slice(0, 10) < today &&
        people.some(function (p) { return p.id === h.personId; });
    }).length;
    var dates = logs.map(function (log) { return log.date || ''; }).sort();
    return { id: member.id, name: member.name, role: member.role, counts: counts,
      goals: (mine.goals || {})[month] || {}, theme: (mine.themes || {})[month] || '',
      list: people.length, late: late, back: back, last: dates.length ? dates[dates.length - 1] : '' };
  });
}
