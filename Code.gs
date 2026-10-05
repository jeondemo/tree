var VERSION = '2026-10-06';
var DEFAULT_MODEL = 'gemini-flash-latest';
var FOLDER_NAME = '과수 방제 라벨 사진';
var API_BASE = 'https://generativelanguage.googleapis.com/v1beta/';

var TABLES = {
  trees: { sheet: '나무', cols: [['name', '이름'], ['aliases', '라벨에 쓰는 이름']] },
  products: { sheet: '약', cols: [['name', '약 이름'], ['kind', '종류'], ['ingredient', '성분'], ['content', '함량'], ['form', '제형'], ['group', '계통'], ['unit', '단위'], ['gap', '다른 약과 간격(일)'], ['match', '같은 약으로 보는 낱말'], ['note', '주의'], ['source', '출처']] },
  doses: { sheet: '용량', cols: [['id', 'id'], ['product', '약 이름'], ['tree', '나무'], ['target', '병해충'], ['dilution', '희석배수'], ['doseText', '용량(글)'], ['phi', '수확 전 일수'], ['phiNote', '시기 메모'], ['max', '연 횟수'], ['src', '근거'], ['checked', '확인한 날'], ['was', '이전 값']] },
  mix: { sheet: '혼용', cols: [['id', 'id'], ['a', '약 1'], ['b', '약 2'], ['verdict', '판정'], ['note', '메모']] },
  plan: { sheet: '계획', cols: [['id', 'id'], ['tree', '나무'], ['kind', '종류'], ['product', '약 이름'], ['text', '할 일'], ['note', '메모'], ['s', '시작(월-일)'], ['e', '끝(월-일)'], ['label', '기간 표시'], ['seq', '차수'], ['gap', '간격(일)'], ['routine', '갈 때마다'], ['post', '수확 뒤']] },
  kit: { sheet: '내약', cols: [['id', 'id'], ['name', '약 이름'], ['ingredient', '성분'], ['content', '함량'], ['form', '제형'], ['unit', '단위'], ['covers', '대신하는 약'], ['status', '상태'], ['checked', '확인한 날'], ['label', '라벨 메모'], ['added', '올린 날']] },
  photos: { sheet: '사진', cols: [['id', 'id'], ['kitId', '약 id'], ['fileId', '드라이브 파일'], ['thumb', '작은 그림'], ['added', '올린 날']] },
  logs: { sheet: '기록', cols: [['id', 'id'], ['date', '뿌린 날'], ['tree', '나무'], ['product', '약 이름'], ['base', '계획의 약'], ['water', '물(L)'], ['key', '계획 id'], ['memo', '메모'], ['created', '적은 때']] },
  checks: { sheet: '체크', cols: [['id', 'id'], ['date', '한 날']] }
};

function doGet() {
  return ContentService.createTextOutput('과수 방제 앱 API가 동작 중입니다. 버전 ' + VERSION).setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  var out;
  try {
    var req = JSON.parse(e.postData.contents);
    checkPin_(req.pin);
    if (req.action === 'load') out = load_();
    else if (req.action === 'ops') out = ops_(req.ops);
    else if (req.action === 'readLabel') out = readLabel_(req);
    else if (req.action === 'delPhotos') out = delPhotos_(req.fileIds);
    else if (req.action === 'setSetting') out = setSetting_(req.key, req.value);
    else throw new Error('모르는 요청: ' + req.action);
    out.ok = true;
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err), code: (err && err.code) || '' };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function prop_(k) {
  return PropertiesService.getScriptProperties().getProperty(k) || '';
}

function fail_(code, message) {
  var e = new Error(message);
  e.code = code;
  return e;
}

function checkPin_(pin) {
  var want = prop_('APP_PIN');
  if (want && String(pin || '') !== want) throw fail_('PIN', 'PIN이 맞지 않습니다.');
}

function ss_() {
  var id = prop_('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw fail_('SETUP', 'setup을 먼저 실행하세요.');
  return ss;
}

function cell_(v) {
  if (v === null || v === undefined || v === false) return '';
  if (v === true) return 'TRUE';
  if (Array.isArray(v)) return v.join(',');
  if (typeof v === 'object') return JSON.stringify(v);
  return String(v);
}

function searchOn_() {
  return prop_('SEARCH') !== 'off';
}

function status_() {
  return {
    version: VERSION,
    hasKey: !!prop_('GEMINI_API_KEY'),
    model: prop_('GEMINI_MODEL') || DEFAULT_MODEL,
    search: searchOn_(),
    pinSet: !!prop_('APP_PIN')
  };
}

function load_() {
  var ss = ss_();
  var tables = {};
  var missing = [];
  Object.keys(TABLES).forEach(function (k) {
    var def = TABLES[k];
    var sh = ss.getSheetByName(def.sheet);
    if (!sh) { tables[k] = []; missing.push(def.sheet); return; }
    var values = sh.getDataRange().getDisplayValues();
    var head = values.length ? values[0] : [];
    var idx = {};
    def.cols.forEach(function (c) {
      var i = head.indexOf(c[1]);
      if (i < 0) i = head.indexOf(c[0]);
      if (i >= 0) idx[c[0]] = i;
    });
    var rows = [];
    for (var r = 1; r < values.length; r++) {
      var row = {};
      var any = false;
      def.cols.forEach(function (c) {
        var v = idx[c[0]] === undefined ? '' : values[r][idx[c[0]]];
        row[c[0]] = v;
        if (v !== '') any = true;
      });
      if (any) rows.push(row);
    }
    tables[k] = rows;
  });
  return { tables: tables, status: status_(), missing: missing };
}

function openTable_(ss, def) {
  var sh = ss.getSheetByName(def.sheet);
  if (!sh) throw fail_('SETUP', def.sheet + ' 시트가 없습니다. setup을 먼저 실행하세요.');
  var width = Math.max(sh.getLastColumn(), def.cols.length);
  var head = sh.getRange(1, 1, 1, width).getDisplayValues()[0];
  var colOf = {};
  def.cols.forEach(function (c) {
    var i = head.indexOf(c[1]);
    if (i < 0) i = head.indexOf(c[0]);
    if (i >= 0) colOf[c[0]] = i + 1;
  });
  var keyCol = colOf[def.cols[0][0]];
  if (!keyCol) throw fail_('SETUP', def.sheet + ' 시트의 첫 줄 제목이 바뀌었습니다.');
  var last = sh.getLastRow();
  var keys = [];
  if (last > 1) {
    sh.getRange(2, keyCol, last - 1, 1).getDisplayValues().forEach(function (r) { keys.push(String(r[0])); });
  }
  return { sh: sh, def: def, colOf: colOf, keys: keys, width: width };
}

function putRow_(st, row) {
  var keyName = st.def.cols[0][0];
  var id = String(row[keyName]);
  if (!id || id === 'undefined') throw new Error(st.def.sheet + ': id가 없는 줄은 저장할 수 없습니다.');
  var at = st.keys.indexOf(id);
  var rowNum, vals;
  if (at >= 0) {
    rowNum = at + 2;
    vals = st.sh.getRange(rowNum, 1, 1, st.width).getDisplayValues()[0];
  } else {
    rowNum = st.keys.length + 2;
    vals = [];
    for (var i = 0; i < st.width; i++) vals.push('');
    st.keys.push(id);
  }
  st.def.cols.forEach(function (c) {
    if (st.colOf[c[0]] && row[c[0]] !== undefined) vals[st.colOf[c[0]] - 1] = cell_(row[c[0]]);
  });
  st.sh.getRange(rowNum, 1, 1, st.width).setNumberFormat('@').setValues([vals]);
}

function delRow_(st, id) {
  var at = st.keys.indexOf(String(id));
  if (at < 0) return;
  st.sh.deleteRow(at + 2);
  st.keys.splice(at, 1);
}

function ops_(ops) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = ss_();
    var open = {};
    (ops || []).forEach(function (op) {
      var def = TABLES[op.table];
      if (!def) throw new Error('모르는 표: ' + op.table);
      var st = open[op.table] || (open[op.table] = openTable_(ss, def));
      if (op.type === 'del') delRow_(st, op.id);
      else putRow_(st, op.row || {});
    });
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return { count: (ops || []).length };
}

function setSetting_(key, value) {
  if (key !== 'SEARCH') throw new Error('바꿀 수 없는 설정입니다.');
  PropertiesService.getScriptProperties().setProperty('SEARCH', value === 'off' ? 'off' : 'on');
  return { status: status_() };
}

function folder_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('FOLDER_ID');
  var folder = null;
  if (id) {
    try {
      folder = DriveApp.getFolderById(id);
      if (folder.isTrashed()) folder = null;
    } catch (e) {
      folder = null;
    }
  }
  if (!folder) {
    folder = DriveApp.createFolder(FOLDER_NAME);
    props.setProperty('FOLDER_ID', folder.getId());
  }
  return folder;
}

function inFolder_(file, folderId) {
  var parents = file.getParents();
  while (parents.hasNext()) {
    if (parents.next().getId() === folderId) return true;
  }
  return false;
}

function delPhotos_(fileIds) {
  var folderId = folder_().getId();
  var n = 0;
  (fileIds || []).forEach(function (id) {
    try {
      var f = DriveApp.getFileById(id);
      if (inFolder_(f, folderId)) { f.setTrashed(true); n++; }
    } catch (e) {}
  });
  return { trashed: n };
}

function readLabel_(req) {
  var photos = req.photos || [];
  if (!photos.length) throw new Error('사진이 없습니다.');
  if (photos.length > 8) throw new Error('한 번에 8장까지 올릴 수 있습니다.');
  var out = { photos: [], label: null, search: null, sources: [], queries: [], model: '', warnings: [], readError: '' };
  var folder = folder_();
  var folderId = folder.getId();
  var stamp = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMdd_HHmmss');
  var parts = [];
  photos.forEach(function (p, i) {
    var mime = p.mime || 'image/jpeg';
    var data = p.data || '';
    var fileId = p.fileId || '';
    if (fileId && !data) {
      var old = DriveApp.getFileById(fileId);
      if (!inFolder_(old, folderId)) throw new Error('라벨 폴더 밖의 파일은 읽지 않습니다.');
      var b = old.getBlob();
      mime = b.getContentType() || mime;
      data = Utilities.base64Encode(b.getBytes());
    } else if (data) {
      var blob = Utilities.newBlob(Utilities.base64Decode(data), mime, '라벨_' + stamp + '_' + (i + 1) + '.jpg');
      fileId = folder.createFile(blob).getId();
    } else {
      throw new Error('사진 내용이 비어 있습니다.');
    }
    out.photos.push({ fileId: fileId });
    parts.push({ inlineData: { mimeType: mime, data: data } });
  });
  if (!prop_('GEMINI_API_KEY')) {
    out.warnings.push('NO_KEY');
    out.readError = 'Gemini 키가 아직 없습니다. 사진은 저장했습니다. 값을 직접 적어 넣을 수 있습니다.';
    return out;
  }
  var crops = (req.crops && req.crops.length) ? req.crops : ['복숭아', '체리(양앵두)', '자두', '감', '대추', '딸기', '배롱나무'];
  try {
    var r = geminiAuto_({
      contents: [{ role: 'user', parts: [{ text: labelPrompt_(crops) }].concat(parts) }],
      generationConfig: { responseMimeType: 'application/json' }
    });
    out.model = r.model;
    out.label = extractJson_(textOf_(r.json));
    if (!out.label) out.readError = '라벨 글자를 읽지 못했습니다. 글자가 또렷하게 나온 사진으로 다시 올려 주세요.';
  } catch (err) {
    out.readError = '라벨 읽기 실패: ' + String((err && err.message) || err);
    return out;
  }
  if (out.label && out.label.name && searchOn_() && req.search !== false) {
    try {
      var s = searchLabel_(out.label, crops);
      out.search = s.data;
      out.sources = s.sources;
      out.queries = s.queries;
      if (s.warning) out.warnings.push(s.warning);
    } catch (err2) {
      out.warnings.push('SEARCH_FAIL: ' + String((err2 && err2.message) || err2));
    }
  } else if (!searchOn_()) {
    out.warnings.push('SEARCH_OFF');
  }
  return out;
}

function labelPrompt_(crops) {
  return [
    '너는 한국 농약 포장지(라벨) 사진을 그대로 옮겨 적는 일을 한다.',
    '첨부한 사진들은 같은 약 한 가지의 포장지다. 사진에 실제로 보이는 글자만 옮겨라.',
    '알고 있는 지식으로 빈칸을 채우지 마라. 보이지 않거나 흐려서 확실하지 않은 값은 null로 둔다.',
    '',
    '적용 병해충과 사용량 표에서는 다음 작물의 줄만 옮긴다: ' + crops.join(', ') + '.',
    '표에 이 작물이 없으면 rows는 빈 배열로 둔다. 다른 작물의 값을 대신 넣지 마라.',
    '',
    '아래 모양의 JSON 하나만 출력한다.',
    '{',
    '  "name": "상표명 (예: 모스피란)",',
    '  "maker": "제조사 또는 판매사",',
    '  "ingredient": "유효성분의 한글 일반명만. 제형이나 함량은 빼고 (예: 아세타미프리드)",',
    '  "content": "유효성분 함량 (예: 8%)",',
    '  "form": "제형 (예: 수화제, 액상수화제, 입상수화제, 유제)",',
    '  "kind": "살균, 살충, 살비, 제초, 기타 중 하나",',
    '  "group": "작용기작 표시 기호 (예: 4a, 사1, 다5). 없으면 빈 문자열",',
    '  "rows": [',
    '    {',
    '      "crop": "라벨에 적힌 작물 이름 그대로",',
    '      "target": "적용 병해충",',
    '      "doseText": "라벨에 적힌 사용량 그대로 (예: 2,000배, 물 20L당 10g)",',
    '      "dilution": "희석배수를 숫자로. 물 20L당 약량만 있으면 20000을 약량으로 나눈 값. 범위(50~100배)거나 알 수 없으면 null",',
    '      "per20L": "물 20L당 약량과 단위 (예: 10g, 10ml). 없으면 빈 문자열",',
    '      "phi": "안전사용기준의 수확 전 일수를 숫자로 (수확 21일 전까지면 21). 날짜가 아니면 null",',
    '      "phiText": "사용 시기 원문 (예: 수확 21일 전까지, 개화 전, 발아 전)",',
    '      "max": "사용 횟수 제한을 숫자로 (3회 이내면 3). 없으면 null"',
    '    }',
    '  ],',
    '  "cautions": ["혼용, 약해, 다른 약과의 간격처럼 뿌릴 때 지켜야 할 주의 문장을 짧게. 사진에 있는 것만"],',
    '  "unreadable": ["읽지 못했거나 잘려서 안 보이는 부분에 대한 짧은 메모"]',
    '}',
    'dilution, phi, max는 숫자나 null이어야 한다. 따옴표로 감싼 글자로 쓰지 마라.'
  ].join('\n');
}

function searchLabel_(label, crops) {
  var who = [label.name, label.ingredient, label.content, label.form, label.maker].filter(function (x) { return x; }).join(' ');
  var prompt = [
    '한국에서 판매하는 농약 "' + who + '"의 등록 사용기준을 검색해서 확인해라.',
    '제조사 제품 페이지, 농촌진흥청 농약안전정보시스템(psis.rda.go.kr), 농사로를 먼저 본다. 블로그나 쇼핑몰 글은 근거로 삼지 않는다.',
    '다음 작물만 찾는다: ' + crops.join(', ') + '.',
    '검색한 페이지에 실제로 적힌 값만 쓴다. 찾지 못한 값은 null로 둔다. 추측하지 마라.',
    '',
    '마지막에 아래 모양의 JSON을 ```json 코드 블록 하나로 출력한다.',
    '{',
    '  "rows": [',
    '    { "crop": "작물", "target": "적용 병해충", "dilution": 2000, "phi": 21, "max": 3, "source": "값을 확인한 페이지의 사이트 이름" }',
    '  ],',
    '  "notes": "혼용이나 약해처럼 눈에 띄는 주의가 있으면 한두 문장"',
    '}',
    'dilution은 희석배수 숫자, phi는 수확 전 일수 숫자, max는 연 사용횟수 숫자다.'
  ].join('\n');
  var r = geminiAuto_({ contents: [{ role: 'user', parts: [{ text: prompt }] }], tools: [{ google_search: {} }] });
  var text = textOf_(r.json);
  var cand = (r.json.candidates && r.json.candidates[0]) || {};
  var gm = cand.groundingMetadata || {};
  var sources = [];
  (gm.groundingChunks || []).forEach(function (c) {
    if (c.web && c.web.uri) sources.push({ title: c.web.title || '', uri: c.web.uri });
  });
  var queries = gm.webSearchQueries || [];
  if (!sources.length) return { data: null, sources: [], queries: queries, warning: 'SEARCH_NO_SOURCE' };
  var data = extractJson_(text);
  if (!data && text) {
    var r2 = geminiAuto_({
      contents: [{ role: 'user', parts: [{ text: '다음 글에 적힌 농약 사용기준을 {"rows":[{"crop":"","target":"","dilution":0,"phi":0,"max":0,"source":""}],"notes":""} 모양의 JSON으로만 옮겨라. 글에 없는 값은 null로 둔다.\n\n' + text }] }],
      generationConfig: { responseMimeType: 'application/json' }
    });
    data = extractJson_(textOf_(r2.json));
  }
  return { data: data, sources: sources.slice(0, 8), queries: queries, warning: data ? '' : 'SEARCH_PARSE' };
}

function textOf_(json) {
  var cand = (json && json.candidates && json.candidates[0]) || null;
  if (!cand || !cand.content || !cand.content.parts) {
    var why = (cand && cand.finishReason) || (json && json.promptFeedback && json.promptFeedback.blockReason) || '';
    if (why) throw new Error('Gemini가 답을 내지 않았습니다 (' + why + ').');
    return '';
  }
  var out = '';
  cand.content.parts.forEach(function (p) {
    if (p.text && !p.thought) out += p.text;
  });
  return out;
}

function extractJson_(text) {
  var s = String(text || '');
  var fence = s.match(/```(?:json)?\s*([\s\S]*?)```/i);
  if (fence) s = fence[1];
  var a = s.indexOf('{');
  var b = s.lastIndexOf('}');
  if (a < 0 || b <= a) return null;
  try {
    return JSON.parse(s.slice(a, b + 1));
  } catch (e) {
    return null;
  }
}

function gemini_(body, model) {
  var key = prop_('GEMINI_API_KEY');
  if (!key) throw fail_('NO_KEY', 'Gemini 키가 없습니다.');
  var url = API_BASE + 'models/' + encodeURIComponent(model) + ':generateContent';
  var res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': key },
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });
  var code = res.getResponseCode();
  var text = res.getContentText();
  if (code === 200) return JSON.parse(text);
  var message = text.slice(0, 300);
  try {
    var j = JSON.parse(text);
    if (j.error && j.error.message) message = j.error.message;
  } catch (e) {}
  if (code === 404) throw fail_('MODEL', '모델을 찾을 수 없습니다 (' + model + ').');
  if (code === 400) throw fail_('BAD', 'Gemini 400: ' + message);
  if (code === 429 || code === 500 || code === 503) throw fail_('BUSY', 'Gemini ' + code + ': ' + message);
  throw fail_('HTTP', 'Gemini ' + code + ': ' + message);
}

function geminiTry_(body, model) {
  try {
    return gemini_(body, model);
  } catch (e) {
    if (e.code === 'BUSY') {
      Utilities.sleep(2500);
      return gemini_(body, model);
    }
    if (e.code === 'BAD' && body.generationConfig) {
      var plain = {};
      Object.keys(body).forEach(function (k) { if (k !== 'generationConfig') plain[k] = body[k]; });
      return gemini_(plain, model);
    }
    throw e;
  }
}

function geminiAuto_(body) {
  var props = PropertiesService.getScriptProperties();
  var model = props.getProperty('GEMINI_MODEL') || DEFAULT_MODEL;
  try {
    return { json: geminiTry_(body, model), model: model };
  } catch (e) {
    if (e.code !== 'MODEL') throw e;
    var next = pickModel_();
    if (!next || next === model) throw e;
    props.setProperty('GEMINI_MODEL', next);
    return { json: geminiTry_(body, next), model: next };
  }
}

function pickModel_() {
  var res = UrlFetchApp.fetch(API_BASE + 'models?pageSize=200', {
    method: 'get',
    headers: { 'x-goog-api-key': prop_('GEMINI_API_KEY') },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) return '';
  var models = (JSON.parse(res.getContentText()).models || []).filter(function (m) {
    return (m.supportedGenerationMethods || []).indexOf('generateContent') >= 0;
  }).map(function (m) { return String(m.name || '').replace(/^models\//, ''); });
  var best = '';
  var bestVer = -1;
  models.forEach(function (name) {
    var m = name.match(/^gemini-(\d+(?:\.\d+)?)-flash$/);
    if (m && parseFloat(m[1]) > bestVer) { bestVer = parseFloat(m[1]); best = name; }
  });
  if (best) return best;
  var loose = models.filter(function (name) {
    return /flash/.test(name) && !/(lite|image|tts|live|audio|embedding)/.test(name);
  });
  return loose.length ? loose[loose.length - 1] : '';
}

function setup() {
  if (typeof SEED === 'undefined') throw new Error('Seed.gs 파일이 없습니다. 파일을 하나 더 만들어 Seed.gs 내용을 붙여 넣은 뒤 다시 실행하세요.');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('구글 시트에서 확장 프로그램 > Apps Script로 연 편집기에서 실행하세요.');
  var props = PropertiesService.getScriptProperties();
  props.setProperty('SHEET_ID', ss.getId());
  var made = [];
  Object.keys(TABLES).forEach(function (k) {
    var def = TABLES[k];
    var sh = ss.getSheetByName(def.sheet);
    if (!sh) sh = ss.insertSheet(def.sheet);
    if (sh.getLastRow() > 0) return;
    fillSheet_(sh, def, SEED[k] || []);
    made.push(def.sheet + ' ' + (SEED[k] || []).length + '줄');
  });
  var folder = folder_();
  if (!props.getProperty('APP_PIN')) props.setProperty('APP_PIN', String(Math.floor(100000 + Math.random() * 900000)));
  if (!props.getProperty('SEARCH')) props.setProperty('SEARCH', 'on');
  Logger.log('새로 채운 시트: ' + (made.length ? made.join(', ') : '없음 (이미 있던 내용은 그대로 둠)'));
  Logger.log('라벨 사진 폴더: ' + folder.getUrl());
  Logger.log('앱 PIN: ' + props.getProperty('APP_PIN') + '  (휴대폰에서 처음 열 때 한 번 입력)');
  Logger.log('Gemini 키: ' + (props.getProperty('GEMINI_API_KEY') ? '들어 있음' : '아직 없음. 프로젝트 설정 > 스크립트 속성에 GEMINI_API_KEY 추가'));
}

function fillSheet_(sh, def, seedRows) {
  var n = def.cols.length;
  sh.clear();
  sh.getRange(1, 1, 1, n).setNumberFormat('@').setValues([def.cols.map(function (c) { return c[1]; })]).setFontWeight('bold');
  sh.setFrozenRows(1);
  var rows = seedRows.map(function (row) {
    return def.cols.map(function (c) { return cell_(row[c[0]]); });
  });
  if (rows.length) sh.getRange(2, 1, rows.length, n).setNumberFormat('@').setValues(rows);
}

function showInfo() {
  var props = PropertiesService.getScriptProperties();
  Logger.log('버전: ' + VERSION);
  Logger.log('앱 PIN: ' + (props.getProperty('APP_PIN') || '없음 (누구나 열 수 있음)'));
  Logger.log('Gemini 키: ' + (props.getProperty('GEMINI_API_KEY') ? '들어 있음' : '없음'));
  Logger.log('모델: ' + (props.getProperty('GEMINI_MODEL') || DEFAULT_MODEL));
  Logger.log('검색: ' + (searchOn_() ? '켜짐' : '꺼짐'));
  Logger.log('라벨 사진 폴더: ' + folder_().getUrl());
  Logger.log('웹 앱 주소: ' + (ScriptApp.getService().getUrl() || '아직 배포 전'));
}

function testGemini() {
  var r = geminiAuto_({ contents: [{ role: 'user', parts: [{ text: '한 낱말로만 답해라. 복숭아는 과일인가 채소인가?' }] }] });
  Logger.log('모델: ' + r.model);
  Logger.log('답: ' + textOf_(r.json));
}

function testSearch() {
  var s = searchLabel_({ name: '모스피란', ingredient: '아세타미프리드', content: '8%', form: '수화제' }, ['복숭아', '감']);
  Logger.log('검색어: ' + JSON.stringify(s.queries));
  Logger.log('근거 ' + s.sources.length + '곳: ' + JSON.stringify(s.sources.slice(0, 3)));
  Logger.log('값: ' + JSON.stringify(s.data));
  Logger.log('메모: ' + (s.warning || '없음'));
}

function resetBaseTables() {
  if (typeof SEED === 'undefined') throw new Error('Seed.gs 파일이 없습니다.');
  var ss = ss_();
  ['trees', 'products', 'doses', 'mix', 'plan'].forEach(function (k) {
    var def = TABLES[k];
    var sh = ss.getSheetByName(def.sheet) || ss.insertSheet(def.sheet);
    fillSheet_(sh, def, SEED[k] || []);
  });
  Logger.log('나무, 약, 용량, 혼용, 계획 시트를 처음 값으로 되돌렸습니다. 내약, 사진, 기록, 체크는 그대로입니다.');
}
