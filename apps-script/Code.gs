// OAB seçim sayımı: canlı paylaşım sunucusu (Google Apps Script).
// Bu kodu bir Google E-Tablonun Apps Script editörüne yapıştırın.
// PIN'i kendinize göre değiştirin. Sitedeki PIN ile aynı olmalıdır.
var PIN = 'DEGISTIR';

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

function sheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function readState_() {
  var v = sheet_('Durum').getRange('A1').getValue();
  return v ? JSON.parse(v) : null;
}

function doGet(e) {
  var p = e.parameter || {};
  if (p.pin !== PIN) return out_({ ok: false, error: 'pin' });
  var st = readState_();
  if (!st) return out_({ ok: true, rev: 0, state: null });
  var since = parseInt(p.since, 10);
  if (since >= 0 && st.rev <= since) return out_({ ok: true, rev: st.rev });
  return out_({ ok: true, rev: st.rev, state: st });
}

function doPost(e) {
  var p = e.parameter || {};
  if (p.pin !== PIN) return out_({ ok: false, error: 'pin' });
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var body = JSON.parse(e.postData.contents);
    var st = body.state;
    if (!st || !st.names || !st.votes) return out_({ ok: false, error: 'bad' });
    var cur = readState_();
    if (cur && (st.rev | 0) <= (cur.rev | 0)) {
      return out_({ ok: false, error: 'stale', rev: cur.rev, state: cur });
    }
    sheet_('Durum').getRange('A1').setValue(JSON.stringify(st));
    record_(st);
    return out_({ ok: true, rev: st.rev });
  } finally {
    lock.releaseLock();
  }
}

// Sonuçları okunabilir bir tablo olarak "Sonuç" sayfasına yazar.
function record_(st) {
  var sh = sheet_('Sonuç');
  var rows = [['Aday No', 'Aday', 'Oy']];
  var total = 0;
  for (var i = 0; i < st.names.length; i++) {
    rows.push([i + 1, st.names[i], st.votes[i]]);
    total += st.votes[i];
  }
  sh.clearContents();
  sh.getRange(1, 1, rows.length, 3).setValues(rows);
  sh.getRange(1, 5, 5, 2).setValues([
    ['Geçerli pusula', st.ballots],
    ['Geçersiz pusula', st.invalid],
    ['Toplam oy', total],
    ['Durum', st.done ? 'Sayım tamamlandı' : 'Sayım sürüyor'],
    ['Son güncelleme', new Date()]
  ]);
}
