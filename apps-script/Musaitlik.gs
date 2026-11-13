/**
 * BNC Sapanca — web sitesi müsaitlik + talep servisi
 *
 * 1) Sitedeki takvime YALNIZCA dolu gecelerin tarihlerini verir.
 *    Bir gece DOLU sayılır: ay sekmesinde o günün ADI-SOYADI / TELEFON / KİŞİ SAYISI / KAPORA
 *    hücrelerinden biri doluysa veya boyanmışsa, YA DA "Web Talepleri" sekmesinde ONAY kutusu
 *    işaretli bir talebin geceleriyse.
 * 2) Sitede gönderilen rezervasyon talepleri "Web Talepleri" sekmesine yeni satır olarak düşer.
 *    Talep takvimi kapatmaz; ONAY kutusunu işaretleyince geceler kapanır.
 *
 * Kurulum: kurTetikleyici ve anahtarOlustur fonksiyonlarını BİR KEZ çalıştırın,
 * sonra Dağıt → Dağıtımları yönet → Yeni sürüm.
 */

const BUNGALOVLAR = { 'BNC-01': '1', 'BNC-02': '2' };
const AYLAR = { 'ocak': 1, 'şubat': 2, 'subat': 2, 'mart': 3, 'nisan': 4, 'mayıs': 5, 'mayis': 5, 'haziran': 6,
  'temmuz': 7, 'ağustos': 8, 'agustos': 8, 'eylül': 9, 'eylul': 9, 'ekim': 10, 'kasım': 11, 'kasim': 11, 'aralık': 12, 'aralik': 12 };
const DOLULUK_SATIRLARI = ['ADI-SOYADI', 'ADI SOYADI', 'TELEFON', 'KAPORA', 'KİŞİ SAYISI'];
const BOS_RENKLER = ['#ffffff', '#fff', 'white', ''];
const TALEP_SEKME = 'Web Talepleri';
const TALEP_BASLIK = ['Geliş zamanı', 'Bungalov', 'Giriş', 'Çıkış', 'Gece', 'Varış', 'Ayrılış', 'Misafir', 'Ad soyad', 'Telefon', 'E-posta', 'Not', 'ONAY'];

function doGet() {
  const cache = CacheService.getScriptCache();
  let json = cache.get('musaitlik');
  if (!json) json = yenile_();
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

// Hesapla ve önbelleğe koy. Dakikada bir otomatik, ayrıca tabloda her değişiklikte yeniden hesaplanır.
function yenile_() {
  const json = JSON.stringify(doluGeceler_());
  CacheService.getScriptCache().put('musaitlik', json, 600);
  return json;
}

function doluGeceler_() {
  const sonuc = { '1': [], '2': [] };
  const bugun = new Date(); bugun.setHours(0, 0, 0, 0);
  SpreadsheetApp.getActiveSpreadsheet().getSheets().forEach(sh => {
    if (gecmisSekme_(sh.getName()) || sh.getName() === TALEP_SEKME) return;
    const rng = sh.getDataRange();
    const deger = rng.getDisplayValues();
    const ham = rng.getValues();
    const renk = rng.getBackgrounds();
    const birlesik = sh.getRange(1, 1, Math.max(1, sh.getLastRow()), 1).getMergedRanges();

    for (let r = 0; r < deger.length; r++) {
      // "TARİH" başlık satırını bul (genelde B sütunu)
      const tarihSutun = deger[r].findIndex(v => normal_(v) === 'TARİH' || normal_(v) === 'TARIH');
      if (tarihSutun < 0) continue;

      // Bloğun bungalov adı: A sütununda, bu satırdan sonraki ilk dolu hücre (birleşik hücre dahil)
      let kod = null, blokSon = r;
      for (let k = r; k < Math.min(deger.length, r + 12); k++) {
        if (k > r && deger[k].some(v => normal_(v) === 'TARİH' || normal_(v) === 'TARIH')) break;
        blokSon = k;
        const a = String(deger[k][0] || '').trim().toUpperCase();
        if (a && BUNGALOVLAR[a]) kod = BUNGALOVLAR[a];
      }
      if (!kod) {
        birlesik.forEach(m => {
          const ust = m.getRow() - 1, alt = ust + m.getNumRows() - 1;
          if (ust <= r && alt >= r) {
            const a = String(m.getDisplayValue() || '').trim().toUpperCase();
            if (BUNGALOVLAR[a]) kod = BUNGALOVLAR[a];
          }
        });
      }
      if (!kod) continue;

      // Doluluğu gösteren satırlar
      const satirlar = [];
      for (let k = r + 1; k <= blokSon; k++) {
        const ad = normal_(deger[k][tarihSutun]);
        if (DOLULUK_SATIRLARI.indexOf(ad) >= 0) satirlar.push(k);
      }

      for (let c = tarihSutun + 1; c < deger[r].length; c++) {
        const tarih = (ham[r][c] instanceof Date) ? gun_(ham[r][c]) : tarihCoz_(deger[r][c]);
        if (!tarih || tarih < bugun) continue;
        const yazili = satirlar.some(k => String(deger[k][c] || '').trim() !== '');
        const boyali = satirlar.some(k => BOS_RENKLER.indexOf(String(renk[k][c] || '').toLowerCase()) < 0);
        if (yazili || boyali) {
          const iso = Utilities.formatDate(tarih, 'Europe/Istanbul', 'yyyy-MM-dd');
          if (sonuc[kod].indexOf(iso) < 0) sonuc[kod].push(iso);
        }
      }
      r = blokSon;
    }
  });
  onayliTalepler_(sonuc, bugun);
  sonuc['1'].sort(); sonuc['2'].sort();
  sonuc.updated = new Date().toISOString();
  return sonuc;
}

// "6 Kasım 2026 Cuma", "06.11.2026", "6/11/2026" → Date
function tarihCoz_(metin) {
  const t = String(metin || '').trim().toLocaleLowerCase('tr-TR');
  let m = t.match(/^(\d{1,2})\s+([a-zçğıöşü]+)\s+(\d{4})/);
  if (m && AYLAR[m[2]]) return new Date(Number(m[3]), AYLAR[m[2]] - 1, Number(m[1]));
  m = t.match(/^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{4})/);
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  return null;
}
// "EKİM 2026" gibi geçmiş ay sekmelerini okumaya gerek yok
function gecmisSekme_(ad) {
  const m = String(ad || '').trim().toLocaleLowerCase('tr-TR').match(/^([a-zçğıöşü]+)\s+(\d{4})$/);
  if (!m || !AYLAR[m[1]]) return false;
  const simdi = new Date();
  return Number(m[2]) * 12 + AYLAR[m[1]] < simdi.getFullYear() * 12 + simdi.getMonth() + 1;
}
function gun_(d) { const x = new Date(d); x.setHours(0, 0, 0, 0); return x; }
function normal_(v) { return String(v || '').trim().toLocaleUpperCase('tr-TR'); }

// ---------- Sitedeki talep formu ----------
// Site (Vercel /api/talep) buraya gizli anahtarla POST eder; talep "Web Talepleri" sekmesine yazılır.
function doPost(e) {
  const cevap = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
  let v;
  try { v = JSON.parse(e.postData.contents); } catch (err) { return cevap({ ok: false }); }
  const anahtar = PropertiesService.getScriptProperties().getProperty('TALEP_ANAHTAR');
  if (!anahtar || v.anahtar !== anahtar) return cevap({ ok: false });
  const kilit = LockService.getScriptLock();
  kilit.waitLock(20000);
  try {
    const sh = talepSekmesi_();
    sh.appendRow([new Date(), v.bungalov, v.giris, v.cikis, v.gece, v.varis, v.ayrilis, v.misafir, v.ad, v.telefon, v.eposta, v.not, false]);
    const r = sh.getLastRow();
    sh.getRange(r, 3, 1, 2).setNumberFormat('dd.mm.yyyy');
    sh.getRange(r, 13).insertCheckboxes();
  } finally { kilit.releaseLock(); }
  return cevap({ ok: true });
}

function talepSekmesi_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(TALEP_SEKME);
  if (!sh) {
    sh = ss.insertSheet(TALEP_SEKME, 0);
    sh.getRange(1, 1, 1, TALEP_BASLIK.length).setValues([TALEP_BASLIK]).setFontWeight('bold').setBackground('#0B3D3B').setFontColor('#E6FFFB');
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('dd.mm.yyyy hh:mm');
    sh.getRange('B2:B').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['No:1', 'No:2', 'Fark etmez'], true).build());
    sh.setColumnWidths(1, TALEP_BASLIK.length, 120);
    sh.getRange(1, 13).setNote('İşaretleyince bu talebin geceleri sitede DOLU görünür. Bungalov "Fark etmez" ise önce No:1 veya No:2 seçin.');
  }
  return sh;
}

// ONAY kutusu işaretli talepleri dolu gecelere ekler (giriş dahil, çıkış hariç)
function onayliTalepler_(sonuc, bugun) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TALEP_SEKME);
  if (!sh || sh.getLastRow() < 2) return;
  sh.getRange(2, 1, sh.getLastRow() - 1, TALEP_BASLIK.length).getValues().forEach(r => {
    if (r[12] !== true) return;
    const kod = ({ 'No:1': '1', 'No:2': '2' })[String(r[1]).trim()];
    if (!kod) return;
    const g = r[2] instanceof Date ? gun_(r[2]) : isoTarih_(r[2]), c = r[3] instanceof Date ? gun_(r[3]) : isoTarih_(r[3]);
    if (!g || !c) return;
    for (let d = new Date(g); d < c; d.setDate(d.getDate() + 1)) {
      if (d < bugun) continue;
      const iso = Utilities.formatDate(d, 'Europe/Istanbul', 'yyyy-MM-dd');
      if (sonuc[kod].indexOf(iso) < 0) sonuc[kod].push(iso);
    }
  });
}
function isoTarih_(v) {
  const m = String(v || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : tarihCoz_(v);
}

// BİR KEZ çalıştırın: siteyle paylaşılacak gizli anahtarı üretir. Günlükteki anahtarı
// Vercel'de TALEP_ANAHTAR ortam değişkenine yapıştırın.
function anahtarOlustur() {
  const pr = PropertiesService.getScriptProperties();
  let k = pr.getProperty('TALEP_ANAHTAR');
  if (!k) { k = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, ''); pr.setProperty('TALEP_ANAHTAR', k); }
  talepSekmesi_();
  Logger.log('TALEP_ANAHTAR = ' + k);
}

// Kontrol: Çalıştır → Yürütme günlüğünde dolu geceler görünür
function testEt() {
  const d = doluGeceler_();
  Logger.log('BNC-01 dolu geceler (' + d['1'].length + '): ' + d['1'].join(', '));
  Logger.log('BNC-02 dolu geceler (' + d['2'].length + '): ' + d['2'].join(', '));
}

// Tabloda değişiklik yapınca önbelleği hemen temizle (site en geç ~30 sn içinde günceller).
function onEdit() { yenile_(); }
// Biçim/renk değişikliklerinde de önbelleği temizlemek için bir kez çalıştırın: kurTetikleyici
function onChangeTemizle() { yenile_(); }
function zamanliYenile() { yenile_(); }
// BİR KEZ çalıştırın: renk/biçim değişikliğinde ve her dakika otomatik yeniler
function kurTetikleyici() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (['onChangeTemizle', 'zamanliYenile'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onChangeTemizle').forSpreadsheet(SpreadsheetApp.getActive()).onChange().create();
  ScriptApp.newTrigger('zamanliYenile').timeBased().everyMinutes(1).create();
  yenile_();
}
