/**
 * BNC Sapanca — web sitesi müsaitlik servisi
 *
 * Bu kod, işletmenin rezervasyon tablosuna (Google E-Tablolar) eklenir ve
 * bncsapancabungalow.com takvimine YALNIZCA dolu gecelerin tarihlerini verir.
 * Müşteri adı, telefon, kapora gibi hiçbir kişisel bilgi dışarı çıkmaz.
 *
 * Tablo düzeni (ay sekmeleri, ör. "KASIM 2026"):
 *   - Her bungalov bir blok: üst satırda "TARİH" ve sağa doğru günler
 *     (ör. "6 Kasım 2026 Cuma"), altında ADI-SOYADI, TELEFON ... satırları.
 *   - Blok, A sütununda "BNC-01" / "BNC-02" yazan birleşik hücreyle adlanır.
 *   - Bir gece DOLU sayılır: o günün sütunu boyalıysa (beyaz değilse)
 *     VEYA ADI-SOYADI / TELEFON / KAPORA hücrelerinden biri doluysa.
 *
 * Kurulum (bir kez):
 *   1) Tabloda: Uzantılar → Apps Script. Bu dosyanın tamamını yapıştırın, kaydedin.
 *   2) Üstteki menüden "testEt" fonksiyonunu seçip ▶ Çalıştır'a basın, izin verin.
 *      Günlük kayıtlarda bulunan dolu geceler listelenir; kontrol edin.
 *   3) Dağıt → Yeni dağıtım → Tür: Web uygulaması
 *        Yürütme yetkisi: Ben   ·   Erişim: Herkes
 *      → Dağıt. Çıkan "Web uygulaması URL'sini" kopyalayın.
 *   4) Vercel → bnc-sapanca-site → Settings → Environment Variables:
 *        Name: MUSAITLIK_URL   Value: (kopyaladığınız URL)   → Save → Redeploy.
 */

const BUNGALOVLAR = { 'BNC-01': '1', 'BNC-02': '2' };
const AYLAR = { 'ocak': 1, 'şubat': 2, 'subat': 2, 'mart': 3, 'nisan': 4, 'mayıs': 5, 'mayis': 5, 'haziran': 6,
  'temmuz': 7, 'ağustos': 8, 'agustos': 8, 'eylül': 9, 'eylul': 9, 'ekim': 10, 'kasım': 11, 'kasim': 11, 'aralık': 12, 'aralik': 12 };
const DOLULUK_SATIRLARI = ['ADI-SOYADI', 'ADI SOYADI', 'TELEFON', 'KAPORA', 'KİŞİ SAYISI'];
const BOS_RENKLER = ['#ffffff', '#fff', 'white', ''];

function doGet() {
  const cache = CacheService.getScriptCache();
  let json = cache.get('musaitlik');
  if (!json) {
    json = JSON.stringify(doluGeceler_());
    cache.put('musaitlik', json, 30); // 30 saniye
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function doluGeceler_() {
  const sonuc = { '1': [], '2': [] };
  const bugun = new Date(); bugun.setHours(0, 0, 0, 0);
  SpreadsheetApp.getActiveSpreadsheet().getSheets().forEach(sh => {
    const rng = sh.getDataRange();
    const deger = rng.getDisplayValues();
    const ham = rng.getValues();
    const renk = rng.getBackgrounds();
    const birlesik = sh.getRange(1, 1, sh.getMaxRows(), 1).getMergedRanges();

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
function gun_(d) { const x = new Date(d); x.setHours(0, 0, 0, 0); return x; }
function normal_(v) { return String(v || '').trim().toLocaleUpperCase('tr-TR'); }

// Kurulumdan önce kontrol için: Çalıştır → Yürütme günlüğünde dolu geceler görünür.
function testEt() {
  const d = doluGeceler_();
  Logger.log('BNC-01 dolu geceler (' + d['1'].length + '): ' + d['1'].join(', '));
  Logger.log('BNC-02 dolu geceler (' + d['2'].length + '): ' + d['2'].join(', '));
}

// Tabloda değişiklik yapınca önbelleği hemen temizle (site en geç ~2 dk içinde günceller).
function onEdit() { CacheService.getScriptCache().remove('musaitlik'); }
// Biçim/renk değişikliklerinde de önbelleği temizlemek için bir kez çalıştırın: kurTetikleyici
function onChangeTemizle_() { CacheService.getScriptCache().remove('musaitlik'); }
function kurTetikleyici() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'onChangeTemizle_') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('onChangeTemizle_').forSpreadsheet(SpreadsheetApp.getActive()).onChange().create();
}
