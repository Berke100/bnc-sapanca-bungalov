// Sitedeki rezervasyon talebini işletmenin Google E-Tablosuna ("Web Talepleri" sekmesi) yazar.
// Apps Script'e giden istek gizli anahtarla (TALEP_ANAHTAR) imzalanır; tarayıcı bu anahtarı görmez.
const ISO = /^\d{4}-\d{2}-\d{2}$/;
const kirp = (v, n) => String(v == null ? "" : v).replace(/[\u0000-\u001f]/g, " ").trim().slice(0, n);
// tabloda formül çalıştırılmasın (= + - @ ile başlayan metinler)
const guvenli = v => (/^[=+\-@]/.test(v) ? "'" + v : v);

module.exports = async (req, res) => {
  res.setHeader("Content-Type", "application/json; charset=utf-8");
  res.setHeader("Cache-Control", "no-store");
  if (req.method !== "POST") return res.status(405).send('{"ok":false}');
  const url = process.env.MUSAITLIK_URL, anahtar = process.env.TALEP_ANAHTAR;
  if (!url || !anahtar) return res.status(200).send('{"ok":false,"reason":"not-configured"}');

  let b = req.body;
  if (typeof b === "string") { try { b = JSON.parse(b); } catch (e) { b = null; } }
  if (!b || typeof b !== "object") return res.status(400).send('{"ok":false}');
  if (b.website) return res.status(200).send('{"ok":true}');           // bot tuzağı
  const bugun = new Date(Date.now() + 3 * 3600e3).toISOString().slice(0, 10);
  if (!ISO.test(b.giris) || !ISO.test(b.cikis) || b.giris < bugun || b.cikis <= b.giris) return res.status(400).send('{"ok":false}');
  const gece = Math.round((Date.parse(b.cikis) - Date.parse(b.giris)) / 864e5);
  if (gece < 1 || gece > 60) return res.status(400).send('{"ok":false}');
  const ad = kirp(b.ad, 80), tel = kirp(b.telefon, 30);
  if (ad.length < 3 || tel.replace(/\D/g, "").length < 6) return res.status(400).send('{"ok":false}');

  const veri = {
    anahtar,
    bungalov: ({ "1": "No:1", "2": "No:2" })[b.bungalov] || "Fark etmez",
    giris: b.giris, cikis: b.cikis, gece,
    varis: kirp(b.varis, 5), ayrilis: kirp(b.ayrilis, 5),
    misafir: guvenli(kirp(b.misafir, 40)),
    ad: guvenli(ad), telefon: guvenli(tel),
    eposta: guvenli(kirp(b.eposta, 120)), not: guvenli(kirp(b.not, 500))
  };
  try {
    const ctl = new AbortController(); const t = setTimeout(() => ctl.abort(), 25000);
    const r = await fetch(url, { method: "POST", headers: { "Content-Type": "text/plain;charset=utf-8" }, body: JSON.stringify(veri), redirect: "follow", signal: ctl.signal });
    clearTimeout(t);
    const j = await r.json().catch(() => ({}));
    return res.status(200).send(JSON.stringify({ ok: !!j.ok }));
  } catch (e) {
    return res.status(502).send('{"ok":false}');
  }
};
