// Sitedeki takvim için dolu geceleri döndürür.
// Kaynak: işletmenin Google E-Tablosuna eklenen Apps Script web uygulaması (MUSAITLIK_URL ortam değişkeni).
// Yalnızca tarihler döner; müşteri adı/telefonu gibi kişisel veriler bu uçtan asla geçmez.
const ISO = /^\d{4}-\d{2}-\d{2}$/;
const temizle = a => (Array.isArray(a) ? a : []).filter(x => typeof x === "string" && ISO.test(x)).slice(0, 2000);

module.exports = async (req, res) => {
  res.setHeader("Content-Type", "application/json; charset=utf-8");
  const url = process.env.MUSAITLIK_URL;
  if (!url) {
    res.setHeader("Cache-Control", "no-store");
    return res.status(200).send(JSON.stringify({ ok: false, reason: "not-configured", "1": [], "2": [] }));
  }
  try {
    const ctl = new AbortController();
    const t = setTimeout(() => ctl.abort(), 8000);
    const r = await fetch(url, { redirect: "follow", signal: ctl.signal });
    clearTimeout(t);
    if (!r.ok) throw new Error("HTTP " + r.status);
    const d = await r.json();
    res.setHeader("Cache-Control", "public, s-maxage=30, stale-while-revalidate=30");
    return res.status(200).send(JSON.stringify({ ok: true, updated: d.updated || null, "1": temizle(d["1"]), "2": temizle(d["2"]) }));
  } catch (e) {
    res.setHeader("Cache-Control", "no-store");
    return res.status(502).send(JSON.stringify({ ok: false, reason: "upstream", "1": [], "2": [] }));
  }
};
