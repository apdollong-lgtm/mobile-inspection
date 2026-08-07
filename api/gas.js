const GAS_URL = process.env.GAS_API_URL
  || 'https://script.google.com/macros/s/AKfycbyUf_8BaiMdxqwnukbAHrDEoDp268iPh6EnUt5qsMVbyaKRQuxR9gCJ8gWE1TngvFN8/exec';

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  try {
    let response;
    if (req.method === 'GET') {
      const qs = new URLSearchParams(req.query).toString();
      const url = qs ? `${GAS_URL}?${qs}` : GAS_URL;
      response = await fetch(url, { redirect: 'follow' });
    } else {
      response = await fetch(GAS_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(req.body || {}),
        redirect: 'follow',
      });
    }

    const text = await response.text();
    res.setHeader('Content-Type', 'application/json; charset=utf-8');
    return res.status(response.status).send(text);
  } catch (err) {
    return res.status(500).json({ ok: false, error: String(err.message || err) });
  }
};
