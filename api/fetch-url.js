module.exports = async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }
  try {
    const { url } = typeof req.body === 'string' ? JSON.parse(req.body) : req.body;
    if (!url) return res.status(400).json({ error: 'URL required' });

    const response = await fetch(url, {
      headers: {
        'User-Agent': 'Mozilla/5.0 (compatible; SBR-Research-Bot/1.0)',
        'Accept': 'text/html,application/xhtml+xml',
        'Accept-Language': 'ja,en-US;q=0.9',
      },
      signal: AbortSignal.timeout(10000),
    });

    if (!response.ok) {
      return res.status(200).json({ text: "", error: `HTTP ${response.status}` });
    }

    const html = await response.text();

    // HTMLからテキスト抽出
    const text = html
      .replace(/<script[\s\S]*?<\/script>/gi, '')
      .replace(/<style[\s\S]*?<\/style>/gi, '')
      .replace(/<[^>]+>/g, ' ')
      .replace(/&nbsp;/g, ' ')
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/\s+/g, ' ')
      .trim()
      .slice(0, 8000); // 最大8000文字

    res.status(200).json({ text, url });
  } catch (err) {
    res.status(200).json({ text: "", error: err.message });
  }
};
