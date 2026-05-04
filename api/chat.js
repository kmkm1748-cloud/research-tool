module.exports = async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }
  try {
    const body = typeof req.body === 'string' ? JSON.parse(req.body) : req.body;
    
    let response, data;
    for (let attempt = 0; attempt < 3; attempt++) {
      response = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': process.env.ANTHROPIC_API_KEY,
          'anthropic-version': '2023-06-01',
        },
        body: JSON.stringify(body),
      });
      data = await response.json();
      if (response.status !== 429) break;
      await new Promise(r => setTimeout(r, (attempt + 1) * 2000));
    }
    res.status(response.status).json(data);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
};
