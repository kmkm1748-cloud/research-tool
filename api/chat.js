module.exports = async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }
  try {
    const body = typeof req.body === 'string' ? JSON.parse(req.body) : req.body;

    const requestBody = {
      ...body,
      model: "claude-haiku-4-5-20251001",
    };

    async function callAPI(reqBody) {
      let response, data;
      for (let attempt = 0; attempt < 3; attempt++) {
        response = await fetch('https://api.anthropic.com/v1/messages', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'x-api-key': process.env.ANTHROPIC_API_KEY,
            'anthropic-version': '2023-06-01',
          },
          body: JSON.stringify(reqBody),
        });
        data = await response.json();
        if (response.status !== 429) break;
        await new Promise(r => setTimeout(r, (attempt + 1) * 2000));
      }
      return { response, data };
    }

    let { response, data } = await callAPI(requestBody);

    // web_searchのツール呼び出しループ処理
    let loopCount = 0;
    while (data.stop_reason === 'tool_use' && loopCount < 5) {
      loopCount++;

      const toolUseBlocks = data.content.filter(b => b.type === 'tool_use');
      if (toolUseBlocks.length === 0) break;

      const toolResults = toolUseBlocks.map(block => ({
        type: 'tool_result',
        tool_use_id: block.id,
        content: [],
      }));

      const nextMessages = [
        ...(requestBody.messages || []),
        { role: 'assistant', content: data.content },
        { role: 'user', content: toolResults },
      ];

      const nextReqBody = {
        ...requestBody,
        messages: nextMessages,
      };

      const result = await callAPI(nextReqBody);
      response = result.response;
      data = result.data;

      if (response.status !== 200) break;
    }

    if (data.content) {
      data.content = data.content.map(block => {
        if (block.type === 'text' && block.text) {
          block.text = block.text.replace(/<\/?cite[^>]*>/gi, '');
        }
        return block;
      });
    }

    res.status(response.status).json(data);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
};
