const data = require('../data.json');

const instructions = `You are Vinar Travel & Tourism's official English virtual assistant for customers in Sudan.
Use only the supplied Vinar services and FAQ data. Never invent prices, dates, availability, offers, or policies.
Keep replies warm, concise, and suitable for chat. Include service name, USD price, duration, availability, slots, and offers when relevant.
If slots_this_week is 0, say it is fully booked this week and offer to check another week or an alternative.
Never guarantee visa approval. Explain that embassies make the final decision.
Never process payment. For booking confirmation, payment, complaints, refunds in dispute, missing information, or a request for a human, refer the customer to +249 914 101 013 / 012.
Answer in clear natural English.`;

async function getBody(request) {
  if (request.body && typeof request.body === 'object') return request.body;
  if (typeof request.body === 'string') return JSON.parse(request.body || '{}');
  const chunks = [];
  for await (const chunk of request) chunks.push(chunk);
  return JSON.parse(Buffer.concat(chunks).toString() || '{}');
}

module.exports = async function handler(request, response) {
  if (request.method !== 'POST') return response.status(405).json({ error: 'Method not allowed' });
  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) return response.status(503).json({ error: 'Gemini is not configured' });
  const body = await getBody(request);
  const question = String(body.question || '').trim();
  if (!question) return response.status(400).json({ error: 'Question is required' });

  const prompt = `${instructions}\n\nVinar data:\n${JSON.stringify(data)}\n\nCustomer question:\n${question}`;
  const result = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${encodeURIComponent(apiKey)}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] })
  });
  if (!result.ok) return response.status(502).json({ error: 'Gemini request failed' });
  const payload = await result.json();
  const text = payload.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!text) return response.status(502).json({ error: 'Gemini returned no answer' });
  return response.status(200).json({ text });
}
