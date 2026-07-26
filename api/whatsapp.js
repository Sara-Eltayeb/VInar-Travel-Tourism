const data = require('../data.json');

const instructions = `You are Vinar Travel & Tourism's official English WhatsApp assistant for customers in Sudan.
Use only the supplied Vinar services and FAQ data. Never invent prices, dates, availability, offers, or policies.
Keep replies warm, concise, and suitable for WhatsApp. Include service name, USD price, duration, availability, slots, and offers when relevant.
If slots_this_week is 0, say it is fully booked this week and offer an alternative.
Never guarantee visa approval. Embassies make the final decision.
Never process payment. For booking confirmation, payment, complaints, refunds in dispute, missing information, or a human request, refer the customer to +249 914 101 013 / 012.
Answer in clear natural English.`;

async function readBody(request) {
  if (request.body && typeof request.body === 'object') return request.body;
  if (typeof request.body === 'string') return JSON.parse(request.body || '{}');
  const chunks = [];
  for await (const chunk of request) chunks.push(chunk);
  return JSON.parse(Buffer.concat(chunks).toString() || '{}');
}

async function createReply(question) {
  const result = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${encodeURIComponent(process.env.GEMINI_API_KEY)}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ contents: [{ parts: [{ text: `${instructions}\n\nVinar data:\n${JSON.stringify(data)}\n\nCustomer message:\n${question}` }] }] })
  });
  if (!result.ok) throw new Error('Gemini request failed');
  const payload = await result.json();
  return payload.candidates?.[0]?.content?.parts?.[0]?.text || 'Please contact a Vinar advisor at +249 914 101 013 / 012.';
}

async function sendWhatsAppMessage(to, body) {
  const version = process.env.META_GRAPH_VERSION || 'v21.0';
  const result = await fetch(`https://graph.facebook.com/${version}/${process.env.WHATSAPP_PHONE_NUMBER_ID}/messages`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${process.env.WHATSAPP_ACCESS_TOKEN}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ messaging_product: 'whatsapp', recipient_type: 'individual', to, type: 'text', text: { preview_url: false, body } })
  });
  if (!result.ok) throw new Error(`WhatsApp send failed: ${result.status}`);
}

module.exports = async function handler(request, response) {
  if (request.method === 'GET') {
    const query = request.query || {};
    if (query['hub.verify_token'] !== process.env.WHATSAPP_VERIFY_TOKEN) return response.status(403).send('Forbidden');
    return response.status(200).send(query['hub.challenge'] || '');
  }
  if (request.method !== 'POST') return response.status(405).send('Method not allowed');
  if (!process.env.GEMINI_API_KEY || !process.env.WHATSAPP_ACCESS_TOKEN || !process.env.WHATSAPP_PHONE_NUMBER_ID) return response.status(503).json({ error: 'WhatsApp connector is not configured' });

  const payload = await readBody(request);
  const messages = payload.entry?.flatMap(entry => entry.changes || []).flatMap(change => change.value?.messages || []) || [];
  await Promise.all(messages.filter(message => message.type === 'text' && message.from).map(async message => {
    const reply = await createReply(message.text.body);
    await sendWhatsAppMessage(message.from, reply);
  }));
  return response.status(200).json({ received: true });
};
