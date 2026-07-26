async function bodyOf(request) {
  if (request.body && typeof request.body === 'object') return request.body;
  if (typeof request.body === 'string') return JSON.parse(request.body || '{}');
  const chunks = [];
  for await (const chunk of request) chunks.push(chunk);
  return JSON.parse(Buffer.concat(chunks).toString() || '{}');
}

const tools = [
  { name: 'get_weather', description: 'Get current weather for a travel destination.', inputSchema: { type: 'object', properties: { destination: { type: 'string', description: 'City or destination name' } }, required: ['destination'] } },
  { name: 'convert_currency', description: 'Convert an amount from USD to another currency using a public live exchange rate.', inputSchema: { type: 'object', properties: { amount: { type: 'number' }, currency: { type: 'string', description: 'Three-letter currency code, such as SDG or EUR' } }, required: ['amount', 'currency'] } }
];

async function callTool(name, args) {
  if (name === 'convert_currency') {
    const currency = String(args.currency || 'SDG').toUpperCase();
    const amount = Number(args.amount) || 0;
    const response = await fetch('https://open.er-api.com/v6/latest/USD');
    const data = await response.json();
    const rate = data.rates?.[currency];
    if (!rate) throw new Error(`No exchange rate found for ${currency}`);
    const value = Number((amount * rate).toFixed(2));
    return { text: `${amount.toLocaleString()} USD is approximately ${value.toLocaleString()} ${currency}.`, structuredContent: { amount, currency, rate, value } };
  }
  if (name === 'get_weather') {
    const destination = String(args.destination || '').trim();
    const locationResponse = await fetch(`https://geocoding-api.open-meteo.com/v1/search?name=${encodeURIComponent(destination)}&count=1&language=en&format=json`);
    const location = (await locationResponse.json()).results?.[0];
    if (!location) throw new Error('Destination not found');
    const weatherResponse = await fetch(`https://api.open-meteo.com/v1/forecast?latitude=${location.latitude}&longitude=${location.longitude}&current=temperature_2m,weather_code,wind_speed_10m&timezone=auto`);
    const current = (await weatherResponse.json()).current;
    const result = { destination: location.name, country: location.country, temperature_c: current.temperature_2m, weather_code: current.weather_code, wind_kmh: current.wind_speed_10m };
    return { text: `${location.name}, ${location.country}: ${current.temperature_2m}°C, wind ${current.wind_speed_10m} km/h.`, structuredContent: result };
  }
  throw new Error(`Unknown tool: ${name}`);
}

module.exports = async function handler(request, response) {
  if (request.method !== 'POST') return response.status(405).json({ error: 'MCP requires POST' });
  const rpc = await bodyOf(request);
  response.setHeader('MCP-Protocol-Version', '2025-03-26');
  if (rpc.method === 'notifications/initialized') return response.status(202).end();
  if (rpc.method === 'initialize') return response.status(200).json({ jsonrpc: '2.0', id: rpc.id, result: { protocolVersion: '2025-03-26', capabilities: { tools: {} }, serverInfo: { name: 'vinar-travel-tools', version: '1.0.0' } } });
  if (rpc.method === 'tools/list') return response.status(200).json({ jsonrpc: '2.0', id: rpc.id, result: { tools } });
  if (rpc.method === 'tools/call') {
    try {
      const result = await callTool(rpc.params?.name, rpc.params?.arguments || {});
      return response.status(200).json({ jsonrpc: '2.0', id: rpc.id, result: { content: [{ type: 'text', text: result.text }], structuredContent: result.structuredContent } });
    } catch (error) {
      return response.status(200).json({ jsonrpc: '2.0', id: rpc.id, result: { isError: true, content: [{ type: 'text', text: error.message }] } });
    }
  }
  return response.status(200).json({ jsonrpc: '2.0', id: rpc.id, error: { code: -32601, message: 'Method not found' } });
};
