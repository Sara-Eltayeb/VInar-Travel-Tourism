const GOOGLE_SHEET_ID = '1KZR98ZMug3CP6-ztfJ9LdSddJNXy5a8-NJW_8SMdSCs';
const WORKBOOK_SOURCES = [
  'data.json',
  `https://docs.google.com/spreadsheets/d/${GOOGLE_SHEET_ID}/gviz/tq?tqx=out:json&gid=2056820402`,
  `https://docs.google.com/spreadsheets/d/${GOOGLE_SHEET_ID}/gviz/tq?tqx=out:json&sheet=FAQ`,
  'https://1drv.ms/x/c/19b2686eee879b15/IQD1elhkALFnRK8yfklM69vkAcRaUh4T97Qn9he6cGXJZQQ?e=Ve8lBx',
  'https://studentncirl-my.sharepoint.com/:x:/r/personal/x25134680_student_ncirl_ie/_layouts/15/doc2.aspx?action=edit&sourcedoc=%7Baa91b898-7ebf-43d1-9768-359c3b5ff4e7%7D&wdExp=TEAMS-TREATMENT&web=1&TeamsCID=e2df6189-66f5-4c6a-b36d-4ce29808e0c6'
];
const services = [];
const faqs = [];
let selectedFilter = 'All';
let dataReady = Promise.resolve();

const clean = value => String(value ?? '').trim();
const normalize = value => clean(value).toLowerCase().replace(/[^a-z0-9\s]/g, ' ');
const STOP_WORDS = new Set(['a', 'an', 'and', 'are', 'can', 'do', 'for', 'how', 'i', 'is', 'me', 'my', 'of', 'please', 'the', 'to', 'what', 'where', 'which', 'with', 'you', 'your']);
const words = value => normalize(value).split(/\s+/).filter(word => word.length > 2 && !STOP_WORDS.has(word));
const escapeHtml = value => clean(value).replace(/[&<>"']/g, char => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[char]);
const money = value => clean(value) ? (clean(value).includes('$') ? clean(value) : `$${clean(value)}`) : 'Price on request';

function rowToService(row) {
  return { id: row.service_id, category: row.category, type: row.type, price: row.price_usd, duration: row.duration, booking: row.requires_booking, availability: row.availability, slots: row.slots_this_week, offer: row.special_offer, name: row.service_name, description: row.description };
}

function rowsToData(rows) {
  const headers = rows[0].map(value => clean(value).toLowerCase().replace(/\s+/g, '_'));
  const records = rows.slice(1).map(row => Object.fromEntries(headers.map((header, index) => [header, clean(row[index])])))
    .filter(row => Object.values(row).some(Boolean));
  const serviceHeaders = ['service_id', 'service_name', 'category'];
  if (serviceHeaders.some(header => headers.includes(header))) return { services: records, faqs: [] };
  return { services: [], faqs: records };
}

function renderServices() {
  const visible = services.filter(service => selectedFilter === 'All' || service.category === selectedFilter);
  document.querySelector('#serviceCount').textContent = `${services.length} service${services.length === 1 ? '' : 's'}`;
  const navCount = document.querySelector('#navCount');
  if (navCount) navCount.textContent = services.length || '—';
  document.querySelector('#serviceList').innerHTML = visible.slice(0, 8).map(service => `<div class="service-row"><div><strong>${escapeHtml(service.name)}</strong><small>${escapeHtml(service.category)} · ${escapeHtml(service.duration)}</small></div><div><div class="price">${escapeHtml(money(service.price))}</div><div class="tag">${escapeHtml(service.slots || 'Check slots')}</div></div></div>`).join('') || '<div class="empty-state">No live services are available to display yet.</div>';
}

function mergeData(data) {
  services.push(...(data.services || []).map(rowToService));
  faqs.push(...(data.faqs || []));
}

async function parseResponse(response) {
  const contentType = response.headers.get('content-type') || '';
  const payload = await response.text();
  if (payload.includes('google.visualization.Query.setResponse')) {
    const start = payload.indexOf('(');
    const end = payload.lastIndexOf(')');
    const table = JSON.parse(payload.slice(start + 1, end)).table;
    const headers = table.cols.map(column => clean(column.label).toLowerCase().replace(/\s+/g, '_'));
    const rows = table.rows.map(row => row.c.map(cell => cell ? clean(cell.f ?? cell.v) : ''));
    return rowsToData(headers.some(header => header === 'service_id' || header === 'faq_id') ? [headers, ...rows] : rows);
  }
  if (contentType.includes('json')) return JSON.parse(payload);
  if (contentType.includes('csv')) return rowsToData((await response.text()).trim().split(/\r?\n/).map(line => line.split(',')));
  if (typeof XLSX === 'undefined') throw new Error('XLSX parser unavailable');
  const workbook = XLSX.read(await (await fetch(response.url)).arrayBuffer(), { type: 'array' });
  const data = { services: [], faqs: [] };
  workbook.SheetNames.forEach(sheetName => {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: '' });
    const sheetData = rowsToData(rows);
    if (sheetName.toLowerCase().includes('faq')) data.faqs.push(...sheetData.faqs);
    else data.services.push(...sheetData.services);
  });
  return data;
}

async function loadLiveData() {
  document.querySelector('#syncStatus').textContent = 'Connecting';
  try {
    const sources = window.VINAR_DATA_URL ? [window.VINAR_DATA_URL] : WORKBOOK_SOURCES;
    services.splice(0, services.length);
    faqs.splice(0, faqs.length);
    let loaded = false;
    for (const [index, source] of sources.entries()) {
      try {
        const response = await fetch(source, { mode: 'cors', credentials: 'omit', cache: 'no-store' });
        if (!response.ok) continue;
        mergeData(await parseResponse(response));
        loaded = true;
        if (window.VINAR_DATA_URL || source === 'data.json' || (index >= 2 && loaded)) break;
      } catch (sourceError) {
        if (index >= 1 && loaded) break;
      }
    }
    if (!loaded) throw new Error('No workbook source was accessible');
    document.querySelector('#dataStatus').textContent = 'Data updated just now';
    document.querySelector('#syncStatus').textContent = 'Synced';
  } catch (error) {
    document.querySelector('#dataStatus').textContent = 'SharePoint access required';
    document.querySelector('#syncStatus').textContent = 'Needs access';
  }
  renderServices();
}

function score(query, value) {
  const target = normalize(value);
  return words(query).reduce((total, word) => {
    if (target.includes(word)) return total + 1;
    const close = target.split(/\s+/).some(candidate => candidate.length > 2 && levenshtein(word, candidate) <= (word.length > 5 ? 2 : 1));
    return total + (close ? 0.75 : 0);
  }, 0);
}

function levenshtein(first, second) {
  const row = Array.from({ length: second.length + 1 }, (_, index) => index);
  for (let i = 1; i <= first.length; i += 1) {
    let previous = row[0];
    row[0] = i;
    for (let j = 1; j <= second.length; j += 1) {
      const current = row[j];
      row[j] = Math.min(row[j] + 1, row[j - 1] + 1, previous + (first[i - 1] === second[j - 1] ? 0 : 1));
      previous = current;
    }
  }
  return row[second.length];
}

function answer(question) {
  const query = normalize(question);
  const serviceMatches = services.map(service => ({ service, score: score(query, `${service.category} ${service.type} ${service.name} ${service.description}`) })).filter(item => item.score > 0).sort((a, b) => b.score - a.score).slice(0, 4).map(item => item.service);
  const faqMatches = faqs.map(faq => ({ faq, score: score(query, faq.question) * 3 + score(query, faq.category) })).filter(item => item.score > 0).sort((a, b) => b.score - a.score);
  const asksForService = ['package', 'honeymoon', 'umrah', 'hajj', 'flight', 'hotel', 'tour', 'transfer', 'visa', 'service', 'price', 'cost', 'offer', 'available'].some(term => score(query, term) > 0);
  if (serviceMatches.length && asksForService) return { text: `I found ${serviceMatches.length === 1 ? 'this option' : 'these options'} in Vinar’s live directory:`, cards: serviceMatches };
  if (faqMatches[0] && faqMatches[0].score >= 3) return { text: faqMatches[0].faq.answer };
  if (!serviceMatches.length) return { text: 'I could not find an answer in Vinar’s connected Services & Packages or FAQ data. A human advisor can check custom routes and special requests. Call +249 914 101 013 / 012.' };
  return { text: `I found ${serviceMatches.length === 1 ? 'this option' : 'these options'} in Vinar’s live directory:`, cards: serviceMatches };
}

async function loadCurrency() {
  const amount = Number(document.querySelector('#currencyAmount').value) || 0;
  const currency = document.querySelector('#currencyCode').value;
  const output = document.querySelector('#currencyResult');
  output.textContent = 'Checking live rate...';
  try {
    const mcp = await callMcpTool('convert_currency', { amount, currency });
    if (mcp?.value !== undefined) {
      output.textContent = `${amount.toLocaleString()} USD ≈ ${Number(mcp.value).toLocaleString()} ${currency}`;
      return;
    }
    const response = await fetch('https://open.er-api.com/v6/latest/USD');
    const data = await response.json();
    output.textContent = `${amount.toLocaleString()} USD ≈ ${(amount * data.rates[currency]).toLocaleString(undefined, { maximumFractionDigits: 2 })} ${currency}`;
  } catch (error) {
    output.textContent = 'Live rate unavailable';
  }
}

async function loadWeather() {
  const destination = document.querySelector('#destinationInput').value.trim();
  const output = document.querySelector('#weatherResult');
  if (!destination) return;
  output.textContent = 'Finding destination...';
  try {
    const mcp = await callMcpTool('get_weather', { destination });
    if (mcp?.destination) {
      output.innerHTML = `<strong>${escapeHtml(mcp.destination)}</strong><span>${Math.round(mcp.temperature_c)}°C · Wind ${Math.round(mcp.wind_kmh)} km/h</span>`;
      return;
    }
    const locationResponse = await fetch(`https://geocoding-api.open-meteo.com/v1/search?name=${encodeURIComponent(destination)}&count=1&language=en&format=json`);
    const location = (await locationResponse.json()).results?.[0];
    if (!location) throw new Error('Destination not found');
    const weatherResponse = await fetch(`https://api.open-meteo.com/v1/forecast?latitude=${location.latitude}&longitude=${location.longitude}&current=temperature_2m,weather_code,wind_speed_10m&timezone=auto`);
    const weather = await weatherResponse.json();
    output.innerHTML = `<strong>${escapeHtml(location.name)}</strong><span>${Math.round(weather.current.temperature_2m)}°C · ${weatherDescription(weather.current.weather_code)}</span>`;
  } catch (error) {
    output.textContent = 'Weather unavailable for this destination';
  }
}

async function callMcpTool(name, args) {
  try {
    const response = await fetch('/api/mcp', { method: 'POST', headers: { 'Content-Type': 'application/json', Accept: 'application/json' }, body: JSON.stringify({ jsonrpc: '2.0', id: Date.now(), method: 'tools/call', params: { name, arguments: args } }) });
    if (!response.ok) return null;
    return (await response.json()).result?.structuredContent || null;
  } catch (error) {
    return null;
  }
}

function weatherDescription(code) {
  if (code === 0) return 'Clear sky';
  if ([1, 2, 3].includes(code)) return 'Partly cloudy';
  if ([45, 48].includes(code)) return 'Foggy';
  if ([51, 53, 55, 56, 57].includes(code)) return 'Light rain';
  if ([61, 63, 65, 80, 81, 82].includes(code)) return 'Rainy';
  if ([71, 73, 75, 77, 85, 86].includes(code)) return 'Snowy';
  if ([95, 96, 99].includes(code)) return 'Thunderstorms';
  return 'Check local forecast';
}

async function answerWithGemini(question) {
  try {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 8000);
    const response = await fetch('/api/chat', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ question }), signal: controller.signal });
    clearTimeout(timeout);
    if (!response.ok) return null;
    const result = await response.json();
    return result.text ? { text: result.text } : null;
  } catch (error) {
    return null;
  }
}

async function send(question) {
  if (!question.trim()) return;
  const conversation = document.querySelector('#conversation');
  conversation.replaceChildren();
  conversation.insertAdjacentHTML('beforeend', `<div class="message user">${escapeHtml(question)}</div>`);
  await dataReady;
  const localResult = answer(question);
  const result = localResult;
  const cards = result.cards ? `<div class="answer-card">${result.cards.map(service => `<div><strong>${escapeHtml(service.name)}</strong><small>${escapeHtml(service.category)} · ${escapeHtml(service.duration)} · ${escapeHtml(service.availability)}</small></div><div><b>${escapeHtml(money(service.price))}</b><span>${escapeHtml(service.offer || (service.slots === '0' ? 'Fully booked this week' : `${service.slots || 'Check'} slots this week`))}</span></div>`).join('')}</div>` : '';
  setTimeout(() => { conversation.insertAdjacentHTML('beforeend', `<div class="message bot"><strong>Vinar Travel</strong><br>${escapeHtml(result.text)}${cards}</div>`); conversation.lastElementChild.scrollIntoView({ behavior: 'smooth', block: 'nearest' }); }, 180);
}

document.querySelector('#chatForm').addEventListener('submit', event => { event.preventDefault(); const input = document.querySelector('#questionInput'); send(input.value); input.value = ''; });
document.querySelectorAll('[data-question]').forEach(button => button.addEventListener('click', () => send(button.dataset.question)));
document.querySelectorAll('.filter').forEach(button => button.addEventListener('click', () => { selectedFilter = button.dataset.filter; document.querySelectorAll('.filter').forEach(item => item.classList.toggle('active', item === button)); renderServices(); }));
document.querySelector('#refreshData').addEventListener('click', loadLiveData);
document.querySelector('#currencyAmount').addEventListener('input', loadCurrency);
document.querySelector('#currencyCode').addEventListener('change', loadCurrency);
document.querySelector('#weatherButton').addEventListener('click', loadWeather);
document.querySelector('#destinationInput').addEventListener('keydown', event => { if (event.key === 'Enter') loadWeather(); });
dataReady = loadLiveData();
loadCurrency();
