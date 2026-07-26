const SOURCE_URL = 'https://studentncirl-my.sharepoint.com/:x:/r/personal/x25134680_student_ncirl_ie/Documents/Finar_Services_and_FAQ.xlsx?download=1';
const services = [];
const faqs = [];
let selectedFilter = 'All';

const clean = value => String(value ?? '').trim();
const escapeHtml = value => clean(value).replace(/[&<>"']/g, char => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[char]);
const money = value => clean(value) ? (clean(value).includes('$') ? clean(value) : `$${clean(value)}`) : 'Price on request';
const normalize = value => clean(value).toLowerCase();

function rowToService(row) {
  return { id: row.service_id, category: row.category, type: row.type, price: row.price_usd, duration: row.duration, booking: row.requires_booking, availability: row.availability, slots: row.slots_this_week, offer: row.special_offer, name: row.service_name, description: row.description };
}

function renderServices() {
  const visible = services.filter(service => selectedFilter === 'All' || service.category === selectedFilter);
  document.querySelector('#serviceCount').textContent = `${services.length} service${services.length === 1 ? '' : 's'}`;
  document.querySelector('#navCount').textContent = services.length || '—';
  document.querySelector('#serviceList').innerHTML = visible.slice(0, 8).map(service => `<div class="service-row"><div><strong>${escapeHtml(service.name)}</strong><small>${escapeHtml(service.category)} · ${escapeHtml(service.duration)}</small></div><div><div class="price">${escapeHtml(money(service.price))}</div><div class="tag">${escapeHtml(service.slots || 'Check slots')}</div></div></div>`).join('') || '<div class="empty-state">No live services are available to display yet.</div>';
}

function parseRows(rows) {
  const headers = rows[0].map(clean);
  return rows.slice(1).map(row => Object.fromEntries(headers.map((header, index) => [header, clean(row[index])])));
}

function parseCsv(text) {
  return text.trim().split(/\r?\n/).map(line => line.split(',').map(value => value.replace(/^"|"$/g, '')));
}

async function loadLiveData() {
  document.querySelector('#syncStatus').textContent = 'Connecting';
  try {
    const response = await fetch(window.VINAR_DATA_URL || SOURCE_URL, { mode: 'cors' });
    if (!response.ok) throw new Error('Source unavailable');
    const contentType = response.headers.get('content-type') || '';
    if (!contentType.includes('json') && !contentType.includes('csv')) throw new Error('Workbook requires a public CSV/JSON export');
    const payload = await response.text();
    const data = JSON.parse(payload);
    services.splice(0, services.length, ...(data.services || []).map(rowToService));
    faqs.splice(0, faqs.length, ...(data.faqs || []));
    document.querySelector('#dataStatus').textContent = 'Data updated just now';
    document.querySelector('#syncStatus').textContent = 'Synced';
  } catch (error) {
    document.querySelector('#dataStatus').textContent = 'Live source needs access';
    document.querySelector('#syncStatus').textContent = 'Check access';
  }
  renderServices();
}

function answer(question) {
  const query = normalize(question);
  const faq = faqs.find(item => normalize(item.question).includes(query) || query.includes(normalize(item.question)));
  if (faq) return { text: faq.answer };
  const matches = services.filter(service => [service.category, service.type, service.name, service.description].some(value => normalize(value).includes(query) || query.split(/\s+/).some(word => word.length > 3 && normalize(value).includes(word))));
  if (!matches.length) return { text: 'I could not find that in Vinar’s connected services or FAQ data. A human advisor can check custom routes, special requests or anything not listed here. Call +249 914 101 013 / 012.' };
  return { text: `I found ${matches.length === 1 ? 'this option' : 'these options'} in Vinar’s live directory:`, cards: matches.slice(0, 4) };
}

function send(question) {
  if (!question.trim()) return;
  const conversation = document.querySelector('#conversation');
  conversation.insertAdjacentHTML('beforeend', `<div class="message user">${escapeHtml(question)}</div>`);
  const result = answer(question);
  const cards = result.cards ? `<div class="answer-card">${result.cards.map(service => `<div><strong>${escapeHtml(service.name)}</strong><small>${escapeHtml(service.category)} · ${escapeHtml(service.duration)} · ${escapeHtml(service.availability)}</small></div><div><b>${escapeHtml(money(service.price))}</b><span>${escapeHtml(service.offer || (service.slots === '0' ? 'Fully booked this week' : `${service.slots || 'Check'} slots this week`))}</span></div>`).join('')}</div>` : '';
  setTimeout(() => { conversation.insertAdjacentHTML('beforeend', `<div class="message bot"><strong>Vinar Travel</strong><br>${escapeHtml(result.text)}${cards}</div>`); conversation.lastElementChild.scrollIntoView({ behavior: 'smooth', block: 'nearest' }); }, 180);
}

document.querySelector('#chatForm').addEventListener('submit', event => { event.preventDefault(); const input = document.querySelector('#questionInput'); send(input.value); input.value = ''; });
document.querySelectorAll('[data-question]').forEach(button => button.addEventListener('click', () => send(button.dataset.question)));
document.querySelectorAll('.filter').forEach(button => button.addEventListener('click', () => { selectedFilter = button.dataset.filter; document.querySelectorAll('.filter').forEach(item => item.classList.toggle('active', item === button)); renderServices(); }));
document.querySelector('#refreshData').addEventListener('click', loadLiveData);
loadLiveData();
