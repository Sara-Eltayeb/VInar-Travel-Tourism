const SHAREPOINT_WORKBOOK_URL = 'https://studentncirl-my.sharepoint.com/:x:/r/personal/x25134680_student_ncirl_ie/_layouts/15/doc2.aspx?action=edit&sourcedoc=%7Baa91b898-7ebf-43d1-9768-359c3b5ff4e7%7D&wdExp=TEAMS-TREATMENT&web=1&TeamsCID=e2df6189-66f5-4c6a-b36d-4ce29808e0c6';
const services = [];
const faqs = [];
let selectedFilter = 'All';

const clean = value => String(value ?? '').trim();
const normalize = value => clean(value).toLowerCase().replace(/[^a-z0-9\s]/g, ' ');
const words = value => normalize(value).split(/\s+/).filter(word => word.length > 2);
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
  if (serviceHeaders.some(header => headers.includes(header))) return { services: records.map(rowToService), faqs: [] };
  return { services: [], faqs: records };
}

function renderServices() {
  const visible = services.filter(service => selectedFilter === 'All' || service.category === selectedFilter);
  document.querySelector('#serviceCount').textContent = `${services.length} service${services.length === 1 ? '' : 's'}`;
  document.querySelector('#navCount').textContent = services.length || '—';
  document.querySelector('#serviceList').innerHTML = visible.slice(0, 8).map(service => `<div class="service-row"><div><strong>${escapeHtml(service.name)}</strong><small>${escapeHtml(service.category)} · ${escapeHtml(service.duration)}</small></div><div><div class="price">${escapeHtml(money(service.price))}</div><div class="tag">${escapeHtml(service.slots || 'Check slots')}</div></div></div>`).join('') || '<div class="empty-state">No live services are available to display yet.</div>';
}

function setData(data) {
  services.splice(0, services.length, ...(data.services || []).map(rowToService));
  faqs.splice(0, faqs.length, ...(data.faqs || []));
}

async function parseResponse(response) {
  const contentType = response.headers.get('content-type') || '';
  if (contentType.includes('json')) return response.json();
  if (contentType.includes('csv')) return rowsToData((await response.text()).trim().split(/\r?\n/).map(line => line.split(',')));
  if (typeof XLSX === 'undefined') throw new Error('XLSX parser unavailable');
  const workbook = XLSX.read(await response.arrayBuffer(), { type: 'array' });
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
    const source = window.VINAR_DATA_URL || SHAREPOINT_WORKBOOK_URL;
    const response = await fetch(source, { mode: 'cors', credentials: 'include', cache: 'no-store' });
    if (!response.ok) throw new Error(`Source returned ${response.status}`);
    setData(await parseResponse(response));
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
  return words(query).reduce((total, word) => total + (target.includes(word) ? (target === normalize(query) ? 4 : 1) : 0), 0);
}

function answer(question) {
  const query = normalize(question);
  const faqMatches = faqs.map(faq => ({ faq, score: score(query, `${faq.category} ${faq.question} ${faq.answer}`) })).filter(item => item.score > 0).sort((a, b) => b.score - a.score);
  if (faqMatches[0] && faqMatches[0].score >= 2) return { text: faqMatches[0].faq.answer };
  const serviceMatches = services.map(service => ({ service, score: score(query, `${service.category} ${service.type} ${service.name} ${service.description}`) })).filter(item => item.score > 0).sort((a, b) => b.score - a.score).slice(0, 4).map(item => item.service);
  if (!serviceMatches.length) return { text: 'I could not find an answer in Vinar’s connected Services & Packages or FAQ data. A human advisor can check custom routes and special requests. Call +249 914 101 013 / 012.' };
  return { text: `I found ${serviceMatches.length === 1 ? 'this option' : 'these options'} in Vinar’s live directory:`, cards: serviceMatches };
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
