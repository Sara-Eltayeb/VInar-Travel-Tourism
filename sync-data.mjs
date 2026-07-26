import { writeFile } from 'node:fs/promises';

const sheetId = '1KZR98ZMug3CP6-ztfJ9LdSddJNXy5a8-NJW_8SMdSCs';
const sources = {
  services: `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:json&gid=2056820402`,
  faqs: `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:json&sheet=FAQ`
};

const clean = value => String(value ?? '').trim();

async function readTab(url, kind) {
  const response = await fetch(url);
  if (!response.ok) throw new Error(`${kind} tab returned ${response.status}`);
  const text = await response.text();
  const start = text.indexOf('(');
  const end = text.lastIndexOf(')');
  const table = JSON.parse(text.slice(start + 1, end)).table;
  const labels = table.cols.map(column => clean(column.label).toLowerCase().replace(/\s+/g, '_'));
  const rows = table.rows.map(row => row.c.map(cell => cell ? clean(cell.f ?? cell.v) : ''));
  const hasHeaders = labels.includes('service_id') || labels.includes('faq_id');
  const allRows = hasHeaders ? [labels, ...rows] : rows;
  const headers = allRows[0].map(value => clean(value).toLowerCase().replace(/\s+/g, '_'));
  return allRows.slice(1).map(row => Object.fromEntries(headers.map((header, index) => [header, clean(row[index])])))
    .filter(row => Object.values(row).some(Boolean));
}

const [services, faqs] = await Promise.all([
  readTab(sources.services, 'Services & Packages'),
  readTab(sources.faqs, 'FAQ')
]);

await writeFile('data.json', `${JSON.stringify({ services, faqs }, null, 2)}\n`);
console.log(`Synced ${services.length} services and ${faqs.length} FAQs.`);
