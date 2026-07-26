# Vinar Travel & Tourism

English-first Vinar travel assistant for GitHub Pages.

## Live data

The interface is intentionally empty until it can read the connected source. This prevents it from inventing prices, offers, availability, or FAQ answers.

The browser expects `window.VINAR_DATA_URL` to point to a public JSON endpoint with this shape:

```json
{
  "services": [{
    "service_id": "...",
    "category": "...",
    "type": "...",
    "price_usd": "...",
    "duration": "...",
    "requires_booking": "...",
    "availability": "...",
    "slots_this_week": "...",
    "special_offer": "...",
    "service_name": "...",
    "description": "..."
  }],
  "faqs": [{ "faq_id": "...", "category": "...", "question": "...", "answer": "..." }]
}
```

The SharePoint workbook link is kept as the configured source, but it requires authenticated access and cannot be read directly by a public GitHub Pages browser. Use a small authenticated backend or a public export endpoint for production data access.

Booking confirmation and payment are always handled by a human Vinar advisor.
