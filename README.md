# Vinar Travel & Tourism

English-first Vinar travel assistant for GitHub Pages.

## Live data

The interface is intentionally empty until it can read the connected source. This prevents it from inventing prices, offers, availability, or FAQ answers.

The app now targets the public Google Sheet first:

`https://docs.google.com/spreadsheets/d/1KZR98ZMug3CP6-ztfJ9LdSddJNXy5a8-NJW_8SMdSCs/edit?gid=2056820402`

It reads the `Services & Packages` data through the sheet's GViz endpoint and separately reads the `FAQ` tab. The OneDrive workbook remains configured as a fallback:

`https://1drv.ms/x/c/19b2686eee879b15/IQD1elhkALFnRK8yfklM69vkAcRaUh4T97Qn9he6cGXJZQQ?e=Ve8lBx`

It falls back to the original SharePoint workbook here:

`https://studentncirl-my.sharepoint.com/:x:/r/personal/x25134680_student_ncirl_ie/_layouts/15/doc2.aspx?action=edit&sourcedoc=%7Baa91b898-7ebf-43d1-9768-359c3b5ff4e7%7D&wdExp=TEAMS-TREATMENT&web=1&TeamsCID=e2df6189-66f5-4c6a-b36d-4ce29808e0c6`

The browser can parse an anonymously accessible `.xlsx` response using SheetJS. For the current authenticated SharePoint workbook, set `window.VINAR_DATA_URL` to a small authenticated proxy/API that returns this shape:

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

The workbook currently returns `401 Unauthorized` outside the Microsoft account session. GitHub Pages cannot safely carry Microsoft credentials or bypass that access control. Make the workbook anonymous-read or provide the authenticated proxy before expecting live answers in the public deployment.

Booking confirmation and payment are always handled by a human Vinar advisor.
