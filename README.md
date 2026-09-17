# UVC Procurement App (free, works on your phone)

**Flow:** MRF email → RFQ (official UVC template) → WhatsApp to suppliers → record quotes → best quote to manager → on "proceed" → Goods Received Note (official UVC template).

The app fills your real UVC forms: RFQ numbers like `UVC/RFQ/2026/75`, GRN numbers like `UVC/GRN/2026/1287`, and carries the UVC Quote Reference / Project / Location from the MRF straight into the GRN.

---

## 🚀 Recommended: run it on your phone (free)

### One-time setup (~20 minutes)
1. **Create a free GitHub account** at github.com (5 min).
2. Create a new repository, e.g. `uvc-procurement`, and upload **all files from this folder** — `app.py`, `requirements.txt`, and the whole `templates/` folder (3 files). (Use "uploading an existing file" → drag all files in.)
3. Go to **share.streamlit.io** → sign in with GitHub → **New app** → pick your repo, main file `app.py` → **Deploy**. Free.
4. You get a link like `https://uvc-procurement.streamlit.app`.
5. **On your phone:** open the link in Chrome/Safari → menu → **Add to Home Screen**. Now it launches like a real app, full-screen.

### Why this works great on a phone
- WhatsApp buttons open WhatsApp directly with the message prefilled — attach the RFQ file and send.
- On the phone you can photograph a supplier's quotation and upload the photo straight into tab 3.
- Outlook login happens once; the app remembers.

### Free Outlook connection (one-time, 5 min)
1. portal.azure.com → sign in with your UVC work account.
2. Search **Microsoft Entra ID** → **App registrations** → **New registration** → name `UVCProcurement` → register.
3. Copy the **Application (client) ID** → paste it in the app's sidebar ("Azure Client ID").
4. **API permissions** → Add → Microsoft Graph → **Delegated** → `Mail.Read`, `Mail.Send` → **Grant admin consent** if the button appears.

## Alternative: run on your PC, use phone on same Wi-Fi
1. Install Python (python.org, tick "Add Python to PATH").
2. In this folder: `pip install -r requirements.txt` then `streamlit run app.py`.
3. It prints a "Network URL" (e.g. http://192.168.1.5:8501) — open that on your phone while connected to the same Wi-Fi. PC must stay on.

---

## Daily use (2 minutes)
1. **Tab 1 — MRF:** tap "fetch latest MRF mails from Outlook" (or upload the .docx), check the extracted items, confirm.
2. **Tab 2 — RFQ:** enter the next sequence number → one download button + one WhatsApp button per supplier. Tap, attach, send. "Mark RFQ sent" bumps the sequence for next time.
3. **Tab 3 — Quotations:** when quotes come back on WhatsApp, enter each supplier's prices. Best quote is auto-highlighted.
4. **Tab 4 — Approval:** one tap emails your manager. When he replies "proceed UVC/RFQ/2026/75", tap "Check manager approval".
5. **Tab 5 — GRN:** pre-filled from the MRF + winning quote (Quote Reference, Project/Site, items, quantities, OK/Short/Damaged). Generate → download → print or forward.

## Notes
- Supplier phones: international format, digits only (e.g. `971501234567`).
- MRFs must be the **.docx** version of the form (the old .doc format can't be read — the app will tell you).
- On the free cloud, quotes are kept while the app is open; documents are always downloadable. If you want quotes to survive between sessions, tell me and I'll add free Google Sheets storage.
