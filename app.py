# UVC Procurement App — fills the official UVC templates
# Flow: MRF (Outlook) -> RFQ -> WhatsApp suppliers -> Quotations -> Manager approval -> GRN
# Free stack: Streamlit + Microsoft Graph (Outlook) + WhatsApp wa.me links
# Phone-ready: deploy free on Streamlit Cloud, open in phone browser, "Add to Home Screen".
import os, re, io, json, base64, urllib.parse, datetime
import streamlit as st
import requests, msal
from docx import Document

GRAPH = "https://graph.microsoft.com/v1.0"
SCOPES = ["Mail.Read", "Mail.Send", "offline_access"]
TPL_DIR = os.path.dirname(os.path.abspath(__file__))
TPL_RFQ = os.path.join(TPL_DIR, "SAMPLE_RFQ.docx")
TPL_GRN = os.path.join(TPL_DIR, "SAMPLE_GRN.docx")
DATA_DIR = "procurement_data"
os.makedirs(DATA_DIR, exist_ok=True)

DEFAULT_CFG = {
    "client_id": "",
    "manager_email": "",
    "buyer_name": "",
    "buyer_phone": "",
    "approval_keyword": "proceed",
    "next_rfq_seq": "75",
    "next_grn_seq": "1287",
    "company": "UVC Technical Services L.L.C",
}


# ---------------- config / persistence ----------------
def cfg_path():
    return os.path.join(DATA_DIR, "config.json")


def load_cfg():
    cfg = dict(DEFAULT_CFG)
    if os.path.exists(cfg_path()):
        with open(cfg_path(), "r") as f:
            cfg.update(json.load(f))
    return cfg


def save_cfg(cfg):
    with open(cfg_path(), "w") as f:
        json.dump(cfg, f, indent=2)


def suppliers_path():
    return os.path.join(DATA_DIR, "suppliers.csv")


def load_suppliers():
    import pandas as pd

    cols = ["name", "attention", "phone", "email"]
    if os.path.exists(suppliers_path()):
        df = pd.read_csv(suppliers_path(), dtype=str)
        for c in cols:
            if c not in df.columns:
                df[c] = ""
        return df[cols].fillna("")
    df = pd.DataFrame(columns=cols)
    for c in cols:
        df[c] = df[c].astype(str)
    return df


def save_suppliers(df):
    df.to_csv(suppliers_path(), index=False)


def quotes_path(rfq_no):
    return os.path.join(DATA_DIR, f"quotes_{rfq_no.replace('/','_')}.json")


def load_quotes(rfq_no):
    p = quotes_path(rfq_no)
    if os.path.exists(p):
        with open(p, "r") as f:
            return json.load(f)
    return []


def save_quotes(rfq_no, q):
    with open(quotes_path(rfq_no), "w") as f:
        json.dump(q, f, indent=2)


# ---------------- Outlook (Microsoft Graph, free) ----------------
def get_token(client_id):
    cache = msal.SerializableTokenCache()
    cf = os.path.join(DATA_DIR, "token_cache.bin")
    if os.path.exists(cf):
        cache.deserialize(open(cf).read())
    app = msal.PublicClientApplication(
        client_id,
        authority="https://login.microsoftonline.com/common",
        token_cache=cache,
    )
    result = None
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
    if not result or "access_token" not in (result or {}):
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise RuntimeError("Device login failed to start: " + str(flow))
        st.warning("**Sign in to Outlook** (browser will open):\n\n" + flow["message"])
        result = app.acquire_token_by_device_flow(flow)
    open(cf, "w").write(cache.serialize())
    if "access_token" not in result:
        raise RuntimeError("Login failed: " + str(result.get("error_description")))
    return result["access_token"]


def headers(t):
    return {"Authorization": f"Bearer {t}"}


def list_inbox(token, only_mrf=True, top=20):
    # Fetch only UNREAD emails, and expand to see attachment names
    url = f"{GRAPH}/me/mailFolders/inbox/messages?$top={top}&$orderby=receivedDateTime desc&$filter=hasAttachments eq true and isRead eq false&$expand=attachments($select=name)"
    r = requests.get(url, headers=headers(token))
    r.raise_for_status()

    valid_mails = []
    for m in r.json().get("value", []):
        for att in m.get("attachments", []):
            name = (att.get("name") or "").lower()
            # Check if 'mrf' is in the file name and it's a Word/Text doc
            if (not only_mrf) or ("mrf" in name and name.endswith((".docx", ".txt"))):
                valid_mails.append(m)
                break  # Found a match, keep this email
    return valid_mails


def get_attachment(token, msg_id):
    r = requests.get(
        f"{GRAPH}/me/messages/{msg_id}/attachments", headers=headers(token)
    )
    r.raise_for_status()
    for att in r.json().get("value", []):
        n = (att.get("name") or "").lower()
        if n.endswith((".docx", ".txt")):
            return att.get("name"), base64.b64decode(att["contentBytes"])
    return None, None


def send_email(token, to, subject, body, files=None):
    atts = []
    for name, data in files or []:
        atts.append(
            {
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": name,
                "contentBytes": base64.b64encode(data).decode(),
            }
        )
    payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": "Text", "content": body},
            "toRecipients": [{"emailAddress": {"address": to}}],
            "attachments": atts,
        }
    }
    r = requests.post(f"{GRAPH}/me/sendMail", headers=headers(token), json=payload)
    r.raise_for_status()
    return True


def check_manager_approval(token, manager_email, keyword, ref_no):
    url = (
        f"{GRAPH}/me/messages?$top=25&$orderby=receivedDateTime desc"
        f"&$filter=from/emailAddress/address eq '{manager_email.lower()}'"
    )
    r = requests.get(url, headers=headers(token))
    r.raise_for_status()
    for m in r.json().get("value", []):
        body = (
            ((m.get("body") or {}).get("content") or "")
            + " "
            + (m.get("subject") or "")
        )
        if keyword.lower() in body.lower() and ref_no.lower() in body.lower():
            return True, m.get("subject", "")
    return False, ""


# ---------------- docx helpers ----------------
def set_cell(cell, text):
    text = str(text)
    for p in cell.paragraphs[1:]:
        p._element.getparent().remove(p._element)
    p = cell.paragraphs[0]
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    lines = text.split("\n")
    p.add_run(lines[0])
    for line in lines[1:]:
        cell.add_paragraph(line)


def fill_header_table(doc, mapping):
    for row in doc.tables[0].rows:
        cs = row.cells
        for i in range(0, len(cs) - 1, 2):
            label = cs[i].text.strip().replace("\n", " ")
            for key, val in mapping.items():
                if key.lower() in label.lower() and val is not None:
                    set_cell(cs[i + 1], val)


def find_items_table(doc):
    for t in doc.tables:
        if "Sl No" in t.rows[0].cells[0].text:
            return t
    return None


def fill_items(table, items, col_map, ncols):
    while len(table.rows) - 1 < len(items):
        table.add_row()
    for ri, it in enumerate(items, start=1):
        cs = table.rows[ri].cells
        for ci in range(ncols):
            cs[ci].text = ""
        for col, key in col_map.items():
            if key in it and it[key] is not None:
                set_cell(cs[col], it[key])


def doc_to_bytes(doc):
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


# ---------------- MRF parsing (official UVC form) ----------------
def parse_mrf_docx(data):
    try:
        doc = Document(io.BytesIO(data))
        fields = {}
        if doc.tables:
            for row in doc.tables[0].rows:
                cs = row.cells
                for i in range(0, len(cs) - 1, 2):
                    fields[cs[i].text.strip().replace("\n", " ")] = cs[i + 1].text.strip()

        def getf(prefix):
            for k, v in fields.items():
                if k.lower().startswith(prefix.lower()):
                    return v
            return ""

        items = []
        t = find_items_table(doc)
        if t:
            for row in t.rows[1:]:
                cs = [c.text.strip() for c in row.cells]
                if len(cs) >= 5 and cs[0].isdigit() and cs[1]:
                    items.append(
                        {
                            "Sl": cs[0],
                            "Item": cs[1],
                            "Sample": cs[2] if len(cs) > 2 else "",
                            "Unit": cs[3] if len(cs) > 3 else "",
                            "Qty": cs[4] if len(cs) > 4 else "",
                            "ReqDate": cs[5] if len(cs) > 5 else "",
                        }
                    )
            
            notes, paras = "", [p.text for p in doc.paragraphs]
            for i, p in enumerate(paras):
                if p.strip().startswith("Notes"):
                    notes = paras[i + 1].strip() if i + 1 < len(paras) else ""
                    break
                    
            mrf = {
                "mrf_no": getf("MRF No"),
                "site": getf("Delivery To"),
                "location": getf("Location"),
                "oa_building": getf("OA / Building"),
                "quote_ref": getf("UVC Quote Reference"),
                "requested_by": getf("Requested By"),
                "lpo_ref": getf("LPO / WO Ref"),
                "priority": getf("Priority"),
                "date": getf("Date"),
                "items": items,
                "notes": notes,
                "project_site": " / ".join(
                    x for x in [getf("Delivery To"), getf("OA / Building")] if x
                ),
            }
            if items:
                return mrf

        # Fallback if no items table or no items extracted
        full_text = "\n".join([p.text for p in doc.paragraphs])
        for table in doc.tables:
            for row in table.rows:
                full_text += "\n" + " | ".join(cell.text.replace("\n", " ") for cell in row.cells)
        return parse_mrf_text(full_text)
    except Exception:
        # Final fallback: just extract whatever text we can and send to text parser
        doc = Document(io.BytesIO(data))
        full_text = "\n".join([p.text for p in doc.paragraphs])
        return parse_mrf_text(full_text)


LINE_RE = re.compile(
    r"^\s*(?:item\s*)?(\d{1,3})?[\.\)\-:]?\s*(.+?)\s+(\d+(?:[.,]\d+)?)\s*"
    r"(pcs?|nos?|units?|ea|kg|mtrs?|meters?|sets?|boxes?|rolls?|bags?|sheets?)?\s*$",
    re.I,
)


def parse_mrf_text(text):
    items = []
    for raw in text.splitlines():
        line = raw.strip().strip("|").strip()
        if len(line) < 4:
            continue
        m = LINE_RE.match(line)
        if m and not re.fullmatch(r"[\d\s\.\)\-]+", line):
            desc = re.sub(r"\s{2,}", " ", m.group(2)).strip(" -–—:|")
            if len(desc) >= 3:
                items.append(
                    {
                        "Sl": str(len(items) + 1),
                        "Item": desc,
                        "Sample": "",
                        "Unit": (m.group(4) or "").lower(),
                        "Qty": m.group(3),
                        "ReqDate": "",
                    }
                )
    return {
        "mrf_no": "",
        "site": "",
        "location": "",
        "oa_building": "",
        "quote_ref": "",
        "requested_by": "",
        "lpo_ref": "",
        "priority": "",
        "date": "",
        "items": items,
        "notes": "",
        "project_site": "",
    }


def parse_mrf(name, data):
    name_lower = name.lower()
    if name_lower.endswith(".docx"):
        return parse_mrf_docx(data)
    elif name_lower.endswith(".pdf"):
        import pypdf
        try:
            reader = pypdf.PdfReader(io.BytesIO(data))
            text = "\n".join([page.extract_text() for page in reader.pages if page.extract_text()])
            return parse_mrf_text(text)
        except Exception as e:
            raise Exception(f"Failed to extract text from PDF: {str(e)}")
    elif name_lower.endswith(".doc"):
        raise Exception("Older '.doc' files are not supported because they are binary files. Please save the file as a '.docx' or '.pdf' and try again.")
    
    # Try decoding as text, but gracefully catch decoding errors
    try:
        text = data.decode("utf-8")
        return parse_mrf_text(text)
    except UnicodeDecodeError:
        raise Exception("This file appears to be a binary format that the app cannot read directly. Please save it as a .docx, .pdf, or plain .txt file and try again.")


# ---------------- document builders (official UVC templates) ----------------
def build_rfq(supplier, rfq_no, date, quote_by, delivery_period, items, extra_rows=0):
    doc = Document(TPL_RFQ)
    fill_header_table(
        doc,
        {
            "Supplier Name": supplier.get("name", ""),
            "Attention": supplier.get("attention", ""),
            "Email": supplier.get("email", ""),
            "Contact No.": supplier.get("phone", ""),
            "RFQ No.": rfq_no,
            "Date": date,
            "Quotation Required By": quote_by,
            "Required Delivery Period": delivery_period,
        },
    )
    fill_items(
        find_items_table(doc),
        [
            {
                "Sl": str(i + 1),
                "Item": it["Item"],
                "Unit": it.get("Unit", ""),
                "Qty": it.get("Qty", ""),
            }
            for i, it in enumerate(items)
        ],
        {0: "Sl", 1: "Item", 2: "Unit", 3: "Qty"},
        4,
    )
    return doc_to_bytes(doc)


def build_grn(
    grn_no, date, received_at, location, quote_ref, project_site, lpo_ref, items
):
    doc = Document(TPL_GRN)
    fill_header_table(
        doc,
        {
            "Received At": received_at,
            "Location": location,
            "GRN No.": grn_no,
            "Date": date,
            "UVC Quote Reference": quote_ref,
            "Project / Site": project_site,
            "LPO / SO Ref": lpo_ref,
        },
    )
    fill_items(
        find_items_table(doc),
        [
            {
                "Sl": str(i + 1),
                "Item": it["Item"],
                "Unit": it.get("Unit", ""),
                "Ordered": it.get("Qty Ordered", ""),
                "Received": it.get("Qty Received", it.get("Qty Ordered", "")),
                "Condition": it.get("Condition", "OK"),
            }
            for i, it in enumerate(items)
        ],
        {0: "Sl", 1: "Item", 2: "Unit", 3: "Ordered", 4: "Received", 5: "Condition"},
        6,
    )
    return doc_to_bytes(doc)


# ---------------- WhatsApp ----------------
def wa_link(phone, message):
    return f"https://wa.me/{re.sub(r'\D', '', str(phone))}?text={urllib.parse.quote(message)}"


def rfq_wa_message(rfq_no, supplier_name, items):
    lines = [
        f"Dear {supplier_name},",
        "",
        f"Please find attached RFQ {rfq_no} for your best price.",
        "",
        "Items:",
    ]
    for i, it in enumerate(items, 1):
        lines.append(f"{i}. {it['Item']} — {it.get('Qty','')} {it.get('Unit','')}")
    lines += [
        "",
        "Kindly send your signed & stamped quotation quoting the RFQ number.",
        "Regards, UVC Procurement",
    ]
    return "\n".join(lines)


# ---------------- UI ----------------
def ui():
    st.set_page_config(page_title="UVC Procurement", page_icon="🧾", layout="wide")
    st.title("🧾 UVC Procurement")
    st.caption("MRF → RFQ → WhatsApp → Quotations → Manager approval → GRN")
    cfg = load_cfg()

    with st.sidebar:
        st.header("⚙️ Settings")
        cfg["client_id"] = st.text_input(
            "Azure Client ID",
            cfg["client_id"],
            help="Free: portal.azure.com → Entra ID → App registrations → New. Copy Application (client) ID. Add delegated Graph permissions Mail.Read + Mail.Send.",
        )
        cfg["manager_email"] = st.text_input("Manager's email", cfg["manager_email"])
        cfg["buyer_name"] = st.text_input(
            "Your name (UVC Procurement)", cfg["buyer_name"]
        )
        cfg["buyer_phone"] = st.text_input(
            "Your WhatsApp (intl digits)", cfg["buyer_phone"]
        )
        cfg["approval_keyword"] = st.text_input(
            "Approval keyword", cfg["approval_keyword"]
        )
        st.divider()
        st.subheader("🔢 Sequences")
        cfg["next_rfq_seq"] = st.text_input("Next RFQ sequence no.", cfg.get("next_rfq_seq", "75"))
        cfg["next_grn_seq"] = st.text_input("Next GRN sequence no.", cfg.get("next_grn_seq", "1287"))
        
        save_cfg(cfg)
        st.divider()
        st.subheader("📇 Suppliers")
        sup = load_suppliers()
        sup = st.data_editor(sup, num_rows="dynamic", use_container_width=True)
        save_suppliers(sup)
        st.caption("Phone: country code + number, digits only (e.g. 971501234567).")
        st.divider()
        token = st.session_state.get("token")
        if st.button("🔌 Connect Outlook", use_container_width=True):
            if not cfg["client_id"]:
                st.error("Paste your Azure Client ID first.")
            else:
                try:
                    token = get_token(cfg["client_id"])
                    st.session_state["token"] = token
                    st.success("Connected!")
                except Exception as e:
                    st.error(str(e))
        if token:
            st.success("✅ Outlook connected")

    t1, t2, t3, t4, t5 = st.tabs(
        ["1️⃣ MRF", "2️⃣ RFQ + WhatsApp", "3️⃣ Quotations", "4️⃣ Approval", "5️⃣ GRN"]
    )

    # ---- Tab 1 ----
    with t1:
        st.subheader("Material Requisition Form")
        up = st.file_uploader(
            "Upload MRF (any text document file)", type=None
        )
        token = st.session_state.get("token")
        if token:
            if st.button("📥 Or fetch latest MRF mails from Outlook"):
                try:
                    st.session_state["mails"] = list_inbox(token, only_mrf=True)
                except Exception as e:
                    st.error(str(e))
            for m in st.session_state.get("mails", []):
                if st.button(
                    f"📧 {m.get('subject','')} — {(m.get('from') or {}).get('emailAddress',{}).get('address','')}",
                    key="m_" + m["id"],
                ):
                    try:
                        name, data = get_attachment(token, m["id"])
                        if data is None:
                            st.warning("No .docx/.txt attachment found.")
                        else:
                            st.session_state["mrf"] = parse_mrf(name, data)
                            st.success(f"Loaded MRF from: {name}")
                    except Exception as e:
                        st.error(str(e))
        if up is not None:
            try:
                st.session_state["mrf"] = parse_mrf(up.name, up.getvalue())
                st.success(f"Parsed: {up.name}")
            except Exception as e:
                st.error(
                    "Could not parse this file. If it is an old .doc file, ask the requester to send the .docx version of the UVC form. ("
                    + str(e)
                    + ")"
                )
        mrf = st.session_state.get("mrf")
        if mrf:
            c1, c2 = st.columns(2)
            c1.write(
                f"**MRF No:** {mrf['mrf_no'] or '—'}  |  **Site:** {mrf['site'] or '—'}  |  **Location:** {mrf['location'] or '—'}"
            )
            c2.write(
                f"**Quote Ref:** {mrf['quote_ref'] or '—'}  |  **Priority:** {mrf['priority'] or '—'}"
            )
            import pandas as pd

            df = pd.DataFrame(mrf["items"])
            if len(df):
                edited = st.data_editor(
                    df, num_rows="dynamic", use_container_width=True
                )
                mrf["items"] = edited.to_dict("records")
                if st.button("✅ Confirm items → create RFQ draft", type="primary"):
                    st.session_state["rfq_items"] = [
                        r for r in mrf["items"] if str(r.get("Item", "")).strip()
                    ]
                    st.success(
                        f"{len(st.session_state['rfq_items'])} items ready — go to tab 2."
                    )
            else:
                st.warning("No items extracted. Check the MRF format.")

    # ---- Tab 2 ----
    with t2:
        st.subheader("RFQ → WhatsApp suppliers")
        items = st.session_state.get("rfq_items")
        mrf = st.session_state.get("mrf", {})
        if not items:
            st.info("Confirm MRF items in tab 1 first.")
        else:
            today = datetime.date.today()
            c1, c2, c3 = st.columns(3)
            seq = c1.text_input("RFQ sequence no.", cfg["next_rfq_seq"])
            rfq_no = f"UVC/RFQ/{today:%Y}/{seq}"
            c2.text_input("RFQ No. (full)", rfq_no, disabled=True)
            q_by = c3.text_input(
                "Quotation required by",
                (today + datetime.timedelta(days=3)).strftime("%d/%m/%Y"),
            )
            d_period = st.text_input(
                "Required delivery period",
                mrf.get("items", [{}])[0].get("ReqDate", "")
                or "As per MRF required date",
            )
            st.dataframe(items, use_container_width=True)
            sup = load_suppliers()
            st.warning(
                "WhatsApp can't auto-attach files (free & safe). Tap each supplier → attach their RFQ → send."
            )
            for idx, s in sup.iterrows():
                if not str(s.get("phone", "")).strip():
                    continue
                docx_bytes = build_rfq(
                    s, rfq_no, today.strftime("%d/%m/%Y"), q_by, d_period, items
                )
                with st.expander(f"🟢 {s['name']}"):
                    safe = re.sub(r"[^A-Za-z0-9]+", "_", s["name"])
                    st.download_button(
                        f"⬇️ Download RFQ for {s['name']}",
                        docx_bytes,
                        file_name=f"{rfq_no.replace('/','-')}_{safe}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"dl_{idx}",
                    )
                    st.link_button(
                        "Open WhatsApp chat",
                        wa_link(s["phone"], rfq_wa_message(rfq_no, s["name"], items)),
                        use_container_width=True,
                    )
            if st.button("💾 Mark RFQ sent (bump sequence)"):
                cfg["next_rfq_seq"] = str(int(seq) + 1)
                save_cfg(cfg)
                st.session_state["rfq_no"] = rfq_no
                st.success(f"RFQ {rfq_no} registered — record quotes in tab 3.")

    # ---- Tab 3 ----
    with t3:
        st.subheader("Supplier quotations")
        items = st.session_state.get("rfq_items")
        rfq_no = st.session_state.get("rfq_no", "")
        rfq_no = st.text_input("RFQ No.", rfq_no, placeholder="UVC/RFQ/2026/75")
        if rfq_no:
            st.session_state["rfq_no"] = rfq_no
        if not items or not rfq_no:
            st.info("Complete tabs 1–2 first.")
        else:
            quotes = load_quotes(rfq_no)
            sup = load_suppliers()

            def qty(it):
                try:
                    return float(str(it.get("Qty", "0")).replace(",", ""))
                except:
                    return 0.0

            with st.form("add_quote", clear_on_submit=True):
                sname = st.selectbox(
                    "Supplier",
                    (
                        sup["name"].tolist()
                        if len(sup)
                        else ["(add suppliers in sidebar)"]
                    ),
                )
                c1, c2 = st.columns(2)
                delivery = c1.text_input("Delivery lead time", "e.g. 7 days")
                terms = c2.text_input("Payment terms", "e.g. 30 days")
                prices = {}
                cols = st.columns(2)
                for i, it in enumerate(items):
                    with cols[i % 2]:
                        prices[it["Item"]] = st.number_input(
                            f"{it['Item']} (qty {qty(it):g} {it.get('Unit','')}) — unit price AED",
                            min_value=0.0,
                            step=0.5,
                            key=f"p_{sname}_{i}",
                        )
                if st.form_submit_button("💾 Save quotation"):
                    rec = {
                        "supplier": sname,
                        "delivery": delivery,
                        "terms": terms,
                        "prices": prices,
                    }
                    quotes = [q for q in quotes if q["supplier"] != sname] + [rec]
                    save_quotes(rfq_no, quotes)
                    st.success(f"Saved quotation from {sname}")
            if quotes:
                import pandas as pd

                rows = [
                    {
                        "Supplier": q["supplier"],
                        "Total (AED)": round(
                            sum(
                                float(q["prices"].get(it["Item"], 0)) * qty(it)
                                for it in items
                            ),
                            2,
                        ),
                        "Delivery": q["delivery"],
                        "Terms": q["terms"],
                    }
                    for q in quotes
                ]
                qdf = pd.DataFrame(rows).sort_values("Total (AED)")
                st.dataframe(qdf, use_container_width=True)
                best = qdf.iloc[0]
                st.success(
                    f"🏆 Best quote: **{best['Supplier']}** — {best['Total (AED)']:,.2f} AED ({best['Delivery']})"
                )
                st.session_state["best_quote"] = best.to_dict()

    # ---- Tab 4 ----
    with t4:
        st.subheader("Send to manager → wait for approval")
        rfq_no = st.session_state.get("rfq_no", "")
        best = st.session_state.get("best_quote")
        if not rfq_no or not best:
            st.info("Record quotations in tab 3 first.")
        else:
            st.write(
                f"**{rfq_no}** — best: **{best['Supplier']}** at **{best['Total (AED)']:,.2f} AED**"
            )
            col_a, col_b = st.columns(2)
            if col_a.button("📧 Send best quotation to manager", type="primary"):
                token = st.session_state.get("token")
                if not token or not cfg["manager_email"]:
                    st.error("Connect Outlook (sidebar) and set manager's email.")
                else:
                    body = (
                        f"Dear Manager,\n\nPlease find below the best quotation received for {rfq_no}.\n\n"
                        f"Supplier : {best['Supplier']}\nTotal    : {best['Total (AED)']:,.2f} AED\n"
                        f"Delivery : {best['Delivery']}\nTerms    : {best['Terms']}\n\n"
                        f"Please reply '{cfg['approval_keyword'].upper()} {rfq_no}' to approve and issue the GRN.\n\n"
                        f"Regards,\n{cfg['buyer_name']} — UVC Procurement"
                    )
                    try:
                        send_email(
                            token,
                            cfg["manager_email"],
                            f"Quotation for approval — {rfq_no}",
                            body,
                        )
                        st.success(
                            "Sent. Wait for manager's reply, then check approval."
                        )
                    except Exception as e:
                        st.error(str(e))
            if col_b.button("🔎 Check manager approval"):
                token = st.session_state.get("token")
                if not token:
                    st.error("Connect Outlook first.")
                else:
                    try:
                        ok, subj = check_manager_approval(
                            token, cfg["manager_email"], cfg["approval_keyword"], rfq_no
                        )
                        st.session_state["approved"] = ok
                        (
                            st.success("✅ APPROVED — go to tab 5.")
                            if ok
                            else st.warning(
                                f"No '{cfg['approval_keyword']}' reply found for {rfq_no} yet."
                            )
                        )
                    except Exception as e:
                        st.error(str(e))
            if st.session_state.get("approved"):
                st.success("Status: APPROVED ✅")

    # ---- Tab 5 ----
    with t5:
        st.subheader("Goods Received Note")
        items = st.session_state.get("rfq_items")
        mrf = st.session_state.get("mrf", {})
        best = st.session_state.get("best_quote")
        if not items:
            st.info("Complete tabs 1–4 first.")
        else:
            approved = st.session_state.get("approved", False)
            if not approved:
                st.warning("Not approved yet — get manager approval in tab 4.")
            today = datetime.date.today()
            c1, c2 = st.columns(2)
            gseq = c1.text_input("GRN sequence no.", cfg["next_grn_seq"])
            grn_no = f"UVC/GRN/{today:%Y}/{gseq}"
            c2.text_input("GRN No. (full)", grn_no, disabled=True)
            c1, c2 = st.columns(2)
            received_at = c1.selectbox("Received At", ["Site", "Store"], index=0)
            location = c2.text_input("Location", mrf.get("location", ""))
            c1, c2 = st.columns(2)
            quote_ref = c1.text_input("UVC Quote Reference", mrf.get("quote_ref", ""))
            project_site = c2.text_input("Project / Site", mrf.get("project_site", ""))
            lpo_ref = st.text_input("LPO / SO Ref", mrf.get("lpo_ref", ""))
            import pandas as pd

            gdf = pd.DataFrame(
                [
                    {
                        "Item": it["Item"],
                        "Unit": it.get("Unit", ""),
                        "Qty Ordered": it.get("Qty", ""),
                        "Qty Received": it.get("Qty", ""),
                        "Condition": "OK",
                    }
                    for it in items
                ]
            )
            edited = st.data_editor(
                gdf,
                num_rows="dynamic",
                use_container_width=True,
                column_config={
                    "Condition": st.column_config.SelectboxColumn(
                        options=["OK", "Short", "Damaged"]
                    )
                },
            )
            if st.button("📄 Generate GRN", type="primary", disabled=not approved):
                data = build_grn(
                    grn_no,
                    today.strftime("%d/%m/%Y"),
                    received_at,
                    location,
                    quote_ref,
                    project_site,
                    lpo_ref,
                    edited.to_dict("records"),
                )
                st.session_state["grn_bytes"] = data
                cfg["next_grn_seq"] = str(int(gseq) + 1)
                save_cfg(cfg)
                st.success(f"GRN {grn_no} created.")
            if st.session_state.get("grn_bytes"):
                st.download_button(
                    "⬇️ Download GRN (Word)",
                    st.session_state["grn_bytes"],
                    file_name=f"{grn_no.replace('/','-')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )


if __name__ == "__main__":
    ui()
