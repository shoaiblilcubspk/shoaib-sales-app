import streamlit as st
import pandas as pd
import re
from io import BytesIO
from xlwt import Workbook
from datetime import datetime
from jinja2 import Template
from xhtml2pdf import pisa
import requests

st.set_page_config(page_title="Shoaib Sales Tools", layout="centered")

# ----------- SIDEBAR -----------
with st.sidebar:
    st.markdown("## 📞 Contact & Feedback")
    st.write("**App Maker:** Muhammad Shoaib")
    st.write("**Email:** shoaibzaynah@gmail.com")
    st.write("**Phone & WhatsApp:** 0302-7245937")
    st.markdown("---")
    st.markdown("**Feedback ya apne office ke liye aisa tool banwane ke liye Shoaib se rabta karein!**")
    st.info("Note: For custom tools, feedback, or business solutions, contact Shoaib anytime.")

# ----------- API FUNCTIONS -----------
def get_advice():
    try:
        res = requests.get("https://api.adviceslip.com/advice", timeout=5)
        if res.status_code == 200:
            return res.json()['slip']['advice']
    except:
        pass
    return "Think and wonder, wonder and think. — Dr. Seuss"

def get_dollar_rate():
    try:
        res = requests.get("https://open.er-api.com/v6/latest/USD", timeout=5)
        if res.status_code == 200:
            data = res.json()
            return round(data['rates']['PKR'], 2)
    except:
        pass
    return 300  # fallback default

# ----------- MAIN TITLE -----------
st.title("Shoaib Sales Tools")

# ----------- 1. WHATSAPP ORDER TO EXCEL -----------
st.header("📋 WhatsApp Order → Excel (.xls) Generator")

courier = st.selectbox("Select Courier Format:", ["PostEx", "BlueEx"])
whatsapp_data = st.text_area("Paste WhatsApp orders below:", height=180)

DEFAULT_NOTE = "kids clothes - If the number is unavailable, please reach out via WhatsApp"

POSTEX_COLUMNS = [
    "Order Reference Number", "Order Amount", "Order Detail", "Customer Name", "Customer Phone",
    "Order Address", "City", "Items", "Airway Bill Copies", "Notes",
    "Address Code", "Return Address Code", "Order Type (Normal/Reversed/Replacement/Overland)", "Booking Weight"
]

BLUEEX_COLUMNS = [
    "Consignee Name", "Consignee Address", "Consignee Contact No", "Consignee Email", "Product Name",
    "COD", "Pieces", "Weight", "Destination", "Customer Reference", "Customer Comment", "Store Id"
]

def parse_orders(raw_data):
    chunks = raw_data.strip().split("📦")
    orders = []
    for chunk in chunks:
        if not chunk.strip():
            continue
        order = {}
        lines = chunk.strip().splitlines()
        for line in lines:
            if ":" not in line:
                continue
            key, value = line.split(":", 1)
            key = key.strip().lower()
            value = value.strip()
            if key == "name":
                order["name"] = value
            elif key == "phone":
                order["phone"] = re.sub(r"[^\d]", "", value)
            elif key == "city":
                order["city"] = value
            elif key == "address":
                order["address"] = value
            elif key == "id":
                order["id"] = value
            elif key == "total":
                order["amount"] = re.sub(r"[^\d]", "", value)
            elif key == "items":
                order["items"] = value
            elif key == "order type":
                order["order_type"] = value
            elif key == "airway bill copy":
                order["awb"] = value
            elif key == "note":
                order["note"] = value
        orders.append(order)
    return orders

def build_postex_df(orders):
    rows = []
    for o in orders:
        note = o.get("note", "")
        full_note = f"{note} — {DEFAULT_NOTE}" if note else DEFAULT_NOTE
        rows.append([
            o.get("id", ""),
            o.get("amount", "0"),
            DEFAULT_NOTE,
            o.get("name", ""),
            o.get("phone", ""),
            o.get("address", ""),
            o.get("city", ""),
            o.get("items", "1"),
            o.get("awb", "1"),
            full_note,
            "", "",
            o.get("order_type", "Normal"),
            "0.5"
        ])
    return pd.DataFrame(rows, columns=POSTEX_COLUMNS)

def build_blueex_df(orders):
    rows = []
    for o in orders:
        note = o.get("note", "")
        full_note = f"{note} — {DEFAULT_NOTE}" if note else DEFAULT_NOTE
        rows.append([
            o.get("name", ""),
            o.get("address", ""),
            o.get("phone", ""),
            "",
            "Kids Clothes",
            o.get("amount", "0"),
            o.get("items", "1"),
            "0.5",
            o.get("city", ""),
            o.get("id", ""),
            full_note,
            ""
        ])
    return pd.DataFrame(rows, columns=BLUEEX_COLUMNS)

def create_xls(df):
    wb = Workbook()
    ws = wb.add_sheet("Orders")
    for col_index, col_name in enumerate(df.columns):
        ws.write(0, col_index, col_name)
    for row_index, row in enumerate(df.itertuples(index=False), start=1):
        for col_index, cell in enumerate(row):
            ws.write(row_index, col_index, str(cell))
    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

if st.button("Generate Excel File"):
    if not whatsapp_data.strip():
        st.warning("Please paste some order data first.")
    else:
        orders = parse_orders(whatsapp_data)
        df = build_postex_df(orders) if courier == "PostEx" else build_blueex_df(orders)
        output_file = create_xls(df)
        filename = f"{courier.lower()}_orders_{datetime.now().strftime('%Y%m%d')}.xls"
        st.download_button(
            label=f"Download {courier} .xls File",
            data=output_file,
            file_name=filename,
            mime="application/vnd.ms-excel"
        )

# ----------- PROMPT EXPANDER -----------
with st.expander("📋 WhatsApp Parcel Data Formatting Prompt – FINAL JULY 2025"):
    st.code(
"""## 📋 WhatsApp Parcel Data Formatting Prompt – FINAL JULY 2025
### 📦 PARCEL OUTPUT FORMAT
📦 Parcel #  
Name: [Customer Name]  
Phone: [03XXXXXXXXX]  
City: [Valid City Name, auto-corrected if needed]  
Address: [Roman Urdu only. If Urdu or mix, convert to Roman Urdu. Use English for "near", "opposite", "beside", "behind", "in front of" etc. City name at end.]  
ID: [Order ID]  
Total: [Amount]  
Items: [Number of Items]  
Order Type: [Normal / Replacement]  
Airway Bill Copy: [1 or 0]  
Note: [Leave blank if not given]
➡️ Repeat this format for Parcel 2, Parcel 3, ...
---
## ✅ FORMATTING RULES
---
### 👤 NAME  
- Properly capitalized: `usman malik` → `Usman Malik`
---
### 📞 PHONE  
- Always format as: **03XXXXXXXXX**  
- If multiple numbers given:  
  - First number → Phone field  
  - Additional numbers → Move to end of **Address**: `(Alternate: 03XXXXXXXXX)`
---
### 🏙️ CITY DETECTION & VALIDATION  
1. **If city list is not provided**, use AI to detect and correct city names from Pakistan’s major cities (e.g. Lahore, Karachi, Islamabad, Faisalabad, Multan, Rawalpindi, Sialkot, Gujranwala, Peshawar, Quetta, etc.).
2. **If city spelling is wrong** (e.g. "Lahor", "Karach", "Fsd", "Isb"), auto-correct to the proper city name.
3. **If city is missing or invalid**, try to extract from address.
4. **If city cannot be determined**, leave city empty and flag in summary.
---
### 🏠 ADDRESS  
- If address is in Urdu or mixed, **convert to Roman Urdu**.
- For location words, always use **English**:  
  - "near" (not "qareeb")  
  - "opposite" (not "samne")  
  - "beside" (not "saath")  
  - "behind" (not "peechay")  
  - "in front of" (not "samne")  
- **City name must always be at the end** (e.g. `dha phase 1 lahore`)
- Add alternate number at end if moved from Phone
- Avoid redundant city names and extra info
---
### 💵 TOTAL / AMOUNT  
- If value is missing → Set: `Total: 0`  
- If payment is **Advance / Paid / Transfer / Non-COD** → Also set `Total: 0` and log  
- Always use plain number (no “Rs.” prefix)
---
### 📦 ITEMS  
- Should be numeric  
- If missing → Default to `Items: 3` and log in summary
---
### 📝 NOTE  
- Use if there’s a customer note (e.g. “Call before delivery”)  
- Leave blank if nothing was provided
---
### 🔄 RETURN / REPLACEMENT/CHANGING DETECTION  
If any of these keywords appear:  
> **Return**, **Replacement**, **Replacing**, **Exchange**, **Changing**, **Change**, **Reverse**, **Wrong Item**, **Wrong Product**, **Sending Back**
✅ **Order Type must always be:**  
- `Order Type: Replacement`  
- (Never "Normal" in these cases)
✅ Log in **summary** as:  
- `📦 Parcel X — Mentioned "Replacement"`  
- `📦 Parcel X — Mentioned "Changing"`  
- etc.
---
## 🧾 NON-COD ORDERS
If input includes keywords:  
> Advance Paid, Already Paid, Bank Transfer, Non-COD  
Then:  
- Set `Total: 0`  
- Keep `Order Type: Normal` (unless above keywords for replacement etc.)  
- Log `📦 Parcel X — Marked as Non-COD (Total set to 0)` in summary
---
## 🛠 DEFAULT ASSUMPTIONS (Handle Missing Fields)
If missing:  
- **Total / Amount** → Set: `Total: 0`  
- **Items** → Set: `Items: 3`  
- **Airway Bill Copy** → Set: `1`  
- **Note** → Leave blank  
- Log all such defaults in the summary clearly.
---
## 📌 SUMMARY FORMAT — AT END OF ALL ORDER OUTPUT
Summarize all flags and adjustments like this:
📦 Parcel 1 — Total missing → set to 0, Items missing → set to 3  
📦 Parcel 2 — Mentioned "Return"  
📦 Parcel 4 — City “Wadala Sandwa” replaced with “Gujranwala” (from address)  
📦 Parcel 5 — Marked as Non-COD (Total set to 0)  
📦 Parcel 6 — Invalid phone format corrected  
📦 Parcel 8 — Airway Bill Copy missing → set to 0  
📦 Parcel 9 — City spelling corrected from "Lahor" to "Lahore"  
📦 Parcel 10 — Address converted from Urdu to Roman Urdu  
➕ **Total Parcels: [X]**
---
### 🏠 ADDRESS EXAMPLES
**Input:**  
Address: پاکت نامه نازی، خالو، تحصیل غازی، ضلع ہری پور, Haripur
**Output:**  
Address: Pakt Nama Nazi, Khalu, Tehsil Ghazi, Zila Haripur, Haripur
---
**Input:**  
Address: گھر کے قریب جامع مسجد، نزد چوک، لاہور
**Output:**  
Address: Ghar ke near Jamia Masjid, near Chowk, Lahore
---
**Input:**  
Address: opposite school, street 5, Multan
**Output:**  
Address: Opposite School, Street 5, Multan
---
**Input:**  
Address: samne bank, pichlay gali, Karachi
**Output:**  
Address: In front of Bank, Behind Gali, Karachi
---""",
        language="markdown"
    )

# ----------- 2. LILCUBS DAILY REPORT -----------
st.header("📋 LilCubs Daily Report")

dollar_rate_api = get_dollar_rate()
dollar_rate = st.number_input("Dollar Rate (PKR):", value=dollar_rate_api)
sales_data = st.text_area("Paste your sales data:", height=180)

with st.expander("📝 AI Data Cleaning Prompt"):
    st.code(
        """Align and clean the following sales data into this format (one per line, semicolon at end):

Name sale [amount] return [amount] add ya ad [amount] usd ya phir [amount] pkr ;

Data:
[paste your messy/random data here]""",
        language="markdown"
    )

pattern = re.compile(
    r"(\w+)\s+sale\s+(\d+)\s+return\s+(\d+)\s+(?:add\s+ya\s+ad|add|ad)\s+([\d.]+)\s+usd",
    re.IGNORECASE
)

if st.button("Generate File"):
    matches = pattern.findall(sales_data)
    if not matches:
        st.error("No valid data found. Please check your input format.")
    else:
        advice = get_advice()
        df = pd.DataFrame(matches, columns=["Name", "Sale", "Return", "Advertise"])
        df["Sale"] = df["Sale"].astype(float)
        df["Return"] = df["Return"].astype(float)
        df["Advertise"] = df["Advertise"].astype(float)
        df["Net Total"] = df["Sale"] - df["Return"]

        total_row = {
            "Name": "Total",
            "Sale": df["Sale"].sum(),
            "Return": df["Return"].sum(),
            "Advertise": df["Advertise"].sum(),
            "Net Total": df["Net Total"].sum()
        }
        df_totals = pd.DataFrame([total_row])
        df_final = pd.concat([df, df_totals], ignore_index=True)

        try:
            return_percent = (df["Return"].sum() / df["Sale"].sum()) * 100 if df["Sale"].sum() > 0 else 0
        except:
            return_percent = 0

        # Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_final[["Name", "Sale", "Return", "Advertise", "Net Total"]].to_excel(writer, index=False, sheet_name='Report')
            worksheet = writer.sheets['Report']
            worksheet.write(len(df_final)+2, 0, "Total Sales:")
            worksheet.write(len(df_final)+2, 1, total_row["Sale"])
            worksheet.write(len(df_final)+3, 0, "Total Return:")
            worksheet.write(len(df_final)+3, 1, total_row["Return"])
            worksheet.write(len(df_final)+4, 0, "Total Add:")
            worksheet.write(len(df_final)+4, 1, total_row["Advertise"])
            worksheet.write(len(df_final)+5, 0, "Total Net:")
            worksheet.write(len(df_final)+5, 1, total_row["Net Total"])
            worksheet.write(len(df_final)+6, 0, "Total Ad Spent (PKR):")
            worksheet.write(len(df_final)+6, 1, round(total_row["Advertise"] * dollar_rate, 2))
            worksheet.write(len(df_final)+7, 0, "Return Percentage:")
            worksheet.write(len(df_final)+7, 1, f"{round(return_percent,2)}%")
            worksheet.write(len(df_final)+9, 0, "Copyright by (Shoaib Z)")
        output.seek(0)
        st.download_button(label="Download Excel File (.xlsx)", data=output.getvalue(), file_name="sales_report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # PDF
        pdf_template = Template("""
        <html>
          <body>
            <h2 style="text-align: center;">LilCubs Daily Report</h2>
            <p style="text-align: center; margin-bottom: 0;"><strong>Date:</strong> {{ date }}<br>
            <strong>Dollar Rate:</strong> {{ dollar_rate }}</p>
            <p style="text-align: center; font-style: italic; color: #444; margin-top: 10px; margin-bottom: 20px;">
              {{ advice }}
            </p>
            <table border='1' cellpadding='5' cellspacing='0' width="100%" style="border-collapse:collapse;">
              <tr>
                {% for col in cols %}<th>{{ col }}</th>{% endfor %}
              </tr>
              {% for row in data %}
              <tr>
                {% for item in row %}<td>{{ item }}</td>{% endfor %}
              </tr>
              {% endfor %}
            </table>
            <br>
            <p><strong>Total Sales:</strong> {{ total_sales }}<br>
            <strong>Total Return:</strong> {{ total_return }}<br>
            <strong>Total Add:</strong> {{ total_add }}<br>
            <strong>Total Net:</strong> {{ total_net }}<br>
            <strong>Total Ad Spent (PKR):</strong> {{ total_add_pkr }}<br>
            <strong>Return Percentage:</strong> {{ return_percent }}%</p>
            <p style="text-align:center;font-size:10px;color:gray; margin-top:40px;">Copyright by (Shoaib Z)</p>
          </body>
        </html>
        """)
        html = pdf_template.render(
            cols=["Name", "Sale", "Return", "Advertise", "Net Total"],
            data=df_final[["Name", "Sale", "Return", "Advertise", "Net Total"]].values.tolist(),
            dollar_rate=dollar_rate,
            total_sales=total_row["Sale"],
            total_return=total_row["Return"],
            total_add=total_row["Advertise"],
            total_net=total_row["Net Total"],
            total_add_pkr=round(total_row["Advertise"] * dollar_rate, 2),
            return_percent=round(return_percent, 2),
            date=datetime.now().strftime("%A, %d %B, %Y"),
            advice=advice
        )
        pdf_output = BytesIO()
        pisa.CreatePDF(html, dest=pdf_output)
        pdf_output.seek(0)
        st.download_button("Download PDF (A4)", data=pdf_output.getvalue(), file_name="sales_report.pdf", mime="application/pdf")