"""
Choice Restaurant Parser — Flask Backend
Запуск: python app.py
Відкрий: http://localhost:5000
"""

from flask import Flask, request, jsonify, send_file, send_from_directory
import requests as req
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
import re, time, os, io

app = Flask(__name__, static_folder=".")

API_KEY = os.environ.get("API_KEY", "")   # береться зі змінної середовища

# ─── PLACES API ───────────────────────────────────────────────────────────────

def text_search(query, city, page_token=None):
    url = "https://maps.googleapis.com/maps/api/place/textsearch/json"
    params = {"query": f"{query} in {city}", "key": API_KEY, "type": "restaurant"}
    if page_token:
        params["pagetoken"] = page_token
    r = req.get(url, params=params, timeout=10)
    return r.json()


def get_place_details(place_id):
    url = "https://maps.googleapis.com/maps/api/place/details/json"
    params = {
        "place_id": place_id, "key": API_KEY,
        "fields": "name,formatted_address,formatted_phone_number,website,rating,user_ratings_total,types,url"
    }
    r = req.get(url, params=params, timeout=10)
    return r.json().get("result", {})


def find_instagram(website_url):
    if not website_url:
        return ""
    try:
        r = req.get(website_url, headers={"User-Agent": "Mozilla/5.0"}, timeout=5)
        matches = re.findall(
            r'https?://(?:www\.)?instagram\.com/([A-Za-z0-9._]+)(?:[/?"\']|$)', r.text
        )
        clean = [m for m in matches if not m.startswith(("p/", "reel", "explore", "sharer"))]
        if clean:
            return f"https://instagram.com/{clean[0]}"
    except Exception:
        pass
    return ""


IGNORE_TYPES = {"point_of_interest", "establishment", "food", "premise", "store"}
CUISINE_MAP = {
    "restaurant": "Ресторан", "cafe": "Кафе", "bar": "Бар",
    "bakery": "Пекарня", "meal_takeaway": "Takeaway", "meal_delivery": "Delivery",
    "pizza": "Піцерія", "night_club": "Нічний клуб", "brewery": "Пивоварня",
}

def get_cuisine(types_list):
    result = [CUISINE_MAP.get(t, t.replace("_", " ").title())
              for t in (types_list or []) if t not in IGNORE_TYPES]
    return ", ".join(result[:3])


def detect_platform(website):
    platforms = ["choiceqr", "glovo", "bolt", "uber", "wolt", "pyszne", "lieferando"]
    for p in platforms:
        if p in (website or "").lower():
            return p.capitalize()
    return ""


# ─── EXCEL ───────────────────────────────────────────────────────────────────

HEADERS = ["Назва", "Адреса", "Телефон", "Сайт", "Instagram",
           "Рейтинг", "Відгуків", "Тип кухні", "Онлайн-замовлення", "Google Maps"]

def build_excel(rows, city):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"{city} Restaurants"
    ws.freeze_panes = "A2"

    col_widths = [28, 38, 16, 32, 32, 10, 10, 22, 20, 38]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    for col, h in enumerate(HEADERS, 1):
        c = ws.cell(row=1, column=col, value=h)
        c.font = Font(bold=True, color="FFFFFF", size=11)
        c.fill = PatternFill("solid", fgColor="1A1A2E")
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    for ri, row in enumerate(rows, 2):
        fill = "F0F4FF" if ri % 2 == 0 else "FFFFFF"
        for ci, val in enumerate(row, 1):
            c = ws.cell(row=ri, column=ci, value=val)
            c.fill = PatternFill("solid", fgColor=fill)
            c.alignment = Alignment(vertical="center", wrap_text=True)
            if ci in (4, 5, 10) and val and str(val).startswith("http"):
                c.hyperlink = val
                c.font = Font(color="0563C1", underline="single")
        ws.row_dimensions[ri].height = 20

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ─── ROUTES ──────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return send_from_directory(".", "index.html")


@app.route("/search", methods=["POST"])
def search():
    data = request.json
    city     = data.get("city", "").strip()
    category = data.get("category", "restaurants").strip()
    limit    = min(int(data.get("limit", 20)), 60)

    if not city:
        return jsonify({"error": "Вкажіть місто"}), 400
    if not API_KEY or API_KEY == "ВАШ_GOOGLE_PLACES_API_KEY":
        return jsonify({"error": "API ключ не вказано в app.py"}), 400

    place_ids = []
    page_token = None
    for _ in range(3):
        resp = text_search(category, city, page_token)
        if resp.get("status") not in ("OK", "ZERO_RESULTS"):
            return jsonify({"error": f"Google API: {resp.get('status')} — {resp.get('error_message','')}"}), 500
        place_ids.extend([r["place_id"] for r in resp.get("results", [])])
        page_token = resp.get("next_page_token")
        if not page_token or len(place_ids) >= limit:
            break
        time.sleep(2)

    place_ids = place_ids[:limit]
    results = []
    for pid in place_ids:
        d = get_place_details(pid)
        results.append({
            "name":     d.get("name", ""),
            "address":  d.get("formatted_address", ""),
            "phone":    d.get("formatted_phone_number", ""),
            "website":  d.get("website", ""),
            "instagram": find_instagram(d.get("website", "")),
            "rating":   d.get("rating", ""),
            "reviews":  d.get("user_ratings_total", ""),
            "cuisine":  get_cuisine(d.get("types", [])),
            "platform": detect_platform(d.get("website", "")),
            "maps_url": d.get("url", ""),
        })
        time.sleep(0.3)

    return jsonify({"count": len(results), "results": results})


@app.route("/export", methods=["POST"])
def export():
    data    = request.json
    rows    = data.get("rows", [])
    city    = data.get("city", "export")
    excel_rows = [
        [r["name"], r["address"], r["phone"], r["website"], r["instagram"],
         r["rating"], r["reviews"], r["cuisine"], r["platform"], r["maps_url"]]
        for r in rows
    ]
    buf = build_excel(excel_rows, city)
    filename = f"restaurants_{city.lower().replace(' ', '_')}.xlsx"
    return send_file(buf, as_attachment=True, download_name=filename,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


if __name__ == "__main__":
    print("🚀 Відкрий http://localhost:5000")
    app.run(debug=True, port=5000)
