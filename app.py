import os
from datetime import datetime
from flask import Flask, render_template_string, request, redirect, url_for, session, send_file
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'ukay_live_secret_key_2026')


def get_user_file(username):
    """Generates a separate Excel filename for each store based on current date (YYYY-MM-DD)."""
    today_str = datetime.now().strftime("%Y-%m-%d")
    safe_username = "".join(c for c in username if c.isalnum() or c in ('_', '-')).lower()
    return f"ukay_inventory_{safe_username}_{today_str}.xlsx"


def init_excel(filepath):
    """Initializes the Excel workbook if it doesn't exist yet for today."""
    if not os.path.exists(filepath):
        wb = Workbook()
        ws = wb.active
        ws.title = "Inventory"
        
        # Header Row
        headers = ["No.", "Code", "Buyer Name", "Item", "Price"]
        ws.append(headers)
        
        # Style Header Row
        header_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True)
        thin_border = Border(
            left=Side(style='thin', color='000000'),
            right=Side(style='thin', color='000000'),
            top=Side(style='thin', color='000000'),
            bottom=Side(style='thin', color='000000')
        )

        for col_num in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border

        wb.save(filepath)


def rebuild_and_format_excel(filepath):
    """Applies borders, auto-fits columns, auto-increments #, and adds SUM/COUNTA formulas."""
    if not os.path.exists(filepath):
        return

    wb = load_workbook(filepath)
    ws = wb["Inventory"]

    # Extract data rows (ignoring header and any previous summary rows)
    raw_rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row and row[1] is not None and str(row[0]).strip() != "TOTAL ITEMS SOLD:":
            # Extract Code, Name, Item, Price
            raw_rows.append((row[1], row[2], row[3], float(row[4]) if row[4] else 0.0))

    # Reset sheet
    ws.delete_rows(1, ws.max_row)

    # Re-add Headers
    headers = ["#", "Code", "Buyer Name", "Item", "Price"]
    ws.append(headers)

    # Styling definitions
    thin_border = Border(
        left=Side(style='thin', color='D3D3D3'),
        right=Side(style='thin', color='D3D3D3'),
        top=Side(style='thin', color='D3D3D3'),
        bottom=Side(style='thin', color='D3D3D3')
    )
    header_fill = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    bold_font = Font(name="Calibri", size=11, bold=True)

    # Apply Header Styles
    for col_num in range(1, 6):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Add Data Rows with Auto-Increment Number
    current_row = 2
    for idx, (code, name, item, price) in enumerate(raw_rows, start=1):
        ws.append([idx, code, name, item, price])
        
        # Apply cell borders and number formatting
        ws.cell(row=current_row, column=1).alignment = Alignment(horizontal="center")
        ws.cell(row=current_row, column=2).alignment = Alignment(horizontal="center")
        ws.cell(row=current_row, column=5).number_format = '₱#,##0.00'

        for col_num in range(1, 6):
            ws.cell(row=current_row, column=col_num).border = thin_border
            
        current_row += 1

    last_data_row = current_row - 1

    if last_data_row >= 2:
        # Blank row before summary
        ws.append([])
        summary_start_row = current_row + 1

        # Total Items Sold Row
        ws.cell(row=summary_start_row, column=3, value="TOTAL ITEMS SOLD:").font = bold_font
        ws.cell(row=summary_start_row, column=3).alignment = Alignment(horizontal="right")
        sold_cell = ws.cell(row=summary_start_row, column=5, value=f"=COUNTA(A2:A{last_data_row})")
        sold_cell.font = bold_font

        # Total Sales Row
        ws.cell(row=summary_start_row + 1, column=3, value="TOTAL SALES:").font = bold_font
        ws.cell(row=summary_start_row + 1, column=3).alignment = Alignment(horizontal="right")
        sales_cell = ws.cell(row=summary_start_row + 1, column=5, value=f"=SUM(E2:E{last_data_row})")
        sales_cell.font = bold_font
        sales_cell.number_format = '₱#,##0.00'

    # Auto-adjust column widths
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    wb.save(filepath)


def read_user_items(filepath):
    """Reads non-summary items to show on web UI table."""
    items = []
    if os.path.exists(filepath):
        wb = load_workbook(filepath, data_only=True)
        ws = wb["Inventory"]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row and row[0] and str(row[0]).isdigit():
                # (#, Code, Buyer, Item, Price)
                items.append((row[0], row[1], row[2], row[3], row[4]))
    return items


HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Ukay Live Inventory</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <style>
        :root {
            --bg-gradient: linear-gradient(135deg, #0f172a 0%, #1e1b4b 100%);
            --card-bg: #1e293b;
            --input-bg: #0f172a;
            --primary: #6366f1;
            --primary-hover: #4f46e5;
            --success: #10b981;
            --success-hover: #059669;
            --accent: #ec4899;
            --text-main: #f8fafc;
            --text-muted: #94a3b8;
            --border-color: #334155;
        }

        * { box-sizing: border-box; margin: 0; padding: 0; font-family: 'Inter', sans-serif; }

        body {
            background: var(--bg-gradient);
            color: var(--text-main);
            min-height: 100vh;
            display: flex;
            justify-content: center;
            align-items: center;
            padding: 20px 15px;
        }

        .container {
            width: 100%;
            max-width: 520px;
            background: var(--card-bg);
            padding: 30px 25px;
            border-radius: 16px;
            box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.5);
            border: 1px solid var(--border-color);
        }

        .header-title { text-align: center; font-size: 24px; font-weight: 700; }
        .header-subtitle { text-align: center; color: var(--text-muted); font-size: 14px; margin: 6px 0 22px; }

        .user-tag {
            text-align: center;
            background: rgba(99, 102, 241, 0.1);
            border: 1px solid rgba(99, 102, 241, 0.2);
            color: #a5b4fc;
            padding: 8px 12px;
            border-radius: 8px;
            font-size: 13px;
            margin-bottom: 20px;
        }
        .user-tag a { color: var(--accent); text-decoration: none; font-weight: 600; margin-left: 5px; }

        .form-group { margin-bottom: 14px; }
        label { font-size: 12px; font-weight: 600; text-transform: uppercase; color: var(--text-muted); display: block; margin-bottom: 6px; }

        input[type="text"], input[type="number"] {
            width: 100%;
            padding: 12px 14px;
            background: var(--input-bg);
            border: 1px solid var(--border-color);
            border-radius: 10px;
            font-size: 15px;
            color: var(--text-main);
        }

        .btn {
            width: 100%;
            padding: 13px;
            border: none;
            border-radius: 10px;
            font-size: 15px;
            font-weight: 600;
            cursor: pointer;
            text-decoration: none;
            display: inline-block;
            text-align: center;
        }
        .btn-green { background: var(--success); color: white; margin-top: 10px; }
        .btn-purple { background: var(--primary); color: white; margin-top: 12px; }

        .alert-success {
            background: rgba(16, 185, 129, 0.15);
            border: 1px solid rgba(16, 185, 129, 0.3);
            color: #6ee7b7;
            padding: 10px;
            border-radius: 8px;
            text-align: center;
            margin-bottom: 16px;
        }

        .table-wrapper { margin-top: 25px; border-top: 1px solid var(--border-color); padding-top: 20px; }
        .table-header-title { font-size: 16px; font-weight: 600; margin-bottom: 12px; display: flex; justify-content: space-between; }
        .badge-count { background: var(--primary); color: white; font-size: 12px; padding: 2px 8px; border-radius: 12px; }

        table { width: 100%; border-collapse: collapse; font-size: 13px; }
        th, td { padding: 8px 6px; text-align: left; border-bottom: 1px solid var(--border-color); }
        th { background: var(--input-bg); color: var(--text-muted); font-size: 11px; }
        .price-text { color: #34d399; font-weight: 600; }
        .date-badge { display: inline-block; font-size: 11px; background: #334155; padding: 2px 6px; border-radius: 4px; margin-left: 5px; }
    </style>
</head>
<body>
    <div class="container">
        {% if not username %}
            <h2 class="header-title">🛍️ Ukay Live Inventory</h2>
            <p class="header-subtitle">Enter store name to start today's session</p>

            <form method="POST" action="/login">
                <div class="form-group">
                    <label>Store Name / Seller ID</label>
                    <input type="text" name="username" placeholder="e.g., Baguio_Ukay_Store" required autofocus>
                </div>
                <button type="submit" class="btn btn-green">Start Today's Session</button>
            </form>
        {% else %}
            <h2 class="header-title">⚡ Live Item Entry</h2>
            <div class="user-tag">
                Store: <strong>{{ username }}</strong> <span class="date-badge">{{ today_date }}</span> | <a href="/logout">Switch Store</a>
            </div>

            {% if msg %}
                <div class="alert-success">✓ {{ msg }}</div>
            {% endif %}

            <form method="POST" action="/add">
                <div class="form-group">
                    <label>Item Code</label>
                    <input type="text" name="code" placeholder="e.g., A01" required autofocus autocomplete="off">
                </div>

                <div class="form-group">
                    <label>Buyer Name</label>
                    <input type="text" name="name" placeholder="e.g., Emilyn" required autocomplete="off">
                </div>

                <div class="form-group">
                    <label>Item Description</label>
                    <input type="text" name="item" placeholder="e.g., Denim Jacket" required autocomplete="off">
                </div>

                <div class="form-group">
                    <label>Price (PHP)</label>
                    <input type="number" step="0.01" name="price" placeholder="e.g., 250" required autocomplete="off">
                </div>

                <button type="submit" class="btn btn-green">Save Item</button>
            </form>

            <a href="/download" class="btn btn-purple">📥 Download Today's Excel File</a>

            {% if items %}
                <div class="table-wrapper">
                    <div class="table-header-title">
                        Today's Saved Items
                        <span class="badge-count">{{ items|length }}</span>
                    </div>
                    <table>
                        <thead>
                            <tr>
                                <th>#</th>
                                <th>Code</th>
                                <th>Buyer</th>
                                <th>Item</th>
                                <th>Price</th>
                            </tr>
                        </thead>
                        <tbody>
                            {% for row in items|reverse %}
                            <tr>
                                <td><strong>{{ row[0] }}</strong></td>
                                <td><strong>{{ row[1] }}</strong></td>
                                <td>{{ row[2] }}</td>
                                <td>{{ row[3] }}</td>
                                <td class="price-text">₱{{ "%.2f"|format(row[4]) }}</td>
                            </tr>
                            {% endfor %}
                        </tbody>
                    </table>
                </div>
            {% endif %}
        {% endif %}
    </div>
</body>
</html>
"""


@app.route('/')
def home():
    username = session.get('username')
    msg = request.args.get('msg', '')
    items = []
    today_date = datetime.now().strftime("%B %d, %Y")

    if username:
        filepath = get_user_file(username)
        init_excel(filepath)
        items = read_user_items(filepath)

    return render_template_string(HTML_TEMPLATE, username=username, msg=msg, items=items, today_date=today_date)


@app.route('/login', methods=['POST'])
def login():
    username = request.form.get('username', '').strip()
    if username:
        session['username'] = username
        filepath = get_user_file(username)
        init_excel(filepath)
    return redirect(url_for('home'))


@app.route('/logout')
def logout():
    session.pop('username', None)
    return redirect(url_for('home'))


@app.route('/add', methods=['POST'])
def add_item():
    username = session.get('username')
    if not username:
        return redirect(url_for('home'))

    code = request.form.get('code')
    name = request.form.get('name')
    item = request.form.get('item')
    price = request.form.get('price')

    filepath = get_user_file(username)
    init_excel(filepath)

    # Append raw row first
    wb = load_workbook(filepath)
    ws = wb["Inventory"]
    ws.append([0, code, name, item, float(price)])
    wb.save(filepath)

    # Rebuild formatting, auto-increment numbers, borders & formulas
    rebuild_and_format_excel(filepath)

    return redirect(url_for('home', msg=f"Saved: #{code} for {name}!"))


@app.route('/download')
def download():
    username = session.get('username')
    if not username:
        return redirect(url_for('home'))

    filepath = get_user_file(username)
    if os.path.exists(filepath):
        return send_file(filepath, as_attachment=True)
    return redirect(url_for('home'))


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
