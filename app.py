import os
import glob
from datetime import datetime, timedelta
from flask import Flask, render_template_string, request, redirect, url_for, session, send_file
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'ukay_live_secret_key_2026')


def get_ph_time():
    """Returns the current datetime in Philippine Standard Time (UTC+8)"""
    return datetime.utcnow() + timedelta(hours=8)


def get_user_file_by_date(username, date_str):
    """Generates file name for a specific YYYY-MM-DD date."""
    safe_username = "".join(c for c in username if c.isalnum() or c in ('_', '-')).lower()
    return f"ukay_inventory_{safe_username}_{date_str}.xlsx"


def get_user_file(username):
    """Generates a separate Excel filename for current date (YYYY-MM-DD)."""
    today_str = get_ph_time().strftime("%Y-%m-%d")
    return get_user_file_by_date(username, today_str)


def init_excel(filepath):
    """Initializes the Excel workbook if it doesn't exist yet."""
    if not os.path.exists(filepath):
        wb = Workbook()
        ws = wb.active
        ws.title = "Inventory"
        
        headers = ["#", "Code", "Buyer Name", "Item", "Price", "Status"]
        ws.append(headers)
        
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
    """Applies borders, auto-fits columns, auto-increments #, and adds SUMIF/COUNTIF formulas."""
    if not os.path.exists(filepath):
        return

    wb = load_workbook(filepath)
    ws = wb["Inventory"]

    raw_rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row and row[1] is not None and str(row[0]).strip() not in ["TOTAL ITEMS SOLD:", "TOTAL PAID SALES:", "TOTAL CANCELLED:", "TOTAL PENDING:"]:
            status = str(row[5]).strip() if len(row) > 5 and row[5] else "Pending"
            raw_rows.append((row[1], row[2], row[3], float(row[4]) if row[4] else 0.0, status))

    ws.delete_rows(1, ws.max_row)

    headers = ["#", "Code", "Buyer Name", "Item", "Price", "Status"]
    ws.append(headers)

    thin_border = Border(
        left=Side(style='thin', color='D3D3D3'),
        right=Side(style='thin', color='D3D3D3'),
        top=Side(style='thin', color='D3D3D3'),
        bottom=Side(style='thin', color='D3D3D3')
    )
    header_fill = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    bold_font = Font(name="Calibri", size=11, bold=True)

    for col_num in range(1, 7):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    current_row = 2
    for idx, (code, name, item, price, status) in enumerate(raw_rows, start=1):
        ws.append([idx, code, name, item, price, status])
        
        ws.cell(row=current_row, column=1).alignment = Alignment(horizontal="center")
        ws.cell(row=current_row, column=2).alignment = Alignment(horizontal="center")
        ws.cell(row=current_row, column=5).number_format = '₱#,##0.00'
        ws.cell(row=current_row, column=6).alignment = Alignment(horizontal="center")

        for col_num in range(1, 7):
            ws.cell(row=current_row, column=col_num).border = thin_border
            
        current_row += 1

    last_data_row = current_row - 1

    if last_data_row >= 2:
        ws.append([])
        summary_start_row = current_row + 1

        # Total Items
        ws.cell(row=summary_start_row, column=3, value="TOTAL ITEMS SOLD:").font = bold_font
        ws.cell(row=summary_start_row, column=3).alignment = Alignment(horizontal="right")
        sold_cell = ws.cell(row=summary_start_row, column=5, value=f'=COUNTIF(F2:F{last_data_row}, "Paid")')
        sold_cell.font = bold_font

        # Total Paid Sales
        ws.cell(row=summary_start_row + 1, column=3, value="TOTAL PAID SALES:").font = bold_font
        ws.cell(row=summary_start_row + 1, column=3).alignment = Alignment(horizontal="right")
        sales_cell = ws.cell(row=summary_start_row + 1, column=5, value=f'=SUMIF(F2:F{last_data_row}, "Paid", E2:E{last_data_row})')
        sales_cell.font = bold_font
        sales_cell.number_format = '₱#,##0.00'

        # Total Cancelled
        ws.cell(row=summary_start_row + 2, column=3, value="TOTAL CANCELLED:").font = bold_font
        ws.cell(row=summary_start_row + 2, column=3).alignment = Alignment(horizontal="right")
        cancel_cell = ws.cell(row=summary_start_row + 2, column=5, value=f'=COUNTIF(F2:F{last_data_row}, "Cancelled")')
        cancel_cell.font = bold_font

        # Total Pending
        ws.cell(row=summary_start_row + 3, column=3, value="TOTAL PENDING:").font = bold_font
        ws.cell(row=summary_start_row + 3, column=3).alignment = Alignment(horizontal="right")
        pending_cell = ws.cell(row=summary_start_row + 3, column=5, value=f'=COUNTIF(F2:F{last_data_row}, "Pending")')
        pending_cell.font = bold_font

    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    wb.save(filepath)


def read_user_items(filepath):
    """Reads inventory rows from an Excel file."""
    items = []
    if os.path.exists(filepath):
        wb = load_workbook(filepath, data_only=True)
        ws = wb["Inventory"]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row and row[0] and str(row[0]).isdigit():
                status = str(row[5]) if len(row) > 5 and row[5] else "Pending"
                items.append({
                    'num': row[0],
                    'code': row[1],
                    'name': row[2],
                    'item': row[3],
                    'price': row[4] or 0.0,
                    'status': status
                })
    return items


def update_item_status_in_excel(filepath, item_num, new_status):
    """Updates the status of a specific row by its auto-increment ID (#)."""
    if not os.path.exists(filepath):
        return
    
    wb = load_workbook(filepath)
    ws = wb["Inventory"]
    for row in range(2, ws.max_row + 1):
        cell_val = ws.cell(row=row, column=1).value
        if cell_val is not None and str(cell_val).isdigit() and int(cell_val) == int(item_num):
            ws.cell(row=row, column=6, value=new_status)
            break
    wb.save(filepath)
    rebuild_and_format_excel(filepath)


def get_monthly_report_data(username, year_str, month_str):
    """Calculates Total Sales, Items Sold, Cancelled Items, and Pending Items for a specified month."""
    safe_username = "".join(c for c in username if c.isalnum() or c in ('_', '-')).lower()
    pattern = f"ukay_inventory_{safe_username}_{year_str}-{month_str}-*.xlsx"
    files = glob.glob(pattern)
    
    total_sales = 0.0
    total_items_sold = 0
    total_cancelled = 0
    total_pending = 0
    file_count = len(files)

    for file in files:
        items = read_user_items(file)
        for item in items:
            if item['status'] == 'Paid':
                total_sales += float(item['price'])
                total_items_sold += 1
            elif item['status'] == 'Cancelled':
                total_cancelled += 1
            elif item['status'] == 'Pending':
                total_pending += 1

    return {
        'total_sales': total_sales,
        'total_items_sold': total_items_sold,
        'total_cancelled': total_cancelled,
        'total_pending': total_pending,
        'file_count': file_count
    }


HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Ukay Live Inventory Management</title>
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
            --warning: #f59e0b;
            --danger: #ef4444;
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
            align-items: {% if not username %}center{% else %}flex-start{% endif %};
            padding: 30px 15px;
        }

        .container {
            width: 100%;
            max-width: {% if not username %}420px{% else %}850px{% endif %};
            background: var(--card-bg);
            padding: 28px 24px;
            border-radius: 16px;
            box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.5);
            border: 1px solid var(--border-color);
        }

        .header-title { text-align: center; font-size: 24px; font-weight: 700; }
        .header-subtitle { text-align: center; color: var(--text-muted); font-size: 14px; margin: 6px 0 20px; }

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
        .user-tag a { color: var(--accent); text-decoration: none; font-weight: 600; margin-left: 8px; }

        /* Navigation Tabs */
        .tabs {
            display: flex;
            gap: 8px;
            border-bottom: 2px solid var(--border-color);
            margin-bottom: 20px;
        }
        .tab-btn {
            padding: 10px 16px;
            background: transparent;
            border: none;
            color: var(--text-muted);
            font-size: 14px;
            font-weight: 600;
            cursor: pointer;
            border-bottom: 3px solid transparent;
            transition: all 0.2s ease;
        }
        .tab-btn.active {
            color: var(--primary);
            border-bottom-color: var(--primary);
        }

        .tab-content { display: none; }
        .tab-content.active { display: block; }

        .form-group { margin-bottom: 14px; }
        label { font-size: 12px; font-weight: 600; text-transform: uppercase; color: var(--text-muted); display: block; margin-bottom: 6px; }

        input[type="text"], input[type="number"], select {
            width: 100%;
            padding: 11px 14px;
            background: var(--input-bg);
            border: 1px solid var(--border-color);
            border-radius: 10px;
            font-size: 14px;
            color: var(--text-main);
        }

        .btn {
            width: 100%;
            padding: 12px;
            border: none;
            border-radius: 10px;
            font-size: 14px;
            font-weight: 600;
            cursor: pointer;
            text-decoration: none;
            display: inline-block;
            text-align: center;
        }
        .btn-green { background: var(--success); color: white; margin-top: 10px; }
        .btn-purple { background: var(--primary); color: white; margin-top: 10px; }
        .btn-small { width: auto; padding: 6px 10px; font-size: 12px; border-radius: 6px; }

        .alert-success {
            background: rgba(16, 185, 129, 0.15);
            border: 1px solid rgba(16, 185, 129, 0.3);
            color: #6ee7b7;
            padding: 10px;
            border-radius: 8px;
            text-align: center;
            margin-bottom: 16px;
        }

        .filter-grid {
            display: grid;
            grid-template-columns: 1fr 1fr 1fr auto;
            gap: 10px;
            align-items: end;
            margin-bottom: 20px;
            background: var(--input-bg);
            padding: 14px;
            border-radius: 10px;
        }

        table { width: 100%; border-collapse: collapse; font-size: 13px; margin-top: 10px; }
        th, td { padding: 10px; text-align: left; border-bottom: 1px solid var(--border-color); }
        th { background: var(--input-bg); color: var(--text-muted); font-size: 11px; }
        
        .price-text { color: #34d399; font-weight: 600; }
        .status-paid { background: rgba(16, 185, 129, 0.2); color: #34d399; padding: 3px 8px; border-radius: 12px; font-size: 11px; font-weight: 600; }
        .status-pending { background: rgba(245, 158, 11, 0.2); color: #fbbf24; padding: 3px 8px; border-radius: 12px; font-size: 11px; font-weight: 600; }
        .status-cancelled { background: rgba(239, 68, 68, 0.2); color: #f87171; padding: 3px 8px; border-radius: 12px; font-size: 11px; font-weight: 600; }

        .action-btns { display: flex; gap: 6px; align-items: center; }

        .stats-grid {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 12px;
            margin-top: 20px;
        }
        .stat-card {
            background: var(--input-bg);
            padding: 16px;
            border-radius: 12px;
            text-align: center;
            border: 1px solid var(--border-color);
        }
        .stat-card h4 { font-size: 11px; color: var(--text-muted); text-transform: uppercase; margin-bottom: 6px; }
        .stat-card .val { font-size: 18px; font-weight: 700; color: #a5b4fc; }
    </style>
</head>
<body>
    <div class="container">
        {% if not username %}
            <h2 class="header-title">🛍️ Ukay Live Inventory</h2>
            <p class="header-subtitle">Enter store name to start session</p>

            <form method="POST" action="/login">
                <div class="form-group">
                    <label>Store Name / Seller ID</label>
                    <input type="text" name="username" placeholder="e.g., Baguio_Ukay_Store" required autofocus>
                </div>
                <button type="submit" class="btn btn-green">Start Session</button>
            </form>
        {% else %}
            <h2 class="header-title">🛍️ Ukay Live Management System</h2>
            <div class="user-tag">
                Store: <strong>{{ username }}</strong> | Today: <strong>{{ today_date }}</strong> | <a href="/logout">Switch Store</a>
            </div>

            {% if msg %}
                <div class="alert-success">✓ {{ msg }}</div>
            {% endif %}

            <!-- Tab Buttons -->
            <div class="tabs">
                <button class="tab-btn {{ 'active' if active_tab == 'add' }}" onclick="switchTab('add')">⚡ Live Entry</button>
                <button class="tab-btn {{ 'active' if active_tab == 'view' }}" onclick="switchTab('view')">📁 View & Edit Inventory</button>
                <button class="tab-btn {{ 'active' if active_tab == 'report' }}" onclick="switchTab('report')">📊 Monthly Report</button>
            </div>

            <!-- TAB 1: Live Entry -->
            <div id="tab-add" class="tab-content {{ 'active' if active_tab == 'add' }}">
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

                    <button type="submit" class="btn btn-green">Save Item (Status: Pending)</button>
                </form>
                
                <a href="/download?date={{ today_str }}" class="btn btn-purple">📥 Download Today's Excel File</a>
            </div>

            <!-- TAB 2: View & Edit Inventory by Date -->
            <div id="tab-view" class="tab-content {{ 'active' if active_tab == 'view' }}">
                <form method="GET" action="/" class="filter-grid">
                    <input type="hidden" name="tab" value="view">
                    <div>
                        <label>Month</label>
                        <select name="view_month">
                            {% for m_num, m_name in months %}
                                <option value="{{ m_num }}" {{ 'selected' if m_num == view_month }}>{{ m_name }}</option>
                            {% endfor %}
                        </select>
                    </div>

                    <div>
                        <label>Day</label>
                        <select name="view_day">
                            {% for d in days %}
                                <option value="{{ d }}" {{ 'selected' if d == view_day }}>{{ d }}</option>
                            {% endfor %}
                        </select>
                    </div>

                    <div>
                        <label>Year</label>
                        <select name="view_year">
                            {% for y in years %}
                                <option value="{{ y }}" {{ 'selected' if y == view_year }}>{{ y }}</option>
                            {% endfor %}
                        </select>
                    </div>

                    <div>
                        <button type="submit" class="btn btn-purple btn-small">Retrieve File</button>
                    </div>
                </form>

                <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom: 10px;">
                    <p style="font-size: 13px; color: var(--text-muted);">
                        Viewing file for: <strong>{{ view_month }}/{{ view_day }}/{{ view_year }}</strong>
                    </p>
                    <a href="/download?date={{ view_year }}-{{ view_month }}-{{ view_day }}" style="font-size: 12px; color: var(--primary);">📥 Download File</a>
                </div>

                {% if retrieved_items %}
                    <table>
                        <thead>
                            <tr>
                                <th>#</th>
                                <th>Code</th>
                                <th>Buyer</th>
                                <th>Item</th>
                                <th>Price</th>
                                <th>Status</th>
                                <th>Action</th>
                            </tr>
                        </thead>
                        <tbody>
                            {% for row in retrieved_items %}
                            <tr>
                                <td><strong>{{ row.num }}</strong></td>
                                <td><strong>{{ row.code }}</strong></td>
                                <td>{{ row.name }}</td>
                                <td>{{ row.item }}</td>
                                <td class="price-text">₱{{ "%.2f"|format(row.price) }}</td>
                                <td>
                                    <span class="{% if row.status == 'Paid' %}status-paid{% elif row.status == 'Cancelled' %}status-cancelled{% else %}status-pending{% endif %}">
                                        {{ row.status }}
                                    </span>
                                </td>
                                <td>
                                    <div class="action-btns">
                                        {% if row.status != 'Paid' %}
                                            <form method="POST" action="/update_status" style="display:inline;">
                                                <input type="hidden" name="item_num" value="{{ row.num }}">
                                                <input type="hidden" name="file_date" value="{{ view_year }}-{{ view_month }}-{{ view_day }}">
                                                <input type="hidden" name="view_month" value="{{ view_month }}">
                                                <input type="hidden" name="view_day" value="{{ view_day }}">
                                                <input type="hidden" name="view_year" value="{{ view_year }}">
                                                <input type="hidden" name="new_status" value="Paid">
                                                <button type="submit" class="btn btn-small" style="background:var(--success); color:white;">Paid</button>
                                            </form>
                                        {% endif %}

                                        {% if row.status != 'Cancelled' %}
                                            <form method="POST" action="/update_status" style="display:inline;">
                                                <input type="hidden" name="item_num" value="{{ row.num }}">
                                                <input type="hidden" name="file_date" value="{{ view_year }}-{{ view_month }}-{{ view_day }}">
                                                <input type="hidden" name="view_month" value="{{ view_month }}">
                                                <input type="hidden" name="view_day" value="{{ view_day }}">
                                                <input type="hidden" name="view_year" value="{{ view_year }}">
                                                <input type="hidden" name="new_status" value="Cancelled">
                                                <button type="submit" class="btn btn-small" style="background:var(--danger); color:white;">Cancel</button>
                                            </form>
                                        {% endif %}
                                    </div>
                                </td>
                            </tr>
                            {% endfor %}
                        </tbody>
                    </table>
                {% else %}
                    <p style="text-align: center; color: var(--text-muted); padding: 30px 0;">No records found for this date.</p>
                {% endif %}
            </div>

            <!-- TAB 3: Monthly Report -->
            <div id="tab-report" class="tab-content {{ 'active' if active_tab == 'report' }}">
                <form method="GET" action="/" class="filter-grid" style="grid-template-columns: 1fr 1fr auto;">
                    <input type="hidden" name="tab" value="report">
                    <div>
                        <label>Select Month</label>
                        <select name="report_month">
                            {% for m_num, m_name in months %}
                                <option value="{{ m_num }}" {{ 'selected' if m_num == report_month }}>{{ m_name }}</option>
                            {% endfor %}
                        </select>
                    </div>

                    <div>
                        <label>Select Year</label>
                        <select name="report_year">
                            {% for y in years %}
                                <option value="{{ y }}" {{ 'selected' if y == report_year }}>{{ y }}</option>
                            {% endfor %}
                        </select>
                    </div>

                    <div>
                        <button type="submit" class="btn btn-purple btn-small">Generate Report</button>
                    </div>
                </form>

                <div class="stats-grid">
                    <div class="stat-card">
                        <h4>Total Paid Sales</h4>
                        <div class="val" style="color:#34d399;">₱{{ "%.2f"|format(report_data.total_sales) }}</div>
                    </div>
                    <div class="stat-card">
                        <h4>Total Items Sold</h4>
                        <div class="val">{{ report_data.total_items_sold }}</div>
                    </div>
                    <div class="stat-card">
                        <h4>Pending Items</h4>
                        <div class="val" style="color:#fbbf24;">{{ report_data.total_pending }}</div>
                    </div>
                    <div class="stat-card">
                        <h4>Total Cancelled</h4>
                        <div class="val" style="color:#f87171;">{{ report_data.total_cancelled }}</div>
                    </div>
                </div>
            </div>
        {% endif %}
    </div>

    <script>
        function switchTab(tabName) {
            document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
            document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
            
            event.target.classList.add('active');
            document.getElementById('tab-' + tabName).classList.add('active');
        }
    </script>
</body>
</html>
"""


@app.route('/')
def home():
    username = session.get('username')
    msg = request.args.get('msg', '')
    active_tab = request.args.get('tab', 'add')

    now = get_ph_time()
    today_str = now.strftime("%Y-%m-%d")
    today_date = now.strftime("%B %d, %Y")

    # Date selections setup
    months = [
        ("01", "January"), ("02", "February"), ("03", "March"), ("04", "April"),
        ("05", "May"), ("06", "June"), ("07", "July"), ("08", "August"),
        ("09", "September"), ("10", "October"), ("11", "November"), ("12", "December")
    ]
    days = [f"{d:02d}" for d in range(1, 32)]
    years = [str(y) for y in range(2024, 2028)]

    # Retrieval defaults
    view_month = request.args.get('view_month', now.strftime("%m"))
    view_day = request.args.get('view_day', now.strftime("%d"))
    view_year = request.args.get('view_year', now.strftime("%Y"))

    report_month = request.args.get('report_month', now.strftime("%m"))
    report_year = request.args.get('report_year', now.strftime("%Y"))

    retrieved_items = []
    report_data = {'total_sales': 0.0, 'total_items_sold': 0, 'total_cancelled': 0, 'total_pending': 0, 'file_count': 0}

    if username:
        # Load View Tab Items
        target_file = get_user_file_by_date(username, f"{view_year}-{view_month}-{view_day}")
        retrieved_items = read_user_items(target_file)

        # Load Report Tab Data
        report_data = get_monthly_report_data(username, report_year, report_month)

    return render_template_string(
        HTML_TEMPLATE,
        username=username,
        msg=msg,
        active_tab=active_tab,
        today_date=today_date,
        today_str=today_str,
        months=months,
        days=days,
        years=years,
        view_month=view_month,
        view_day=view_day,
        view_year=view_year,
        report_month=report_month,
        report_year=report_year,
        retrieved_items=retrieved_items,
        report_data=report_data
    )


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

    wb = load_workbook(filepath)
    ws = wb["Inventory"]
    # Saved as Pending by default
    ws.append([0, code, name, item, float(price), "Pending"])
    wb.save(filepath)

    rebuild_and_format_excel(filepath)

    return redirect(url_for('home', msg=f"Saved: #{code} for {name} (Pending)!", tab='add'))


@app.route('/update_status', methods=['POST'])
def update_status():
    username = session.get('username')
    if not username:
        return redirect(url_for('home'))

    item_num = request.form.get('item_num')
    file_date = request.form.get('file_date')
    new_status = request.form.get('new_status')

    view_month = request.form.get('view_month')
    view_day = request.form.get('view_day')
    view_year = request.form.get('view_year')

    filepath = get_user_file_by_date(username, file_date)
    update_item_status_in_excel(filepath, item_num, new_status)

    return redirect(url_for('home', tab='view', view_month=view_month, view_day=view_day, view_year=view_year, msg=f"Item #{item_num} status updated to {new_status}!"))


@app.route('/download')
def download():
    username = session.get('username')
    if not username:
        return redirect(url_for('home'))

    date_str = request.args.get('date', get_ph_time().strftime("%Y-%m-%d"))
    filepath = get_user_file_by_date(username, date_str)

    if os.path.exists(filepath):
        return send_file(filepath, as_attachment=True)
    return redirect(url_for('home', msg="File not found for that date."))


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
