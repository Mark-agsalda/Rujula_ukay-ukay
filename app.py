import os
import sqlite3
import tempfile
from datetime import datetime
from flask import Flask, render_template_string, request, redirect, url_for, send_file

# Excel export support
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'ukay_live_secret_key_2026')
DB_NAME = "ukayinventory.db"

def get_db_connection():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    """Initializes the database table if it doesn't exist."""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS ukay_inventory (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            mine_code TEXT NOT NULL,
            item_description TEXT NOT NULL,
            price REAL NOT NULL,
            buyer_name TEXT NOT NULL,
            status TEXT DEFAULT 'Mined',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    conn.commit()
    conn.close()

# Auto-initialize database on server startup
with app.app_context():
    init_db()

HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Ukay-Ukay Live Inventory</title>
    <style>
        :root { --bg: #f8fafc; --card: #ffffff; --primary: #2563eb; --accent: #16a34a; --border: #e2e8f0; }
        body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; background: var(--bg); margin: 0; padding: 20px; color: #1e293b; }
        .container { max-width: 1100px; margin: 0 auto; }
        .header-bar { display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; flex-wrap: wrap; gap: 10px; }
        h1 { margin: 0; color: #0f172a; font-size: 1.6rem; }
        .grid { display: grid; grid-template-columns: 320px 1fr; gap: 20px; }
        @media (max-width: 800px) { .grid { grid-template-columns: 1fr; } }
        .card { background: var(--card); padding: 20px; border-radius: 10px; border: 1px solid var(--border); box-shadow: 0 1px 3px rgba(0,0,0,0.05); }
        .form-group { margin-bottom: 12px; }
        label { display: block; font-size: 0.85rem; font-weight: 600; margin-bottom: 4px; }
        input, select { width: 100%; padding: 8px 10px; border: 1px solid var(--border); border-radius: 6px; box-sizing: border-box; font-size: 0.95rem; }
        button, .btn { background: var(--primary); color: white; border: none; padding: 10px; width: 100%; border-radius: 6px; font-weight: 600; cursor: pointer; font-size: 0.95rem; text-decoration: none; display: inline-block; text-align: center; box-sizing: border-box; }
        button:hover, .btn:hover { opacity: 0.9; }
        .btn-excel { background: #16a34a; width: auto; padding: 8px 16px; font-size: 0.85rem; }
        .stats { display: flex; gap: 15px; margin-bottom: 15px; }
        .stat-box { background: #eff6ff; border: 1px solid #bfdbfe; padding: 12px; border-radius: 8px; flex: 1; text-align: center; }
        .stat-box .num { font-size: 1.4rem; font-weight: 700; color: #1e40af; }
        table { width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 0.9rem; }
        th, td { text-align: left; padding: 10px; border-bottom: 1px solid var(--border); }
        th { background: #f1f5f9; font-weight: 600; }
        .badge { padding: 3px 8px; border-radius: 12px; font-size: 0.75rem; font-weight: 600; }
        .badge-Mined { background: #fef3c7; color: #92400e; }
        .badge-Paid { background: #dcfce7; color: #166534; }
        .badge-Shipped { background: #e0e7ff; color: #3730a3; }
        .badge-Cancelled { background: #fee2e2; color: #991b1b; }
        .action-btns { display: flex; gap: 6px; }
        .btn-sm { padding: 4px 8px; font-size: 0.75rem; width: auto; border-radius: 4px; text-decoration: none; }
        .btn-edit { background: #0284c7; color: white; }
        .btn-del { background: #ef4444; color: white; }
    </style>
</head>
<body>

<div class="container">
    <div class="header-bar">
        <h1>👕 Ukay-Ukay Live Inventory Tracker</h1>
        {% if has_excel %}
            <a href="/export/excel" class="btn btn-excel">📊 Export to Excel</a>
        {% endif %}
    </div>
    
    <div class="grid">
        <!-- Input / Edit Form -->
        <div class="card">
            <h3>{% if edit_item %}✏️ Edit Record{% else %}⚡ Quick Add Mined Item{% endif %}</h3>
            <form action="{% if edit_item %}/update/{{ edit_item['id'] }}{% else %}/add{% endif %}" method="POST">
                <div class="form-group">
                    <label>Mine Code / Tag #</label>
                    <input type="text" name="mine_code" placeholder="e.g. M01" value="{{ edit_item['mine_code'] if edit_item else '' }}" required autofocus>
                </div>
                <div class="form-group">
                    <label>Item Description</label>
                    <input type="text" name="item_description" placeholder="e.g. Denim Jacket" value="{{ edit_item['item_description'] if edit_item else '' }}" required>
                </div>
                <div class="form-group">
                    <label>Price (₱)</label>
                    <input type="number" step="0.01" name="price" placeholder="150" value="{{ edit_item['price'] if edit_item else '' }}" required>
                </div>
                <div class="form-group">
                    <label>Buyer Name / Handle</label>
                    <input type="text" name="buyer_name" placeholder="e.g. @MariaClara" value="{{ edit_item['buyer_name'] if edit_item else '' }}" required>
                </div>
                <div class="form-group">
                    <label>Status</label>
                    <select name="status">
                        <option value="Mined" {% if edit_item and edit_item['status']=='Mined' %}selected{% endif %}>Mined (Unpaid)</option>
                        <option value="Paid" {% if edit_item and edit_item['status']=='Paid' %}selected{% endif %}>Paid</option>
                        <option value="Shipped" {% if edit_item and edit_item['status']=='Shipped' %}selected{% endif %}>Shipped</option>
                        <option value="Cancelled" {% if edit_item and edit_item['status']=='Cancelled' %}selected{% endif %}>Cancelled</option>
                    </select>
                </div>
                <button type="submit">{% if edit_item %}Update Record{% else %}Save Mined Item{% endif %}</button>
                {% if edit_item %}
                    <a href="/" style="display:block; text-align:center; margin-top:8px; font-size:0.85rem; color:#64748b; text-decoration:none;">Cancel Edit</a>
                {% endif %}
            </form>
        </div>

        <!-- Inventory Summary & List -->
        <div>
            <div class="stats">
                <div class="stat-box">
                    <div>Total Items</div>
                    <div class="num">{{ items|length }}</div>
                </div>
                <div class="stat-box">
                    <div>Total Revenue</div>
                    <div class="num">₱{{ "%.2f"|format(total_val) }}</div>
                </div>
            </div>

            <div class="card">
                <h3>📋 Recorded Items</h3>
                <table>
                    <thead>
                        <tr>
                            <th>Code</th>
                            <th>Description</th>
                            <th>Price</th>
                            <th>Buyer</th>
                            <th>Status</th>
                            <th>Actions</th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for item in items %}
                        <tr>
                            <td><strong>{{ item['mine_code'] }}</strong></td>
                            <td>{{ item['item_description'] }}</td>
                            <td>₱{{ "%.2f"|format(item['price']) }}</td>
                            <td>{{ item['buyer_name'] }}</td>
                            <td><span class="badge badge-{{ item['status'] }}">{{ item['status'] }}</span></td>
                            <td class="action-btns">
                                <a href="/edit/{{ item['id'] }}" class="btn-sm btn-edit">Edit</a>
                                <a href="/delete/{{ item['id'] }}" class="btn-sm btn-del" onclick="return confirm('Delete record?')">X</a>
                            </td>
                        </tr>
                        {% else %}
                        <tr>
                            <td colspan="6" style="text-align: center; color: #94a3b8; padding: 20px;">No items mined yet. Start adding items!</td>
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>
        </div>
    </div>
</div>

</body>
</html>
'''

@app.route('/')
def index():
    conn = get_db_connection()
    items = conn.execute('SELECT * FROM ukay_inventory ORDER BY id DESC').fetchall()
    total_val = sum(item['price'] for item in items if item['status'] != 'Cancelled')
    conn.close()
    return render_template_string(HTML_TEMPLATE, items=items, total_val=total_val, edit_item=None, has_excel=HAS_OPENPYXL)

@app.route('/add', methods=['POST'])
def add_item():
    code = request.form['mine_code']
    desc = request.form['item_description']
    price = float(request.form['price'])
    buyer = request.form['buyer_name']
    status = request.form['status']

    conn = get_db_connection()
    conn.execute('''
        INSERT INTO ukay_inventory (mine_code, item_description, price, buyer_name, status)
        VALUES (?, ?, ?, ?, ?)
    ''', (code, desc, price, buyer, status))
    conn.commit()
    conn.close()
    return redirect(url_for('index'))

@app.route('/edit/<int:item_id>')
def edit_item(item_id):
    conn = get_db_connection()
    edit_item = conn.execute('SELECT * FROM ukay_inventory WHERE id = ?', (item_id,)).fetchone()
    items = conn.execute('SELECT * FROM ukay_inventory ORDER BY id DESC').fetchall()
    total_val = sum(item['price'] for item in items if item['status'] != 'Cancelled')
    conn.close()
    return render_template_string(HTML_TEMPLATE, items=items, total_val=total_val, edit_item=edit_item, has_excel=HAS_OPENPYXL)

@app.route('/update/<int:item_id>', methods=['POST'])
def update_item(item_id):
    code = request.form['mine_code']
    desc = request.form['item_description']
    price = float(request.form['price'])
    buyer = request.form['buyer_name']
    status = request.form['status']

    conn = get_db_connection()
    conn.execute('''
        UPDATE ukay_inventory
        SET mine_code=?, item_description=?, price=?, buyer_name=?, status=?
        WHERE id=?
    ''', (code, desc, price, buyer, status, item_id))
    conn.commit()
    conn.close()
    return redirect(url_for('index'))

@app.route('/delete/<int:item_id>')
def delete_item(item_id):
    conn = get_db_connection()
    conn.execute('DELETE FROM ukay_inventory WHERE id = ?', (item_id,))
    conn.commit()
    conn.close()
    return redirect(url_for('index'))

@app.route('/export/excel')
def export_excel():
    if not HAS_OPENPYXL:
        return "openpyxl module is missing in requirements.txt", 500

    conn = get_db_connection()
    items = conn.execute('SELECT * FROM ukay_inventory ORDER BY id DESC').fetchall()
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Ukay Inventory"

    # Header styling
    headers = ["ID", "Mine Code", "Description", "Price (PHP)", "Buyer Name", "Status", "Date Mined"]
    ws.append(headers)

    header_fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)

    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    # Rows
    for item in items:
        ws.append([
            item['id'],
            item['mine_code'],
            item['item_description'],
            item['price'],
            item['buyer_name'],
            item['status'],
            item['created_at']
        ])

    # Auto-fit columns
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    # Save to temp directory and serve file
    temp_dir = tempfile.gettempdir()
    file_path = os.path.join(temp_dir, "ukay_inventory_export.xlsx")
    wb.save(file_path)

    return send_file(file_path, as_attachment=True, download_name=f"ukay_inventory_{datetime.now().strftime('%Y%m%d')}.xlsx")

if __name__ == '__main__':
    app.run(debug=True)
