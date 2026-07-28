import os
from flask import Flask, render_template_string, request, redirect, url_for, session, send_file
from openpyxl import Workbook, load_workbook

app = Flask(__name__)
# Uses environment variable if set in cloud, otherwise falls back to key
app.secret_key = os.environ.get('SECRET_KEY', 'ukay_live_secret_key_2026')


def get_user_file(username):
    safe_username = "".join(c for c in username if c.isalnum() or c in ('_', '-')).lower()
    return f"ukay_inventory_{safe_username}.xlsx"


def init_excel(filepath):
    if not os.path.exists(filepath):
        wb = Workbook()
        ws = wb.active
        ws.title = "Inventory"
        ws.append(["Code", "Buyer Name", "Item", "Price"])
        wb.save(filepath)


def read_user_items(filepath):
    items = []
    if os.path.exists(filepath):
        wb = load_workbook(filepath)
        ws = wb["Inventory"]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if any(row):
                items.append(row)
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

        * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
            font-family: 'Inter', sans-serif;
        }

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
            max-width: 480px;
            background: var(--card-bg);
            padding: 30px 25px;
            border-radius: 16px;
            box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.5), 0 8px 10px -6px rgba(0, 0, 0, 0.3);
            border: 1px solid var(--border-color);
        }

        .header-title {
            text-align: center;
            font-size: 24px;
            font-weight: 700;
            color: var(--text-main);
            letter-spacing: -0.5px;
        }

        .header-subtitle {
            text-align: center;
            color: var(--text-muted);
            font-size: 14px;
            margin-top: 6px;
            margin-bottom: 22px;
        }

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

        .user-tag a {
            color: var(--accent);
            text-decoration: none;
            font-weight: 600;
            margin-left: 5px;
        }

        .user-tag a:hover { text-decoration: underline; }

        .form-group {
            margin-bottom: 14px;
        }

        label {
            font-size: 12px;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            color: var(--text-muted);
            display: block;
            margin-bottom: 6px;
        }

        input[type="text"], input[type="number"] {
            width: 100%;
            padding: 12px 14px;
            background: var(--input-bg);
            border: 1px solid var(--border-color);
            border-radius: 10px;
            font-size: 15px;
            color: var(--text-main);
            transition: all 0.2s ease;
        }

        input[type="text"]:focus, input[type="number"]:focus {
            outline: none;
            border-color: var(--primary);
            box-shadow: 0 0 0 3px rgba(99, 102, 241, 0.25);
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
            transition: transform 0.1s ease, background 0.2s ease;
        }

        .btn:active { transform: scale(0.98); }

        .btn-green {
            background: var(--success);
            color: white;
            margin-top: 10px;
            box-shadow: 0 4px 12px rgba(16, 185, 129, 0.25);
        }

        .btn-green:hover { background: var(--success-hover); }

        .btn-purple {
            background: var(--primary);
            color: white;
            margin-top: 12px;
            box-shadow: 0 4px 12px rgba(99, 102, 241, 0.25);
        }

        .btn-purple:hover { background: var(--primary-hover); }

        .alert-success {
            background: rgba(16, 185, 129, 0.15);
            border: 1px solid rgba(16, 185, 129, 0.3);
            color: #6ee7b7;
            padding: 10px;
            border-radius: 8px;
            text-align: center;
            font-size: 14px;
            font-weight: 500;
            margin-bottom: 16px;
        }

        .table-wrapper {
            margin-top: 25px;
            border-top: 1px solid var(--border-color);
            padding-top: 20px;
        }

        .table-header-title {
            font-size: 16px;
            font-weight: 600;
            color: var(--text-main);
            margin-bottom: 12px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }

        .badge-count {
            background: var(--primary);
            color: white;
            font-size: 12px;
            padding: 2px 8px;
            border-radius: 12px;
        }

        table {
            width: 100%;
            border-collapse: collapse;
            font-size: 13px;
        }

        th {
            background: var(--input-bg);
            color: var(--text-muted);
            font-weight: 600;
            text-transform: uppercase;
            font-size: 11px;
            letter-spacing: 0.5px;
            padding: 10px;
            text-align: left;
            border-bottom: 1px solid var(--border-color);
        }

        td {
            padding: 10px;
            border-bottom: 1px solid var(--border-color);
            color: var(--text-main);
        }

        tr:last-child td { border-bottom: none; }

        .price-text {
            color: #34d399;
            font-weight: 600;
        }
    </style>
</head>
<body>
    <div class="container">
        {% if not username %}
            <h2 class="header-title">🛍️ Ukay Inventory</h2>
            <p class="header-subtitle">Enter your store name to start a private session</p>

            <form method="POST" action="/login">
                <div class="form-group">
                    <label>Store Name / Seller ID</label>
                    <input type="text" name="username" placeholder="e.g., Baguio_Ukay_Store" required autofocus>
                </div>
                <button type="submit" class="btn btn-green">Start Session</button>
            </form>
        {% else %}
            <h2 class="header-title">⚡ Live Item Entry</h2>
            <div class="user-tag">
                Store: <strong>{{ username }}</strong> | <a href="/logout">Switch Store</a>
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

            <a href="/download" class="btn btn-purple">📥 Download Excel File</a>

            {% if items %}
                <div class="table-wrapper">
                    <div class="table-header-title">
                        Recent Saved Items
                        <span class="badge-count">{{ items|length }}</span>
                    </div>
                    <table>
                        <thead>
                            <tr>
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
                                <td>{{ row[1] }}</td>
                                <td>{{ row[2] }}</td>
                                <td class="price-text">₱{{ "%.2f"|format(row[3]) }}</td>
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

    if username:
        filepath = get_user_file(username)
        init_excel(filepath)
        items = read_user_items(filepath)

    return render_template_string(HTML_TEMPLATE, username=username, msg=msg, items=items)


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
    ws.append([code, name, item, float(price)])
    wb.save(filepath)

    return redirect(url_for('home', msg=f"Saved: {code} for {name}!"))


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
    # Cloud platforms assign a dynamic PORT environment variable
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)