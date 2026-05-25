<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>GL Reconciliation Engine</title>
    <link href="https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=IBM+Plex+Mono:wght@400;500;600&display=swap" rel="stylesheet">
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <style>
        :root {
            --bg: #090c10;
            --surface: #0f1318;
            --surface2: #161b22;
            --border: #21262d;
            --border-bright: #30363d;
            --accent: #00d4aa;
            --accent-dim: rgba(0, 212, 170, 0.12);
            --accent-glow: rgba(0, 212, 170, 0.25);
            --amber: #f0a500;
            --amber-dim: rgba(240, 165, 0, 0.1);
            --red: #f85149;
            --red-dim: rgba(248, 81, 73, 0.1);
            --text-primary: #e6edf3;
            --text-secondary: #7d8590;
            --text-muted: #484f58;
        }

        * { margin: 0; padding: 0; box-sizing: border-box; }

        body {
            background: var(--bg);
            color: var(--text-primary);
            font-family: 'IBM Plex Mono', monospace;
            min-height: 100vh;
            padding: 40px 20px 80px;
            position: relative;
            overflow-x: hidden;
        }

        /* Background grid */
        body::before {
            content: '';
            position: fixed;
            inset: 0;
            background-image:
                linear-gradient(rgba(0,212,170,0.03) 1px, transparent 1px),
                linear-gradient(90deg, rgba(0,212,170,0.03) 1px, transparent 1px);
            background-size: 40px 40px;
            pointer-events: none;
            z-index: 0;
        }

        /* Glow blob */
        body::after {
            content: '';
            position: fixed;
            top: -200px;
            right: -200px;
            width: 600px;
            height: 600px;
            background: radial-gradient(circle, rgba(0,212,170,0.06) 0%, transparent 70%);
            pointer-events: none;
            z-index: 0;
        }

        .wrapper {
            max-width: 820px;
            margin: 0 auto;
            position: relative;
            z-index: 1;
        }

        /* ── HEADER ── */
        .header {
            margin-bottom: 48px;
            padding-bottom: 32px;
            border-bottom: 1px solid var(--border);
            animation: fadeDown 0.6s ease both;
        }

        .header-eyebrow {
            font-family: 'IBM Plex Mono', monospace;
            font-size: 11px;
            font-weight: 500;
            letter-spacing: 0.2em;
            text-transform: uppercase;
            color: var(--accent);
            margin-bottom: 12px;
            display: flex;
            align-items: center;
            gap: 10px;
        }

        .header-eyebrow::before {
            content: '';
            display: block;
            width: 24px;
            height: 1px;
            background: var(--accent);
        }

        .header h1 {
            font-family: 'Syne', sans-serif;
            font-size: clamp(28px, 5vw, 44px);
            font-weight: 800;
            letter-spacing: -0.02em;
            line-height: 1.1;
            color: var(--text-primary);
        }

        .header h1 span {
            color: var(--accent);
        }

        .header-sub {
            margin-top: 10px;
            font-size: 12px;
            color: var(--text-secondary);
            letter-spacing: 0.02em;
        }

        /* ── CARDS ── */
        .card {
            background: var(--surface);
            border: 1px solid var(--border);
            border-radius: 12px;
            padding: 28px;
            margin-bottom: 16px;
            transition: border-color 0.2s;
            animation: fadeUp 0.5s ease both;
        }

        .card:hover { border-color: var(--border-bright); }

        .card-label {
            font-family: 'Syne', sans-serif;
            font-size: 11px;
            font-weight: 700;
            letter-spacing: 0.15em;
            text-transform: uppercase;
            color: var(--text-muted);
            margin-bottom: 16px;
            display: flex;
            align-items: center;
            gap: 8px;
        }

        .card-label .num {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 20px;
            height: 20px;
            border: 1px solid var(--border-bright);
            border-radius: 4px;
            font-size: 10px;
            color: var(--accent);
        }

        .card-title {
            font-family: 'Syne', sans-serif;
            font-size: 18px;
            font-weight: 700;
            color: var(--text-primary);
            margin-bottom: 6px;
        }

        .card-desc {
            font-size: 11px;
            color: var(--text-secondary);
            margin-bottom: 20px;
            line-height: 1.6;
        }

        /* ── FILE UPLOAD ── */
        .upload-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 16px;
            margin-bottom: 16px;
        }

        @media (max-width: 580px) { .upload-grid { grid-template-columns: 1fr; } }

        .upload-box {
            background: var(--surface2);
            border: 1px dashed var(--border-bright);
            border-radius: 10px;
            padding: 24px 20px;
            transition: border-color 0.2s, background 0.2s;
            animation: fadeUp 0.6s ease both;
        }

        .upload-box:hover {
            border-color: var(--accent);
            background: var(--accent-dim);
        }

        .upload-box.loaded {
            border-color: var(--accent);
            border-style: solid;
            background: var(--accent-dim);
        }

        .upload-title {
            font-family: 'Syne', sans-serif;
            font-size: 13px;
            font-weight: 700;
            color: var(--text-primary);
            margin-bottom: 4px;
        }

        .upload-sub {
            font-size: 10px;
            color: var(--text-muted);
            margin-bottom: 16px;
            letter-spacing: 0.05em;
        }

        .upload-status {
            font-size: 11px;
            color: var(--accent);
            font-weight: 600;
            margin-top: 12px;
            display: none;
            letter-spacing: 0.05em;
        }

        .upload-status.show { display: block; }

        /* Custom file input */
        .file-btn {
            display: inline-block;
            cursor: pointer;
        }

        .file-btn input[type="file"] {
            display: none;
        }

        .file-btn-label {
            display: inline-flex;
            align-items: center;
            gap: 8px;
            padding: 8px 16px;
            background: transparent;
            border: 1px solid var(--border-bright);
            border-radius: 6px;
            font-family: 'IBM Plex Mono', monospace;
            font-size: 11px;
            font-weight: 500;
            color: var(--text-secondary);
            cursor: pointer;
            transition: all 0.2s;
            letter-spacing: 0.05em;
        }

        .file-btn-label:hover {
            border-color: var(--accent);
            color: var(--accent);
        }

        /* ── PARAMS ── */
        .params-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 20px;
        }

        @media (max-width: 580px) { .params-grid { grid-template-columns: 1fr; } }

        .param-group label {
            display: block;
            font-size: 10px;
            font-weight: 600;
            letter-spacing: 0.15em;
            text-transform: uppercase;
            color: var(--text-muted);
            margin-bottom: 10px;
        }

        .param-input-wrap {
            display: flex;
            align-items: center;
            gap: 10px;
        }

        .param-input-wrap input {
            width: 72px;
            padding: 8px 12px;
            background: var(--surface2);
            border: 1px solid var(--border-bright);
            border-radius: 6px;
            font-family: 'IBM Plex Mono', monospace;
            font-size: 14px;
            font-weight: 600;
            color: var(--text-primary);
            text-align: center;
            outline: none;
            transition: border-color 0.2s;
        }

        .param-input-wrap input:focus {
            border-color: var(--accent);
            box-shadow: 0 0 0 3px var(--accent-glow);
        }

        .param-unit {
            font-size: 11px;
            color: var(--text-secondary);
        }

        /* ── DEMO BANNER ── */
        .demo-banner {
            background: var(--amber-dim);
            border: 1px solid rgba(240,165,0,0.2);
            border-radius: 10px;
            padding: 16px 20px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 16px;
            margin-bottom: 16px;
            flex-wrap: wrap;
            animation: fadeUp 0.4s ease both;
        }

        .demo-text {
            font-size: 11px;
            color: var(--amber);
            letter-spacing: 0.03em;
        }

        .demo-text strong {
            font-weight: 600;
        }

        .demo-btn {
            padding: 8px 20px;
            background: var(--amber);
            border: none;
            border-radius: 6px;
            font-family: 'Syne', sans-serif;
            font-size: 11px;
            font-weight: 700;
            letter-spacing: 0.1em;
            text-transform: uppercase;
            color: #090c10;
            cursor: pointer;
            transition: opacity 0.2s, transform 0.1s;
            white-space: nowrap;
        }

        .demo-btn:hover { opacity: 0.9; transform: translateY(-1px); }
        .demo-btn:active { transform: translateY(0); }

        /* ── RUN BUTTON ── */
        .run-btn {
            width: 100%;
            padding: 18px;
            background: var(--accent);
            border: none;
            border-radius: 10px;
            font-family: 'Syne', sans-serif;
            font-size: 14px;
            font-weight: 800;
            letter-spacing: 0.12em;
            text-transform: uppercase;
            color: #090c10;
            cursor: pointer;
            transition: all 0.2s;
            position: relative;
            overflow: hidden;
            margin-top: 8px;
        }

        .run-btn::before {
            content: '';
            position: absolute;
            inset: 0;
            background: linear-gradient(90deg, transparent, rgba(255,255,255,0.15), transparent);
            transform: translateX(-100%);
            transition: transform 0.5s;
        }

        .run-btn:hover::before { transform: translateX(100%); }
        .run-btn:hover { box-shadow: 0 0 32px var(--accent-glow), 0 8px 24px rgba(0,0,0,0.4); }
        .run-btn:active { transform: scale(0.99); }

        /* ── DIVIDER ── */
        .divider {
            display: flex;
            align-items: center;
            gap: 16px;
            margin: 24px 0;
        }
        .divider::before, .divider::after {
            content: '';
            flex: 1;
            height: 1px;
            background: var(--border);
        }
        .divider span {
            font-size: 10px;
            color: var(--text-muted);
            letter-spacing: 0.1em;
            text-transform: uppercase;
        }

        /* ── ANIMATIONS ── */
        @keyframes fadeDown {
            from { opacity: 0; transform: translateY(-16px); }
            to   { opacity: 1; transform: translateY(0); }
        }

        @keyframes fadeUp {
            from { opacity: 0; transform: translateY(16px); }
            to   { opacity: 1; transform: translateY(0); }
        }

        .card:nth-child(1) { animation-delay: 0.1s; }
        .card:nth-child(2) { animation-delay: 0.2s; }
        .card:nth-child(3) { animation-delay: 0.3s; }
    </style>
</head>
<body>
<div class="wrapper">

    <!-- Header -->
    <div class="header">
        <div class="header-eyebrow">Automated Finance Tooling</div>
        <h1>GL <span>Reconciliation</span><br>Engine</h1>
        <p class="header-sub">// Upload CSV records → match ledger to bank → export clean Excel audit report</p>
    </div>

    <!-- Demo Banner -->
    <div class="demo-banner">
        <div class="demo-text"><strong>No data?</strong> Load a built-in scenario to test the engine instantly.</div>
        <button type="button" class="demo-btn" onclick="loadSampleData()">✦ Load Demo Data</button>
    </div>

    <!-- Upload Card -->
    <div class="card">
        <div class="card-label">
            <span class="num">01</span>
            Data Sources
        </div>
        <div class="upload-grid">
            <!-- GL Upload -->
            <div class="upload-box" id="glBox">
                <div class="upload-title">General Ledger</div>
                <div class="upload-sub">.CSV — internal book records</div>
                <label class="file-btn">
                    <input type="file" id="glFile" accept=".csv" onchange="handleFileUpload(event, 'gl')" />
                    <span class="file-btn-label">
                        <svg width="12" height="12" viewBox="0 0 16 16" fill="none"><path d="M8 1v10M4 7l4 4 4-4M2 13h12" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg>
                        Select File
                    </span>
                </label>
                <div class="upload-status" id="glStatus">✓ &nbsp;FILE LOADED</div>
            </div>

            <!-- Bank Upload -->
            <div class="upload-box" id="bankBox">
                <div class="upload-title">Bank Statement</div>
                <div class="upload-sub">.CSV — clearing transactions</div>
                <label class="file-btn">
                    <input type="file" id="bankFile" accept=".csv" onchange="handleFileUpload(event, 'bank')" />
                    <span class="file-btn-label">
                        <svg width="12" height="12" viewBox="0 0 16 16" fill="none"><path d="M8 1v10M4 7l4 4 4-4M2 13h12" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg>
                        Select File
                    </span>
                </label>
                <div class="upload-status" id="bankStatus">✓ &nbsp;FILE LOADED</div>
            </div>
        </div>
    </div>

    <!-- Params Card -->
    <div class="card">
        <div class="card-label">
            <span class="num">02</span>
            Match Parameters
        </div>
        <div class="params-grid">
            <div class="param-group">
                <label>Date Tolerance Window</label>
                <div class="param-input-wrap">
                    <input type="number" id="daysTolerance" value="3" min="0" />
                    <span class="param-unit">days allowed between entries</span>
                </div>
            </div>
            <div class="param-group">
                <label>Rounding Discrepancy</label>
                <div class="param-input-wrap">
                    <span class="param-unit">$</span>
                    <input type="number" id="amountTolerance" value="0.00" step="0.01" min="0" />
                    <span class="param-unit">max delta</span>
                </div>
            </div>
        </div>
    </div>

    <!-- Run Button -->
    <button type="button" id="runBtn" onclick="runReconciliation()" class="run-btn">
        ⟶ &nbsp; Run Reconciliation &amp; Export Excel
    </button>

</div>

<script>
    let glData = [];
    let bankData = [];

    function parseCSV(text) {
        const lines = text.split('\n').map(l => l.trim()).filter(l => l.length > 0);
        if (!lines.length) return [];
        const headers = lines[0].split(',').map(h => h.replace(/^"|"$/g, '').trim().toLowerCase());
        return lines.slice(1).map(line => {
            const values = line.split(/,(?=(?:(?:[^"]*"){2})*[^"]*$)/);
            const row = {};
            headers.forEach((h, i) => {
                let v = values[i] ? values[i].trim() : '';
                row[h] = v.replace(/^"|"$/g, '');
            });
            return row;
        });
    }

    function handleFileUpload(event, type) {
        const file = event.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = function(e) {
            const parsed = parseCSV(e.target.result);
            if (type === 'gl') {
                glData = parsed;
                const el = document.getElementById('glStatus');
                const box = document.getElementById('glBox');
                el.classList.add('show');
                el.innerText = '✓  GL LOADED — ' + parsed.length + ' ENTRIES';
                box.classList.add('loaded');
            } else {
                bankData = parsed;
                const el = document.getElementById('bankStatus');
                const box = document.getElementById('bankBox');
                el.classList.add('show');
                el.innerText = '✓  BANK LOADED — ' + parsed.length + ' ENTRIES';
                box.classList.add('loaded');
            }
        };
        reader.readAsText(file);
    }

    function loadSampleData() {
        glData = [
            { id: 'tx-1001', date: '2026-05-18', amount: '2500.00' },
            { id: 'tx-1002', date: '2026-05-19', amount: '450.75' },
            { id: 'tx-1003', date: '2026-05-20', amount: '1200.00' },
            { id: 'tx-1004', date: '2026-05-21', amount: '85.00' }
        ];
        bankData = [
            { id: 'tx-1001', date: '2026-05-18', amount: '2500.00' },
            { id: 'tx-1002', date: '2026-05-22', amount: '450.75' },
            { id: 'bank-fee-99', date: '2026-05-23', amount: '15.00' },
            { id: 'tx-1004', date: '2026-05-21', amount: '85.00' }
        ];

        ['gl', 'bank'].forEach((t, i) => {
            const el = document.getElementById(t + 'Status');
            const box = document.getElementById(t + 'Box');
            el.classList.add('show');
            el.innerText = '✓  DEMO DATA INJECTED';
            box.classList.add('loaded');
        });
    }

    function s(n) { return n.toFixed(2); }

    function runReconciliation() {
        if (!glData.length || !bankData.length) {
            alert('Load data first — use demo or upload your CSVs.');
            return;
        }

        const daysTol = parseInt(document.getElementById('daysTolerance').value) || 0;
        const amtTol  = parseFloat(document.getElementById('amountTolerance').value) || 0;

        const cleanRows = arr => arr.map(row => {
            const keys = Object.keys(row);
            const idKey   = keys.find(k => /id|ref|code|desc/.test(k)) || keys[0];
            const dateKey = keys.find(k => /date|time|post/.test(k)) || keys[1];
            const amtKey  = keys.find(k => /amt|amount|val|balance/.test(k)) || keys[2];
            const raw = parseFloat(String(row[amtKey] || '0').replace(/[^0-9.-]/g, ''));
            return {
                id: String(row[idKey] || '').trim(),
                dateStr: String(row[dateKey] || ''),
                date: new Date(row[dateKey] || Date.now()),
                amount: isNaN(raw) ? 0 : raw
            };
        });

        const cleanGL   = cleanRows(glData);
        const cleanBank = cleanRows(bankData);

        const matched = [];
        const unmatchedGL   = [...cleanGL];
        const unmatchedBank = [...cleanBank];

        for (let i = unmatchedGL.length - 1; i >= 0; i--) {
            const g = unmatchedGL[i];
            const bi = unmatchedBank.findIndex(b => {
                const amtOk  = Math.abs(g.amount - b.amount) <= amtTol;
                const dateOk = Math.abs(g.date - b.date) <= daysTol * 86400000;
                const idOk   = g.id !== '' && g.id.toLowerCase() === b.id.toLowerCase();
                return amtOk && (idOk || dateOk);
            });
            if (bi !== -1) {
                const b = unmatchedBank.splice(bi, 1)[0];
                unmatchedGL.splice(i, 1);
                matched.push({
                    'GL Transaction ID': g.id,
                    'GL Amount': g.amount,
                    'Bank Transaction ID': b.id,
                    'Bank Amount': b.amount,
                    'Date (GL)': g.dateStr,
                    'Date (Bank)': b.dateStr,
                    'Status': 'MATCHED ✓'
                });
            }
        }

        const variances = [];
        unmatchedGL.forEach(x => variances.push({
            'Source': 'GENERAL LEDGER',
            'Transaction ID': x.id,
            'Date': x.dateStr,
            'Amount': x.amount,
            'Issue': 'NOT FOUND IN BANK STATEMENT'
        }));
        unmatchedBank.forEach(x => variances.push({
            'Source': 'BANK STATEMENT',
            'Transaction ID': x.id,
            'Date': x.dateStr,
            'Amount': x.amount,
            'Issue': 'NOT FOUND IN GENERAL LEDGER'
        }));

        const glTotal   = cleanGL.reduce((s, i) => s + i.amount, 0);
        const bankTotal = cleanBank.reduce((s, i) => s + i.amount, 0);

        const summary = [
            { 'Metric': '', 'Value': '' },
            { 'Metric': 'RECONCILIATION SUMMARY', 'Value': '' },
            { 'Metric': '─────────────────────────────', 'Value': '──────────────' },
            { 'Metric': 'General Ledger Total',     'Value': glTotal },
            { 'Metric': 'Bank Statement Total',     'Value': bankTotal },
            { 'Metric': 'Net Variance',             'Value': glTotal - bankTotal },
            { 'Metric': '─────────────────────────────', 'Value': '──────────────' },
            { 'Metric': 'Matched Pairs',            'Value': matched.length },
            { 'Metric': 'Unresolved GL Breaks',     'Value': unmatchedGL.length },
            { 'Metric': 'Unresolved Bank Breaks',   'Value': unmatchedBank.length },
            { 'Metric': 'Match Rate',               'Value': ((matched.length / Math.max(cleanGL.length, 1)) * 100).toFixed(1) + '%' },
            { 'Metric': '─────────────────────────────', 'Value': '──────────────' },
            { 'Metric': 'Report Generated',         'Value': new Date().toLocaleString() },
            { 'Metric': 'Days Tolerance Used',      'Value': daysTol },
            { 'Metric': 'Amount Tolerance Used',    'Value': '$' + amtTol.toFixed(2) }
        ];

        try {
            const wb = XLSX.utils.book_new();

            // ── Sheet styling helper ──
            const makeSheet = (data) => {
                const ws = XLSX.utils.json_to_sheet(data);
                // Auto column widths
                const cols = Object.keys(data[0] || {}).map(k => ({
                    wch: Math.max(k.length, ...data.map(r => String(r[k] ?? '').length))
                }));
                ws['!cols'] = cols;
                return ws;
            };

            XLSX.utils.book_append_sheet(wb, makeSheet(summary),   '📊 Summary');
            XLSX.utils.book_append_sheet(wb, makeSheet(matched),    '✅ Matched');
            XLSX.utils.book_append_sheet(wb, makeSheet(variances),  '⚠️ Variances');

            XLSX.writeFile(wb, `GL_Recon_${new Date().toISOString().slice(0,10)}.xlsx`);
        } catch (err) {
            console.error(err);
            alert('Export failed — check the console for details.');
        }
    }
</script>
</body>
</html>
