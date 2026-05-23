<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Automated GL Reconciliation Engine</title>
    <!-- Tailwind CSS for Modern Visuals -->
    <script src="https://cdn.jsdelivr.net/npm/@tailwindcss/browser@4"></script>
    <!-- SheetJS (The rock-solid browser script for pure Excel generation) -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
</head>
<body class="bg-gray-100 min-h-screen py-12 px-4">

    <div class="max-w-4xl mx-auto bg-white shadow-xl rounded-xl border border-gray-100 p-6">
        <!-- Header -->
        <div class="text-center mb-8">
            <h1 class="text-3xl font-extrabold text-gray-900 tracking-tight">Automated GL Reconciliation Engine</h1>
            <p class="text-sm text-gray-500 mt-2">Upload raw CSV records, process ledger assets, and instantly extract a clean Excel workbook file.</p>
        </div>

        <!-- Sandbox Demo Button -->
        <div class="text-center mb-8 p-4 bg-amber-50 rounded-lg border border-amber-200">
            <p class="text-xs text-amber-800 font-medium mb-2">No data ready? Trigger this button to instantly test with a built-in scenario:</p>
            <button type="button" onclick="loadSampleData()" class="px-5 py-2.5 bg-amber-500 text-white font-bold rounded-lg shadow hover:bg-amber-600 transition text-xs uppercase tracking-wider cursor-pointer">
                ✨ Load Mock Demo Data
            </button>
        </div>

        <!-- File Upload Area -->
        <div class="grid grid-cols-1 md:grid-cols-2 gap-8 mb-8">
            <!-- GL Box -->
            <div class="p-5 bg-gray-50 rounded-xl border border-gray-200">
                <h2 class="text-lg font-bold text-gray-800 mb-1">1. General Ledger Source</h2>
                <p class="text-xs text-gray-400 mb-3">Select your internal ledger records (.csv)</p>
                <input type="file" id="glFile" accept=".csv" onchange="handleFileUpload(event, 'gl')" class="block w-full text-sm text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-md file:border-0 file:text-sm file:font-semibold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100" />
                <div id="glStatus" class="mt-2 text-xs text-green-600 font-bold hidden">✓ File Loaded</div>
            </div>

            <!-- Bank Box -->
            <div class="p-5 bg-gray-50 rounded-xl border border-gray-200">
                <h2 class="text-lg font-bold text-gray-800 mb-1">2. Bank Statement Source</h2>
                <p class="text-xs text-gray-400 mb-3">Select your bank statement transactions (.csv)</p>
                <input type="file" id="bankFile" accept=".csv" onchange="handleFileUpload(event, 'bank')" class="block w-full text-sm text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-md file:border-0 file:text-sm file:font-semibold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100" />
                <div id="bankStatus" class="mt-2 text-xs text-green-600 font-bold hidden">✓ File Loaded</div>
            </div>
        </div>

        <!-- Parameters Controls -->
        <div class="p-4 bg-blue-50/50 rounded-xl border border-blue-100 mb-8 grid grid-cols-1 sm:grid-cols-2 gap-4">
            <div>
                <label class="block text-xs font-bold text-blue-900 tracking-wide uppercase mb-1">Date Matching Window:</label>
                <div class="flex items-center gap-2">
                    <input type="number" id="daysTolerance" value="3" min="0" class="w-16 p-1 bg-white border border-blue-200 rounded text-sm font-semibold text-blue-900 text-center" />
                    <span class="text-xs text-blue-700 font-medium">Days allowed between entries</span>
                </div>
            </div>
            <div>
                <label class="block text-xs font-bold text-blue-900 tracking-wide uppercase mb-1">Allowed Rounding Discrepancy:</label>
                <div class="flex items-center gap-2">
                    <span class="text-xs text-blue-700 font-medium">$</span>
                    <input type="number" id="amountTolerance" value="0.00" step="0.01" min="0" class="w-20 p-1 bg-white border border-blue-200 rounded text-sm font-semibold text-blue-900 text-center" />
                </div>
            </div>
        </div>

        <!-- Processing Button -->
        <button type="button" id="runBtn" onclick="runReconciliation()" class="w-full py-4 bg-blue-700 text-white font-bold rounded-xl shadow-md tracking-wide hover:bg-blue-800 transition duration-200 uppercase text-sm cursor-pointer">
            Run Reconciliation Engine & Export Excel
        </button>
    </div>

    <!-- Core Parsing and Processing Engine -->
    <script>
        let glData = [];
        let bankData = [];

        // Parsing engine function
        function parseCSV(text) {
            const lines = text.split('\n').map(line => line.trim()).filter(line => line.length > 0);
            if (lines.length === 0) return [];
            const headers = lines[0].split(',').map(h => h.replace(/^"|"$/g, '').trim().toLowerCase());
            return lines.slice(1).map(line => {
                const values = line.split(/,(?=(?:(?:[^"]*"){2})*[^"]*$)/);
                const rowObj = {};
                headers.forEach((header, index) => {
                    let val = values[index] ? values[index].trim() : '';
                    rowObj[header] = val.replace(/^"|"$/g, '');
                });
                return rowObj;
            });
        }

        // File loading pipeline
        function handleFileUpload(event, type) {
            const file = event.target.files[0];
            if (!file) return;
            const reader = new FileReader();
            reader.onload = function(e) {
                const parsed = parseCSV(e.target.result);
                if (type === 'gl') {
                    glData = parsed;
                    document.getElementById('glStatus').classList.remove('hidden');
                    document.getElementById('glStatus').innerText = "✓ Loaded (" + parsed.length + " lines)";
                } else {
                    bankData = parsed;
                    document.getElementById('bankStatus').classList.remove('hidden');
                    document.getElementById('bankStatus').innerText = "✓ Loaded (" + parsed.length + " lines)";
                }
            };
            reader.readAsText(file);
        }

        // Inject sample scenario array values instantly
        function loadSampleData() {
            glData = [
                { id: 'tx-1001', date: '2026-05-18', amount: '2500.00' },
                { id: 'tx-1002', date: '2026-05-19', amount: '450.75' },
                { id: 'tx-1003', date: '2026-05-20', amount: '1200.00' }, // Missing from statement
                { id: 'tx-1004', date: '2026-05-21', amount: '85.00' }
            ];
            bankData = [
                { id: 'tx-1001', date: '2026-05-18', amount: '2500.00' },
                { id: 'tx-1002', date: '2026-05-22', amount: '450.75' }, // 3-day window match variance
                { id: 'bank-fee-99', date: '2026-05-23', amount: '15.00' }, // Missing from ledger
                { id: 'tx-1004', date: '2026-05-21', amount: '85.00' }
            ];
            document.getElementById('glStatus').classList.remove('hidden');
            document.getElementById('glStatus').innerText = "✓ Mock Ledger Injected";
            document.getElementById('bankStatus').classList.remove('hidden');
            document.getElementById('bankStatus').innerText = "✓ Mock Statement Injected";
            alert("Demo state initialized! Press the blue process action button below to collect your audit download.");
        }

        // Processing matching logic loops
        function runReconciliation() {
            if (!glData.length || !bankData.length) {
                alert('Please upload your data assets or click the mock injection sandbox button first.');
                return;
            }

            const daysTolerance = parseInt(document.getElementById('daysTolerance').value) || 0;
            const amountTolerance = parseFloat(document.getElementById('amountTolerance').value) || 0;

            const cleanRows = (arr) => arr.map(row => {
                const keys = Object.keys(row);
                const idKey = keys.find(k => k.includes('id') || k.includes('ref') || k.includes('code') || k.includes('desc')) || keys[0];
                const dateKey = keys.find(k => k.includes('date') || k.includes('time') || k.includes('post')) || keys[1];
                const amtKey = keys.find(k => k.includes('amt') || k.includes('amount') || k.includes('val') || k.includes('balance')) || keys[2];

                const rawAmt = String(row[amtKey] || '0');
                const cleanAmt = parseFloat(rawAmt.replace(/[^0-9.-]/g, ''));

                return {
                    id: String(row[idKey] || '').trim(),
                    dateStr: String(row[dateKey] || ''),
                    date: new Date(row[dateKey] || Date.now()),
                    amount: isNaN(cleanAmt) ? 0 : cleanAmt
                };
            });

            const cleanGL = cleanRows(glData);
            const cleanBank = cleanRows(bankData);

            const matchedSheetData = [];
            const unmatchedGL = [...cleanGL];
            const unmatchedBank = [...cleanBank];

            for (let i = unmatchedGL.length - 1; i >= 0; i--) {
                const glItem = unmatchedGL[i];
                const bankIndex = unmatchedBank.findIndex(bankItem => {
                    const amountMatches = Math.abs(glItem.amount - bankItem.amount) <= amountTolerance;
                    const dateMatches = Math.abs(glItem.date - bankItem.date) <= (daysTolerance * 24 * 60 * 60 * 1000);
                    const idMatches = glItem.id !== '' && glItem.id.toLowerCase() === bankItem.id.toLowerCase();
                    return amountMatches && (idMatches || dateMatches);
                });

                if (bankIndex !== -1) {
                    const bankItem = unmatchedBank.splice(bankIndex, 1)[0];
                    unmatchedGL.splice(i, 1);
                    matchedSheetData.push({
                        'GL Item ID': glItem.id,
                        'GL Amount': glItem.amount,
                        'Bank Item ID': bankItem.id,
                        'Bank Amount': bankItem.amount,
                        'Match Status': 'Verified Match'
                    });
                }
            }

            // Create output structure array
            const varianceSheetData = [];
            unmatchedGL.forEach(item => {
                varianceSheetData.push({ 'Source Record': 'General Ledger', 'Transaction ID': item.id, 'Date Logged': item.dateStr, 'Unresolved Value': item.amount });
            });
            unmatchedBank.forEach(item => {
                varianceSheetData.push({ 'Source Record': 'Bank Statement', 'Transaction ID': item.id, 'Date Logged': item.dateStr, 'Unresolved Value': item.amount });
            });

            const summarySheetData = [
                { 'Reconciliation Metric': 'Total Book Balance Value (GL)', 'Summary Value': cleanGL.reduce((s, i) => s + i.amount, 0) },
                { 'Reconciliation Metric': 'Total Clearing Statement Value (Bank)', 'Summary Value': cleanBank.reduce((s, i) => s + i.amount, 0) },
                { 'Reconciliation Metric': 'Net System Accounting Variance', 'Summary Value': cleanGL.reduce((s, i) => s + i.amount, 0) - cleanBank.reduce((s, i) => s + i.amount, 0) },
                { 'Reconciliation Metric': 'Perfect Reconciled Pairs Count', 'Summary Value': matchedSheetData.length },
                { 'Reconciliation Metric': 'Unresolved GL Error Breaks', 'Summary Value': unmatchedGL.length },
                { 'Reconciliation Metric': 'Unresolved Bank Error Breaks', 'Summary Value': unmatchedBank.length }
            ];

            // SheetJS Processing Pipeline execution
            try {
                const wb = XLSX.utils.book_new();
                
                const wsSummary = XLSX.utils.json_to_sheet(summarySheetData);
                const wsMatched = XLSX.utils.json_to_sheet(matchedSheetData);
                const wsVariance = XLSX.utils.json_to_sheet(varianceSheetData);

                XLSX.utils.book_append_sheet(wb, wsSummary, "Summary Dashboard");
                XLSX.utils.book_append_sheet(wb, wsMatched, "Matched Transactions");
                XLSX.utils.book_append_sheet(wb, wsVariance, "Unresolved Variances");

                // Execute the instant stream download action
                XLSX.writeFile(wb, `Reconciliation_Report.xlsx`);
            } catch (err) {
                console.error(err);
                alert("The binary execution file sequence ran into an error. Review browser parameters.");
            }
        }
    </script>
</body>
</html>
