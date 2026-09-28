import io
import json
import os
import threading
from datetime import datetime

from flask import Flask, jsonify, render_template_string, request, send_file
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side


app = Flask(__name__)

# Render supplies PORT automatically.
HOST = "0.0.0.0"
PORT = int(os.environ.get("PORT", "10000"))

receiver_lock = threading.Lock()
receiving = False
events = []
devices = {}

MAX_EVENTS = 10000


PAGE = r"""
<!doctype html>
<html>
<head>
    <meta charset="utf-8">
    <title>Keyboard Receiver</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <style>
        * { box-sizing: border-box; }
        body {
            margin: 0;
            font-family: Arial, sans-serif;
            background: #f5f6f8;
            color: #202124;
        }
        header {
            background: white;
            border-bottom: 1px solid #ddd;
            padding: 18px 24px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            gap: 15px;
            position: sticky;
            top: 0;
            z-index: 10;
        }
        h1 { margin: 0; font-size: 22px; }
        .status {
            display: inline-flex;
            align-items: center;
            gap: 8px;
            margin-top: 5px;
            color: #666;
            font-size: 14px;
        }
        .dot {
            width: 10px;
            height: 10px;
            border-radius: 50%;
            display: inline-block;
            background: #888;
        }
        .dot.on { background: #16a34a; }
        .actions {
            display: flex;
            gap: 8px;
            flex-wrap: wrap;
        }
        button, .download {
            border: 0;
            border-radius: 7px;
            padding: 10px 15px;
            font-size: 14px;
            cursor: pointer;
            text-decoration: none;
        }
        #start { background: #16a34a; color: white; }
        #stop { background: #dc2626; color: white; }
        .download { background: #2563eb; color: white; }
        button:disabled {
            opacity: .5;
            cursor: not-allowed;
        }
        main { padding: 22px; }
        .cards {
            display: grid;
            grid-template-columns: repeat(3, minmax(150px, 1fr));
            gap: 14px;
            margin-bottom: 20px;
        }
        .card {
            background: white;
            border: 1px solid #ddd;
            border-radius: 10px;
            padding: 18px;
        }
        .label { color: #777; font-size: 13px; }
        .value { font-size: 26px; margin-top: 5px; }
        .panel {
            background: white;
            border: 1px solid #ddd;
            border-radius: 10px;
            overflow: hidden;
        }
        .panel-title {
            padding: 16px 18px;
            border-bottom: 1px solid #ddd;
            font-weight: 600;
        }
        .table-wrap { overflow-x: auto; }
        table {
            width: 100%;
            border-collapse: collapse;
            min-width: 900px;
        }
        th, td {
            padding: 11px 12px;
            text-align: left;
            border-bottom: 1px solid #eee;
            white-space: nowrap;
            font-size: 13px;
        }
        th { background: #fafafa; }
        .empty {
            padding: 40px;
            text-align: center;
            color: #777;
        }
        @media (max-width: 700px) {
            header { align-items: flex-start; flex-direction: column; }
            .cards { grid-template-columns: 1fr; }
        }
    </style>
</head>
<body>
<header>
    <div>
        <h1>Keyboard Receiver</h1>
        <div class="status">
            <span id="dot" class="dot"></span>
            <span id="status">Not receiving</span>
        </div>
    </div>
    <div class="actions">
        <button id="start" onclick="setReceiving(true)">Start Receiving</button>
        <button id="stop" onclick="setReceiving(false)">Stop Receiver</button>
        <button id="clear" onclick="clearLogs()">Clear Logs</button>
        <a class="download" href="/download-excel">Download Excel</a>
    </div>
</header>

<main>
    <div class="cards">
        <div class="card">
            <div class="label">Devices</div>
            <div class="value" id="deviceCount">0</div>
        </div>
        <div class="card">
            <div class="label">Events</div>
            <div class="value" id="eventCount">0</div>
        </div>
        <div class="card">
            <div class="label">Last Event</div>
            <div class="value" id="lastEvent" style="font-size:16px">—</div>
        </div>
    </div>

    <div class="panel" style="margin-bottom:20px">
        <div class="panel-title">Devices</div>
        <div id="deviceTabs" style="padding:12px;display:flex;gap:8px;flex-wrap:wrap">
            <button onclick="selectDevice('ALL')" style="background:#2563eb;color:white">
                All Devices
            </button>
        </div>
    </div>

    <div class="panel">
        <div class="panel-title">
            Live Event Log
            <span id="selectedDeviceLabel" style="float:right;color:#777;font-weight:normal">
                All Devices
            </span>
        </div>
        <div class="table-wrap">
            <table>
                <thead>
                    <tr>
                        <th>Time</th>
                        <th>Device</th>
                        <th>IP</th>
                        <th>Key</th>
                        <th>KEY_DOWN</th>
                        <th>KEY_UP</th>
                        <th>Duration</th>
                        <th>Count</th>
                        <th>Type</th>
                        <th>Data</th>
                    </tr>
                </thead>
                <tbody id="tbody">
                    <tr><td colspan="10" class="empty">No events received.</td></tr>
                </tbody>
            </table>
        </div>
    </div>
</main>

<script>
let selectedDevice = 'ALL';

function selectDevice(deviceId) {
    selectedDevice = deviceId;
    getStatus();
}

function escapeHtml(value) {
    const div = document.createElement('div');
    div.textContent = value ?? '';
    return div.innerHTML;
}

async function getStatus() {
    const r = await fetch('/api/status', {cache: 'no-store'});
    const data = await r.json();

    document.getElementById('status').textContent =
        data.receiving ? 'Receiving' : 'Not receiving';

    document.getElementById('dot').className =
        data.receiving ? 'dot on' : 'dot';

    document.getElementById('start').disabled = data.receiving;
    document.getElementById('stop').disabled = !data.receiving;

    document.getElementById('deviceCount').textContent = data.device_count;
    document.getElementById('eventCount').textContent = data.event_count;

    document.getElementById('lastEvent').textContent =
        data.last_event || '—';

    const deviceTabs = document.getElementById('deviceTabs');

    deviceTabs.innerHTML =
        '<button onclick="selectDevice(\\'ALL\\')" ' +
        'style="background:#2563eb;color:white;border:0;border-radius:7px;padding:9px 13px;cursor:pointer">' +
        'All Devices</button>';

    for (const d of data.devices) {
        const button = document.createElement('button');
        button.textContent = d.label;
        button.style.border = '0';
        button.style.borderRadius = '7px';
        button.style.padding = '9px 13px';
        button.style.cursor = 'pointer';
        button.style.background =
            selectedDevice === d.id ? '#2563eb' : '#e5e7eb';
        button.style.color =
            selectedDevice === d.id ? 'white' : '#202124';
        button.onclick = () => selectDevice(d.id);
        deviceTabs.appendChild(button);
    }

    const filtered = selectedDevice === 'ALL'
        ? data.events
        : data.events.filter(e => e.device_id === selectedDevice);

    document.getElementById('selectedDeviceLabel').textContent =
        selectedDevice === 'ALL' ? 'All Devices' :
        (data.devices.find(d => d.id === selectedDevice)?.label || selectedDevice);

    const tbody = document.getElementById('tbody');
    tbody.innerHTML = '';

    if (!filtered.length) {
        tbody.innerHTML =
            '<tr><td colspan="10" class="empty">No events for this device.</td></tr>';
        return;
    }

    for (const e of filtered) {
        const tr = document.createElement('tr');

        [
            e.time,
            e.device,
            e.ip,
            e.key,
            e.event === 'KEY_DOWN' ? e.key : '',
            e.event === 'KEY_UP' ? e.key : '',
            e.duration,
            e.count,
            e.type,
            e.data
        ].forEach(value => {
            const td = document.createElement('td');
            td.textContent = value ?? '';
            tr.appendChild(td);
        });

        tbody.appendChild(tr);
    }
}


async function setReceiving(value) {
    try {
        const response = await fetch('/api/receiving', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify({receiving: value})
        });

        if (!response.ok) {
            const body = await response.text();
            alert('Receiver control failed: HTTP ' + response.status + '\n' + body);
            return;
        }

        await getStatus();
    } catch (error) {
        alert('Receiver control failed: ' + error);
    }
}


async function clearLogs() {
    const confirmed = confirm(
        'Clear ALL received logs and device history? This cannot be undone.'
    );

    if (!confirmed) return;

    const response = await fetch('/api/clear', {
        method: 'POST'
    });

    if (response.ok) {
        selectedDevice = 'ALL';
        await getStatus();
    } else {
        alert('Unable to clear logs.');
    }
}


getStatus();
setInterval(getStatus, 1000);
</script>
</body>
</html>
"""


def now_text():
    return datetime.now().strftime("%H:%M:%S")


def clean_text(value, default=""):
    if value is None:
        return default
    return str(value)


def add_events(device, incoming):
    device_name = clean_text(device.get("name"), "UNKNOWN")
    device_ip = clean_text(device.get("ip"), "UNKNOWN")
    device_id = f"{device_name}|{device_ip}"

    with receiver_lock:
        devices[device_id] = {
            "name": device_name,
            "ip": device_ip,
        }

        for item in incoming:
            if not isinstance(item, dict):
                continue

            row = {
                "time": now_text(),
                "device": device_name,
                "ip": device_ip,
                "key": clean_text(item.get("key")),
                "event": clean_text(item.get("event")),
                "duration": clean_text(item.get("duration")),
                "count": clean_text(item.get("count")),
                "type": clean_text(item.get("type")),
                "data": clean_text(item.get("data")),
            }

            events.append(row)

        if len(events) > MAX_EVENTS:
            del events[:-MAX_EVENTS]


@app.get("/")
def index():
    return render_template_string(PAGE)


@app.get("/api/status")
def api_status():
    with receiver_lock:
        latest = events[-1] if events else None
        return jsonify({
            "receiving": receiving,
            "device_count": len(devices),
            "event_count": len(events),
            "last_event": (
                f"{latest['device']} - {latest['key']} "
                f"{latest['event']}"
                if latest else None
            ),
            "devices": [
                {
                    "id": device_id,
                    "label": f"{device['name']} ({device['ip']})",
                }
                for device_id, device in devices.items()
            ],
            "events": [
                dict(row, device_id=f"{row['device']}|{row['ip']}")
                for row in reversed(events[-500:])
            ],
        })


@app.post("/api/receiving")
def api_receiving():
    global receiving

    data = request.get_json(silent=True) or {}
    value = bool(data.get("receiving", False))

    with receiver_lock:
        receiving = value

    return jsonify({
        "status": "ok",
        "receiving": receiving,
    }), 200


@app.post("/events")
def receive_events():
    with receiver_lock:
        active = receiving

    if not active:
        return jsonify({
            "status": "stopped",
            "message": "Receiver is currently stopped."
        }), 503

    payload = request.get_json(silent=True)

    if not isinstance(payload, dict):
        return jsonify({"error": "invalid JSON body"}), 400

    device = payload.get("device") or {}
    incoming = payload.get("events") or []

    if not isinstance(device, dict):
        return jsonify({"error": "device must be an object"}), 400

    if not isinstance(incoming, list):
        return jsonify({"error": "events must be an array"}), 400

    add_events(device, incoming)

    return jsonify({
        "status": "ok",
        "received": len(incoming),
    })


def safe_sheet_name(name):
    invalid = '[]:*?/\\'
    result = "".join("_" if c in invalid else c for c in name)
    return result[:31] or "Device"


def create_excel():
    workbook = Workbook()
    workbook.remove(workbook.active)

    headers = [
        "time",
        "Key",
        "KEY_DOWN",
        "KEY_UP",
        "duration",
        "Count",
        "type",
        "IP",
    ]

    with receiver_lock:
        grouped = {}

        for row in events:
            key = f"{row['device']}|{row['ip']}"
            grouped.setdefault(key, []).append(dict(row))

    if not grouped:
        sheet = workbook.create_sheet("No events")
        sheet.append(headers)
    else:
        for device_key, rows_for_device in grouped.items():
            device = rows_for_device[0]["device"]
            ip = rows_for_device[0]["ip"]

            sheet = workbook.create_sheet(
                safe_sheet_name(f"{device}_{ip}")
            )

            sheet["A1"] = f"{ip} ({device})"
            sheet["A1"].font = Font(size=14)

            for col, header in enumerate(headers, 1):
                cell = sheet.cell(row=2, column=col, value=header)
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal="left")

            for row in rows_for_device:
                # Put the key event into its own KEY_DOWN or KEY_UP column.
                event_name = row["event"].upper()

                key_down = row["key"] if event_name == "KEY_DOWN" else ""
                key_up = row["key"] if event_name == "KEY_UP" else ""

                sheet.append([
                    row["time"],
                    row["key"],
                    key_down,
                    key_up,
                    row["duration"],
                    row["count"],
                    row["type"],
                    row["ip"],
                ])

            widths = [16, 16, 16, 16, 14, 10, 14, 18]
            for i, width in enumerate(widths, 1):
                sheet.column_dimensions[get_column_letter(i)].width = width

            sheet.freeze_panes = "A3"
            sheet.auto_filter.ref = sheet.dimensions

    output = io.BytesIO()
    workbook.save(output)
    workbook.close()
    output.seek(0)
    return output



@app.post("/api/clear")
def clear_logs():
    global events, devices

    with receiver_lock:
        events.clear()
        devices.clear()

    return jsonify({
        "status": "ok",
        "message": "All logs cleared.",
    })


@app.get("/download-excel")
def download_excel():
    output = create_excel()

    return send_file(
        output,
        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        as_attachment=True,
        download_name="keyboard_logs.xlsx",
    )


if __name__ == "__main__":
    print(f"Starting web receiver on {HOST}:{PORT}", flush=True)
    app.run(host=HOST, port=PORT, threaded=True)
