from __future__ import annotations

import base64
import json
from pathlib import Path

import pandas as pd


BASE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = BASE_DIR.parent.parent

MICRO_SCHEDULE_PATH = BASE_DIR / "Micro_Schedule_Generator" / "outputs" / "Micro_Schedule.csv"
OUTPUT_HTML_PATH = BASE_DIR / "spatial_visualizer_micro.html"

FLOOR_PLAN_DIR = PROJECT_DIR / "floor_plans" / "cropped_png"

BACKGROUND_BY_LEVEL = {
    "L -1": FLOOR_PLAN_DIR / "04_Island_ARCH_ConceptB_Level -1_Mar6_page_0_cropped.png",
    "L 0": FLOOR_PLAN_DIR / "04_Island_ARCH_ConceptB_Level 0_Mar6 (1)_page_0_cropped.png",
    "L 1": FLOOR_PLAN_DIR / "04_Island_ARCH_ConceptB_Level 1_Mar6_page_0_cropped.png",
}

TASK_COLORS = {
    "Frame: Columns + Beams": "#c2410c",
    "Floor": "#0284c7",
    "Roof": "#15803d",
}


def image_to_data_uri(path: Path) -> str:
    data = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:image/png;base64,{data}"


def level_sort_key(level: str) -> tuple[float, str]:
    text = str(level).strip()
    if text.lower() == "roof":
        return (9000.0, text)
    if text.upper().startswith("L"):
        try:
            value = float(text.replace("L", "").strip())
            return (value, text)
        except ValueError:
            pass
    digits = "".join(ch for ch in text if ch.isdigit() or ch == "-")
    if digits:
        try:
            return (float(digits), text)
        except ValueError:
            pass
    return (9500.0, text)


def build_payload() -> dict[str, object]:
    df = pd.read_csv(MICRO_SCHEDULE_PATH)
    df["slot_start_dt"] = pd.to_datetime(df["slot_start"])
    df["slot_end_dt"] = pd.to_datetime(df["slot_end"])

    intervals = sorted(df["slot_start"].dropna().astype(str).unique())
    levels = sorted(df["level"].dropna().astype(str).unique(), key=level_sort_key)
    tasks = list(dict.fromkeys(df["task_name"].astype(str).tolist()))

    level_bounds: dict[str, dict[str, float]] = {}
    level_points_by_interval: dict[str, list[list[dict[str, object]]]] = {}

    for level in levels:
        level_df = df[df["level"].astype(str) == level].copy()
        min_x = float(level_df["coord_x_ft"].min())
        max_x = float(level_df["coord_x_ft"].max())
        min_y = float(level_df["coord_y_ft"].min())
        max_y = float(level_df["coord_y_ft"].max())
        padding_x = max((max_x - min_x) * 0.06, 8.0)
        padding_y = max((max_y - min_y) * 0.06, 8.0)
        level_bounds[level] = {
            "min_x": round(min_x - padding_x, 3),
            "max_x": round(max_x + padding_x, 3),
            "min_y": round(min_y - padding_y, 3),
            "max_y": round(max_y + padding_y, 3),
        }

        interval_frames: list[list[dict[str, object]]] = []
        for interval in intervals:
            frame_df = level_df[level_df["slot_start"].astype(str) == interval].copy()
            frame_points: list[dict[str, object]] = []
            for _, row in frame_df.sort_values(["task_name", "order_in_task"]).iterrows():
                frame_points.append(
                    {
                        "element_key": str(row["element_key"]),
                        "task_name": str(row["task_name"]),
                        "task_color": TASK_COLORS.get(str(row["task_name"]), "#b45309"),
                        "x": round(float(row["coord_x_ft"]), 3),
                        "y": round(float(row["coord_y_ft"]), 3),
                        "order_in_task": int(row["order_in_task"]),
                        "scheduled_duration_hr": float(row["scheduled_duration_hr"]),
                        "slot_active_duration_hr": float(row.get("slot_active_duration_hr", 0.5)),
                        "batch_index": int(row["batch_index"]),
                    }
                )
            interval_frames.append(frame_points)
        level_points_by_interval[level] = interval_frames

    interval_summary = []
    for interval in intervals:
        interval_df = df[df["slot_start"].astype(str) == interval].copy()
        interval_summary.append(
            {
                "interval_start": interval,
                "active_elements": int(interval_df["element_key"].nunique()),
                "active_tasks": int(interval_df["task_name"].nunique()),
                "levels_active": sorted(interval_df["level"].dropna().astype(str).unique().tolist(), key=level_sort_key),
            }
        )

    backgrounds = {
        level: image_to_data_uri(path)
        for level, path in BACKGROUND_BY_LEVEL.items()
        if path.exists()
    }

    return {
        "title": "Superstructure Spatial Playback",
        "subtitle": "30-minute BIM element schedule playback generated from Micro_Schedule_Generator/outputs/Micro_Schedule.csv.",
        "intervals": intervals,
        "levels": levels,
        "tasks": tasks,
        "task_colors": {task: TASK_COLORS.get(task, "#b45309") for task in tasks},
        "level_bounds": level_bounds,
        "points_by_level": level_points_by_interval,
        "interval_summary": interval_summary,
        "backgrounds_by_level": backgrounds,
        "total_rows": int(len(df)),
        "total_elements": int(df["element_key"].nunique()),
    }


def build_html(payload: dict[str, object]) -> str:
    payload_json = json.dumps(payload)
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Superstructure Spatial Visualizer</title>
  <style>
    :root {{
      --bg: #efe7d7;
      --panel: rgba(255, 252, 246, 0.94);
      --ink: #1f2937;
      --muted: #6b7280;
      --accent: #9a3412;
      --border: rgba(31, 41, 55, 0.08);
      --shadow: 0 20px 48px rgba(31, 41, 55, 0.12);
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: Georgia, "Times New Roman", serif;
      color: var(--ink);
      background:
        radial-gradient(circle at top left, rgba(154, 52, 18, 0.16), transparent 28%),
        linear-gradient(180deg, #f8f4eb 0%, var(--bg) 100%);
    }}
    .page {{
      max-width: 1320px;
      margin: 0 auto;
      padding: 28px;
      display: grid;
      gap: 20px;
    }}
    .header {{
      display: grid;
      gap: 8px;
    }}
    .eyebrow {{
      text-transform: uppercase;
      letter-spacing: 0.18em;
      font-size: 12px;
      color: var(--accent);
    }}
    h1 {{
      margin: 0;
      font-size: clamp(32px, 5vw, 56px);
      line-height: 0.96;
    }}
    .subtitle {{
      max-width: 860px;
      font-size: 17px;
      color: var(--muted);
      line-height: 1.5;
    }}
    .layout {{
      display: grid;
      grid-template-columns: 320px minmax(0, 1fr);
      gap: 20px;
      align-items: start;
    }}
    .card {{
      background: var(--panel);
      border: 1px solid var(--border);
      border-radius: 24px;
      box-shadow: var(--shadow);
      backdrop-filter: blur(8px);
    }}
    .sidebar {{
      padding: 20px;
      display: grid;
      gap: 18px;
      position: sticky;
      top: 16px;
    }}
    .metrics {{
      display: grid;
      gap: 12px;
    }}
    .metric {{
      padding: 12px 14px;
      border-radius: 18px;
      background: linear-gradient(135deg, rgba(250, 234, 214, 0.86), rgba(255,255,255,0.94));
      border: 1px solid rgba(154, 52, 18, 0.12);
    }}
    .metric-label {{
      font-size: 11px;
      text-transform: uppercase;
      letter-spacing: 0.12em;
      color: var(--muted);
      margin-bottom: 4px;
    }}
    .metric-value {{
      font-size: 24px;
      font-weight: 700;
    }}
    .control {{
      display: grid;
      gap: 8px;
    }}
    .control-label {{
      font-size: 12px;
      text-transform: uppercase;
      letter-spacing: 0.12em;
      color: var(--muted);
    }}
    button, select {{
      border: none;
      border-radius: 999px;
      padding: 10px 14px;
      font: inherit;
      background: white;
      color: var(--ink);
      box-shadow: inset 0 0 0 1px rgba(31, 41, 55, 0.12);
      cursor: pointer;
    }}
    button.primary {{
      background: var(--accent);
      color: white;
      box-shadow: none;
    }}
    input[type="range"] {{
      width: 100%;
      accent-color: var(--accent);
    }}
    .button-row {{
      display: flex;
      gap: 10px;
      flex-wrap: wrap;
    }}
    .tasks {{
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
    }}
    .task-pill {{
      display: inline-flex;
      align-items: center;
      gap: 8px;
      padding: 7px 10px;
      border-radius: 999px;
      background: #fff7ed;
      color: #7c2d12;
      font-size: 13px;
    }}
    .swatch {{
      width: 10px;
      height: 10px;
      border-radius: 999px;
      flex: 0 0 auto;
    }}
    .viz {{
      padding: 18px;
      display: grid;
      gap: 16px;
    }}
    .viz-top {{
      display: flex;
      justify-content: space-between;
      gap: 12px;
      flex-wrap: wrap;
      align-items: end;
    }}
    .readout {{
      display: grid;
      gap: 4px;
    }}
    .readout .label {{
      font-size: 12px;
      text-transform: uppercase;
      letter-spacing: 0.12em;
      color: var(--muted);
    }}
    .readout .value {{
      font-size: 28px;
      font-weight: 700;
    }}
    .viz-stage {{
      position: relative;
    }}
    canvas {{
      width: 100%;
      aspect-ratio: 1.35 / 1;
      border-radius: 22px;
      background: linear-gradient(180deg, rgba(255,255,255,0.94), rgba(247, 241, 230, 0.96));
      border: 1px solid var(--border);
    }}
    .note {{
      font-size: 13px;
      color: var(--muted);
      line-height: 1.5;
    }}
    @media (max-width: 980px) {{
      .layout {{ grid-template-columns: 1fr; }}
      .sidebar {{ position: static; }}
    }}
  </style>
</head>
<body>
  <div class="page">
    <div class="header">
      <div class="eyebrow">Spatial Planning Playback</div>
      <h1>Superstructure Dot Visualizer</h1>
      <div class="subtitle">Animated 30-minute playback of BIM elements from the micro schedule, overlaid on available floor-plan crops by level.</div>
    </div>

    <div class="layout">
      <aside class="card sidebar">
        <div class="metrics">
          <div class="metric">
            <div class="metric-label">Intervals</div>
            <div class="metric-value" id="metric-intervals"></div>
          </div>
          <div class="metric">
            <div class="metric-label">Elements</div>
            <div class="metric-value" id="metric-elements"></div>
          </div>
          <div class="metric">
            <div class="metric-label">Rows</div>
            <div class="metric-value" id="metric-rows"></div>
          </div>
          <div class="metric">
            <div class="metric-label">Active Now</div>
            <div class="metric-value" id="metric-active"></div>
          </div>
        </div>

        <div class="control">
          <div class="control-label">Level</div>
          <select id="level-select"></select>
        </div>

        <div class="control">
          <div class="control-label">Interval</div>
          <input id="interval-slider" type="range" min="0" max="0" value="0" step="1">
        </div>

        <div class="button-row">
          <button class="primary" id="play-btn">Play</button>
          <button id="pause-btn">Pause</button>
        </div>

        <div class="control">
          <div class="control-label">Tasks</div>
          <div class="tasks" id="task-pills"></div>
        </div>

        <div class="note" id="background-note"></div>
      </aside>

      <section class="card viz">
        <div class="viz-top">
          <div class="readout">
            <div class="label">Current Interval</div>
            <div class="value" id="interval-label"></div>
          </div>
          <div class="readout">
            <div class="label">Interval Summary</div>
            <div class="value" id="summary-label"></div>
          </div>
        </div>

        <div class="viz-stage">
          <canvas id="canvas" width="1040" height="760"></canvas>
        </div>

        <div class="note">
          Each dot is an active BIM element in the selected 30-minute interval. Dot color shows task, while dot size and opacity reflect how much of that exact installation overlaps the current 30-minute slot.
        </div>
      </section>
    </div>
  </div>

  <script>
    const payload = {payload_json};
    const canvas = document.getElementById('canvas');
    const ctx = canvas.getContext('2d');
    const intervalSlider = document.getElementById('interval-slider');
    const levelSelect = document.getElementById('level-select');
    const intervalLabel = document.getElementById('interval-label');
    const summaryLabel = document.getElementById('summary-label');
    const metricIntervals = document.getElementById('metric-intervals');
    const metricElements = document.getElementById('metric-elements');
    const metricRows = document.getElementById('metric-rows');
    const metricActive = document.getElementById('metric-active');
    const backgroundNote = document.getElementById('background-note');
    const taskPills = document.getElementById('task-pills');

    let currentIntervalIndex = 0;
    let currentLevel = payload.levels[0];
    let timerId = null;
    const backgroundImages = {{}};

    function fmtInterval(isoString) {{
      return new Date(isoString).toLocaleString([], {{
        year: 'numeric',
        month: 'short',
        day: 'numeric',
        hour: 'numeric',
        minute: '2-digit'
      }});
    }}

    function getBackground(level) {{
      const src = payload.backgrounds_by_level[level];
      if (!src) {{
        return null;
      }}
      if (!backgroundImages[level]) {{
        const img = new Image();
        img.src = src;
        backgroundImages[level] = img;
      }}
      return backgroundImages[level];
    }}

    function xToCanvas(x, bounds, plotLeft, plotWidth) {{
      const ratio = (x - bounds.min_x) / Math.max(bounds.max_x - bounds.min_x, 1);
      return plotLeft + ratio * plotWidth;
    }}

    function yToCanvas(y, bounds, plotTop, plotHeight) {{
      const ratio = (y - bounds.min_y) / Math.max(bounds.max_y - bounds.min_y, 1);
      return plotTop + plotHeight - ratio * plotHeight;
    }}

    function populateUi() {{
      metricIntervals.textContent = payload.intervals.length.toLocaleString();
      metricElements.textContent = payload.total_elements.toLocaleString();
      metricRows.textContent = payload.total_rows.toLocaleString();

      payload.levels.forEach(level => {{
        const option = document.createElement('option');
        option.value = level;
        option.textContent = level;
        levelSelect.appendChild(option);
      }});

      payload.tasks.forEach(task => {{
        const pill = document.createElement('span');
        pill.className = 'task-pill';
        const swatch = document.createElement('span');
        swatch.className = 'swatch';
        swatch.style.background = payload.task_colors[task] || '#b45309';
        pill.appendChild(swatch);
        pill.appendChild(document.createTextNode(task));
        taskPills.appendChild(pill);
      }});

      intervalSlider.max = Math.max(payload.intervals.length - 1, 0);
      intervalSlider.value = '0';
    }}

    function drawAxes(bounds, plotLeft, plotTop, plotWidth, plotHeight) {{
      ctx.strokeStyle = 'rgba(31, 41, 55, 0.1)';
      ctx.lineWidth = 1;
      ctx.strokeRect(plotLeft, plotTop, plotWidth, plotHeight);

      ctx.fillStyle = '#6b7280';
      ctx.font = '12px Georgia, serif';
      ctx.fillText(`X ${{bounds.min_x}} ft`, plotLeft, plotTop + plotHeight + 22);
      ctx.fillText(`X ${{bounds.max_x}} ft`, plotLeft + plotWidth - 72, plotTop + plotHeight + 22);
      ctx.save();
      ctx.translate(18, plotTop + plotHeight);
      ctx.rotate(-Math.PI / 2);
      ctx.fillText(`Y ${{bounds.min_y}} ft`, 0, 0);
      ctx.fillText(`Y ${{bounds.max_y}} ft`, plotHeight - 72, 0);
      ctx.restore();
    }}

    function drawPoints(points, bounds, plotLeft, plotTop, plotWidth, plotHeight, opacity, outline) {{
      points.forEach(point => {{
        const cx = xToCanvas(point.x, bounds, plotLeft, plotWidth);
        const cy = yToCanvas(point.y, bounds, plotTop, plotHeight);
        const slotRatio = Math.max(0.18, Math.min(point.slot_active_duration_hr / 0.5, 1));
        const radius = Math.max(4, 4 + (point.scheduled_duration_hr * 0.8) + (slotRatio * 6));
        ctx.beginPath();
        ctx.arc(cx, cy, radius, 0, Math.PI * 2);
        ctx.fillStyle = point.task_color;
        ctx.globalAlpha = opacity * (0.35 + slotRatio * 0.65);
        ctx.fill();
        ctx.globalAlpha = 1;
        if (outline) {{
          ctx.strokeStyle = 'rgba(255,255,255,0.9)';
          ctx.lineWidth = 2;
          ctx.stroke();
        }}
      }});
    }}

    function draw() {{
      const width = canvas.width;
      const height = canvas.height;
      const plotLeft = 56;
      const plotTop = 48;
      const plotWidth = width - 112;
      const plotHeight = height - 112;
      const bounds = payload.level_bounds[currentLevel];

      ctx.clearRect(0, 0, width, height);
      const background = ctx.createLinearGradient(0, 0, 0, height);
      background.addColorStop(0, 'rgba(255,255,255,0.96)');
      background.addColorStop(1, 'rgba(244, 237, 225, 0.96)');
      ctx.fillStyle = background;
      ctx.fillRect(0, 0, width, height);

      const bgImage = getBackground(currentLevel);
      if (bgImage && bgImage.complete) {{
        ctx.save();
        ctx.globalAlpha = 0.9;
        ctx.drawImage(bgImage, plotLeft, plotTop, plotWidth, plotHeight);
        ctx.restore();
        backgroundNote.textContent = 'Using cropped floor plan background for this level.';
      }} else if (payload.backgrounds_by_level[currentLevel]) {{
        backgroundNote.textContent = 'Loading floor plan background for this level.';
      }} else {{
        backgroundNote.textContent = 'No cropped floor plan background available for this level.';
      }}

      drawAxes(bounds, plotLeft, plotTop, plotWidth, plotHeight);

      for (let lookback = 4; lookback >= 1; lookback -= 1) {{
        const idx = currentIntervalIndex - lookback;
        if (idx < 0) continue;
        const echoPoints = (payload.points_by_level[currentLevel] || [])[idx] || [];
        drawPoints(echoPoints, bounds, plotLeft, plotTop, plotWidth, plotHeight, 0.08 * (5 - lookback), false);
      }}

      const currentPoints = (payload.points_by_level[currentLevel] || [])[currentIntervalIndex] || [];
      drawPoints(currentPoints, bounds, plotLeft, plotTop, plotWidth, plotHeight, 0.82, true);

      ctx.fillStyle = '#1f2937';
      ctx.font = '600 15px Georgia, serif';
      ctx.fillText(`Level ${{currentLevel}}`, plotLeft, 28);
    }}

    function syncReadout() {{
      const iso = payload.intervals[currentIntervalIndex];
      const summary = payload.interval_summary[currentIntervalIndex];
      intervalLabel.textContent = fmtInterval(iso);
      summaryLabel.textContent = `${{summary.active_elements}} elements active`;
      metricActive.textContent = summary.active_elements.toLocaleString();
    }}

    function render() {{
      draw();
      syncReadout();
    }}

    function play() {{
      if (timerId !== null) return;
      timerId = window.setInterval(() => {{
        currentIntervalIndex = (currentIntervalIndex + 1) % payload.intervals.length;
        intervalSlider.value = String(currentIntervalIndex);
        render();
      }}, 500);
    }}

    function pause() {{
      if (timerId !== null) {{
        window.clearInterval(timerId);
        timerId = null;
      }}
    }}

    intervalSlider.addEventListener('input', event => {{
      currentIntervalIndex = Number(event.target.value);
      render();
    }});

    levelSelect.addEventListener('change', event => {{
      currentLevel = event.target.value;
      render();
    }});

    document.getElementById('play-btn').addEventListener('click', play);
    document.getElementById('pause-btn').addEventListener('click', pause);

    populateUi();
    render();
  </script>
</body>
</html>
"""


if __name__ == "__main__":
    payload = build_payload()
    OUTPUT_HTML_PATH.write_text(build_html(payload), encoding="utf-8")
    print(f"Wrote {OUTPUT_HTML_PATH.name}")
