# Renewal Quote Comparator (CCWR This–Last Analysis)

A small app that compares **last year** and **this year** Cisco renewal quote Excel
files and produces a detailed Excel report (value bridge, add/remove, SLA/tier,
duration, quantity, LDOS, and data‑quality sheets). It also includes a
**Price Change Impact Analysis** tab for assessing list price changes on active quotes.

Everything runs **locally on your computer**. No cloud, no AI API, and your quote
files are never uploaded anywhere.

---

## Recent updates

- **Fix (Sep 2026): SLA/service-level changes were misreported as a new serial
  number and instance number added/removed.** Cisco encodes the SLA/tier in the
  service SKU itself (e.g. `CON-SNTP-...` → `CON-L24HR-...`), so an SLA renewal
  on the exact same serial/instance used to be misclassified as a removed old-SKU
  line plus an added new-SKU line. Year-over-year matching now uses **serial
  number + instance number only** (SKU differences on a matched line are
  reported as an SLA/tier change instead), so the Summary's "Higher/Lower SLA /
  Tier value" and the **SLA Changes** sheet now correctly capture these.

---

## How to get the app — choose your option

### Option A — SharePoint portable bundle (recommended for first-time users)

Download the ready-to-run ZIP from the shared SharePoint link:

- **Windows**: download `CCWR-Windows-Portable.zip` → unzip → double-click **`Run App.bat`**
  - Python and all packages are bundled — **no installation needed**
  - If Windows SmartScreen warns you, click **More info → Run anyway**
- **Mac**: download `CCWR-Mac-Portable.zip` → unzip → right-click **`Run App.command`** → **Open**

That's it. Your browser opens the app automatically.

### Option B — GitHub source code (for users with Cisco enterprise GitHub access)

> **Important:** Use Option A for your very first install. Once Python is set up,
> you can switch to this method for easy updates.

1. Go to the GitHub repo and click **Code → Download ZIP**
2. Unzip into the same folder as your previous install (overwrite existing files)
3. Double-click **`Run App.bat`** (Windows) or **`Run App.command`** (Mac) as normal
4. The launcher automatically installs any new packages — you don't need to do anything extra

**Getting updates (Option B):** Repeat steps 1–3 whenever a new version is released.
Any new required packages are installed automatically by the launcher on first run.

---

## What you need first (one time)

You only need **Python 3** installed. The app installs everything else by itself
the first time it runs.

- **Windows** — install from <https://www.python.org/downloads/>. On the first
  installer screen, **tick “Add python.exe to PATH”**, then click *Install Now*.
- **macOS** — Python 3 usually ships with recent macOS. If not, install it from
  <https://www.python.org/downloads/macos/> or run `xcode-select --install`.

To check it worked, open a terminal / command prompt and run:

```bash
python --version
```

(macOS may need `python3 --version`.) You should see `Python 3.x.x`.

---

## Get the code

Ask colleagues to download it directly from GitHub — no manual zipping needed:

1. Open the repo: <https://github.com/camornpi_cisco/CCWR-This-Last-Analysis>
2. Click the green **Code** button → **Download ZIP**.
3. **Unzip** the file.
4. Open the unzipped folder.

> Do **not** upload real quote / SNIFF Excel files to GitHub. You only upload
> them inside the app, on your own machine, after it starts.

---

## Run it — the easy way (double‑click)

### Windows

1. Double‑click **`Run App.bat`**.
2. **If Python is not installed, the launcher installs it for you automatically**
   using the built‑in Windows Package Manager (winget). When it finishes it will
   ask you to **close the window and double‑click `Run App.bat` again** — do that,
   and this time it will start the app.
3. The first successful run creates a local environment and installs packages
   (takes a minute), then opens the app in your browser.
4. If Windows SmartScreen shows *“Windows protected your PC”*: click
   **More info → Run anyway**.

> No manual Python download needed on Windows 10/11 — the launcher handles it.
> If your PC is too old for winget, the launcher shows the exact manual download
> steps (tick **“Add python.exe to PATH”** during install).

### macOS

macOS blocks scripts downloaded from the internet the first time (Gatekeeper).
Pick **one** of these:

**Option A – right‑click open (easiest):**

1. **Right‑click** (or Control‑click) **`Run App.command`** → **Open**.
2. In the warning dialog, click **Open** again. macOS remembers this next time.

**Option B – if macOS still blocks it:**

1. Open **System Settings → Privacy & Security**.
2. Scroll down to the security message about `Run App.command`.
3. Click **Open Anyway**, then run the file again.

**Option C – clear the block for the whole folder (Terminal one‑liner):**

```bash
cd "/path/to/the/unzipped/folder"
xattr -dr com.apple.quarantine .
```

Then double‑click **`Run App.command`** normally.

> Tip: if double‑click opens the file in a text editor instead of running it,
> right‑click → **Open With → Terminal**, or use the Manual Run steps below.

When the app starts, your browser opens automatically at
<http://localhost:8501> (or a nearby port). To stop the app, close the terminal
window it opened, or press **Ctrl + C** in it.

> ### ⚠️ Keep the black terminal window open while you use the app
>
> Double‑clicking the launcher opens a **terminal / command window** (a black or
> white window full of text). **This window IS the app’s engine** — it runs the
> local server that your browser talks to.
>
> - **Leave it open** the whole time you are using the app. You can *minimise* it,
>   but do **not** close it.
> - **If you close that window, the app stops** and the browser page will freeze
>   or show a “connection lost” error.
> - When you are finished, closing the window (or pressing **Ctrl + C** in it) is
>   the correct way to shut the app down.
> - To use the app again later, just double‑click the launcher again — a new
>   window will open. It is normal for this window to appear every time.
> - Scary‑looking red text in that window is usually harmless. Only worry if the
>   browser page itself shows an error.

---

## Run it — the manual way (if the double‑click fails)

Open a terminal / command prompt **in the app folder**, then:

**Windows (Command Prompt or PowerShell):**

```bat
python -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt
.venv\Scripts\streamlit run app.py --browser.gatherUsageStats=false
```

**macOS / Linux:**

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/streamlit run app.py --browser.gatherUsageStats=false
```

---

## Using the app

The app has two tabs: **Renewal Comparison** and **Price Change Impact**.

### Tab 1 — Renewal Comparison

1. In the left sidebar, set the **Save folder** and **File name** for the report.
2. Upload the **Last year** (blue) and **This year** (orange) quote files using the
   upload boxes at the top of the page. You can drop **several files** into one box
   — they are combined automatically.
3. Optionally upload the **SNIFF** files (blue = last year, orange = this year).
4. Click **Generate Comparison** — the button is disabled until at least one Last
   year and one This year quote file are uploaded.
5. The Excel report and PowerPoint deck save automatically to your chosen folder.
   Use **💾 Save to Folder** or **⬇ Download** buttons to save/download manually.

The app warns you if a file is in the wrong slot (e.g. a SNIFF file in a quote
box) or if the two years look swapped.

**Output Excel sheets include:** Summary (value bridge), SW Subscription Analysis,
Line Detail Summary, Serial/Instance Number Changes, SKU Price Changes, SLA Changes,
Line Comparison Detail, LDOS Exceptions, Data Quality Issues, and more.

> **Auto‑numbering:** If the output file already exists, the app appends a number
> (`renewal_comparison (1).xlsx`, `(2).xlsx`, etc.) instead of overwriting. The
> paired `.pptx` file uses the same number.

#### Summary sheet section structure

| Section | Content |
|---|---|
| A / B / C | Last year total / This year total / Net change (bridge identity: C = D + E) |
| D | **Add value** — D.1 Add SN#, D.2 Add Instance#, D.3–D.5 Sub‑TnC adds, D.6.1–D.6.4 Shapley effects (duration/SLA/price/qty) |
| E | **Removal value** — E.1 Remove SN# (E.1.1 LDOS / E.1.2 by customer), E.2 Remove Instance# (E.2.1 LDOS / E.2.2 by customer), E.3–E.5 Sub‑TnC removes, E.6.1–E.6.4 Shapley effects |
| F | **Quantity & item counts — all** (includes zero‑price lines) |
| G | **Quantity & item counts — priced items only** — G.1 TS Add, G.2 TS Remove (G.2.1 due to LDOS, G.2.2 by customer), G.3 This Year Renew to LDOS Date, G.4–G.6 Sub‑TnC add/remove/tier changes |

Each removal row that reaches LDOS (`end_date == ldos_date`) is split out separately
from customer‑driven removals (decommission / tech refresh), both in the dollar
bridge (E) and the quantity counts (G).

#### PowerPoint summary deck

Generated automatically alongside the Excel report (or on demand from the app).
7 slides, built with `python-pptx` on the Cisco theme template:

1. **Cover** — customer name, period, date
2. **Executive Summary** — Last year / This year / Net change KPI cards
3. **Value Bridge – Service (TS)** — waterfall chart + two grouped tables:
   green **Add Value** table (with TOTAL row) and amber **Removal Value** table
   (with TOTAL row)
4. **Value Bridge – SW Subscription (Sub‑TnC)** — same layout as slide 3
5. **Added Items** — table of newly added priced items
6. **Removed Items** — table of removed priced items
7. **SW Subscription Changes** — Sub‑TnC add/remove/tier summary

### Tab 2 — Price Change Impact Analysis

Use this tab to measure how a Cisco list price update (Master Net Change Validation
Report) affects the prorated value of one or more active renewal quotes.

1. Upload the **Master Net Change Validation Report** (`.xlsx`) — the Cisco file
   with columns: Service SKU (C), Final Price (G), Old Price (H).
2. Upload one or more **quote files** to check. Both multi-quote format (header at
   row 3, one row per line) and single-quote format (header at row 32) are
   auto-detected. Workbooks with **multiple tabs** are handled — every sheet with
   a recognisable table is read and tagged with its own sheet/tab name.
3. Enter a **Customer / Quote name** (used in the output filename).
4. Click **Generate Price Impact Report**.

**Output Excel sheets:**

| Sheet | Content |
|---|---|
| **Summary by Quote** | One row per distinct quote (grouped by Quote Name + Quote Number + Sheet/Tab), showing full prorated total, matched prorated, new prorated, impact $ and %, plus a TOTAL row |
| **Summary by SKU** | One row per matched SKU across all quotes, with old/new unit list, % change, matched lines, and prorated change. TOTAL row at the bottom. |
| **Line Detail** | Every line with Quote Name, Quote Number, Source File, Sheet/Tab, SKU, dates, days, prorate factor, old/new unit list, old/new prorated, change, and matched flag. Filterable. |
| **Totals** | Grand totals: full quote prorated (all lines), matched-only prorated, new prorated (matched), and price impact $ / % |

**Prorated formula:** `new_unit = quote_unit × (Final ÷ Old from master)`,
`prorated = new_unit × qty × (days + 1) / 365`.

---

## Troubleshooting

| Problem | Fix |
|---|---|
| `python: command not found` (macOS) | Use `python3` instead, or install Python from python.org. |
| `python is not recognized` (Windows) | Re‑run **`Run App.bat`** — it auto‑installs Python via winget. If that isn’t available, install Python manually and tick **Add python.exe to PATH**. |
| Windows: launcher installed Python but app didn’t start | This is expected on the first install — **close the window and double‑click `Run App.bat` again** so Windows picks up the new Python. |
| macOS “cannot be opened / unidentified developer” | Use **Right‑click → Open**, or **System Settings → Privacy & Security → Open Anyway**, or run the `xattr` one‑liner above. |
| Windows “Windows protected your PC” | Click **More info → Run anyway**. |
| Streamlit asks for an email on first run | Just press **Enter** to skip — telemetry is already disabled. |
| Packages fail to install (corporate network) | You may be behind a proxy. Set `HTTP_PROXY` / `HTTPS_PROXY`, or install from an approved internal mirror. |
| The report file won’t save / “permission denied” | Close the Excel report if it’s open, or pick a different Save folder. |
| Port already in use | Another copy is running. Close the other terminal, or add `--server.port 8502`. |

---

## Testing on Windows without a second computer

You do **not** need a full Windows PC to check it works there:

1. **GitHub Actions (recommended, free, no emulator).** This repo includes a CI
   workflow (`.github/workflows/ci.yml`) that installs the app and imports all
   modules on **Windows, macOS and Linux** automatically on every push. Open the
   **Actions** tab on GitHub to see the green/red result — that confirms the
   Windows install path works.
2. **A Windows virtual machine on your Mac**, if you want to click through the UI:
   - **UTM** — free, <https://mac.getutm.app> (runs Windows 11 ARM).
   - **VMware Fusion** — free for personal use.
   - **Parallels Desktop** — paid, smoothest on Apple‑Silicon Macs.

For most cases the GitHub Actions check is enough and much lighter than a VM.

---

## Better ways to share it (optional)

- **Keep it as‑is (simplest).** Friends download the ZIP and double‑click the
  launcher. Only Python is required.
- **Standalone executable (no Python needed).** Use
  [PyInstaller](https://pyinstaller.org) to bundle the app into a single
  double‑click file. Build it **on each OS** you want to target (a macOS build
  cannot produce a Windows `.exe` — build the Windows version on Windows, or in
  GitHub Actions). Ask and I can add a small launcher + build steps.
- **Internal hosting.** If your team has an approved internal server, the app can
  be hosted there so nobody installs anything. Do **not** use the public Streamlit
  Community Cloud for real Cisco quote data.

---

## What the report calculates

- Auto‑detects the table header row and normalizes different quote export formats.
- Matches year‑over‑year lines by **serial number + instance number**. A change
  in service SKU/service level on an otherwise matched line is treated as an
  SLA/tier change, not an add + remove (service SKU is still used, together with
  serial/instance, to enrich against SNIFF and to detect overlapping split
  orders within a single quote).
- Extended list price is taken from the quote’s **prorated list price** so totals
  match the quote exactly (Cisco prorates as
  `unit list price × quantity × (inclusive coverage days ÷ 365)`).
- Caps coverage value at LDOS where applicable.
- Splits the year‑over‑year change into quantity, duration, unit‑price and
  SLA/tier effects using a **Shapley (symmetric) decomposition** so the
  attribution is fair and order‑independent, and always reconciles to the total.
- Separates hardware/service (TS) from software subscriptions (Sub‑TnC).
- Flags SNIFF coverage exceptions (expired / never‑covered) and data‑quality items.

No AI API is used. All calculations run locally.
