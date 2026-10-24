"""Static checks for the UI pages (static/index.html + static/hebdo.html):
JS syntax, JS<->HTML id wiring, endpoint references, option mappings.

Run:  python check_ui.py
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent
failures: list[str] = []
summaries: list[str] = []

import main  # noqa: E402

routes = {r.path for r in main.app.routes}


def check_page(filename: str) -> None:
    html = (BASE / "static" / filename).read_text(encoding="utf-8")

    # Pages may contain several inline <script> blocks (e.g. the early theme
    # bootstrap in <head> + the main IIFE): check each one, use all of them
    # for the cross-reference checks.
    script_blocks = re.findall(r"<script>(.*?)</script>", html, re.S)
    script = "\n".join(script_blocks)

    # ---- 1. ids referenced by the JS exist in the markup -------------------
    used = set(re.findall(r'\$\("([A-Za-z0-9_]+)"\)', script))
    defined = set(re.findall(r'id="([A-Za-z0-9_]+)"', html))
    missing = sorted(used - defined - {"spin"})  # "spin" is created at runtime
    if missing:
        failures.append(f"{filename}: JS references ids that do not exist: {missing}")
    summaries.append(f"{filename:14s}: {len(used)} js ids referenced, {len(defined)} defined")

    # ---- 2. JS syntax (via esprima if available) ---------------------------
    try:
        import esprima  # type: ignore

        errors = []
        for index, block in enumerate(script_blocks):
            try:
                esprima.parseScript(block)
            except Exception as exc:  # noqa: BLE001
                errors.append(f"block #{index + 1}: {exc}")
        if errors:
            failures.append(f"{filename}: JS syntax error: {'; '.join(errors)}")
            syntax = f"ERROR: {'; '.join(errors)}"
        else:
            syntax = f"parsed OK ({len(script_blocks)} block(s))"
    except ImportError:
        syntax = "skipped (esprima not installed)"
    summaries.append(f"{filename:14s}: js syntax {syntax}")

    # ---- 3. endpoint references match the backend -------------------------
    for path in set(re.findall(r'fetch\("([^"]+)"', script)):
        if path not in routes:
            failures.append(f"{filename}: JS calls an endpoint that does not exist: {path}")


check_page("index.html")
check_page("hebdo.html")

# ---- 4. every data-opt maps to a real Options field (converter page) ------
import converter  # noqa: E402

index_html = (BASE / "static" / "index.html").read_text(encoding="utf-8")
options_fields = {f for f in converter.Options.__dataclass_fields__}
html_keys = set(re.findall(r'data-opt="([a-z_]+)"', index_html))
html_keys |= set(re.findall(r'data-opt-single="([a-z_]+)"', index_html))
unknown = sorted(html_keys - options_fields)
if unknown:
    failures.append(f"data-opt keys unknown to Options: {unknown}")
uncontrolled = sorted(options_fields - html_keys)
if uncontrolled:
    failures.append(f"Options without a UI control: {uncontrolled}")
summaries.append(f"option controls  : {len(html_keys)} wired to Options({len(options_fields)} fields)")

# ---- 5. every hebdo data-col maps to a known source column role -----------
import hebdo  # noqa: E402

hebdo_html = (BASE / "static" / "hebdo.html").read_text(encoding="utf-8")
col_keys = set(re.findall(r'data-col="([a-z_]+)"', hebdo_html))
unknown_cols = sorted(col_keys - set(hebdo.ROLES))
if unknown_cols:
    failures.append(f"data-col keys unknown to hebdo.ROLES: {unknown_cols}")
unwired = sorted(set(hebdo.ROLES) - col_keys)
if unwired:
    failures.append(f"hebdo column roles without a UI control: {unwired}")
summaries.append(f"hebdo columns    : {len(col_keys)} wired to hebdo.ROLES({len(hebdo.ROLES)} roles)")

# ---- 6. the pages link to each other --------------------------------------
if 'href="/hebdo"' not in index_html:
    failures.append("index.html: no link to /hebdo")
hebdo_page = (BASE / "static" / "hebdo.html").read_text(encoding="utf-8")
if 'href="/"' not in hebdo_page:
    failures.append("hebdo.html: no link back to /")

# ---- 7. theme switcher on both pages --------------------------------------
for page_name, page_html in (("index.html", index_html), ("hebdo.html", hebdo_page)):
    if 'id="themeToggle"' not in page_html:
        failures.append(f"{page_name}: no theme switcher (id=themeToggle)")
    if 'id="themeIcon"' not in page_html:
        failures.append(f"{page_name}: theme switcher has no icon span")
    if "localStorage" not in page_html or 'setAttribute("data-theme"' not in page_html:
        failures.append(f"{page_name}: theme is not persisted via localStorage/data-theme")
    if 'localStorage.getItem("theme")' not in page_html:
        failures.append(f"{page_name}: no early theme bootstrap before first paint")
    if '[data-theme="dark"]' not in page_html:
        failures.append(f"{page_name}: no dark-theme CSS block")
    if "[hidden] { display: none !important" not in page_html:
        failures.append(f"{page_name}: no [hidden] display guard "
                        f"(toast/panels stay visible despite the hidden attribute)")

# ---- 8. mandatory, digits-only week input ---------------------------------
if 'aria-required="true"' not in hebdo_page or 'title="Required"' not in hebdo_page:
    failures.append("hebdo.html: week number is not marked as mandatory (*)")
if 'inputmode="numeric"' not in hebdo_page or 'pattern="[0-9]*"' not in hebdo_page:
    failures.append("hebdo.html: week input is not restricted to digits")
if r'value.replace(/[^0-9]/g, "")' not in hebdo_page:
    failures.append("hebdo.html: JS does not strip non-numeric week characters")
if "weekDigits()" not in hebdo_page or "weekOk" not in hebdo_page:
    failures.append("hebdo.html: buttons are not gated on a valid week")
if 'data-col="duration"' not in hebdo_page:
    failures.append("hebdo.html: no column override for the Duree (duration) role")

# ---- 9. zero vs no-data rendering in the preview tables -------------------
if 'n === null || n === undefined ? "-" : String(n)' not in hebdo_page:
    failures.append("hebdo.html: countText does not render null counts as '-'")
if 'n === null || n === undefined ? "num dash" : "num"' not in hebdo_page:
    failures.append("hebdo.html: countCls does not style a real 0 as numeric (not dash)")
if 'rate === null || rate === undefined ? "-" : (rate * 100).toFixed(2) + "%"' not in hebdo_page:
    failures.append("hebdo.html: rateText does not render null rates as '-'")
if 'data-col="sla"' not in hebdo_page or "Statut SLA Retablissement" not in hebdo_page:
    failures.append("hebdo.html: no SLA/status column mapping mentioning "
                    "'Statut SLA Retablissement' (GNOC header)")
if "counts open tickets of the status column" not in hebdo_page:
    failures.append("hebdo.html: output note does not explain the Ouvert status rules")
if "Duree &gt; 0" not in hebdo_page:
    failures.append("hebdo.html: output note does not state the MTTR Duree > 0 filter")
summaries.append("hebdo rules    : theme switcher (2 pages) + mandatory digits-only week + "
                 "zero-vs-'-' preview cells")

for line in summaries:
    print(line)
print(f"routes           : {sorted(p for p in routes if p.startswith('/api') or p in ('/', '/hebdo'))}")

if failures:
    print(f"\nFAILED ({len(failures)}):")
    for f in failures:
        print("  -", f)
    sys.exit(1)
print("\nOK: UI checks passed")
