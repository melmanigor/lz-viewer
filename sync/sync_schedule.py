# -*- coding: utf-8 -*-
"""
לוז התקנות שבועי: מושך מ-Outlook את תכנית ההתקנות ואת מייל הסריאלים מברינקס,
מצליב ביניהם ופותח בדפדפן לוז שבועי מסודר.

הרצה רגילה (Windows עם Outlook פתוח):
    python sync_schedule.py

הרצה מקבצים מקומיים (לבדיקה, בלי Outlook):
    python sync_schedule.py --plan "תכנית התקנות.xlsx" --email-image screenshot.jpg
    python sync_schedule.py --plan "תכנית התקנות.xlsx" --email-html mail.html
"""
import argparse
import base64
import datetime as dt
import html
import json
import os
import re
import sys
import tempfile
import webbrowser
from html.parser import HTMLParser
from pathlib import Path

# ---------------------------------------------------------------- הגדרות
# אפשר לשנות כאן בלי לגעת בשאר הקוד.
DAYS_BACK = 10                                   # כמה ימים אחורה לחפש מיילים
OUTLOOK_FOLDERS = ["Inbox"]                      # תיקיות לחיפוש. "Inbox" = תיבת הדואר הנכנס (כולל תתי-תיקיות)
PLAN_ATTACHMENT_KEYWORDS = ["תכנית התקנות", "תוכנית התקנות"]   # מילה בשם קובץ האקסל
BRINKS_SUBJECT_KEYWORDS = ["סריאלים", "סיראלים", "סריאל", "סיראל"]   # מילה בנושא המייל של הסריאלים (גם בכתיב אחר)
BRINKS_SKIP_SUBJECT = ["פערים"]                  # נושאים שמדלגים עליהם (תשובות למייל הפערים)
BRINKS_SENDER_KEYWORDS = ["brinks", "ברינקס", "zarfati", "zakuta", "צרפתי"]   # מילה בכתובת או בשם השולח (ריק = כל שולח)
OUTPUT_DIR = Path.home() / "Documents" / "לוז התקנות"
# קריאת טבלה מתמונה דורשת מפתח API של Claude במשתנה הסביבה ANTHROPIC_API_KEY.
# אם המייל של ברינקס מכיל טבלה אמיתית (לא תמונה), לא צריך מפתח.
CLAUDE_MODEL = "claude-opus-5-5"
# מייל פערים: נפתח ב-Outlook כטיוטה לאישור שלך. שום דבר לא נשלח בלי שתלחץ "שלח".
GAPS_EMAIL_TO = "Ariel.Zakuta@brinksinc.com; Leonid.Rubinshtein@brinksinc.com"
GAPS_TASKS = ["התקנה", "החלפה"]        # סוגי משימות שנבדקים לפערים במייל
# ------------------------------------------------------------------------

HERE = Path(__file__).resolve().parent
TEMPLATE = HERE / "schedule_template.html"
EXCEL_EPOCH = dt.datetime(1899, 12, 30)


def log(msg):
    try:
        print(msg, flush=True)
    except Exception:
        pass  # pythonw (הרצה ברקע) אין חלון
    try:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        with open(OUTPUT_DIR / "log.txt", "a", encoding="utf-8") as fh:
            fh.write(f"{dt.datetime.now():%d/%m %H:%M} {msg}\n")
    except Exception:
        pass


def notify(title, body):
    """התראה בפינת המסך של Windows. אם לא אפשרי, חלון הודעה רגיל."""
    if os.name != "nt":
        log(f"[התראה] {title}: {body}")
        return
    import subprocess
    esc = lambda s: str(s).replace("'", "''").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    ps = ("[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] > $null;"
          "$x = New-Object Windows.Data.Xml.Dom.XmlDocument;"
          f"$x.LoadXml('<toast><visual><binding template=\"ToastGeneric\"><text>{esc(title)}</text><text>{esc(body)}</text></binding></visual></toast>');"
          "$app = '{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\\WindowsPowerShell\\v1.0\\powershell.exe';"
          "[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($app).Show([Windows.UI.Notifications.ToastNotification]::new($x))")
    try:
        subprocess.run(["powershell", "-NoProfile", "-WindowStyle", "Hidden", "-Command", ps],
                       check=True, timeout=20, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except Exception:
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(0, body, title, 0x40 | 0x40000)
        except Exception:
            log(f"[התראה] {title}: {body}")


# ================================================================ Outlook
def iter_folders(folder):
    yield folder
    for sub in folder.Folders:
        yield from iter_folders(sub)


def outlook_available():
    try:
        import win32com.client  # noqa: F401
        return True
    except ImportError:
        return False


def sender_text(item):
    """שם + כתובת השולח. בחשבון ארגוני הכתובת מגיעה לפעמים בצורה פנימית (/O=EXCHANGELABS/...), אז מבקשים גם את כתובת ה-SMTP."""
    parts = [getattr(item, "SenderName", "") or "", getattr(item, "SenderEmailAddress", "") or ""]
    try:
        eu = item.Sender.GetExchangeUser()
        if eu:
            parts.append(eu.PrimarySmtpAddress or "")
    except Exception:
        pass
    return " ".join(parts).lower()


def brinks_from_item(item, received, workdir, idx):
    images = []
    for i in range(item.Attachments.Count):
        a = item.Attachments.Item(i + 1)
        if (a.FileName or "").lower().endswith((".png", ".jpg", ".jpeg", ".gif", ".bmp")):
            p = workdir / f"brinks_{idx}_{i}_{a.FileName}"
            a.SaveAsFile(str(p))
            images.append(p)
    return {"html": item.HTMLBody or "", "images": images, "received": received, "subject": item.Subject or ""}


def pick_brinks(cands, workdir):
    """מבין המיילים עם "סריאלים לכספות": קודם מה שנשלח מברינקס, ובתוך זה מהחדש לישן.
    בוחרים את הראשון שיש בו טבלה שאפשר לקרוא. אם אין טבלה באף אחד, את החדש שיש בו תמונה."""
    if not cands:
        log("✗ לא נמצא מייל שבנושא שלו " + " / ".join(BRINKS_SUBJECT_KEYWORDS) + f" ב-{DAYS_BACK} הימים האחרונים")
        return None
    cands.sort(key=lambda c: (c[0], c[1]), reverse=True)
    if not cands[0][0]:
        log(f"  (אף מייל סריאלים לא זוהה לפי השולח {BRINKS_SENDER_KEYWORDS}. משתמש לפי הנושא בלבד)")
    with_image = None
    for idx, (by_sender, received, item) in enumerate(cands[:6]):
        b = brinks_from_item(item, received, workdir, idx)
        if rows_from_html(b["html"]):
            log(f"✓ מייל ברינקס: {b['subject']} (מ-{received:%d/%m %H:%M}), עם טבלה")
            return b
        if b["images"] and not with_image:
            with_image = b
        log(f"  - {b['subject']} ({received:%d/%m %H:%M}): אין בו טבלה" + (f", יש {len(b['images'])} תמונות" if b["images"] else ""))
    if with_image:
        log(f"✓ מייל ברינקס: {with_image['subject']} (מ-{with_image['received']:%d/%m %H:%M}), הטבלה בתמונה")
        return with_image
    log("✗ נמצאו מיילי סריאלים, אבל באף אחד מהם אין טבלה או תמונה")
    return None


def fetch_from_outlook(workdir: Path):
    """מחזיר (נתיב אקסל, מידע על מייל ברינקס) מהמיילים האחרונים."""
    try:
        import win32com.client  # pywin32
    except ImportError:
        sys.exit("חסרה הספרייה pywin32. התקן עם: pip install pywin32")

    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    since = dt.datetime.now() - dt.timedelta(days=DAYS_BACK)
    roots = []
    for name in OUTLOOK_FOLDERS:
        roots.append(ns.GetDefaultFolder(6) if name.lower() == "inbox" else ns.Folders.Item(1).Folders[name])

    plan_hit, brinks_cands = None, []  # (received, item)
    for root in roots:
        for folder in iter_folders(root):
            try:
                items = folder.Items
                items.Sort("[ReceivedTime]", True)
                items = items.Restrict("[ReceivedTime] >= '" + since.strftime("%m/%d/%Y %H:%M %p") + "'")
            except Exception:
                continue
            for item in items:
                if getattr(item, "Class", None) != 43:  # 43 = MailItem
                    continue
                rt = item.ReceivedTime
                received = dt.datetime(rt.year, rt.month, rt.day, rt.hour, rt.minute)
                subject = item.Subject or ""
                sender = sender_text(item)
                atts = [item.Attachments.Item(i + 1) for i in range(item.Attachments.Count)]
                if any(k in (a.FileName or "") for a in atts for k in PLAN_ATTACHMENT_KEYWORDS
                       ) and any((a.FileName or "").lower().endswith((".xlsx", ".xlsm", ".xls")) for a in atts):
                    if not plan_hit or received > plan_hit[0]:
                        plan_hit = (received, item)
                # מייל הסריאלים. לא תשובות למייל הפערים שלך (גם בהן מופיע "סריאלים לכספות").
                if any(k in subject for k in BRINKS_SUBJECT_KEYWORDS) and not any(k in subject for k in BRINKS_SKIP_SUBJECT):
                    by_sender = not BRINKS_SENDER_KEYWORDS or any(k.lower() in sender for k in BRINKS_SENDER_KEYWORDS)
                    brinks_cands.append((by_sender, received, item))

    plan_path = None
    if plan_hit:
        item = plan_hit[1]
        for i in range(item.Attachments.Count):
            a = item.Attachments.Item(i + 1)
            if a.FileName.lower().endswith((".xlsx", ".xlsm", ".xls")) and any(k in a.FileName for k in PLAN_ATTACHMENT_KEYWORDS):
                plan_path = workdir / a.FileName
                a.SaveAsFile(str(plan_path))
                log(f"✓ תכנית התקנות: {a.FileName} (מייל מ-{plan_hit[0]:%d/%m %H:%M})")
                break
    else:
        log("✗ לא נמצא מייל עם קובץ תכנית התקנות בימים האחרונים")

    brinks = pick_brinks(brinks_cands, workdir)
    return plan_path, brinks


# ================================================================ Plan (Excel)
def load_plan_rows(path: Path, keep_days_back=21):
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = None
    for sheet in wb.worksheets:
        for row in sheet.iter_rows(max_row=5, values_only=True):
            if any(str(c or "").strip() == "שם אתר" for c in row):
                ws = sheet
                break
        if ws:
            break
    if ws is None:
        sys.exit("לא נמצא גיליון עם עמודת 'שם אתר' בקובץ התכנית")

    def cv(v):
        if isinstance(v, dt.datetime):
            return (v - EXCEL_EPOCH).days
        if isinstance(v, dt.date):
            return (dt.datetime(v.year, v.month, v.day) - EXCEL_EPOCH).days
        return v

    rows = [[cv(c) for c in r] for r in ws.iter_rows(values_only=True)]
    header_idx = next(i for i, r in enumerate(rows) if any(str(c or "").strip() == "שם אתר" for c in r))
    cutoff = (dt.datetime.now() - dt.timedelta(days=keep_days_back) - EXCEL_EPOCH).days
    start = next((i for i, r in enumerate(rows) if i > header_idx and isinstance(r[0], int) and r[0] >= cutoff), header_idx + 1)
    return [rows[header_idx]] + rows[start:]


# ================================================================ Brinks email
COLS = {
    "date": ["תאריך"], "day": ["יום"], "site": ["שם אתר"], "supplier": ["ספק"],
    "device": ["מספר כספת", "מכשיר"], "model": ["דגם"], "ins": ["ביטוח"], "plate": ["פלטת", "פלטה"],
    "task": ["סוג משימה", "משימה"], "serial": ["סריאל"],
}


class TableParser(HTMLParser):
    """אוסף טבלאות HTML כרשימות שורות, עם טיפול ב-rowspan/colspan."""

    def __init__(self):
        super().__init__()
        self.tables, self._stack = [], []

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if tag == "table":
            self._stack.append({"rows": [], "row": None, "cell": None})
        elif not self._stack:
            return
        elif tag == "tr":
            self._stack[-1]["row"] = []
        elif tag in ("td", "th") and self._stack[-1]["row"] is not None:
            self._stack[-1]["cell"] = {"text": "", "rs": int(a.get("rowspan") or 1), "cs": int(a.get("colspan") or 1)}
        elif tag == "br" and self._stack[-1]["cell"] is not None:
            self._stack[-1]["cell"]["text"] += " "

    def handle_endtag(self, tag):
        if not self._stack:
            return
        t = self._stack[-1]
        if tag in ("td", "th") and t["cell"] is not None:
            t["row"].append(t["cell"]); t["cell"] = None
        elif tag == "tr" and t["row"] is not None:
            t["rows"].append(t["row"]); t["row"] = None
        elif tag == "table":
            self.tables.append(self._grid(self._stack.pop()["rows"]))

    def handle_data(self, data):
        if self._stack and self._stack[-1]["cell"] is not None:
            self._stack[-1]["cell"]["text"] += data

    @staticmethod
    def _grid(rows):
        grid, carry = [], {}
        for r in rows:
            out, col, cells = [], 0, list(r)
            while cells or col in carry:
                if col in carry:
                    text, left = carry[col]
                    out.append(text)
                    if left > 1: carry[col] = (text, left - 1)
                    else: del carry[col]
                    col += 1
                    continue
                c = cells.pop(0)
                text = re.sub(r"\s+", " ", html.unescape(c["text"])).strip()
                for _ in range(c["cs"]):
                    out.append(text)
                    if c["rs"] > 1: carry[col] = (text, c["rs"] - 1)
                    col += 1
            grid.append(out)
        return grid


def parse_date(s):
    s = str(s or "").strip()
    m = re.match(r"(\d{1,2})[/.](\d{1,2})[/.](\d{2,4})", s)
    if m:
        y = m.group(3); y = "20" + y if len(y) == 2 else y
        return f"{y}-{int(m.group(2)):02d}-{int(m.group(1)):02d}"
    m = re.match(r"(\d{4})-(\d{1,2})-(\d{1,2})", s)
    return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}" if m else None


def rows_from_html(body: str):
    p = TableParser(); p.feed(body)
    out = []
    for grid in p.tables:
        header = None
        for row in grid:
            if any("שם אתר" in c for c in row) and any("סריאל" in c for c in row):
                header = {k: next((i for i, h in enumerate(row) if any(w in h for w in ws)), -1) for k, ws in COLS.items()}
                continue
            if not header:
                continue
            get = lambda k: row[header[k]] if 0 <= header[k] < len(row) else ""
            d = parse_date(get("date"))
            if not d or not get("site"):
                continue
            out.append({"date": d, "day": get("day"), "site": get("site"), "supplier": get("supplier"),
                        "device": get("device"), "model": get("model"),
                        "ins": get("ins").strip().upper() == "V", "plate": get("plate").strip().upper() == "V",
                        "task": get("task"), "serial": get("serial")})
    return out


def keep_tasks(email_rows):
    """רק התקנה והחלפה. שורות הזזה/פירוק במייל של ברינקס נזרקות כבר בקריאה."""
    return [e for e in email_rows if norm_task(e.get("task")) in GAPS_TASKS]


IMG_PROMPT = """התמונה היא צילום מסך של מייל מחברת ברינקס עם טבלה (אחת או יותר) של סריאלים לכספות.
עמודות הטבלה (מימין לשמאל): תאריך, יום, שם אתר, ספק, מספר כספת/מכשיר, דגם, ביטוח ברינקס, מגן קיר, בסיס, פלטת יציקה, סוג משימה, סריאלי.
תא התאריך והיום ממוזג לכמה שורות: החל אותו על כל השורות בקבוצה. התעלם משורות הכותרת שחוזרות.
העתק כל שורת נתונים בדיוק כפי שהיא כתובה, כולל טקסט אדום שמצורף לשם האתר. תא ריק = מחרוזת ריקה.
החזר JSON בלבד בצורה:
{"rows":[{"date":"YYYY-MM-DD","day":"ראשון","site":"","supplier":"","device":"","model":"","ins":false,"plate":false,"task":"","serial":""}]}"""


def rows_from_images(images):
    key = os.environ.get("ANTHROPIC_API_KEY")
    if not key:
        log("! המייל של ברינקס מכיל רק תמונה. כדי לקרוא אותה הגדר ANTHROPIC_API_KEY,"
            " או גרור את התמונה לאפליקציה ב-claude.ai.")
        return []
    import urllib.request
    content = []
    for p in images:
        ext = p.suffix.lower().lstrip(".")
        mt = {"jpg": "image/jpeg", "jpeg": "image/jpeg", "png": "image/png", "gif": "image/gif"}.get(ext)
        if not mt:
            continue
        content.append({"type": "image", "source": {"type": "base64", "media_type": mt,
                                                     "data": base64.b64encode(p.read_bytes()).decode()}})
    content.append({"type": "text", "text": IMG_PROMPT})
    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages",
        data=json.dumps({"model": CLAUDE_MODEL, "max_tokens": 8000,
                         "messages": [{"role": "user", "content": content}]}).encode(),
        headers={"x-api-key": key, "anthropic-version": "2023-06-01", "content-type": "application/json"})
    with urllib.request.urlopen(req, timeout=180) as r:
        data = json.loads(r.read())
    text = "".join(b.get("text", "") for b in data.get("content", []) if b.get("type") == "text")
    m = re.search(r"\{.*\}", text, re.S)
    rows = json.loads(m.group(0)).get("rows", []) if m else []
    for r in rows:
        r["date"] = parse_date(r.get("date")) or r.get("date")
    return [r for r in rows if r.get("date")]


# ================================================================ Gaps
DAYS_HE = ["שני", "שלישי", "רביעי", "חמישי", "שישי", "שבת", "ראשון"]  # לפי weekday() של פייתון
TASK_WORDS = ["התקנה", "החלפה", "הזזה", "פירוק"]


def clean(v):
    return re.sub(r"\s+", " ", str(v if v is not None else "")).strip()


def dev_no(v):
    return re.sub(r"\D", "", clean(v))


def norm_task(v):
    s = clean(v)
    return next((t for t in TASK_WORDS if t in s), s or "אחר")


def serial_to_date(v):
    if isinstance(v, (int, float)) and 20000 < v < 80000:
        return (EXCEL_EPOCH + dt.timedelta(days=int(v))).strftime("%Y-%m-%d")
    return parse_date(v)


def parse_plan(rows):
    """אותו פענוח כמו בדף: שורות תכנית עם תאריך, אתר, מספר כספת וסוג משימה."""
    header = [clean(h) for h in rows[0]]

    def col(*keys):
        for k in keys:
            if k in header:
                return header.index(k)
        for k in keys:
            for i, h in enumerate(header):
                if h and k in h:
                    return i
        return -1

    c_date, c_site, c_dev, c_model, c_notes = col("תאריך"), col("שם אתר"), col("מספר כספת/מכשיר", "מספר כספת"), col("דגם"), col("הערות")
    c_addr, c_region = col("כתובת"), col("מבצע")
    width = max(len(r) for r in rows[:400])
    counts = [sum(1 for r in rows[1:400] if c < len(r) and re.match(r"^(התקנה|החלפה|הזזה|פירוק)", clean(r[c]))) for c in range(width)]
    if 0 <= c_notes < width:
        counts[c_notes] = 0
    c_task = max(range(width), key=lambda c: counts[c]) if any(counts) else -1
    out, cur = [], None
    for r in rows[1:]:
        g = lambda c: r[c] if 0 <= c < len(r) else None
        if clean(g(c_site)) == "שם אתר":
            continue
        d = serial_to_date(g(c_date))
        if d:
            cur = d
        site = clean(g(c_site))
        if re.match(r"^צוות\s*\d+", site) or not cur or (not site and not dev_no(g(c_dev))):
            continue
        out.append({"date": cur, "site": site, "device": clean(g(c_dev)), "model": clean(g(c_model)), "task": norm_task(g(c_task)),
                    "address": clean(g(c_addr)), "region": clean(g(c_region))})
    return [r for r in out if r["task"] in GAPS_TASKS]   # רק התקנה והחלפה. הזזה, פירוק וכל השאר לא נכנסים בכלל


def site_sim(a, b):
    tok = lambda s: [w for w in re.sub(r"[^\w\s]", " ", re.split(r"-לקוח|–|\(", clean(s))[0]).split() if len(w) > 1]
    A, B = tok(a), set(tok(b))
    return sum(1 for w in A if w in B) / min(len(A), len(B)) if A and B else 0


def match_pairs(plan_rows, email_rows):
    """מתאים כל שורה ממייל ברינקס לשורה בתכנית (לפי מספר כספת, ואם אין - לפי שם האתר).
    מחזיר (plan, email, pairs) כש-pairs היא רשימת (אינדקס בתכנית או -1, שורת מייל)."""
    plan = parse_plan(plan_rows)
    plan, email_rows = apply_overrides(plan, [dict(e) for e in email_rows], load_overrides())
    days_apart = lambda a, b: abs((dt.date.fromisoformat(a) - dt.date.fromisoformat(b)).days)
    used, pairs = set(), []
    for e in email_rows:
        dn, pi = dev_no(e.get("device")), -1
        if dn:
            cands = [i for i, p in enumerate(plan) if i not in used and dev_no(p["device"]) == dn and days_apart(p["date"], e["date"]) <= 7]
            cands.sort(key=lambda i: days_apart(plan[i]["date"], e["date"]))
            pi = cands[0] if cands else -1
        if pi < 0:
            best = 0.6
            for i, p in enumerate(plan):
                if i in used or days_apart(p["date"], e["date"]) > 7:
                    continue
                s = site_sim(p["site"], e.get("site", "")) + (0.01 if p["date"] == e["date"] else 0)
                if s >= best:
                    best, pi = s, i
        if pi >= 0:
            used.add(pi)
        pairs.append((pi, e))
    return plan, pairs


def find_gaps(plan_rows, email_rows):
    """פערים רק במספר כספת (כששני המספרים קיימים ושונים) וביום. האקסל קובע."""
    plan, pairs = match_pairs(plan_rows, email_rows)
    gaps = []
    for pi, e in pairs:
        if pi < 0:
            continue
        p = plan[pi]
        dn = dev_no(e.get("device"))
        if p["task"] not in GAPS_TASKS:
            continue
        dev_diff = bool(dn and dev_no(p["device"]) and dev_no(p["device"]) != dn)
        day_diff = p["date"] != e["date"]
        if dev_diff or day_diff:
            gaps.append({"plan": p, "email": e, "dev": dev_diff, "day": day_diff})
    # באקסל ולא במייל הסריאלים (באותו שבוע): ברינקס צריכים לאשר
    matched = {pi for pi, _ in pairs if pi >= 0}
    weeks = {week_of(e["date"]) for e in email_rows if e.get("date")}
    for i, p in enumerate(plan):
        if i not in matched and p["task"] in GAPS_TASKS and week_of(p["date"]) in weeks:
            gaps.append({"plan": p, "email": {}, "dev": False, "day": False, "missing": True})
    return gaps


def day_label(d):
    x = dt.date.fromisoformat(d)
    return f"{DAYS_HE[x.weekday()]} {x.day}/{x.month}"


def gaps_email_html(gaps):
    td = 'style="border:1px solid #bbb;padding:4px 8px"'
    rows = ""
    for g in gaps:
        p, e = g["plan"], g["email"]
        what = []
        if g["dev"]:
            what.append(f"מספר כספת: אצלכם {html.escape(clean(e.get('device')))}, בתכנית {html.escape(p['device'])}")
        if g["day"]:
            what.append(f"יום: אצלכם {day_label(e['date'])}, בתכנית {day_label(p['date'])}")
        if g.get("missing"):
            what.append("חוסר התאמה: המשימה בתכנית, אבל לא מופיעה בטבלת הסריאלים של אלי")
        rows += (f"<tr><td {td}>{html.escape(p['site'])}</td><td {td}>{html.escape(p['task'])}</td>"
                 f"<td {td}>{day_label(p['date'])}</td><td {td}>{html.escape(p['device'])}</td>"
                 f"<td {td}>{html.escape(clean(e.get('serial')))}</td><td {td}>{'<br>'.join(what)}</td></tr>")
    return (f'<div dir="rtl" style="font-family:Arial,sans-serif;font-size:11pt">'
            f"<p>אריאל, לאוניד שלום,</p><p>מצאתי חוסר התאמה בין טבלת הסריאלים של אלי לבין תכנית ההתקנות. "
            f"הנתונים הנכונים הם לפי התכנית:</p>"
            f'<table style="border-collapse:collapse" dir="rtl"><tr>'
            f"<th {td}>אתר</th><th {td}>משימה</th><th {td}>יום לפי התכנית</th><th {td}>מספר כספת לפי התכנית</th>"
            f"<th {td}>סריאל</th><th {td}>הפער</th></tr>{rows}</table>"
            f"<p>אשמח שתבדקו ותעדכנו.</p><p>תודה,</p></div>")


def open_gaps_draft(gaps):
    """פותח ב-Outlook טיוטה לאישור. לא שולח."""
    import win32com.client
    mail = win32com.client.Dispatch("Outlook.Application").CreateItem(0)
    mail.To = GAPS_EMAIL_TO
    mail.Subject = f"פערים בסריאלים לכספות ({len(gaps)})"
    mail.Display()                       # מציג את החלון קודם, כדי שהחתימה שלך תיכנס
    mail.HTMLBody = gaps_email_html(gaps) + mail.HTMLBody
    log(f"✉ נפתחה טיוטה ל-{GAPS_EMAIL_TO}. בדוק אותה ולחץ 'שלח' ב-Outlook רק אם אתה מאשר.")


# ================================================================ Brinks replies
GAPS_SUBJECT = "פערים בסריאלים לכספות"
REPLY_PROMPT = """אתה עוזר לרכז לוז התקנות כספות. שלחנו לחברת ברינקס מייל עם פערים בין תכנית ההתקנות שלנו לבין טבלת הסריאלים שלהם.
להלן הפערים הפתוחים (JSON):
{gaps}

להלן התשובה שהתקבלה מברינקס (טקסט חופשי, בלי הציטוט של המייל המקורי):
---
{reply}
---

חלץ מהתשובה רק שינויים מפורשים שהם מבקשים או מאשרים, לכל כספת:
- field "date": יום/תאריך אחר להתקנה (value בפורמט YYYY-MM-DD; השבוע הרלוונטי מופיע בפערים).
- field "device": מספר כספת אחר (value = המספר).
- field "serial": סריאל אחר או חדש (value = הסריאל).
- field "confirm_plan": הם מאשרים שהתכנית שלנו נכונה ויתקנו אצלם (value ריק).
device_in_plan = מספר הכספת כפי שמופיע בתכנית שלנו לאותו אתר (מתוך הפערים).
אל תמציא שינויים. אם אין שינוי ברור, החזר רשימה ריקה.
החזר JSON בלבד: {{"changes":[{{"device_in_plan":"","site":"","field":"","value":"","quote":""}}],"summary":""}}"""


def claude_json(content, max_tokens=4000):
    key = os.environ.get("ANTHROPIC_API_KEY")
    if not key:
        return None
    import urllib.request
    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages",
        data=json.dumps({"model": CLAUDE_MODEL, "max_tokens": max_tokens,
                         "messages": [{"role": "user", "content": content}]}).encode(),
        headers={"x-api-key": key, "anthropic-version": "2023-06-01", "content-type": "application/json"})
    with urllib.request.urlopen(req, timeout=180) as r:
        data = json.loads(r.read())
    text = "".join(b.get("text", "") for b in data.get("content", []) if b.get("type") == "text")
    m = re.search(r"\{.*\}", text, re.S)
    return json.loads(m.group(0)) if m else None


def overrides_file():
    return OUTPUT_DIR / "approved_changes.json"


def load_overrides():
    try:
        return json.loads(overrides_file().read_text(encoding="utf-8"))
    except Exception:
        return []


def save_overrides(items):
    overrides_file().write_text(json.dumps(items, ensure_ascii=False, indent=1), encoding="utf-8")


def apply_overrides(plan, email_rows, overrides):
    """שינויים שאישרת: יום/מספר כספת משנים את שורת התכנית, סריאל משנה את שורת המייל."""
    for o in overrides:
        k = dev_no(o.get("device_in_plan"))
        if not k:
            continue
        for p in plan:
            if dev_no(p["device"]) == k:
                if o["field"] == "date" and parse_date(o.get("value")):
                    p["date"] = parse_date(o["value"])
                elif o["field"] == "device" and dev_no(o.get("value")):
                    p["device"] = clean(o["value"])
                p["override"] = True
        target = dev_no(o.get("value")) if o["field"] == "device" else k
        for e in email_rows:
            if o["field"] == "serial" and dev_no(e.get("device")) in (k, target):
                e["serial"] = clean(o.get("value"))
    return plan, email_rows


def reply_text_only(body):
    """מחזיר רק את הטקסט החדש בתשובה, בלי הציטוט של המייל הקודם."""
    cut = re.split(r"\n\s*(?:From:|מאת:|-----Original Message-----|________________________________|On .+ wrote:)", body or "", maxsplit=1)[0]
    return cut.strip()


def fetch_replies(seen_ids):
    """תשובות חדשות מברינקס למייל הפערים."""
    import win32com.client
    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    since = dt.datetime.now() - dt.timedelta(days=DAYS_BACK)
    out = []
    for folder in iter_folders(ns.GetDefaultFolder(6)):
        try:
            items = folder.Items.Restrict("[ReceivedTime] >= '" + since.strftime("%m/%d/%Y %H:%M %p") + "'")
        except Exception:
            continue
        for item in items:
            if getattr(item, "Class", None) != 43 or GAPS_SUBJECT not in (item.Subject or ""):
                continue
            sender = f"{getattr(item, 'SenderName', '')} {getattr(item, 'SenderEmailAddress', '')}".lower()
            if not any(k.lower() in sender for k in BRINKS_SENDER_KEYWORDS):
                continue
            if item.EntryID in seen_ids:
                continue
            rt = item.ReceivedTime
            out.append({"id": item.EntryID, "from": item.SenderName, "received": f"{rt.day:02d}/{rt.month:02d} {rt.hour:02d}:{rt.minute:02d}",
                        "text": reply_text_only(item.Body)})
    return out


def review_replies(replies, gaps):
    """מציג לך כל שינוי שברינקס ביקשו, ושומר רק את מה שאישרת."""
    state_path = OUTPUT_DIR / ".replies_seen.json"
    approved = load_overrides()
    gaps_ctx = [{"site": g["plan"]["site"], "device_in_plan": g["plan"]["device"], "date_in_plan": g["plan"]["date"],
                 "device_in_brinks": clean(g["email"].get("device")), "date_in_brinks": g["email"].get("date", "")} for g in gaps]
    changed = False
    for r in replies:
        log(f"\n✉ תשובה מ-{r['from']} ({r['received']}):\n{r['text'][:1200]}\n")
        res = claude_json([{"type": "text", "text": REPLY_PROMPT.format(gaps=json.dumps(gaps_ctx, ensure_ascii=False), reply=r["text"][:6000])}])
        if res is None:
            log("  (אין ANTHROPIC_API_KEY, אז אני לא מפענח את התשובה. אם צריך לשנות, עדכן את האקסל.)")
            continue
        changes = [c for c in res.get("changes", []) if c.get("field") in ("date", "device", "serial", "confirm_plan")]
        if not changes:
            log("  אין בתשובה שינוי מפורש. " + clean(res.get("summary")))
        for c in changes:
            if c["field"] == "confirm_plan":
                log(f"  ✓ {c.get('site')}: ברינקס מאשרים שהתכנית נכונה.")
                continue
            what = {"date": "יום", "device": "מספר כספת", "serial": "סריאל"}[c["field"]]
            val = day_label(parse_date(c["value"])) if c["field"] == "date" and parse_date(c.get("value")) else c.get("value")
            ans = input(f"  ? {c.get('site')} (כספת {c.get('device_in_plan')}): לשנות {what} ל-{val}? "
                        f"[\"{clean(c.get('quote'))[:80]}\"] (כ/ל): ").strip()
            if ans in ("כ", "כן", "y", "yes", "Y"):
                c.update({"approved_at": f"{dt.datetime.now():%d/%m/%Y %H:%M}", "reply_from": r["from"]})
                approved.append(c)
                changed = True
                log("    ✓ נשמר. הלוז יתעדכן לפי זה.")
            else:
                log("    ✗ לא שונה. האקסל נשאר הקובע.")
    try:
        seen = json.loads(state_path.read_text(encoding="utf-8"))
    except Exception:
        seen = []
    state_path.write_text(json.dumps(seen + [r["id"] for r in replies]), encoding="utf-8")
    if changed:
        save_overrides(approved)
    return changed


def seen_reply_ids():
    try:
        return set(json.loads((OUTPUT_DIR / ".replies_seen.json").read_text(encoding="utf-8")))
    except Exception:
        return set()


# ================================================================ Technician viewer
UPDATE_URL = "https://raw.githubusercontent.com/melmanigor/lz-viewer/main/sync/"
UPDATE_FILES = ["sync_schedule.py", "schedule_template.html", "README.md"]   # viewer_config.json וקבצי ה-bat לא מתעדכנים


SITE_URL = "https://melmanigor.github.io/lz-viewer/"


def open_schedule(local_file):
    """פותח את הלוז. אם הלוז מתפרסם לאתר, פותחים את האתר (אותו מראה כמו באפליקציה, תמיד הגרסה האחרונה)."""
    if VIEWER_CONFIG.exists():
        webbrowser.open(SITE_URL)
    else:
        webbrowser.open(local_file.as_uri())


def self_update():
    """מוריד את הגרסה האחרונה של הקבצים מ-GitHub (חינם). אם הסקריפט עצמו השתנה, מריץ את עצמו מחדש."""
    import urllib.request
    changed = []
    for name in UPDATE_FILES:
        try:
            with urllib.request.urlopen(UPDATE_URL + name, timeout=15) as r:
                new = r.read()
            if r.status != 200 or not new:
                continue
            if name.endswith(".py"):
                compile(new, name, "exec")        # לא מחליפים בקובץ שבור
            old_path = HERE / name
            if old_path.exists() and old_path.read_bytes() == new:
                continue
            tmp = old_path.with_suffix(old_path.suffix + ".new")
            tmp.write_bytes(new)
            os.replace(tmp, old_path)
            changed.append(name)
        except Exception as ex:
            log(f"  (בדיקת עדכון ל-{name} נכשלה, ממשיך עם הגרסה הקיימת: {ex})")
    if changed:
        log("⬆ עודכן אוטומטית: " + ", ".join(changed))
    if "sync_schedule.py" in changed:
        import subprocess   # ב-Windows עדיף להריץ ולחכות מאשר execv
        sys.exit(subprocess.call([sys.executable, str(HERE / "sync_schedule.py"), "--no-update"] + sys.argv[1:]))


VIEWER_CONFIG = HERE / "viewer_config.json"   # כתובת Supabase, מפתח ציבורי וקוד אדמין. לא לשתף.


def publish_to_viewer(plan_rows, email_rows, source):
    """מעלה את הלוז לקישור של הטכנאי. רק הקוד של האדמין (אצלך) יכול לעדכן."""
    try:
        cfg = json.loads(VIEWER_CONFIG.read_text(encoding="utf-8"))
    except Exception:
        log("  (אין viewer_config.json, אז הלוז לא הועלה לקישור של הטכנאי)")
        return
    import urllib.request
    body = json.dumps({"p_token": cfg["admin_token"],
                       "p_data": {"plan": plan_rows, "email": email_rows, "source": source}},
                      ensure_ascii=False, default=str).encode()
    req = urllib.request.Request(cfg["supabase_url"].rstrip("/") + "/rest/v1/rpc/lz_put", data=body,
                                 headers={"apikey": cfg["key"], "Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            r.read()
        log("✓ הלוז עודכן בקישור של הטכנאי")
    except Exception as ex:
        log(f"! העדכון לקישור של הטכנאי נכשל: {ex}")
    archive_weeks(cfg, plan_rows, email_rows, source)


def week_of(d):
    x = dt.date.fromisoformat(d)
    return (x - dt.timedelta(days=(x.weekday() + 1) % 7)).isoformat()   # השבוע מתחיל ביום ראשון


def archive_weeks(cfg, plan_rows, email_rows, source):
    """שומר כל שבוע שיש לו מייל סריאלים בארכיון, לחיפוש לפי אתר, מספר כספת או סריאל."""
    import urllib.request
    plan, pairs = match_pairs(plan_rows, email_rows)
    matched = {pi for pi, _ in pairs if pi >= 0}
    items = []
    for pi, e in pairs:   # האקסל קובע: יום, אתר ומספר כספת מהתכנית; הסריאל מהמייל
        p = plan[pi] if pi >= 0 else {}
        items.append({"date": p.get("date") or e["date"], "site": p.get("site") or clean(e.get("site")),
                      "device": p.get("device") or clean(e.get("device")), "model": p.get("model") or clean(e.get("model")),
                      "task": p.get("task") or norm_task(e.get("task")), "serial": clean(e.get("serial"))})
    for i, p in enumerate(plan):
        if i not in matched:
            items.append({"date": p["date"], "site": p["site"], "device": p["device"], "model": p["model"], "task": p["task"], "serial": ""})
    weeks = sorted({week_of(e["date"]) for e in email_rows if e.get("date")})
    header, done = plan_rows[0], 0
    for w in weeks:
        end = (dt.date.fromisoformat(w) + dt.timedelta(days=7)).isoformat()
        in_w = lambda d: w <= d < end
        wk_plan, cur = [header], None
        for r in plan_rows[1:]:
            d = serial_to_date(r[0] if r else None)
            if d:
                cur = d
            if cur and in_w(cur):
                wk_plan.append(r)
        data = {"plan": wk_plan, "email": [e for e in email_rows if in_w(e["date"])],
                "items": [x for x in items if in_w(x["date"])], "source": source}
        body = json.dumps({"p_token": cfg["admin_token"], "p_week": w, "p_data": data}, ensure_ascii=False, default=str).encode()
        req = urllib.request.Request(cfg["supabase_url"].rstrip("/") + "/rest/v1/rpc/lz_archive", data=body,
                                     headers={"apikey": cfg["key"], "Content-Type": "application/json"})
        try:
            with urllib.request.urlopen(req, timeout=30) as r:
                r.read()
            done += 1
        except Exception as ex:
            log(f"! שמירת שבוע {w} בארכיון נכשלה: {ex}")
    if done:
        log(f"✓ נשמרו בארכיון {done} שבועות")


# ================================================================ Output
def build_page(plan_rows, email_rows, source):
    t = TEMPLATE.read_text(encoding="utf-8")
    dump = lambda o: json.dumps(o, ensure_ascii=False, default=str).replace("</", "<\\/")
    t = t.replace("/*__PLAN__*/[]", dump(plan_rows)).replace("/*__EMAIL__*/[]", dump(email_rows))
    t = t.replace("/*__SOURCE__*/{}", dump(source))
    return t


def main():
    ap = argparse.ArgumentParser(description="לוז התקנות שבועי מ-Outlook")
    ap.add_argument("--plan", help="קובץ תכנית התקנות מקומי (במקום Outlook)")
    ap.add_argument("--email-html", help="קובץ HTML של מייל ברינקס (במקום Outlook)")
    ap.add_argument("--email-image", nargs="*", help="צילום/י מסך של מייל ברינקס (במקום Outlook)")
    ap.add_argument("--auto", action="store_true", help="הרצה ברקע: בלי שאלות, רק התראות")
    ap.add_argument("--force", action="store_true", help="לבנות מחדש גם אם לא היה שינוי")
    ap.add_argument("--no-open", action="store_true", help="לא לפתוח את הדפדפן בסוף")
    ap.add_argument("--no-update", action="store_true", help="לא לבדוק עדכונים ב-GitHub")
    args = ap.parse_args()
    if not args.no_update:
        self_update()

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    workdir = Path(tempfile.mkdtemp(prefix="lz_"))

    manual_email = bool(args.email_html or args.email_image)
    if args.plan or manual_email:
        # העלאה ידנית: מה שנתת כקובץ גובר. מה שלא נתת נמשך מ-Outlook (אם אפשר).
        plan_path = Path(args.plan) if args.plan else None
        brinks = {"html": Path(args.email_html).read_text(encoding="utf-8", errors="ignore") if args.email_html else "",
                  "images": [Path(p) for p in (args.email_image or [])], "received": dt.datetime.now(),
                  "subject": "העלאה ידנית"} if manual_email else None
        if not plan_path or not brinks:
            try:
                o_plan, o_brinks = fetch_from_outlook(workdir)
                plan_path = plan_path or o_plan
                brinks = brinks or o_brinks
                log("✓ החלק שלא הועלה ידנית נמשך מ-Outlook")
            except BaseException as ex:
                log(f"  (לא נמשך מ-Outlook: {ex})")
    else:
        plan_path, brinks = fetch_from_outlook(workdir)

    plan_rows = load_plan_rows(plan_path) if plan_path else [["תאריך", "יום", "שם אתר"]]
    email_rows = []
    if brinks:
        email_rows = rows_from_html(brinks["html"])
        if email_rows:
            log(f"✓ נקראו {len(email_rows)} שורות מהטבלה במייל")
        elif brinks["images"]:
            email_rows = rows_from_images(brinks["images"])
            if email_rows:
                log(f"✓ נקראו {len(email_rows)} שורות מהתמונה")
    email_rows = keep_tasks(email_rows)
    # למה אין סריאלים? נכתב גם בראש הלוז, כדי שלא תצטרך לחפש בלוג.
    email_note = ""
    if not brinks:
        email_note = "לא נמצא מייל סריאלים מברינקס, לכן הכול מסומן 'ממתין לסריאל'."
    elif not email_rows:
        email_note = (f"נמצא המייל '{brinks['subject']}', אבל הטבלה בו היא תמונה ואין מפתח לקריאת תמונות (ANTHROPIC_API_KEY)."
                      if brinks["images"] and not os.environ.get("ANTHROPIC_API_KEY") else
                      f"נמצא המייל '{brinks['subject']}', אבל לא הצלחתי לקרוא ממנו את טבלת הסריאלים.")
    if email_note:
        log("! " + email_note)
        notify("לוז התקנות: אין סריאלים", email_note)

    local = not outlook_available()   # בלי Outlook (למשל בבדיקה): לא מושכים תשובות ולא פותחים טיוטה

    # תשובות חדשות של ברינקס למייל הפערים: מציגים לך כל שינוי ומחילים רק מה שאישרת.
    if not local and plan_path and email_rows:
        try:
            replies = fetch_replies(seen_reply_ids())
            if replies:
                log(f"✉ נמצאו {len(replies)} תשובות חדשות של ברינקס למייל הפערים")
                if args.auto:
                    notify("לוז התקנות: ברינקס ענו", f"{len(replies)} תשובות למייל הפערים מחכות לאישור שלך. הפעל את run.bat.")
                else:
                    review_replies(replies, find_gaps(plan_rows, email_rows))
        except Exception as ex:
            log(f"! בדיקת תשובות ברינקס נכשלה: {ex}")

    # אותה תכנית, אותו מייל ואותם שינויים מאושרים כמו בריצה הקודמת? לא בונים מחדש ולא שואלים שוב.
    import hashlib
    overrides = load_overrides()
    fp = hashlib.sha256((plan_path.read_bytes() if plan_path else b"") + TEMPLATE.read_bytes() + Path(__file__).read_bytes() +
                        json.dumps([email_rows, overrides], ensure_ascii=False, sort_keys=True, default=str).encode()).hexdigest()
    state_file = OUTPUT_DIR / ".last_run.json"
    try:
        last = json.loads(state_file.read_text(encoding="utf-8"))
    except Exception:
        last = {}
    if not args.force and last.get("fp") == fp and Path(last.get("out", "")).exists():
        log(f"= אין שינוי מאז הריצה הקודמת ({last.get('at')}). פותח את הלוז הקיים.")
        if not args.no_open:
            open_schedule(Path(last["out"]))
        return

    now = dt.datetime.now()
    source = {
        "planName": plan_path.name if plan_path else "",
        "emailName": f"{brinks['subject']} ({brinks['received']:%d/%m %H:%M})" if brinks else "",
        "note": email_note,
        "banner": (f"הועלה ידנית ב-{now:%d/%m/%Y %H:%M}." if (args.plan or manual_email) else f"נמשך מ-Outlook ב-{now:%d/%m/%Y %H:%M}.") + (f" כולל {len(overrides)} שינויים שאישרת מתשובות ברינקס." if overrides else ""),
        "overrides": overrides,
    }
    out = OUTPUT_DIR / f"לוז התקנות {now:%Y-%m-%d}.html"
    out.write_text(build_page(plan_rows, email_rows, source), encoding="utf-8")
    log(f"✓ הלוז נשמר: {out}")
    state_file.write_text(json.dumps({"fp": fp, "out": str(out), "at": f"{now:%d/%m %H:%M}"}, ensure_ascii=False), encoding="utf-8")
    if plan_path:   # גם בהעלאה ידנית, וגם בלי סריאלים (הטכנאי רואה לפחות את הלוז): מתעדכן אצל הטכנאי ונשמר בארכיון
        publish_to_viewer(plan_rows, email_rows, {k: v for k, v in source.items() if k != "banner"})
    if not args.no_open:
        open_schedule(out)

    # כספת לאילת באחד הימים: צריך להכין אותה עד יום ראשון של אותו שבוע
    if plan_path:
        today = dt.date.today().isoformat()
        eilat = [p for p in parse_plan(plan_rows) if "אילת" in f"{p['site']} {p.get('address', '')}" and p["date"] >= today]
        for w in sorted({week_of(p["date"]) for p in eilat}):
            ps = [p for p in eilat if week_of(p["date"]) == w]
            sun = dt.date.fromisoformat(w)
            log(f"🔔 כספת לאילת ({len(ps)}): להכין עד יום ראשון {sun.day}/{sun.month}")
            for p in ps:
                log(f"   - {day_label(p['date'])} {p['site']} | {p['task']} | כספת {p['device'] or '—'}")
            if args.auto:   # בהרצה רגילה ההודעה קופצת באתר אחרי הכניסה
                notify(f"כספת לאילת: להכין עד יום ראשון {sun.day}/{sun.month}",
                       ", ".join(f"{p['site']} ({day_label(p['date'])})" for p in ps[:3]))

    if plan_path and email_rows:
        # התקנה/החלפה שבאקסל ולא במייל הסריאלים: נשארות בלוז באותו יום, ומודיעים לך.
        plan, pairs = match_pairs(plan_rows, email_rows)
        matched = {pi for pi, _ in pairs if pi >= 0}
        weeks = {week_of(e["date"]) for e in email_rows if e.get("date")}
        missing = [p for i, p in enumerate(plan) if i not in matched and week_of(p["date"]) in weeks]
        if missing:
            log(f"! {len(missing)} משימות מהאקסל לא מופיעות במייל הסריאלים (נוספו ללוז באותו יום):")
            for p in missing:
                log(f"   - {p['date'][8:]}/{p['date'][5:7]} {p['site']} | {p['task']} | כספת {p['device'] or '—'}")
        gaps = find_gaps(plan_rows, email_rows)
        if not gaps:
            log("✓ אין פערים במספרי כספות או בימים. לא הוכן מייל.")
        else:
            log(f"! נמצאו {len(gaps)} פערים:")
            for g in gaps:
                log(f"   - {g['plan']['site']}" + (" | מספר כספת שונה" if g["dev"] else "") + (" | יום שונה" if g["day"] else "")
                    + (" | חסר במייל הסריאלים" if g.get("missing") else ""))
            names = ", ".join(g["plan"]["site"] for g in gaps[:3]) + ("…" if len(gaps) > 3 else "")
            notify(f"לוז התקנות: {len(gaps)} פערים מול ברינקס", names)
            if args.auto:
                log("  (הרצה ברקע: לאישור מייל הפערים הפעל את run.bat)")
            elif local:
                prev = OUTPUT_DIR / f"מייל פערים {now:%Y-%m-%d}.html"
                prev.write_text("<!doctype html><meta charset='utf-8'>" + gaps_email_html(gaps), encoding="utf-8")
                log(f"  (מצב מקומי: המייל נשמר לתצוגה ב-{prev} ולא נפתח ב-Outlook)")
            else:
                ans = input(f"להכין מייל פערים ל-{GAPS_EMAIL_TO}? הוא ייפתח ב-Outlook לאישור ולא יישלח לבד (כ/ל): ").strip()
                if ans in ("כ", "כן", "y", "yes", "Y"):
                    open_gaps_draft(gaps)
                else:
                    log("לא הוכן מייל.")


if __name__ == "__main__":
    main()
