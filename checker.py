"""FAX注文用紙 店舗情報・QRコード 校正ロジック

原稿Excel(「店舗情報」シート)を、マスタ(FAXDMDB「住所TELFAX・その他」シート)と照合する。
Streamlit に依存しないので、単体でテストできる。
"""
from __future__ import annotations

import csv
import io
import re
import unicodedata
from dataclasses import dataclass, field

import openpyxl

# ---------------------------------------------------------------------------
# 設定
# ---------------------------------------------------------------------------
TARGET_SHEET = "店舗情報"
WEB_SHEET = "Web申込"
MASTER_SHEET = "住所TELFAX・その他"

# 照合する項目: (表示名, 原稿の見出し, マスタの見出し, 種類)
FIELDS = [
    ("正式店舗名", "正式店舗名", "正式店舗名", "name"),
    ("電話番号", "記載用電話番号", "記載用電話番号", "phone"),
    ("FAX番号", "記載用FAX番号", "記載用FAX番号", "phone"),
    ("住所", "記載住所", "記載住所", "text"),
    ("住所２", "住所２", "住所２", "text"),
    ("DX価格", "DX価格", "DX価格", "price"),
    ("HC価格", "HC価格", "HC価格", "price"),
    ("味噌汁（文言）", "味噌汁（文言）", "味噌汁（文言）", "note"),
    ("受付時間（文言）", "受付時間（文言）", "受付時間（文言）", "note"),
    ("QRコードURL", "QRコード用URL", "QRコード用URL", "url"),
]

# マスタで DX/HC 価格の見出しが空欄のとき、「住所２」列から何列右にあるか
MASTER_PRICE_OFFSET = {"DX価格": 3, "HC価格": 4}

# マスタのセルに書き込まれがちな変更メモの目印
NOTE_MARKERS = ("配信", "削除", "→", "⇒", "月号", "変更", "記載しない", "※")

SEVERITY_ORDER = {"高": 0, "中": 1, "低": 2, "情報": 3}


# ---------------------------------------------------------------------------
# 文字の正規化
# ---------------------------------------------------------------------------
_DASHES = "‐‑‒–—―−ーｰ－"


def norm(v) -> str:
    """全角半角・ハイフン・空白のゆれをそろえる"""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = unicodedata.normalize("NFKC", str(v))
    s = re.sub(r"\s+", " ", s).strip()
    return s


def show(v) -> str:
    """画面表示用(原文のまま。数値の .0 と前後の空白だけ除く)"""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def norm_phone(v) -> str:
    s = norm(v)
    for d in _DASHES:
        s = s.replace(d, "-")
    return s


def nospace(s: str) -> str:
    return re.sub(r"\s+", "", s)


def digits(s: str) -> str:
    return re.sub(r"\D", "", s)


# ---------------------------------------------------------------------------
# 読み込み
# ---------------------------------------------------------------------------
def rows_from_csv_text(text: str) -> list[list[str]]:
    return [list(r) for r in csv.reader(io.StringIO(text))]


def rows_from_xlsx(data: bytes, sheet: str | None) -> list[list]:
    wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    ws = wb[sheet] if sheet and sheet in wb.sheetnames else wb.worksheets[0]
    return [list(r) for r in ws.iter_rows(values_only=True)]


def find_header_row(rows: list[list], key: str = "正式店舗名") -> int:
    for i, r in enumerate(rows[:30]):
        if any(norm(c) == key for c in r):
            return i
    raise ValueError(f"見出し「{key}」の行が見つかりません")


def header_index(header: list, name: str) -> int | None:
    for i, c in enumerate(header):
        if norm(c) == norm(name):
            return i
    return None


def cell(row: list, idx: int | None):
    if idx is None or idx >= len(row):
        return None
    return row[idx]


@dataclass
class Master:
    by_id: dict[str, list[dict]] = field(default_factory=dict)
    default_uketsuke: str = ""
    row_count: int = 0


def load_master(rows: list[list]) -> Master:
    h = find_header_row(rows)
    header = rows[h]

    idx = {}
    for _, _, mname, _ in FIELDS:
        idx[mname] = header_index(header, mname)
    # 店舗ID: 見出しがあればそれ、なければ「店舗名」の1つ左
    id_col = header_index(header, "店舗ID")
    if id_col is None:
        name_col = header_index(header, "店舗名")
        if name_col is None or name_col == 0:
            raise ValueError("マスタの店舗ID列が特定できません")
        id_col = name_col - 1
    # 価格: 見出しが空欄なら住所２からの位置で決める
    addr2 = idx.get("住所２")
    for k, off in MASTER_PRICE_OFFSET.items():
        if idx.get(k) is None and addr2 is not None:
            idx[k] = addr2 + off

    # 受付時間の既定文言(見出しの1行上)
    default_uketsuke = ""
    u = idx.get("受付時間（文言）")
    if h > 0 and u is not None:
        default_uketsuke = norm(cell(rows[h - 1], u))

    m = Master(default_uketsuke=default_uketsuke)
    for r in rows[h + 1:]:
        sid = norm(cell(r, id_col))
        if not sid:
            continue
        rec = {k: cell(r, i) for k, i in idx.items()}
        m.by_id.setdefault(sid, []).append(rec)
        m.row_count += 1
    return m


@dataclass
class TargetRow:
    excel_row: int
    store_id: str
    store_name: str
    values: dict


def load_target(data: bytes) -> tuple[list[TargetRow], dict[str, str]]:
    """原稿Excelを読む。戻り値: (店舗情報の行, Web申込の {店舗ID: URL})"""
    wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    if TARGET_SHEET not in wb.sheetnames:
        raise ValueError(f"原稿に「{TARGET_SHEET}」シートがありません(シート: {', '.join(wb.sheetnames)})")
    rows = [list(r) for r in wb[TARGET_SHEET].iter_rows(values_only=True)]
    h = find_header_row(rows)
    header = rows[h]
    id_col = header_index(header, "店舗ID")
    name_col = header_index(header, "店舗名")
    if id_col is None:
        raise ValueError("原稿の「店舗ID」列が見つかりません")
    cols = {tname: header_index(header, tname) for _, tname, _, _ in FIELDS}

    out = []
    for i, r in enumerate(rows[h + 1:], start=h + 2):
        sid = norm(cell(r, id_col))
        if not sid:
            continue
        out.append(TargetRow(
            excel_row=i,
            store_id=sid,
            store_name=norm(cell(r, name_col)),
            values={t: cell(r, c) for t, c in cols.items()},
        ))

    web = {}
    if WEB_SHEET in wb.sheetnames:
        wrows = [list(r) for r in wb[WEB_SHEET].iter_rows(values_only=True)]
        for wh, r in enumerate(wrows[:10]):
            if header_index(r, "店舗ID") is not None and header_index(r, "URL") is not None:
                wi, wu = header_index(r, "店舗ID"), header_index(r, "URL")
                for rr in wrows[wh + 1:]:
                    sid = norm(cell(rr, wi))
                    if sid:
                        web[sid] = norm(cell(rr, wu))
                break
    return out, web


# ---------------------------------------------------------------------------
# 照合
# ---------------------------------------------------------------------------
@dataclass
class Finding:
    severity: str
    excel_row: int
    store_id: str
    store_name: str
    item: str
    draft: str
    master: str
    comment: str

    def as_dict(self):
        return {
            "重要度": self.severity,
            "行": self.excel_row,
            "店舗ID": self.store_id,
            "店舗名": self.store_name,
            "項目": self.item,
            "マスタ(正)": self.master,
            "Excel(原稿)": self.draft,
            "コメント": self.comment,
        }


def has_note(s: str) -> bool:
    return any(k in s for k in NOTE_MARKERS)


def pick_master(cands: list[dict], t: TargetRow) -> tuple[dict, str]:
    if len(cands) == 1:
        return cands[0], ""
    tn = nospace(norm(t.values.get("正式店舗名")))
    for c in cands:
        if nospace(norm(c.get("正式店舗名"))) == tn:
            return c, f"マスタに同じ店舗IDが{len(cands)}行あり、店舗名が一致する行と照合"
    return cands[0], f"マスタに同じ店舗IDが{len(cands)}行あり、1行目と照合"


def check_url(u: str) -> str | None:
    if not u:
        return None
    if not re.match(r"^https?://[\x21-\x7e]+$", u):
        return "URLの形式が不正(空白・全角文字・http(s)以外が含まれる)"
    return None


def compare_field(kind: str, label: str, d_raw, m_raw, default: str = ""):
    """1項目を比べて (重要度, コメント) を返す。問題なしは None"""
    if kind == "phone":
        d, m = norm_phone(d_raw), norm_phone(m_raw)
    else:
        d, m = norm(d_raw), norm(m_raw)
    if kind == "note" and not m and default:
        m = default

    if d == m:
        return None

    if kind == "url":
        if not d:
            return ("高", "ExcelのQRコードURLが空欄") if m else None
        return ("高", "QRコードURLがマスタと違う")

    if kind == "phone":
        if not d and m:
            return ("高", f"Excelの{label}が空欄")
        if digits(d) and digits(d) == digits(m):
            return ("低", "数字は同じ。ハイフン・空白の位置だけ違う")
        if has_note(m) or not digits(m):
            return ("中", "マスタに文字メモあり。内容を確認")
        return ("高", f"{label}がマスタと違う")

    if kind == "price":
        if not m:
            return ("中", "マスタの価格が空欄。Excelの価格で正しいか確認")
        if not d:
            return ("高", f"Excelの{label}が空欄")
        return ("高", f"{label}がマスタと違う")

    if nospace(d) == nospace(m):
        return ("低", "空白の有無・位置だけ違う")

    if kind == "note":
        if not d and "削除" in m:
            return ("情報", "マスタに削除の指示あり。空欄で正しいか念のため確認")
        if has_note(m):
            return ("中", "マスタに変更メモあり。最新の指示に沿っているか確認")
        if not d:
            return ("中", "Excelが空欄、マスタには文言あり")
        return ("中", "マスタとExcelで文言が違う")

    # name / text
    if not d and m:
        return ("高", f"Excelの{label}が空欄")
    if d and nospace(d) in nospace(m) and has_note(m):
        return ("中", "マスタに変更メモあり。最新の指示に沿っているか確認")
    if has_note(m):
        return ("中", "マスタに変更メモあり。内容を確認")
    return ("高", f"{label}がマスタと違う")


def run_check(targets: list[TargetRow], master: Master, web: dict[str, str]) -> list[Finding]:
    res: list[Finding] = []

    for t in targets:
        def add(sev, item, d, m, comment):
            res.append(Finding(sev, t.excel_row, t.store_id, t.store_name, item, show(d), show(m), comment))

        cands = master.by_id.get(t.store_id)
        if not cands:
            add("高", "店舗ID", t.store_id, "", "マスタにこの店舗IDがない")
            continue
        mrec, dup_note = pick_master(cands, t)
        if dup_note:
            add("情報", "店舗ID", t.store_id, t.store_id, dup_note)

        for label, tname, mname, kind in FIELDS:
            d_raw, m_raw = t.values.get(tname), mrec.get(mname)
            r = compare_field(kind, label, d_raw, m_raw,
                              master.default_uketsuke if mname == "受付時間（文言）" else "")
            if r:
                m_show = show(m_raw) or (master.default_uketsuke if mname == "受付時間（文言）" else "")
                add(r[0], label, d_raw, m_show, r[1])

        # QRコードURL 単体チェック
        url = norm(t.values.get("QRコード用URL"))
        bad = check_url(url)
        if bad:
            add("高", "QRコードURL", url, norm(mrec.get("QRコード用URL")), bad)

        # Web申込のURLと突合
        if web:
            if t.store_id not in web:
                add("低", "QRコードURL(Web申込)", url, "", "Web申込シートにこの店舗IDがない")
            elif web[t.store_id] and web[t.store_id] != url:
                add("高", "QRコードURL(Web申込)", url, web[t.store_id], "Web申込シートのURLと違う")

    res.sort(key=lambda f: (SEVERITY_ORDER[f.severity], f.excel_row))
    return res


# ---------------------------------------------------------------------------
# 店舗別チェック表(1店舗1行、チェック項目を列に)
# ---------------------------------------------------------------------------
# 表の列: 列名 -> その列にまとめる指摘の「項目」
MATRIX_COLUMNS = {
    "正式店舗名": ["正式店舗名"],
    "TEL": ["電話番号"],
    "FAX": ["FAX番号"],
    "住所": ["住所", "住所２"],
    "価格": ["DX価格", "HC価格"],
    "味噌汁文言": ["味噌汁（文言）"],
    "受付時間": ["受付時間（文言）"],
    "QRコード": ["QRコードURL", "QRコードURL(Web申込)"],
}
MARK = {"高": "❌", "中": "⚠️", "低": "🔸", "情報": "ℹ️"}
OK_MARK = "✅"


def _short(s: str, n: int = 18) -> str:
    s = s.replace("\n", " ")
    return s if len(s) <= n else s[:n] + "…"


def finding_column(f: Finding) -> str | None:
    for col, items in MATRIX_COLUMNS.items():
        if f.item in items:
            return col
    return None


def build_matrix(targets: list[TargetRow], findings: list[Finding]):
    """戻り値: (行のリスト, 各セルの重要度 {(行番号, 列名): 重要度})"""
    item_to_col = {it: col for col, items in MATRIX_COLUMNS.items() for it in items}
    by_row: dict[int, list[Finding]] = {}
    for f in findings:
        by_row.setdefault(f.excel_row, []).append(f)

    rows, sev_map = [], {}
    for i, t in enumerate(targets):
        fs = by_row.get(t.excel_row, [])
        r = {"行": t.excel_row, "店舗ID": t.store_id, "店舗名(Excel)": t.store_name}
        missing = any(f.item == "店舗ID" and f.severity == "高" for f in fs)
        for col in MATRIX_COLUMNS:
            if missing:
                r[col] = "－ マスタになし"
                sev_map[(i, col)] = "高"
                continue
            hits = [f for f in fs if item_to_col.get(f.item) == col and f.severity != "情報"]
            info = [f for f in fs if item_to_col.get(f.item) == col and f.severity == "情報"]
            if not hits:
                r[col] = OK_MARK
                sev_map[(i, col)] = "OK"
                continue
            worst = min(hits, key=lambda f: SEVERITY_ORDER[f.severity])
            r[col] = MARK[worst.severity]
            sev_map[(i, col)] = worst.severity
        n_bad = sum(1 for col in MATRIX_COLUMNS if sev_map[(i, col)] in ("高", "中", "低"))
        r["判定"] = OK_MARK if n_bad == 0 else f"{MARK[min((sev_map[(i, c)] for c in MATRIX_COLUMNS if sev_map[(i, c)] in SEVERITY_ORDER and sev_map[(i, c)] != '情報'), key=SEVERITY_ORDER.get)]} {n_bad}件"
        rows.append(r)
    return rows, sev_map


# ---------------------------------------------------------------------------
# 結果のExcel出力
# ---------------------------------------------------------------------------
CELL_COLORS = {"高": "F8D7DA", "中": "FFF3CD", "低": "E2E3E5", "情報": "D1ECF1", "OK": "D4EDDA"}


def findings_to_xlsx(findings: list[Finding], matrix=None) -> bytes:
    from openpyxl.styles import Alignment, Font, PatternFill

    wb = openpyxl.Workbook()
    ws = wb.active

    if matrix is not None:
        mrows, sev_map = matrix
        ws.title = "店舗別チェック表"
        head = ["行", "店舗ID", "店舗名(Excel)", "判定"] + list(MATRIX_COLUMNS)
        ws.append(head)
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="444444")
        for i, r in enumerate(mrows):
            ws.append([r[k] for k in head])
            for j, col in enumerate(MATRIX_COLUMNS, start=5):
                ws.cell(ws.max_row, j).fill = PatternFill("solid", fgColor=CELL_COLORS[sev_map[(i, col)]])
        for i, w in enumerate([6, 8, 24, 9] + [22] * len(MATRIX_COLUMNS), start=1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
        for row in ws.iter_rows(min_row=2):
            for c in row:
                c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.freeze_panes = "E2"
        ws = wb.create_sheet()

    ws.title = "指摘一覧"
    cols = ["重要度", "行", "店舗ID", "店舗名", "項目", "マスタ(正)", "Excel(原稿)", "コメント"]
    ws.append(cols)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="444444")
    fills = {"高": "F8D7DA", "中": "FFF3CD", "低": "E2E3E5", "情報": "D1ECF1"}
    for f in findings:
        d = f.as_dict()
        ws.append([d[k] for k in cols])
        ws.cell(ws.max_row, 1).fill = PatternFill("solid", fgColor=fills[f.severity])
    widths = [7, 6, 8, 26, 16, 40, 40, 40]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A2"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
