import re
import urllib.request

import pandas as pd
import streamlit as st

import checker as c

DEFAULT_MASTER_URL = (
    "https://docs.google.com/spreadsheets/d/1hP_nKsdwXvu78tbprwwgwMejyk31vsZmtvZNHTOP_CQ/edit?gid=1864789094"
)

st.set_page_config(page_title="FAX注文用紙 校正ツール", page_icon="📠", layout="wide")


# ---------------------------------------------------------------------------
# マスタ取得
# ---------------------------------------------------------------------------
def csv_export_url(url: str) -> str:
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", url)
    if not m:
        raise ValueError("スプレッドシートのURLではありません")
    g = re.search(r"[#&?]gid=(\d+)", url)
    gid = g.group(1) if g else "0"
    return f"https://docs.google.com/spreadsheets/d/{m.group(1)}/export?format=csv&gid={gid}"


@st.cache_data(ttl=600, show_spinner="マスタを読み込み中…")
def fetch_master_rows(url: str):
    req = urllib.request.Request(csv_export_url(url), headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=30) as r:
        body = r.read()
    text = body.decode("utf-8-sig")
    if text.lstrip().startswith("<"):
        raise PermissionError(
            "スプレッドシートを読めませんでした。共有設定を「リンクを知っている全員」が「閲覧者」にしてください。"
        )
    return c.rows_from_csv_text(text)


with st.sidebar:
    st.header("マスタ")
    how = st.radio("取得方法", ["スプレッドシートから読み込む", "ファイルをアップロード"])
    master_rows = None
    master_err = None
    if how == "スプレッドシートから読み込む":
        url = st.text_input("マスタのURL(住所TELFAX・その他シート)", DEFAULT_MASTER_URL)
        st.caption("※ シートの共有設定は「リンクを知っている全員」が「閲覧者」にしてください。URL末尾の gid でシートを指定します。")
        if st.button("🔄 マスタを今すぐ取り直す"):
            fetch_master_rows.clear()
        try:
            master_rows = fetch_master_rows(url)
        except Exception as e:  # noqa: BLE001
            master_err = str(e)
    else:
        up = st.file_uploader("マスタ(Excel / CSV)", type=["xlsx", "csv"])
        if up:
            try:
                if up.name.lower().endswith(".csv"):
                    master_rows = c.rows_from_csv_text(up.getvalue().decode("utf-8-sig"))
                else:
                    master_rows = c.rows_from_xlsx(up.getvalue(), c.MASTER_SHEET)
            except Exception as e:  # noqa: BLE001
                master_err = str(e)

    master = None
    if master_rows is not None:
        try:
            master = c.load_master(master_rows)
            st.success(f"マスタ読み込み済み:{len(master.by_id)}店舗")
        except Exception as e:  # noqa: BLE001
            master_err = str(e)
    if master_err:
        st.error(master_err)

    with st.expander("判定ルール"):
        st.markdown(
            """
**照合する項目**:正式店舗名/電話番号/FAX番号/住所/住所２/DX・HC価格/味噌汁(文言)/受付時間(文言)/QRコードURL

**QRコードURL**:マスタと一致するか、URLの形式が正しいか、原稿内の「Web申込」シートのURLと一致するか

**受付時間**:マスタが空欄なら「前日21時までにご注文ください」として照合

**重要度**
- 高=誤りの可能性大(電話・FAX・URL・価格の不一致、原稿の空欄、マスタにない店舗ID)
- 中=要確認(マスタに「→〇〇配信～削除」などの変更メモがある、文言違い)
- 低=空白やハイフンの位置だけ違う
- 情報=参考(マスタで削除指示済み、同じ店舗IDが複数行 など)
"""
        )


# ---------------------------------------------------------------------------
# 本体
# ---------------------------------------------------------------------------
st.title("📠 FAX注文用紙 校正ツール")
st.write("原稿Excelの「店舗情報」シートを、マスタ(FAXDMDB「住所TELFAX・その他」)と照合します。")

files = st.file_uploader("原稿Excel(複数可)", type=["xlsx"], accept_multiple_files=True)
go = st.button("校正する", type="primary", disabled=not (files and master))

if go:
    for f in files:
        st.subheader(f"📄 {f.name}")
        try:
            targets, web = c.load_target(f.getvalue())
        except Exception as e:  # noqa: BLE001
            st.error(f"原稿を読めませんでした:{e}")
            continue

        findings = c.run_check(targets, master, web)
        counts = {s: sum(1 for x in findings if x.severity == s) for s in c.SEVERITY_ORDER}

        cols = st.columns(5)
        cols[0].metric("店舗数", len(targets))
        for col, s in zip(cols[1:], c.SEVERITY_ORDER):
            col.metric(s, counts[s])

        if not findings:
            st.success("指摘はありません 🎉")
            continue

        df = pd.DataFrame([x.as_dict() for x in findings])
        show = st.multiselect(
            "表示する重要度", list(c.SEVERITY_ORDER), default=["高", "中"], key=f"sev_{f.name}"
        )
        view = df[df["重要度"].isin(show)]

        colors = {"高": "#f8d7da", "中": "#fff3cd", "低": "#e2e3e5", "情報": "#d1ecf1"}
        st.dataframe(
            view.style.apply(
                lambda r: [f"background-color: {colors[r['重要度']]}" if k == "重要度" else "" for k in r.index],
                axis=1,
            ),
            use_container_width=True,
            hide_index=True,
        )
        st.download_button(
            "結果をExcelでダウンロード",
            c.findings_to_xlsx(findings),
            file_name=f"校正結果_{f.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_{f.name}",
        )
