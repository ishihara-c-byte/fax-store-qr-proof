import re
import urllib.request

import pandas as pd
import streamlit as st

import checker as c

DEFAULT_MASTER_URL = (
    "https://docs.google.com/spreadsheets/d/1hP_nKsdwXvu78tbprwwgwMejyk31vsZmtvZNHTOP_CQ/edit?gid=1864789094"
)

st.set_page_config(page_title="FAXDMエリアシートチェックアプリ", page_icon="📠", layout="wide")


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

**判定**
- 間違い:電話・FAX・QRコードURL・価格がマスタと違う、Excelが空欄、マスタにない店舗ID
- 要確認:マスタに「→〇〇配信～削除」などのメモがある、文言が違う
- 空白・ハイフン違い:中身は同じで空白やハイフンの位置だけ違う(通常は非表示)
"""
        )


# ---------------------------------------------------------------------------
# 本体
# ---------------------------------------------------------------------------
st.title("FAXDMエリアシートチェックアプリ")
st.write("Excelの「店舗情報」シートを、マスタ(FAXDMDB「住所TELFAX・その他」)と照合します。")

files = st.file_uploader("チェックするExcel(複数可)", type=["xlsx"], accept_multiple_files=True)
go = st.button("チェックする", type="primary", disabled=not (files and master))

if go:
    st.session_state["results"] = []
    for f in files:
        try:
            targets, web = c.load_target(f.getvalue())
            st.session_state["results"].append((f.name, targets, c.run_check(targets, master, web), None))
        except Exception as e:  # noqa: BLE001
            st.session_state["results"].append((f.name, [], [], str(e)))

colors = {k: f"#{v}" for k, v in c.CELL_COLORS.items()}

for name, targets, findings, err in st.session_state.get("results", []):
    st.subheader(f"📄 {name}")
    if err:
        st.error(f"Excelを読めませんでした:{err}")
        continue

    mrows, sev_map = c.build_matrix(targets, findings)
    n_bad = sum(1 for r in mrows if r["判定"] != c.OK_MARK)
    if n_bad == 0:
        st.success(f"{len(targets)}店舗すべて、マスタと一致しました 🎉")
    else:
        st.markdown(f"**{len(targets)}店舗中 {n_bad}店舗に違いがあります**")

    st.caption("✅ 問題なし　❌ 間違いの可能性大　⚠️ 要確認　🔸 空白・ハイフンの違いだけ　／　違うマスは マスタ と Excel を上下に並べ、違う文字に色を付けています(青＝マスタ、赤＝Excel)")
    cc1, cc2 = st.columns([3, 1])
    pick = cc1.multiselect(
        "チェックする項目(列)", list(c.MATRIX_COLUMNS), default=list(c.MATRIX_COLUMNS), key=f"cols_{name}"
    )
    only_ng = cc2.checkbox("違いがある店舗だけ", value=False, key=f"ng_{name}")

    st.markdown(c.matrix_html(targets, findings, pick, only_ng), unsafe_allow_html=True)

    with st.expander("📋 指摘を1件ずつ一覧で見る"):
        df = pd.DataFrame([x.as_dict() for x in findings if x.severity != "情報"])
        if df.empty:
            st.write("指摘はありません")
        else:
            st.dataframe(df, use_container_width=True, hide_index=True)

    st.download_button(
        "結果をExcelでダウンロード",
        c.findings_to_xlsx(findings, (mrows, sev_map)),
        file_name=f"チェック結果_{name}",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"dl_{name}",
    )
