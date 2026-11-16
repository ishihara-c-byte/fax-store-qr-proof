import html
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

        colors = {k: f"#{v}" for k, v in c.CELL_COLORS.items()}
        matrix = c.build_matrix(targets, findings)
        tab1, tab2 = st.tabs(["🗂 店舗別チェック表", "📋 指摘一覧"])

        with tab1:
            st.caption("✅ 問題なし　❌ 誤りの可能性大　⚠️ 要確認　🔸 空白・ハイフンの違いだけ　／　違いの中身は表の下の枠で確認できます")
            cc1, cc2 = st.columns([3, 1])
            pick = cc1.multiselect(
                "チェックする項目(列)", list(c.MATRIX_COLUMNS), default=list(c.MATRIX_COLUMNS), key=f"cols_{f.name}"
            )
            only_ng = cc2.checkbox("指摘がある店舗だけ", value=False, key=f"ng_{f.name}")
            mrows, sev_map = matrix
            mdf = pd.DataFrame(mrows)
            keep_idx = [
                i for i in range(len(mrows))
                if not only_ng or any(sev_map[(i, col)] in ("高", "中", "低") for col in pick)
            ]
            mview = mdf.loc[keep_idx, ["行", "店舗ID", "店舗名(Excel)", "判定"] + pick]

            def paint(r):
                return [
                    f"background-color: {colors[sev_map[(r.name, k)]]}" if k in c.MATRIX_COLUMNS else ""
                    for k in r.index
                ]

            st.dataframe(mview.style.apply(paint, axis=1), use_container_width=True, hide_index=True, height=600)

            # ---- 違いの詳細(店舗ごとに枠で囲んで、マスタとExcelを並べる) ----
            st.markdown("#### 違いの詳細")
            by_row = {}
            for x in findings:
                col = c.finding_column(x)
                if col in pick and x.severity in ("高", "中", "低"):
                    by_row.setdefault(x.excel_row, []).append(x)
                elif x.item == "店舗ID" and x.severity == "高":
                    by_row.setdefault(x.excel_row, []).append(x)
            if not by_row:
                st.success("選んだ項目に違いはありません 🎉")
            for t in targets:
                xs = by_row.get(t.excel_row)
                if not xs:
                    continue
                worst = min(xs, key=lambda x: c.SEVERITY_ORDER[x.severity])
                with st.container(border=True):
                    st.markdown(f"**{c.MARK[worst.severity]} {t.store_name}**　（店舗ID {t.store_id}／Excel {t.excel_row}行目）")
                    trs = []
                    for x in xs:
                        bg = colors[x.severity]
                        m_ = html.escape(x.master) or '<span style="color:#999">（空欄）</span>'
                        e_ = html.escape(x.draft) or '<span style="color:#999">（空欄）</span>'
                        trs.append(
                            f"<tr><td style='white-space:nowrap'>{c.MARK[x.severity]} {html.escape(x.item)}</td>"
                            f"<td>{m_}</td><td style='background:{bg}'>{e_}</td>"
                            f"<td style='color:#555'>{html.escape(x.comment)}</td></tr>"
                        )
                    st.markdown(
                        "<table style='width:100%;border-collapse:collapse'>"
                        "<tr style='background:#f0f2f6'><th style='width:14%'>項目</th><th style='width:30%'>マスタ（正）</th>"
                        "<th style='width:30%'>Excel（原稿）</th><th>コメント</th></tr>"
                        + "".join(trs) + "</table>",
                        unsafe_allow_html=True,
                    )

        with tab2:
            df = pd.DataFrame([x.as_dict() for x in findings])
            show = st.multiselect(
                "表示する重要度", list(c.SEVERITY_ORDER), default=["高", "中"], key=f"sev_{f.name}"
            )
            view = df[df["重要度"].isin(show)]
            st.dataframe(
                view.style.apply(
                    lambda r: [f"background-color: {colors[r['重要度']]}" if k == "重要度" else "" for k in r.index],
                    axis=1,
                ),
                use_container_width=True,
                hide_index=True,
            )

        st.download_button(
            "結果をExcelでダウンロード(店舗別チェック表+指摘一覧)",
            c.findings_to_xlsx(findings, matrix),
            file_name=f"校正結果_{f.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_{f.name}",
        )
