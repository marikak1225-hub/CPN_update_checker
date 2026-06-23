import streamlit as st
import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from concurrent.futures import ThreadPoolExecutor, as_completed

st.title("URL更新チェック")

# ✅ KW入力（改行で複数指定※OR条件）
kw_input = st.text_area("特定KW（1行1キーワード!!改行で複数指定※OR条件）")

uploaded_file = st.file_uploader("Excelアップロード", type=["xlsx"])
st.caption("D列に検索用URLが貼ってあれば何でもOKです")

start = st.button("チェック開始")

def check_url(row_idx, url, keywords):
    try:
        if not url:
            return (row_idx, "URLなし")

        url = str(url).strip()
        if not url.startswith("http"):
            url = "https://" + url

        headers = {"User-Agent": "Mozilla/5.0"}

        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()

        soup = BeautifulSoup(response.text, "html.parser")
        page_text = soup.get_text().lower()

        hit_keywords = [kw for kw in keywords if kw.lower() in page_text]

        if hit_keywords:
            return (row_idx, ",".join(hit_keywords))
        else:
            return (row_idx, "該当なし")

    except Exception:
        return (row_idx, "取得失敗")

# ===== 実行 =====
if uploaded_file and kw_input and start:

    # ✅ 改行で分割（←ここが今回のキモ）
    keywords = [
        kw.strip()
        for kw in kw_input.split("\n")
        if kw.strip()
    ]

    st.write(f"入力KW数：{len(keywords)}件")

    wb = load_workbook(uploaded_file)
    ws = wb.active

    max_row = ws.max_row

    progress = st.progress(0)

    futures = []
    results = {}

    max_workers = 10

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        for i in range(2, max_row + 1):
            url = ws[f"G{i}"].value
            futures.append(executor.submit(check_url, i, url, keywords))

        for i, future in enumerate(as_completed(futures)):
            row_idx, result = future.result()
            results[row_idx] = result
            progress.progress((i + 1) / len(futures))

    # ✅ Excelに書き込み（文字列固定）
    for row_idx, result in results.items():
        cell = ws[f"B{row_idx}"]
        cell.value = str(result)

    st.success("チェック完了🚀")

    # ✅ 表示用データ
    display_data = []
    hit_count = 0
    fail_count = 0


    for i in range(2, max_row + 1):
        url = ws[f"G{i}"].value
        result = str(ws[f"B{i}"].value)

        if result == "取得失敗":
            fail_count += 1
        elif result != "該当なし":
            hit_count += 1

        display_data.append({
            "行": i,
            "URL": url,
            "判定結果": result
        })

    st.subheader("チェック結果")
    st.dataframe(display_data)

    st.write(f"✅ 該当あり：{hit_count}件")
    st.write(f"❌ 取得失敗：{fail_count}件")
    st.write(f"📊 全体：{len(display_data)}件")

    # ✅ ダウンロード
    output_file = "result.xlsx"
    wb.save(output_file)

    with open(output_file, "rb") as f:
        st.download_button(
            label="Excelダウンロード",
            data=f,
            file_name=output_file
        )