import streamlit as st
import openpyxl
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
from io import BytesIO
import pandas as pd

# --- 1. 定数定義 ---
TERMS = ["1期(4-5月)", "2期(6-8月)", "3期(9-12月)", "4期(1-3月)"]
MONTH_AGES_0Y = [
    "57日～3か月未満", "3か月～6か月未満", "6か月～9か月未満",
    "9か月～12か月未満", "1歳～1歳3か月未満", "1歳3か月～2歳未満"
]

# --- 2. Excel作成関数（年間計画用） ---
def create_annual_excel(age, config, orientation):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"年間指導計画({age})"
    
    # スタイル
    thin = Side(style='thin')
    border = Border(top=thin, bottom=thin, left=thin, right=thin)
    header_fill = PatternFill(start_color="F2F2F2", fill_type="solid")
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    top_left_align = Alignment(horizontal='left', vertical='top', wrap_text=True)

    # ページ設定
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE if orientation == "横" else ws.ORIENTATION_PORTRAIT
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1

    # 列幅の初期設定
    ws.column_dimensions['A'].width = 15
    for c in ['B', 'C', 'D', 'E']: ws.column_dimensions[c].width = 25

    # --- A. ヘッダー・印鑑欄 ---
    ws.merge_cells("A1:C1")
    ws['A1'] = f"年間指導計画 ({age})"
    ws['A1'].font = Font(bold=True, size=16)
    
    # 印鑑枠
    ws.cell(row=1, column=4, value="担任").border = border
    ws.cell(row=1, column=5, value="園長").border = border
    ws.cell(row=2, column=4).border = border
    ws.cell(row=2, column=5).border = border
    for c in [4,5]: ws.cell(row=1, column=c).alignment = center_align

    # --- B. 上段：共通固定項目 ---
    row = 3
    fixed_items = [("年間目標", "年間目標"), ("健康・安全・災害", "健康・安全")]
    if age == "5歳児":
        fixed_items += [("幼児期の終わりまでに育ってほしい姿10項目", "10項目"), ("小学校との連携", "小学校連携")]

    for label, key in fixed_items:
        ws.merge_cells(f"A{row}:A{row+1}")
        ws.cell(row=row, column=1, value=label).fill = header_fill
        ws.merge_cells(f"B{row}:E{row+1}")
        ws.cell(row=row, column=2, value=config['values'].get(key, ""))
        row += 2

    # --- C. 中段：4期別メインエリア ---
    # 期のヘッダー
    ws.cell(row=row, column=1, value="項目 / 期").fill = header_fill
    for i, t_name in enumerate(TERMS):
        ws.cell(row=row, column=i+2, value=t_name).fill = header_fill
        ws.cell(row=row, column=i+2).alignment = center_align
    row += 1

    # メイン項目
    items = config['mid_items']
    for item in items:
        ws.cell(row=row, column=1, value=item).fill = header_fill
        for i, t_name in enumerate(TERMS):
            ws.cell(row=row, column=i+2, value=config['values'].get(f"{item}_{t_name}", ""))
        ws.row_dimensions[row].height = 100
        row += 1

    # --- D. 下段：反省・評価 ---
    # 期ごとの反省
    ws.cell(row=row, column=1, value="自己評価・反省(期)").fill = header_fill
    for i, t_name in enumerate(TERMS):
        ws.cell(row=row, column=i+2, value=config['values'].get(f"反省_{t_name}", ""))
    row += 1

    # 年間を通した反省（横いっぱい）
    ws.merge_cells(f"A{row}:E{row}")
    ws.cell(row=row, column=1, value="年間を通した自己評価・反省").fill = header_fill
    row += 1
    ws.merge_cells(f"A{row}:E{row+1}")
    ws.cell(row=row, column=1, value=config['values'].get("年間反省", ""))
    ws.row_dimensions[row].height = 100

    # 全体へのスタイル適用
    for r in ws.iter_rows(min_row=1, max_row=row+1, min_col=1, max_col=5):
        for cell in r:
            cell.border = border
            if cell.alignment.horizontal is None:
                cell.alignment = top_left_align if cell.column > 1 else center_align

    output = BytesIO()
    wb.save(output)
    return output.getvalue()

# --- 3. Streamlit UI ---
st.title("📛 保育指導計画 作成・連動システム")

# 共有の年齢選択をサイドバーに
age = st.sidebar.selectbox("対象年齢", ["0歳児", "1歳児", "2歳児", "3歳児", "4歳児", "5歳児"])
mode = st.sidebar.radio("作成する書類", ["年間指導計画", "月間指導計画"])
orient = st.sidebar.radio("用紙向き", ["横", "縦"])

if mode == "年間指導計画":
    st.header(f"📅 {age} 年間指導計画")
    
    # 年齢に応じたデフォルト項目設定
    default_items = "園児の姿\nねらい\n養護（生命・情緒）\n教育（5領域）\n環境構成・援助\n保護者支援\n早朝・延長保育\n行事"
    if age == "0歳児":
        default_items = "月齢別・園児の姿\nねらい\n養護（生命・情緒）\n環境構成・援助\n保護者支援\n行事"

    with st.sidebar.expander("項目のカスタマイズ"):
        custom_items = st.text_area("項目名（改行区切り）", default_items)
        mid_item_list = custom_items.split('\n')

    user_values = {}
    t1, t2, t3 = st.tabs(["📌 基本情報", "📝 各期の計画", "📊 反省・評価"])

    with t1:
        st.subheader("年間を通じた目標")
        user_values["年間目標"] = st.text_area("年間目標", height=100)
        user_values["健康・安全"] = st.text_area("健康・安全・災害対策", height=100)
        if age == "5歳児":
            st.divider()
            user_values["10項目"] = st.text_area("幼児期の終わりまでに育ってほしい姿10項目")
            user_values["小学校連携"] = st.text_area("小学校教育との接続・連携")

    with t2:
        if age == "0歳児":
            st.warning("0歳児は【月齢別】の視点を含めて入力してください")
        
        # 4列レイアウトで期ごとに入力
        cols = st.columns(4)
        for i, term in enumerate(TERMS):
            with cols[i]:
                st.markdown(f"### {term}")
                for item in mid_item_list:
                    user_values[f"{item}_{term}"] = st.text_area(f"{item}", key=f"{item}_{term}", height=120)

    with t3:
        st.subheader("自己評価・反省")
        cols = st.columns(4)
        for i, term in enumerate(TERMS):
            user_values[f"反省_{term}"] = cols[i].text_area(f"{term}の反省", key=f"rev_{term}")
        user_values["年間反省"] = st.text_area("年間を通した総括", height=150)

    # Excel生成
    st.divider()
    if st.button("🚀 年間指導計画Excelを作成"):
        config = {'mid_items': mid_item_list, 'values': user_values}
        excel_data = create_annual_excel(age, config, orient)
        st.download_button("📥 ダウンロード", excel_data, f"{age}_年間計画_{orient}.xlsx")

elif mode == "月間指導計画":
    st.header(f"📝 {age} 月間指導計画")
    st.info("ここに以前の月案コードを統合します。年間計画で入力した『ねらい』等をボタン一つで呼び出せるようになります。")
    # ※ここに前回作成した月案のコードを配置