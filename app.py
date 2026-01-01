import streamlit as st
import openpyxl
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
from io import BytesIO
import datetime
import pandas as pd # プレビュー用

# --- 1. 定型文データの定義 ---
TEIKEI_DATA = {
    "0歳児": {
        "ねらい": ["安心できる保育士との関係の中で心地よく過ごす。", "離乳食を意欲的に食べ、満足感を味わう。", "身の回りのものに興味を持ち、手を伸ばして遊ぶ。"],
        "養護:生命": ["一人一人の生理的欲求を満たし、健康に過ごす。", "室温や湿度に留意し、心地よく眠れるようにする。"],
        "養護:情緒": ["特定の保育士との関わりの中で、甘えたい気持ちを満たす。", "泣く、笑うなどの感情の表出を受け止めてもらう。"],
        "環境構成": ["清潔で安全なハイハイスペースを確保する。", "音の鳴る玩具や感触の違う布を用意する。"],
        "家庭連携": ["家庭での睡眠時間や食事の様子を細かく共有する。", "体調の変化に留意し、早めの連絡をお願いする。"]
    },
    "1歳児": {
        "ねらい": ["保育士に見守られながら、自分でしようとする気持ちを持つ。", "探索活動を十分に楽しむ。", "簡単な言葉のやり取りを喜ぶ。"],
        "教育:健康": ["保育士と一緒に手を洗おうとする。", "戸外で体を十分に動かして遊ぶ。"],
        "教育:人間関係": ["保育士を仲立ちとして、友達に興味を持つ。", "自分の好きな玩具で遊ぶことを楽しむ。"],
        "環境構成": ["自分で玩具を選べるよう、低い棚に配置する。", "安心して探索できる場所を整える。"],
        "家庭連携": ["自分でやりたい気持ちを大切にしてもらうよう伝える。", "靴のサイズ確認をお願いする。"]
    },
    # 必要に応じて他の年齢も追加
}
DEFAULT_TEXTS = ["（定型文を選択、または直接入力）", "自分で入力する"]

# --- 2. Excel作成関数 ---
def create_final_excel(age, target_month, config, num_weeks, orientation):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "指導計画表"
    
    thin = Side(style='thin')
    border = Border(top=thin, bottom=thin, left=thin, right=thin)
    header_fill = PatternFill(start_color="F2F2F2", fill_type="solid")
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    top_left_align = Alignment(horizontal='left', vertical='top', wrap_text=True)
    
    total_cols = 1 + num_weeks
    
    # --- ヘッダー ---
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols-2 if total_cols>2 else 1)
    ws['A1'] = f"【指導計画】 {target_month} ({age})"
    ws['A1'].font = Font(bold=True, size=14)
    ws.cell(row=1, column=total_cols-1, value="園長").border = border
    ws.cell(row=1, column=total_cols, value="担任").border = border
    
    # --- 上段項目 ---
    top_labels = [config['l_top1'], config['l_top2'], config['l_top3']]
    
    # 簡易レイアウト（左・中央・右）
    mid_end_col = total_cols - 1
    # 左
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=2)
    ws.cell(row=3, column=1, value=top_labels[0])
    ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=2)
    ws.cell(row=4, column=1, value=config['values'].get(top_labels[0], ""))
    # 中央
    ws.merge_cells(start_row=3, start_column=3, end_row=3, end_column=mid_end_col)
    ws.cell(row=3, column=3, value=top_labels[1])
    ws.merge_cells(start_row=4, start_column=3, end_row=4, end_column=mid_end_col)
    ws.cell(row=4, column=3, value=config['values'].get(top_labels[1], ""))
    # 右
    ws.merge_cells(start_row=3, start_column=mid_end_col+1, end_row=3, end_column=total_cols)
    ws.cell(row=3, column=mid_end_col+1, value=top_labels[2])
    ws.merge_cells(start_row=4, start_column=mid_end_col+1, end_row=4, end_column=total_cols)
    ws.cell(row=4, column=mid_end_col+1, value=config['values'].get(top_labels[2], ""))

    # --- 中段 (項目 / 週) ---
    ws.cell(row=5, column=1, value="項目 / 週")
    for i in range(1, num_weeks + 1):
        ws.cell(row=5, column=i+1, value=f"第{i}週")

    mid_labels = [config[f'l_mid{r}'] for r in range(6, 16)]
    last_row = 15
    for r_idx, label in enumerate(mid_labels, start=6):
        ws.cell(row=r_idx, column=1, value=label)
        for w_idx in range(1, num_weeks + 1):
            key = f"{label}_週{w_idx}"
            ws.cell(row=r_idx, column=w_idx+1, value=config['values'].get(key, ""))
            
    # --- 下段 (反省) ---
    reflection_row_h = last_row + 1
    reflection_row_c = last_row + 2
    ws.merge_cells(start_row=reflection_row_h, start_column=1, end_row=reflection_row_h, end_column=total_cols)
    ws.cell(row=reflection_row_h, column=1, value="今月の振り返り・反省")
    ws.merge_cells(start_row=reflection_row_c, start_column=1, end_row=reflection_row_c, end_column=total_cols)
    ws.cell(row=reflection_row_c, column=1, value=config['values'].get("reflection", ""))

    # --- スタイル ---
    for row in ws.iter_rows(min_row=1, max_row=reflection_row_c, min_col=1, max_col=total_cols):
        for cell in row:
            if not (cell.row == 1 and cell.column >= total_cols - 1): # ハンコ欄以外
                cell.border = border
            
            if cell.row in [3, 5, reflection_row_h] or (cell.column == 1 and 6 <= cell.row <= last_row):
                 cell.alignment = center_align
                 cell.fill = header_fill
            elif cell.row == 1:
                pass
            else:
                cell.alignment = top_left_align

    # --- ページ設定 ---
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    
    if orientation == "横":
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.column_dimensions['A'].width = 16
        rem_width = 110
    else:
        ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
        ws.column_dimensions['A'].width = 12
        rem_width = 75

    week_col_width = rem_width / num_weeks
    for i in range(1, num_weeks + 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i + 1)].width = week_col_width

    # 高さ調整
    ws.row_dimensions[1].height = 30
    ws.row_dimensions[4].height = 60
    for r in range(6, last_row + 1): ws.row_dimensions[r].height = 60
    ws.row_dimensions[reflection_row_c].height = 90
    
    ws.page_margins.left = 0.4
    ws.page_margins.right = 0.4
    ws.page_margins.top = 0.4
    ws.page_margins.bottom = 0.4

    output = BytesIO()
    wb.save(output)
    return output.getvalue()

# --- 3. Streamlit 画面構成 ---
st.set_page_config(page_title="指導計画プロ", layout="wide")
st.title("🖨️ 指導計画作成＆プレビュー")

with st.sidebar:
    st.header("⚙️ 設定")
    age = st.selectbox("年齢", list(TEIKEI_DATA.keys()) + ["2歳児", "3歳児", "4歳児", "5歳児"])
    month = st.date_input("対象月").strftime("%Y年%m月")
    
    st.divider()
    orientation = st.radio("用紙向き", ["横", "縦"], horizontal=True)
    weeks_option = st.radio("週数", ["4週", "5週"], horizontal=True)
    num_weeks = 5 if weeks_option == "5週" else 4
    
    st.divider()
    with st.expander("項目名の編集"):
        l_top = {1: st.text_input("上段1", "前月の振り返り"), 2: st.text_input("上段2", "今月の目標"), 3: st.text_input("上段3", "家庭連携")}
        l_mid = {r: st.text_input(f"中段{r}", val) for r, val in zip(range(6, 16), ["ねらい", "養護:生命", "養護:情緒", "教育:健康", "教育:人間関係", "教育:環境", "教育:言葉", "教育:表現", "環境構成", "小学校連携"])}

# タブ設定（プレビュータブを追加）
tab_labels = [f"第{i}週" for i in range(1, num_weeks + 1)] + ["共通・反省", "👀 全体プレビュー"]
tabs = st.tabs(tab_labels)

age_data = TEIKEI_DATA.get(age, {})
user_values = {}

# --- 入力画面 ---
# 各週
for i in range(num_weeks):
    with tabs[i]:
        st.caption(f"{month} 第{i+1}週の内容を入力")
        cols = st.columns(2)
        for idx, (row_num, label) in enumerate(l_mid.items()):
            col = cols[0] if idx < 5 else cols[1]
            user_values[f"{label}_週{i+1}"] = col.selectbox(
                f"{label}", age_data.get(label, DEFAULT_TEXTS), key=f"w{i+1}_{row_num}"
            )

# 共通項目
with tabs[num_weeks]: # 共通・反省タブ
    st.subheader("共通項目")
    c1, c2 = st.columns(2)
    with c1: user_values[l_top[1]] = st.text_area(l_top[1], height=80)
    with c2: user_values[l_top[2]] = st.text_area(l_top[2], height=80)
    user_values[l_top[3]] = st.selectbox(f"{l_top[3]} (定型文)", age_data.get("家庭連携", DEFAULT_TEXTS))
    
    st.divider()
    st.subheader("今月の振り返り・反省")
    user_values["reflection"] = st.text_area("反省・特記事項", height=120)

# --- プレビュー画面 (NEW!) ---
with tabs[num_weeks + 1]: # 最後のタブ
    st.subheader(f"📄 {month} {age} 指導計画プレビュー")
    st.info("※ ここで全体のバランスを確認できます（実際のExcelレイアウトとは多少異なります）")
    
    # 1. 上段項目の表示
    st.markdown(f"**【{l_top[1]}】** {user_values.get(l_top[1], '')}")
    st.markdown(f"**【{l_top[2]}】** {user_values.get(l_top[2], '')}")
    st.markdown(f"**【{l_top[3]}】** {user_values.get(l_top[3], '')}")
    
    st.divider()
    
    # 2. 中段項目の表表示 (Pandasを使用)
    preview_data = []
    for label in l_mid.values():
        row = {"項目": label}
        for i in range(1, num_weeks + 1):
            row[f"第{i}週"] = user_values.get(f"{label}_週{i}", "")
        preview_data.append(row)
    
    df = pd.DataFrame(preview_data)
    st.dataframe(df, hide_index=True, use_container_width=True)
    
    st.divider()
    
    # 3. 反省欄
    st.markdown(f"**【今月の振り返り・反省】**")
    st.warning(user_values.get("reflection", "（未入力）"))

# --- 生成ボタン ---
config = {
    'l_top1': l_top[1], 'l_top2': l_top[2], 'l_top3': l_top[3],
    **{f'l_mid{r}': val for r, val in l_mid.items()},
    'values': user_values
}

st.sidebar.divider()
if st.sidebar.button("🚀 Excelをダウンロード"):
    excel_data = create_final_excel(age, month, config, num_weeks, orientation)
    st.sidebar.download_button(
        label="📥 ファイル保存", 
        data=excel_data, 
        file_name=f"{month}_{age}_計画表({orientation}).xlsx"
    )