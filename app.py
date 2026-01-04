import streamlit as st
import openpyxl
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
from io import BytesIO
import pandas as pd
import datetime

# --- 0. ページ設定 (これは必ず一番最初に書くルール) ---
st.set_page_config(page_title="保育指導計画システム", layout="wide")

# --- 1. 定数・データ定義 ---
# 年間計画用
TERMS = ["1期(4-5月)", "2期(6-8月)", "3期(9-12月)", "4期(1-3月)"]
MONTH_RANGES_0Y = [
    "57日～3か月未満", "3か月～6か月未満", "6か月～9か月未満",
    "9か月～12か月未満", "1歳～1歳3か月未満", "1歳3か月～2歳未満"
]

# 月案用定型文 (一部抜粋)
TEIKEI_DATA = {
    "0歳児": {
        "ねらい": ["安心できる保育士との関係の中で心地よく過ごす。", "離乳食を意欲的に食べ、満足感を味わう。", "身の回りのものに興味を持ち、手を伸ばして遊ぶ。"],
        "養護:生命": ["一人一人の生理的欲求を満たし、健康に過ごす。", "室温や湿度に留意し、心地よく眠れるようにする。"],
        "養護:情緒": ["特定の保育士との関わりの中で、甘えたい気持ちを満たす。", "泣く、笑うなどの感情の表出を受け止めてもらう。"],
        "家庭連携": ["家庭での睡眠時間や食事の様子を細かく共有する。", "体調の変化に留意し、早めの連絡をお願いする。"]
    },
    "1歳児": {
        "ねらい": ["保育士に見守られながら、自分でしようとする気持ちを持つ。", "探索活動を十分に楽しむ。", "簡単な言葉のやり取りを喜ぶ。"],
        "教育:健康": ["保育士と一緒に手を洗おうとする。", "戸外で体を十分に動かして遊ぶ。"],
        "家庭連携": ["自分でやりたい気持ちを大切にしてもらうよう伝える。", "靴のサイズ確認をお願いする。"]
    },
    # 必要に応じて他年齢も追加
}
DEFAULT_TEXTS = ["（定型文を選択、または直接入力）", "自分で入力する"]

# --- 2. Excel作成関数群 ---

# A. 年間計画Excel作成
def create_annual_excel(age, config, orientation):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"年間指導計画({age})"
    
    thin = Side(style='thin')
    border = Border(top=thin, bottom=thin, left=thin, right=thin)
    header_fill = PatternFill(start_color="F2F2F2", fill_type="solid")
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    top_left_align = Alignment(horizontal='left', vertical='top', wrap_text=True)

    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE if orientation == "横" else ws.ORIENTATION_PORTRAIT
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1

    ws.column_dimensions['A'].width = 15
    for c in ['B', 'C', 'D', 'E']: ws.column_dimensions[c].width = 25

    # ヘッダー
    ws.merge_cells("A1:C1")
    ws['A1'] = f"年間指導計画 ({age})"
    ws['A1'].font = Font(bold=True, size=16)
    ws.cell(row=1, column=4, value="担任").border = border
    ws.cell(row=1, column=5, value="園長").border = border
    ws.cell(row=2, column=4).border = border
    ws.cell(row=2, column=5).border = border

    # 上段固定項目
    row = 3
    fixed_items = [("年間目標", "年間目標"), ("健康・安全・災害", "健康・安全")]
    if age == "5歳児":
        fixed_items += [("幼児期の終わりまでに育ってほしい姿10項目", "10項目"), ("小学校との連携", "小学校連携")]

    for label, key in fixed_items:
        ws.merge_cells(f"A{row}:A{row+1}")
        ws.cell(row=row, column=1, value=label).fill = header_fill
        ws.cell(row=row, column=1).alignment = center_align
        ws.cell(row=row, column=1).border = border
        ws.cell(row=row+1, column=1).border = border
        
        ws.merge_cells(f"B{row}:E{row+1}")
        c = ws.cell(row=row, column=2, value=config['values'].get(key, ""))
        c.alignment = top_left_align
        c.border = border
        # 結合セルの罫線処理（簡易）
        for r_b in range(row, row+2):
            for c_b in range(2, 6):
                ws.cell(row=r_b, column=c_b).border = border
        row += 2

    # 中段メイン
    ws.cell(row=row, column=1, value="項目 / 期").fill = header_fill
    ws.cell(row=row, column=1).border = border
    for i, t_name in enumerate(TERMS):
        c = ws.cell(row=row, column=i+2, value=t_name)
        c.fill = header_fill
        c.alignment = center_align
        c.border = border
    row += 1

    items = config['mid_items']
    for item in items:
        ws.cell(row=row, column=1, value=item).fill = header_fill
        ws.cell(row=row, column=1).border = border
        ws.cell(row=row, column=1).alignment = center_align
        for i, t_name in enumerate(TERMS):
            c = ws.cell(row=row, column=i+2, value=config['values'].get(f"{item}_{t_name}", ""))
            c.alignment = top_left_align
            c.border = border
        ws.row_dimensions[row].height = 100
        row += 1

    # 下段反省
    ws.cell(row=row, column=1, value="自己評価・反省(期)").fill = header_fill
    ws.cell(row=row, column=1).border = border
    for i, t_name in enumerate(TERMS):
        c = ws.cell(row=row, column=i+2, value=config['values'].get(f"反省_{t_name}", ""))
        c.border = border
        c.alignment = top_left_align
    row += 1

    ws.merge_cells(f"A{row}:E{row}")
    c = ws.cell(row=row, column=1, value="年間を通した自己評価・反省")
    c.fill = header_fill
    c.alignment = center_align
    c.border = border
    for i in range(2, 6): ws.cell(row=row, column=i).border = border
    row += 1
    
    ws.merge_cells(f"A{row}:E{row+1}")
    c = ws.cell(row=row, column=1, value=config['values'].get("年間反省", ""))
    c.alignment = top_left_align
    c.border = border
    for r_b in range(row, row+2):
        for c_b in range(1, 6):
            ws.cell(row=r_b, column=c_b).border = border
    ws.row_dimensions[row].height = 100

    output = BytesIO()
    wb.save(output)
    return output.getvalue()

# B. 月案Excel作成
def create_monthly_excel(age, target_month, config, num_weeks, orientation):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "指導計画表"
    
    thin = Side(style='thin')
    border = Border(top=thin, bottom=thin, left=thin, right=thin)
    header_fill = PatternFill(start_color="F2F2F2", fill_type="solid")
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    top_left_align = Alignment(horizontal='left', vertical='top', wrap_text=True)
    
    total_cols = 1 + num_weeks
    
    # ヘッダー
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols-2 if total_cols>2 else 1)
    ws['A1'] = f"【指導計画】 {target_month} ({age})"
    ws['A1'].font = Font(bold=True, size=14)
    ws.cell(row=1, column=total_cols-1, value="園長").border = border
    ws.cell(row=1, column=total_cols, value="担任").border = border
    
    # 上段
    top_labels = [config['l_top1'], config['l_top2'], config['l_top3']]
    # 簡易配置
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=2)
    ws.cell(row=3, column=1, value=top_labels[0])
    ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=2)
    ws.cell(row=4, column=1, value=config['values'].get(top_labels[0], ""))
    
    mid_end_col = total_cols - 1
    ws.merge_cells(start_row=3, start_column=3, end_row=3, end_column=mid_end_col)
    ws.cell(row=3, column=3, value=top_labels[1])
    ws.merge_cells(start_row=4, start_column=3, end_row=4, end_column=mid_end_col)
    ws.cell(row=4, column=3, value=config['values'].get(top_labels[1], ""))
    
    ws.merge_cells(start_row=3, start_column=mid_end_col+1, end_row=3, end_column=total_cols)
    ws.cell(row=3, column=mid_end_col+1, value=top_labels[2])
    ws.merge_cells(start_row=4, start_column=mid_end_col+1, end_row=4, end_column=total_cols)
    ws.cell(row=4, column=mid_end_col+1, value=config['values'].get(top_labels[2], ""))

    # 中段
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
            
    # 下段
    reflection_row_h = last_row + 1
    reflection_row_c = last_row + 2
    ws.merge_cells(start_row=reflection_row_h, start_column=1, end_row=reflection_row_h, end_column=total_cols)
    ws.cell(row=reflection_row_h, column=1, value="今月の振り返り・反省")
    ws.merge_cells(start_row=reflection_row_c, start_column=1, end_row=reflection_row_c, end_column=total_cols)
    ws.cell(row=reflection_row_c, column=1, value=config['values'].get("reflection", ""))

    # スタイル
    for row in ws.iter_rows(min_row=1, max_row=reflection_row_c, min_col=1, max_col=total_cols):
        for cell in row:
            if not (cell.row == 1 and cell.column >= total_cols - 1):
                cell.border = border
            if cell.row in [3, 5, reflection_row_h] or (cell.column == 1 and 6 <= cell.row <= last_row):
                 cell.alignment = center_align
                 cell.fill = header_fill
            elif cell.row == 1: pass
            else: cell.alignment = top_left_align

    # ページ設定
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToPage = True
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

    ws.row_dimensions[1].height = 30
    ws.row_dimensions[4].height = 60
    for r in range(6, last_row + 1): ws.row_dimensions[r].height = 60
    ws.row_dimensions[reflection_row_c].height = 90
    
    ws.page_margins.left = 0.4; ws.page_margins.right = 0.4
    ws.page_margins.top = 0.4; ws.page_margins.bottom = 0.4

    output = BytesIO()
    wb.save(output)
    return output.getvalue()

# --- 3. メイン処理 ---

# セッション状態の初期化
if 'annual_data' not in st.session_state:
    st.session_state['annual_data'] = {}

st.title("📛 保育指導計画 作成・連動システム")

# サイドバー共通設定
age = st.sidebar.selectbox("対象年齢", ["0歳児", "1歳児", "2歳児", "3歳児", "4歳児", "5歳児"])
mode = st.sidebar.radio("作成する書類", ["年間指導計画", "月間指導計画"])
orient = st.sidebar.radio("用紙向き", ["横", "縦"])

# ==========================================
# モードA：年間指導計画
# ==========================================
if mode == "年間指導計画":
    st.header(f"📅 {age} 年間指導計画")
    
    # 項目設定
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
            st.info("0歳児：月齢別の入力も可能です")
        
        cols = st.columns(4)
        for i, term in enumerate(TERMS):
            with cols[i]:
                st.markdown(f"### {term}")
                for item in mid_item_list:
                    # キーを生成して入力
                    val = st.text_area(f"{item}", key=f"{item}_{term}", height=120)
                    user_values[f"{item}_{term}"] = val
                    
                    # ★ここでセッションに保存（連動用）★
                    if term not in st.session_state['annual_data']:
                        st.session_state['annual_data'][term] = {}
                    st.session_state['annual_data'][term][item] = val

    with t3:
        st.subheader("自己評価・反省")
        cols = st.columns(4)
        for i, term in enumerate(TERMS):
            user_values[f"反省_{term}"] = cols[i].text_area(f"{term}の反省", key=f"rev_{term}")
        user_values["年間反省"] = st.text_area("年間を通した総括", height=150)

    st.divider()
    if st.button("🚀 年間指導計画Excelを作成"):
        config = {'mid_items': mid_item_list, 'values': user_values}
        excel_data = create_annual_excel(age, config, orient)
        st.download_button("📥 ダウンロード", excel_data, f"{age}_年間計画_{orient}.xlsx")
        st.success("作成しました！入力データは月案への連動用に一時保存されました。")

# ==========================================
# モードB：月間指導計画 (連動機能付き)
# ==========================================
elif mode == "月間指導計画":
    st.header(f"📝 {age} 月間指導計画")
    
    # 月案設定
    month_date = st.date_input("対象月", key="monthly_date")
    month_str = month_date.strftime("%Y年%m月")
    target_month_val = month_date.month
    
    weeks_option = st.radio("週の数", ["4週", "5週"], horizontal=True, key="monthly_weeks")
    num_weeks = 5 if weeks_option == "5週" else 4
    
    with st.sidebar.expander("月案項目の編集"):
        l_top = {1: st.text_input("上段1", "前月の振り返り"), 2: st.text_input("上段2", "今月の目標"), 3: st.text_input("上段3", "家庭連携")}
        l_mid = {r: st.text_input(f"中段{r}", val) for r, val in zip(range(6, 16), ["ねらい", "養護:生命", "養護:情緒", "教育:健康", "教育:人間関係", "教育:環境", "教育:言葉", "教育:表現", "環境構成", "小学校連携"])}

    # ★連動ボタン★
    st.info("💡 年間計画を作成済みの場合、以下のボタンで目標を引用できます")
    if st.button("✨ 年間計画から今期の『ねらい』を引用"):
        # 期の判定
        if target_month_val in [4, 5]: current_term = TERMS[0]
        elif target_month_val in [6, 7, 8]: current_term = TERMS[1]
        elif target_month_val in [9, 10, 11, 12]: current_term = TERMS[2]
        else: current_term = TERMS[3]
        
        # データ取得
        if current_term in st.session_state['annual_data']:
            # 年間計画の「ねらい」という項目を探す
            fetched = st.session_state['annual_data'][current_term].get("ねらい", "")
            if fetched:
                st.session_state["target_aim_input"] = fetched
                st.success(f"【{current_term}】のねらいを読み込みました！")
            else:
                st.warning(f"{current_term}のデータはありますが、『ねらい』が空欄です。")
        else:
            st.error(f"まだ{current_term}の年間計画データが保存されていません。年間計画タブで入力してください。")

    # 入力タブ
    tabs = st.tabs([f"第{i}週" for i in range(1, num_weeks + 1)] + ["共通・反省", "👀 プレビュー"])
    
    age_data = TEIKEI_DATA.get(age, {})
    user_values = {}
    
    # 共通項目（連動データ受け入れ）
    with tabs[num_weeks]:
        st.subheader("共通項目")
        c1, c2 = st.columns(2)
        with c1: user_values[l_top[1]] = st.text_area(l_top[1], height=80)
        
        # ここに連動データが入る
        default_aim = st.session_state.get("target_aim_input", "")
        with c2: user_values[l_top[2]] = st.text_area(l_top[2], value=default_aim, height=80, help="年間計画から引用できます")
        
        user_values[l_top[3]] = st.selectbox(f"{l_top[3]} (定型文)", age_data.get("家庭連携", DEFAULT_TEXTS))
        st.divider()
        user_values["reflection"] = st.text_area("今月の振り返り・反省", height=120)

    # 各週入力
    for i in range(num_weeks):
        with tabs[i]:
            st.caption(f"{month_str} 第{i+1}週")
            cols = st.columns(2)
            for idx, (row_num, label) in enumerate(l_mid.items()):
                col = cols[0] if idx < 5 else cols[1]
                user_values[f"{label}_週{i+1}"] = col.selectbox(f"{label}", age_data.get(label, DEFAULT_TEXTS), key=f"w{i+1}_{row_num}")

    # プレビュー
    with tabs[num_weeks + 1]:
        st.subheader("プレビュー")
        df_data = []
        for label in l_mid.values():
            row = {"項目": label}
            for i in range(1, num_weeks + 1): row[f"第{i}週"] = user_values.get(f"{label}_週{i}", "")
            df_data.append(row)
        st.dataframe(pd.DataFrame(df_data), use_container_width=True)

    # Excel生成
    st.sidebar.divider()
    if st.sidebar.button("🚀 月案Excelをダウンロード"):
        config = {
            'l_top1': l_top[1], 'l_top2': l_top[2], 'l_top3': l_top[3],
            **{f'l_mid{r}': val for r, val in l_mid.items()},
            'values': user_values
        }
        excel_data = create_monthly_excel(age, month_str, config, num_weeks, orient)
        st.sidebar.download_button("📥 ファイル保存", excel_data, f"{month_str}_{age}_月案_{orient}.xlsx")# --- (前略：インポート、定数、年間・月間Excel関数はそのまま保持) ---

# C. 週案Excel作成関数
def create_weekly_excel(age, config, orientation):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "週案"
    
    thin = Side(style='thin')
    border = Border(top=thin, bottom=thin, left=thin, right=thin)
    header_fill = PatternFill(start_color="F2F2F2", fill_type="solid")
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    top_left_align = Alignment(horizontal='left', vertical='top', wrap_text=True)

    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE if orientation == "横" else ws.ORIENTATION_PORTRAIT
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1

    # ヘッダー (週のねらい)
    ws.merge_cells("A1:D1")
    ws['A1'] = f"【週案】 {config['week_range']} ({age})"
    ws['A1'].font = Font(bold=True, size=14)
    
    ws.merge_cells("A2:A3")
    ws['A2'] = "週のねらい"
    ws['A2'].fill = header_fill
    ws.merge_cells("B2:D3")
    ws['B2'] = config['values'].get("weekly_aim", "")
    
    # 曜日ヘッダー
    headers = ["曜日・日付", "活動予定", "配慮事項・援助", "準備物"]
    for i, h in enumerate(headers):
        cell = ws.cell(row=4, column=i+1, value=h)
        cell.fill = header_fill
        cell.alignment = center_align
        cell.border = border

    # 曜日データ (月～土)
    days = ["月", "火", "水", "木", "金", "土"]
    row_idx = 5
    for day in days:
        # 曜日・日付
        ws.cell(row=row_idx, column=1, value=f"{day}\n({config['values'].get(f'date_{day}', '')})").border = border
        # 内容
        ws.cell(row=row_idx, column=2, value=config['values'].get(f"activity_{day}", "")).border = border
        ws.cell(row=row_idx, column=3, value=config['values'].get(f"care_{day}", "")).border = border
        ws.cell(row=row_idx, column=4, value=config['values'].get(f"tool_{day}", "")).border = border
        
        ws.row_dimensions[row_idx].height = 80
        row_idx += 1

    # 列幅調整
    ws.column_dimensions['A'].width = 15
    ws.column_dimensions['B'].width = 35
    ws.column_dimensions['C'].width = 35
    ws.column_dimensions['D'].width = 20

    # 全体スタイル
    for r in ws.iter_rows(min_row=1, max_row=row_idx-1, min_col=1, max_col=4):
        for cell in r:
            cell.border = border
            if cell.alignment.horizontal is None:
                cell.alignment = top_left_align if cell.column > 1 else center_align

    output = BytesIO()
    wb.save(output)
    return output.getvalue()

# --- メイン処理のサイドバーメニューに「週案」を追加 ---
# mode = st.sidebar.radio("作成する書類", ["年間指導計画", "月間指導計画", "週案"])

# ==========================================
# モードC：週案
# ==========================================
if mode == "週案":
    st.header(f"📅 {age} 週間指導計画（週案）")
    
    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input("週の開始日（月曜日）", value=datetime.date.today())
    
    # 月案からの連動
    st.info("💡 月案の『第〇週のねらい』を引用できます")
    if st.button("✨ 月案から今週のねらいを引用"):
        # セッションから月案データを探す（簡易実装例）
        # 本来は月案保存時に st.session_state['monthly_data'] に入れる処理が必要です
        st.warning("月案データとの連動機能：月案側で『保存』した内容をここに反映するロジックを次ステップで実装可能です")

    st.divider()
    user_values = {}
    user_values["weekly_aim"] = st.text_area("週のねらい", height=100)
    
    st.subheader("日ごとの計画")
    days = ["月", "火", "水", "木", "金", "土"]
    
    # プレビュー兼入力用の表形式レイアウト
    
    
    for i, day in enumerate(days):
        current_date = start_date + datetime.timedelta(days=i)
        date_str = current_date.strftime("%m/%d")
        user_values[f"date_{day}"] = date_str
        
        with st.expander(f"【{day}】 {date_str} の内容"):
            c1, c2, c3 = st.columns([2, 2, 1])
            user_values[f"activity_{day}"] = c1.text_area("活動予定", key=f"act_{day}")
            user_values[f"care_{day}"] = c2.text_area("配慮事項・援助", key=f"care_{day}")
            user_values[f"tool_{day}"] = c3.text_area("準備物", key=f"tool_{day}")

    if st.button("🚀 週案Excelを作成"):
        config = {
            'week_range': f"{start_date.strftime('%Y/%m/%d')} ～",
            'values': user_values
        }
        excel_data = create_weekly_excel(age, config, orient)
        st.download_button("📥 ダウンロード", excel_data, f"{age}_週案_{date_str}.xlsx")