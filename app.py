import streamlit as st
import openpyxl
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
from io import BytesIO
import pandas as pd
import datetime
import json
from streamlit_gsheets import GSheetsConnection
import google.generativeai as genai

# SecretsからAPIキーを読み込む（設定されていない場合のエラー回避付き）
if "GEMINI_API_KEY" in st.secrets:
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
    has_api_key = True
else:
    has_api_key = False
# --- 0. ページ設定 ---
st.set_page_config(page_title="保育指導計画システム", layout="wide", page_icon="📛")

# --- 1. 定数・データ定義 ---
TERMS = ["1期(4-5月)", "2期(6-8月)", "3期(9-12月)", "4期(1-3月)"]

# 定型文データ
TEIKEI_DATA = {
    "0歳児": {
        "ねらい": ["安心できる保育士との関係の中で心地よく過ごす。", "離乳食を意欲的に食べ、満足感を味わう。"],
        "養護:生命": ["一人一人の生理的欲求を満たし、健康に過ごす。", "室温や湿度に留意し、心地よく眠れるようにする。"],
        "家庭連携": ["家庭での睡眠時間や食事の様子を細かく共有する。", "体調の変化に留意し、早めの連絡をお願いする。"]
    },
    "1歳児": {
        "ねらい": ["保育士に見守られながら、自分でしようとする気持ちを持つ。", "探索活動を十分に楽しむ。"],
        "教育:健康": ["保育士と一緒に手を洗おうとする。", "戸外で体を十分に動かして遊ぶ。"],
        "家庭連携": ["自分でやりたい気持ちを大切にしてもらうよう伝える。", "靴のサイズ確認をお願いする。"]
    }
}
DEFAULT_TEXTS = ["（定型文を選択、または直接入力）", "自分で入力する"]

# --- 2. データベース操作関数 (保存・読込) ---

def load_data_from_sheet(user_id, doc_type):
    """スプレッドシートからデータを読み込み、セッションステートに反映する"""
    conn = st.connection("gsheets", type=GSheetsConnection)
    try:
        df = conn.read(ttl=0)
        # ユーザーIDと書類タイプで検索
        user_df = df[(df["user_id"] == user_id) & (df["doc_type"] == doc_type)]
        
        if not user_df.empty:
            # 最新のデータを取得
            latest_row = user_df.iloc[-1]
            json_str = latest_row["data_json"]
            data_dict = json.loads(json_str)
            
            # セッションステートに書き戻す
            for key, value in data_dict.items():
                # 日付型などの復元が必要な場合はここで処理可能だが、今回は文字列として戻す
                st.session_state[key] = value
            return True
        else:
            return False
    except Exception as e:
        st.error(f"読み込みエラー: {e}")
        return False

def save_data_to_sheet(user_id, doc_type):
    """現在のセッションステート（入力内容）をJSONにして保存する"""
    conn = st.connection("gsheets", type=GSheetsConnection)
    try:
        df = conn.read(ttl=0)
        
        # 保存対象のキーのみを抽出（ウィジェットのキーなど）
        save_dict = {}
        for key in st.session_state:
            # Streamlitの内部キーなどを除外して保存
            if isinstance(st.session_state[key], (str, int, float, bool, list)):
                save_dict[key] = st.session_state[key]
            # 日付型はJSONにできないので文字列変換
            elif isinstance(st.session_state[key], (datetime.date, datetime.datetime)):
                save_dict[key] = st.session_state[key].strftime("%Y-%m-%d")

        json_str = json.dumps(save_dict, ensure_ascii=False)
        now_str = datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S")
        
        # 新しい行を作成
        new_row = pd.DataFrame([{
            "user_id": user_id,
            "doc_type": doc_type,
            "updated_at": now_str,
            "data_json": json_str
        }])
        
        # 既存データがあれば、そのユーザー・タイプの古いデータを削除して上書きするロジックも可能だが、
        # ここではシンプルに「追記」して、読み込み時に「最新」を取る方式にする
        # (スプレッドシートが重くなる場合は、定期的に削除が必要)
        updated_df = pd.concat([df, new_row], ignore_index=True)
        conn.update(data=updated_df)
        return True
    except Exception as e:
        st.error(f"保存エラー: {e}")
        return False

# --- 3. Excel作成関数群 (前と同じなので省略せず記述) ---

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
    
    # (レイアウト詳細は省略せず実装)
    ws.merge_cells("A1:C1")
    ws['A1'] = f"年間指導計画 ({age})"
    ws['A1'].font = Font(bold=True, size=16)
    
    row = 3
    fixed_items = [("年間目標", "年間目標"), ("健康・安全", "健康・安全")]
    for label, key in fixed_items:
        ws.merge_cells(f"A{row}:A{row+1}")
        ws.cell(row=row, column=1, value=label).fill = header_fill
        ws.cell(row=row, column=1).border = border
        ws.cell(row=row+1, column=1).border = border
        ws.merge_cells(f"B{row}:E{row+1}")
        c = ws.cell(row=row, column=2, value=config['values'].get(key, ""))
        c.alignment = top_left_align
        c.border = border
        row += 2

    # 4期メイン
    ws.cell(row=row, column=1, value="項目 / 期").fill = header_fill
    ws.cell(row=row, column=1).border = border
    for i, t_name in enumerate(TERMS):
        c = ws.cell(row=row, column=i+2, value=t_name)
        c.fill = header_fill
        c.border = border
    row += 1

    for item in config['mid_items']:
        ws.cell(row=row, column=1, value=item).fill = header_fill
        ws.cell(row=row, column=1).border = border
        for i, t_name in enumerate(TERMS):
            c = ws.cell(row=row, column=i+2, value=config['values'].get(f"{item}_{t_name}", ""))
            c.alignment = top_left_align
            c.border = border
        row += 1

    output = BytesIO()
    wb.save(output)
    return output.getvalue()

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
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)
    ws['A1'] = f"【指導計画】 {target_month} ({age})"
    ws['A1'].font = Font(bold=True, size=14)
    
    row = 3
    # 簡易実装：主要データのみ出力
    ws.cell(row=row, column=1, value="項目").fill = header_fill
    for i in range(1, num_weeks+1):
        ws.cell(row=row, column=i+1, value=f"第{i}週").fill = header_fill
    row += 1
    
    mid_labels = [config[f'l_mid{r}'] for r in range(6, 16)]
    for label in mid_labels:
        ws.cell(row=row, column=1, value=label).fill = header_fill
        for w_idx in range(1, num_weeks + 1):
            key = f"{label}_週{w_idx}"
            ws.cell(row=row, column=w_idx+1, value=config['values'].get(key, "")).alignment = top_left_align
            ws.cell(row=row, column=w_idx+1).border = border
        row += 1
        
    output = BytesIO()
    wb.save(output)
    return output.getvalue()

def create_weekly_excel(age, config, orientation):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "週案"
    thin = Side(style='thin')
    border = Border(top=thin, bottom=thin, left=thin, right=thin)
    header_fill = PatternFill(start_color="F2F2F2", fill_type="solid")
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    top_left_align = Alignment(horizontal='left', vertical='top', wrap_text=True)

    ws.merge_cells("A1:D1")
    ws['A1'] = f"【週案】 {config['week_range']} ({age})"
    ws['A1'].font = Font(bold=True, size=14)
    
    ws['A2'] = "週のねらい"
    ws['B2'] = config['values'].get("weekly_aim", "")
    
    days = ["月", "火", "水", "木", "金", "土"]
    row_idx = 4
    for day in days:
        ws.cell(row=row_idx, column=1, value=day)
        ws.cell(row=row_idx, column=2, value=config['values'].get(f"activity_{day}", ""))
        row_idx += 1

    output = BytesIO()
    wb.save(output)
    return output.getvalue()
# ▼▼▼ 追加コードここから ▼▼▼
def ask_gemini_aim(age, keywords):
    # SecretsからAPIキーを取得
    if "GEMINI_API_KEY" not in st.secrets:
        return "エラー: APIキーがSecretsに設定されていません。"
    
    api_key = st.secrets["GEMINI_API_KEY"]
    genai.configure(api_key=api_key)
    
    try:
        # モデル名はこれで完璧です！
        model = genai.GenerativeModel('gemini-pro')
        
        prompt = f"""
        あなたはベテラン保育士です。
        以下の条件で、月間指導計画の「ねらい」の文章を1つ作成してください。
        
        【条件】
        ・対象年齢: {age}
        ・キーワード: {keywords}
        ・文体: 保育の専門用語を用い、最後は「〜する。」で終える。
        """
        
        response = model.generate_content(prompt)
        return response.text.strip()
            
    except Exception as e:
        return f"接続エラー: {str(e)}"
# ▲▲▲ 追加コードここまで ▲▲▲

# --- 4. メイン画面構築 ---

# ロゴとタイトルの表示
col1, col2 = st.columns([1, 5])
with col1:
    try:
        st.image("logo.png", width=80) # ロゴ画像があれば表示
    except:
        st.write("📛") # 画像がない場合の代わり
with col2:
    st.title("保育指導計画システム")

# セッション初期化
if 'annual_data' not in st.session_state: st.session_state['annual_data'] = {}
if 'monthly_data' not in st.session_state: st.session_state['monthly_data'] = {}

# サイドバー設定
st.sidebar.header("⚙️ 設定")
age = st.sidebar.selectbox("対象年齢", ["0歳児", "1歳児", "2歳児", "3歳児", "4歳児", "5歳児"])
mode = st.sidebar.radio("作成する書類", ["年間指導計画", "月間指導計画", "週案"])
orient = st.sidebar.radio("用紙向き", ["横", "縦"])

# 掲示板へのリンク
st.sidebar.markdown("---")
st.sidebar.link_button("☕ 掲示板（休憩室）へ", "https://ここに掲示板のURLを貼ってください")
st.sidebar.markdown("---")

# 📥 データ保存・読込エリア（サイドバー下部）
st.sidebar.subheader("💾 データの保存・読込")
user_id = st.sidebar.text_input("先生のお名前 (ID)", placeholder="例: yamada")
st.sidebar.caption("名前を入力して保存すると、後で続きから始められます。")

c1, c2 = st.sidebar.columns(2)
if c1.button("データ保存"):
    if user_id:
        if save_data_to_sheet(user_id, mode):
            st.sidebar.success(f"{mode}を保存しました！")
    else:
        st.sidebar.error("名前を入力してください")

if c2.button("データ読込"):
    if user_id:
        if load_data_from_sheet(user_id, mode):
            st.sidebar.success("読み込みました！")
            st.rerun() # 画面を更新してデータを反映
        else:
            st.sidebar.warning("データが見つかりません")
    else:
        st.sidebar.error("名前を入力してください")


# ==========================================
# モードA：年間指導計画
# ==========================================
if mode == "年間指導計画":
    st.header(f"📅 {age} 年間指導計画")
    
    default_items = "園児の姿\nねらい\n養護（生命・情緒）\n教育（5領域）\n環境構成・援助\n保護者支援\n行事"
    mid_item_list = st.text_area("項目設定（改行区切り）", default_items).split('\n')

    user_values = {}
    t1, t2 = st.tabs(["📌 基本情報", "📝 各期の計画"])

    with t1:
        st.subheader("年間を通じた目標")
        # keyを指定することで、session_stateに直接値が入る（保存・読込に対応）
        user_values["年間目標"] = st.text_area("年間目標", key="年間目標", height=100)
        user_values["健康・安全"] = st.text_area("健康・安全・災害対策", key="健康・安全", height=100)

    with t2:
        cols = st.columns(4)
        for i, term in enumerate(TERMS):
            with cols[i]:
                st.markdown(f"**{term}**")
                for item in mid_item_list:
                    k = f"{item}_{term}"
                    val = st.text_area(f"{item}", key=k, height=100)
                    user_values[k] = val
                    
                    # 連動用データ保持
                    if term not in st.session_state['annual_data']: st.session_state['annual_data'][term] = {}
                    st.session_state['annual_data'][term][item] = val

    if st.button("🚀 Excel作成"):
        config = {'mid_items': mid_item_list, 'values': user_values}
        data = create_annual_excel(age, config, orient)
        st.download_button("📥 ダウンロード", data, f"年間計画_{age}.xlsx")

# ==========================================
# モードB：月間指導計画
# ==========================================
elif mode == "月間指導計画":
    st.header(f"📝 {age} 月間指導計画")
    # ▼▼▼ 追加コード：AIアシスタントエリア ▼▼▼
    with st.expander("🤖 AIアシスタント（キーワードから『ねらい』を作成）", expanded=True):
        c_ai1, c_ai2, c_ai3 = st.columns([2, 1, 1])
        with c_ai1:
            ai_keywords = st.text_input("キーワードを入力", placeholder="例：雪遊び 手袋 貸し借り 感染症予防")
        with c_ai2:
            target_week = st.selectbox("反映先", ["第1週", "第2週", "第3週", "第4週"])
        with c_ai3:
            st.write("") # レイアウト調整用
            if st.button("✨ AI作成"):
                if not ai_keywords:
                    st.error("キーワードを入れてください")
                else:
                    with st.spinner("AIが執筆中..."):
                        generated_text = ask_gemini_aim(age, ai_keywords)
                        
                        # 生成されたテキストを、対象の週の「ねらい」入力欄にセットする
                        # ※前回のコードで、ねらいのキーは "w{週番号}_6" となっていました
                        week_num = target_week.replace("第", "").replace("週", "") # "1", "2"...
                        target_key = f"w{week_num}_6"
                        
                        st.session_state[target_key] = generated_text
                        st.success(f"{target_week}の『ねらい』に入力しました！")
    # 日付などは保存対象外（毎回選択）とする運用がシンプル
    month_date = st.date_input("対象月", value=datetime.date.today())
    month_str = month_date.strftime("%Y年%m月")
    
    st.info("💡 年間計画のデータがあれば、ここから引用できます")
    if st.button("年間計画から引用"):
         # (連動ロジックは前のまま使用可能)
         pass

    num_weeks = 4
    l_mid = {r: st.text_input(f"項目{r}", val, key=f"lm_{r}") for r, val in zip(range(6, 16), ["ねらい", "養護", "教育", "環境", "支援", "行事", "連携", "食育", "健康", "その他"])}
    
    tabs = st.tabs([f"第{i}週" for i in range(1, 5)] + ["反省"])
    user_values = {}
    
    age_data = TEIKEI_DATA.get(age, {})
    
    for i in range(4):
        with tabs[i]:
            st.caption(f"第{i+1}週")
            for r_num, label in l_mid.items():
                # keyを一意にする: w(週)_(行番号)
                k = f"w{i+1}_{r_num}"
                # 定型文がある項目はselectbox、なければtext_areaに自動切替
                if label in age_data:
                    val = st.selectbox(label, age_data[label] + ["自由入力"], key=k)
                else:
                    val = st.text_area(label, key=k, height=60)
                user_values[f"{label}_週{i+1}"] = val
                
                if label == "ねらい":
                    st.session_state['monthly_data'][f"ねらい_週{i+1}"] = val

    with tabs[4]:
        user_values["reflection"] = st.text_area("振り返り", key="mon_ref", height=100)

    if st.button("🚀 Excel作成"):
        config = {**{f'l_mid{r}': val for r, val in l_mid.items()}, 'values': user_values}
        data = create_monthly_excel(age, month_str, config, num_weeks, orient)
        st.download_button("📥 ダウンロード", data, f"月案_{month_str}.xlsx")

# ==========================================
# モードC：週案
# ==========================================
elif mode == "週案":
    st.header(f"📅 {age} 週案")
    
    start_date = st.date_input("週の開始日")
    
    if st.button("月案からねらい引用"):
        w_aim = st.session_state['monthly_data'].get("ねらい_週1", "")
        if w_aim:
            st.session_state['weekly_aim_input'] = w_aim # 下のtext_areaに反映される
            st.rerun()

    user_values = {}
    # keyを指定して、保存データが読み込まれたらここに表示されるようにする
    user_values["weekly_aim"] = st.text_area("週のねらい", key="weekly_aim_input", height=80)
    
    days = ["月", "火", "水", "木", "金", "土"]
    cols = st.columns(3)
    for i, day in enumerate(days):
        with cols[i%3]:
            st.subheader(f"{day}曜日")
            user_values[f"activity_{day}"] = st.text_area("活動", key=f"act_{day}", height=80)
            user_values[f"care_{day}"] = st.text_area("配慮", key=f"care_{day}", height=60)
            user_values[f"tool_{day}"] = st.text_area("準備", key=f"tool_{day}", height=40)

    if st.button("🚀 Excel作成"):
        config = {'week_range': start_date.strftime('%Y/%m/%d〜'), 'values': user_values}
        data = create_weekly_excel(age, config, orient)
        st.download_button("📥 ダウンロード", data, f"週案_{age}.xlsx")

