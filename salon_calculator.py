import streamlit as st
import pandas as pd

# --- 頁面基本設定 ---
st.set_page_config(page_title="美業業績與薪資計算系統", page_icon="💰", layout="wide")

st.title("💰 美業業績與薪資計算系統")
st.markdown("快速上傳「預約項目.xlsx」結算【泡泡】的原價業績")

# --- 定義常數與預設值 ---
if 'base_salary' not in st.session_state:
    st.session_state.base_salary = 20000

if 'dynamic_prices' not in st.session_state:
    st.session_state.dynamic_prices = {
        70: [200, 250, 300],
        80: [200, 250, 300],
        100: [300, 350, 400],
        110: [350, 400, 450]
    }

# --- 定義新的計價規則 ---
ITEM_CONFIG = {
    # 動態定價項目 (只記錄基礎分鐘數，金額由動態面板控制)
    '水光槍': {'type': 'dynamic', 'mins': 80},
    '水飛梭': {'type': 'dynamic', 'mins': 80},
    '刷酸煥膚': {'type': 'dynamic', 'mins': 80},
    'VITAC駐顏': {'type': 'dynamic', 'mins': 100},
    'PDRN煥膚': {'type': 'dynamic', 'mins': 100},
    '藻針': {'type': 'dynamic', 'mins': 100},
    '藻晶': {'type': 'dynamic', 'mins': 100},
    '矽晶喚顏': {'type': 'dynamic', 'mins': 110},
    '胜肽緊緻': {'type': 'dynamic', 'mins': 80},
    'EMS膠原雪花': {'type': 'dynamic', 'mins': 100},
    '藻針背部護理': {'type': 'dynamic', 'mins': 80},
    '刷酸背部護理': {'type': 'dynamic', 'mins': 70},
    '美白背部護理': {'type': 'dynamic', 'mins': 70},

    # 固定金額項目
    '音波': {'type': 'fixed', 'mins': 40, 'price': 200},
    '韓國音波': {'type': 'fixed', 'mins': 40, 'price': 200},
    '腋下': {'type': 'fixed', 'mins': 15, 'price': 60},
    '烙腮鬍': {'type': 'fixed', 'mins': 15, 'price': 60},
    '腹部肚臍': {'type': 'fixed', 'mins': 15, 'price': 60},
    '上/下手臂': {'type': 'fixed', 'mins': 20, 'price': 100},
    '全手臂': {'type': 'fixed', 'mins': 30, 'price': 150},
    '上/下腿': {'type': 'fixed', 'mins': 20, 'price': 120},
    '全腿': {'type': 'fixed', 'mins': 40, 'price': 200},
    '練習': {'type': 'fixed', 'mins': 0, 'price': 200},
    '比基尼線': {'type': 'fixed', 'mins': 15, 'price': 120},
    'VIO全除': {'type': 'fixed', 'mins': 20, 'price': 180},
    '熱石滑罐1堂': {'type': 'fixed', 'mins': 0, 'price': 150},
    '熱石滑罐2堂': {'type': 'fixed', 'mins': 40, 'price': 300},
    '駐顏、醫美安瓶、原液': {'type': 'fixed', 'mins': 0, 'price': 20},
    '美胸舒緩': {'type': 'fixed', 'mins': 0, 'price': 180},
    '肩頸管理': {'type': 'fixed', 'mins': 0, 'price': 100},
    '頭皮舒壓': {'type': 'fixed', 'mins': 0, 'price': 100},
    '鏟皮刀': {'type': 'fixed', 'mins': 0, 'price': 50},
    '臉部撥筋、EMS、粉晶石': {'type': 'fixed', 'mins': 0, 'price': 150},
    '眼周拉提': {'type': 'fixed', 'mins': 0, 'price': 80},
    '蛋肌除毛': {'type': 'fixed', 'mins': 0, 'price': 100},
    '頸拉提': {'type': 'fixed', 'mins': 0, 'price': 80},

    # 新增項目 (劃分在練習類別內)
    '小平臉': {'type': 'fixed', 'mins': 0, 'price': 300},
    '小平臉+脖子': {'type': 'fixed', 'mins': 0, 'price': 400}
}

# --- 側邊欄設定區 ---
with st.sidebar:
    st.header("⚙️ 設定區")

    st.subheader("設定設計師底薪")
    st.session_state.base_salary = st.number_input("泡泡的底薪", min_value=0, value=st.session_state.base_salary, step=1000)

    st.divider()

    st.subheader("動態階梯業績設定")
    # 將每個分鐘數獨立成一個區塊
    for mins in sorted(st.session_state.dynamic_prices.keys()):
        with st.expander(f"{mins} 分鐘項目金額 (L1 / L2 / L3)", expanded=True):
            cols = st.columns(3)
            for i in range(3):
                new_val = cols[i].number_input(
                    f"Level {i+1}",
                    value=st.session_state.dynamic_prices[mins][i],
                    key=f"dyn_{mins}_{i}",
                    step=50
                )
                st.session_state.dynamic_prices[mins][i] = new_val

# --- 主畫面：檔案上傳 ---
uploaded_file = st.file_uploader("匯入資料表 (.xlsx, .xls)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # 使用 Pandas 讀取 Excel，指定 engine 為 openpyxl
        df = pd.read_excel(uploaded_file, engine='openpyxl')

        # 檢查必要欄位
        required_cols = ['設計師', '類別', '項目', '數量']
        if not all(col in df.columns for col in required_cols):
            st.error(f"檔案缺少必要欄位！請確認包含：{', '.join(required_cols)}")
        else:
            # 清理資料：移除欄位名稱和資料內容前後的空白
            df.columns = df.columns.str.strip()
            df['設計師'] = df['設計師'].astype(str).str.strip()
            df['類別'] = df['類別'].astype(str).str.strip()
            df['項目'] = df['項目'].astype(str).str.strip()
            # 確保數量為數字，若有錯誤填入 0
            df['數量'] = pd.to_numeric(df['數量'], errors='coerce').fillna(0).astype(int)

            # 過濾條件：只看「泡泡」且排除「現場諮詢」
            filtered_df = df[(df['設計師'] == '泡泡') & (df['類別'] != '現場諮詢')].copy()

            if filtered_df.empty:
                st.warning("這份檔案中沒有找到符合計算條件的資料！")
            else:
                # --- 計算邏輯 ---
                total_minutes = 0
                parsed_rows = []
                unmatched_items = set()

                # 第一輪掃描：計算總分鐘數與匹配設定
                for index, row in filtered_df.iterrows():
                    item_name = row['項目']
                    category = row['類別']
                    qty = row['數量']

                    config = ITEM_CONFIG.get(item_name)

                    # 處理同名衝突：水光槍
                    if item_name == '水光槍' and category == '附加課程':
                        config = {'type': 'fixed', 'mins': 0, 'price': 20}

                    if config:
                        total_minutes += config['mins'] * qty
                        parsed_rows.append({'row_data': row, 'config': config, 'qty': qty})
                    else:
                        unmatched_items.add(item_name)
                        parsed_rows.append({'row_data': row, 'config': None, 'qty': qty})

                # 決定 Level
                level_idx = 0
                level_name = "Level 1 (1-65人頭)"
                if total_minutes > 7200:
                    level_idx = 2
                    level_name = "Level 3 (91人頭以上)"
                elif total_minutes > 5200:
                    level_idx = 1
                    level_name = "Level 2 (66-90人頭)"

                # 第二輪掃描：計算金額
                total_revenue = 0
                items_count = 0
                category_summary = {} # 記錄類別與其下的項目 { '類別': { 'total': 0, 'items': { '項目': {'qty':0, 'revenue':0} } } }

                for row_info in parsed_rows:
                    row = row_info['row_data']
                    config = row_info['config']
                    qty = row_info['qty']
                    category = row['類別']
                    item_name = row['項目']

                    item_revenue = 0
                    items_count += qty

                    if config:
                        if config['type'] == 'dynamic':
                            mins = config['mins']
                            # 從 session_state 取出動態設定的價格
                            prices = st.session_state.dynamic_prices.get(mins, [0, 0, 0])
                            item_revenue = prices[level_idx] * qty
                        else:
                            item_revenue = config['price'] * qty

                    total_revenue += item_revenue

                    # 組織報表資料結構
                    if category not in category_summary:
                        category_summary[category] = {'total': 0, 'items': {}}

                    category_summary[category]['total'] += item_revenue

                    if item_name not in category_summary[category]['items']:
                        category_summary[category]['items'][item_name] = {'qty': 0, 'revenue': 0}

                    category_summary[category]['items'][item_name]['qty'] += qty
                    category_summary[category]['items'][item_name]['revenue'] += item_revenue


                # --- 顯示結果 ---
                st.divider()
                st.header("📊 業績計算結果")

                # 未匹配警告
                if unmatched_items:
                    st.error(f"⚠️ 注意：以下項目未設定計價規則，金額無法計算：{', '.join(unmatched_items)}")

                # 重點數據 Metrics
                col1, col2, col3, col4, col5 = st.columns(5)
                col1.metric("總服務項次", f"{items_count} 項")
                col2.metric("總分鐘數", f"{total_minutes} m", delta=level_name, delta_color="off")
                col3.metric("換算人頭數", f"{total_minutes / 80:.1f} 人")
                col4.metric("計算後總抽成", f"NT$ {total_revenue:,}")

                total_salary = st.session_state.base_salary + total_revenue
                col5.metric("總薪資 (底薪+抽成)", f"NT$ {total_salary:,}")

                st.subheader("服務項目類別抽成佔比")

                # 類別詳細列表 (類似手風琴效果)
                # 依總金額排序類別
                sorted_categories = sorted(category_summary.items(), key=lambda x: x[1]['total'], reverse=True)

                for cat_name, cat_data in sorted_categories:
                    with st.expander(f"📁 {cat_name} - NT$ {cat_data['total']:,}"):
                        # 依項目金額排序
                        sorted_items = sorted(cat_data['items'].items(), key=lambda x: x[1]['revenue'], reverse=True)

                        # 將細項轉換成 DataFrame 顯示成表格比較好看
                        item_df = pd.DataFrame([
                            {"項目名稱": name, "數量": data['qty'], "抽成金額": f"NT$ {data['revenue']:,}"}
                            for name, data in sorted_items
                        ])
                        st.dataframe(item_df, use_container_width=True, hide_index=True)

    except Exception as e:
        st.error(f"處理檔案時發生錯誤: {e}")