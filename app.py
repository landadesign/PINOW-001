import streamlit as st
import pandas as pd
from datetime import datetime
import io
from PIL import Image, ImageDraw, ImageFont
import numpy as np

# ページ設定
st.set_page_config(page_title="PINO精算アプリケーション", layout="wide")

# 定数
RATE_PER_KM = 15
DAILY_ALLOWANCE = 200

def create_expense_table_image(df, name):
    # 画像サイズとフォント設定
    width = 1200
    row_height = 40
    header_height = 60
    padding = 30
    title_height = 50
    
    # 全行数を計算（タイトル + ヘッダー + データ行 + 合計行 + 注釈）
    total_rows = len(df) + 3
    height = title_height + header_height + (total_rows * row_height) + padding * 2
    
    # 画像作成
    img = Image.new('RGB', (width, height), 'white')
    draw = ImageDraw.Draw(img)
    
    # フォント設定
    font = ImageFont.load_default()
    
    # タイトル描画
    title = f"{name}様 1月 社内通貨（交通費）清算額"
    draw.text((padding, padding), title, fill='black', font=font)
    
    # ヘッダー
    headers = ['日付', '経路', '距離(km)', '交通費', '運転手当']
    x_positions = [padding, padding + 80, padding + 800, padding + 900, padding + 1000]
    
    y = padding + title_height
    for header, x in zip(headers, x_positions):
        draw.text((x, y), header, fill='black', font=font)
    
    # 罫線
    line_y = y + header_height - 5
    draw.line([(padding, line_y), (width - padding, line_y)], fill='black', width=1)
    
    # データ行
    y = padding + title_height + header_height
    for _, row in df.iterrows():
        # 日付
        draw.text((x_positions[0], y), str(row['date']), fill='black', font=font)
        
        # 経路（長い場合は折り返し）
        route_text = str(row['route'])
        draw.text((x_positions[1], y), route_text, fill='black', font=font)
        
        # 距離
        distance_text = f"{row['total_distance']:.1f}"
        draw.text((x_positions[2], y), distance_text, fill='black', font=font)
        
        # 交通費
        fee_text = f"{int(row['transportation_fee']):,}"
        draw.text((x_positions[3], y), fee_text, fill='black', font=font)
        
        # 運転手当
        allowance_text = f"{int(row['allowance']):,}"
        draw.text((x_positions[4], y), allowance_text, fill='black', font=font)
        
        y += row_height
    
    # 合計行の罫線
    line_y = y - 5
    draw.line([(padding, line_y), (width - padding, line_y)], fill='black', width=1)
    
    # 合計行
    total_text = f"合計金額: {int(df['total'].sum()):,} 円"
    draw.text((width - padding - 200, y + 10), total_text, fill='black', font=font)
    
    # 注釈
    note_text = "※2025年1月分給与にて清算しました。"
    draw.text((padding, height - padding - 20), note_text, fill='black', font=font)
    
    # 計算日時
    calc_date = datetime.now().strftime("計算日時: %Y/%m/%d")
    draw.text((padding, height - padding - 40), calc_date, fill='black', font=font)
    
    # 画像をバイト列に変換
    img_byte_arr = io.BytesIO()
    img.save(img_byte_arr, format='PNG')
    img_byte_arr = img_byte_arr.getvalue()
    
    return img_byte_arr

def parse_expense_data(text):
    try:
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        data = []
        current_name = None
        
        # 日付ごとのデータを一時保存
        daily_routes = {}
        
        for line in lines:
            if '様' in line:
                # 新しい担当者の処理開始時に前の担当者のデータを集計
                if current_name and daily_routes:
                    for date, routes in daily_routes.items():
                        total_distance = sum(route['distance'] for route in routes)
                        transportation_fee = int(total_distance * RATE_PER_KM)  # 切り捨て
                        data.append({
                            'name': current_name,
                            'date': date,
                            'routes': routes,
                            'total_distance': total_distance,
                            'transportation_fee': transportation_fee,
                            'allowance': DAILY_ALLOWANCE,  # 1日1回のみ
                            'total': transportation_fee + DAILY_ALLOWANCE
                        })
                    daily_routes = {}
                
                current_name = line.replace('様', '').strip()
                continue
            
            parts = line.split()
            if len(parts) >= 2 and current_name:
                date = parts[0]
                route = ' '.join(parts[1:])
                route_points = route.split('→')
                distance = (len(route_points) - 1) * 5.0
                
                if date not in daily_routes:
                    daily_routes[date] = []
                daily_routes[date].append({
                    'route': route,
                    'distance': distance
                })
        
        # 最後の担当者のデータを処理
        if current_name and daily_routes:
            for date, routes in daily_routes.items():
                total_distance = sum(route['distance'] for route in routes)
                transportation_fee = int(total_distance * RATE_PER_KM)  # 切り捨て
                data.append({
                    'name': current_name,
                    'date': date,
                    'routes': routes,
                    'total_distance': total_distance,
                    'transportation_fee': transportation_fee,
                    'allowance': DAILY_ALLOWANCE,
                    'total': transportation_fee + DAILY_ALLOWANCE
                })
        
        if data:
            df = pd.DataFrame(data)
            df = df.sort_values(['name', 'date'])  # 日付順にソート
            return df
        
        st.error("データが見つかりませんでした。正しい形式で入力してください。")
        return None
        
    except Exception as e:
        st.error(f"エラーが発生しました: {str(e)}")
        return None

def main():
    st.title("PINO精算アプリケーション")
    
    # データ入力
    input_text = st.text_area("精算データを貼り付けてください", height=200)
    
    if st.button("データを解析"):
        if input_text:
            df = parse_expense_data(input_text)
            if df is not None:
                st.session_state['expense_data'] = df
                st.success("データを解析しました！")
    
    # データ表示と精算書生成
    if 'expense_data' in st.session_state:
        df = st.session_state['expense_data']
        unique_names = df['name'].unique().tolist()
        
        # 個人別の詳細表示
        tabs = st.tabs(unique_names)
        
        for i, name in enumerate(unique_names):
            with tabs[i]:
                person_data = df[df['name'] == name].copy()
                
                # タイトル表示
                st.markdown(f"### {name}様 2024年12月25日～2025年1月 社内通貨（交通費）清算額")
                
                # データ表示用のリストを作成
                display_rows = []
                for _, row in person_data.iterrows():
                    for route_data in row['routes']:
                        display_rows.append({
                            '日付': row['date'],
                            '経路': route_data['route'],
                            '合計距離(km)': row['total_distance'] if route_data == row['routes'][0] else '',
                            '交通費（距離×15P）(円)': row['transportation_fee'] if route_data == row['routes'][0] else '',
                            '運転手当(円)': row['allowance'] if route_data == row['routes'][0] else '',
                            '合計(円)': row['total'] if route_data == row['routes'][0] else ''
                        })
                
                display_df = pd.DataFrame(display_rows)
                
                # 合計行を追加
                totals = pd.DataFrame([{
                    '日付': '合計',
                    '経路': '',
                    '合計距離(km)': '',
                    '交通費（距離×15P）(円)': '',
                    '運転手当(円)': '',
                    '合計(円)': person_data['total'].sum()
                }])
                display_df = pd.concat([display_df, totals])
                
                # 数値列を適切な型に変換
                numeric_columns = ['合計距離(km)', '交通費（距離×15P）(円)', '運転手当(円)', '合計(円)']
                for col in numeric_columns:
                    display_df[col] = pd.to_numeric(display_df[col], errors='coerce')
                
                # データフレーム表示
                st.dataframe(
                    display_df.style.format({
                        '合計距離(km)': lambda x: f'{x:.1f}' if pd.notnull(x) else '',
                        '交通費（距離×15P）(円)': lambda x: f'{int(x):,}' if pd.notnull(x) else '',
                        '運転手当(円)': lambda x: f'{int(x):,}' if pd.notnull(x) else '',
                        '合計(円)': lambda x: f'{int(x):,}' if pd.notnull(x) else ''
                    }),
                    use_container_width=True,
                    hide_index=True
                )
                
                # 注釈表示
                st.markdown("※2025年1月分給与にて清算しました。")
                
                # 画像生成とダウンロードボタン
                img_bytes = create_expense_table_image(person_data, name)
                st.download_button(
                    label="精算書をダウンロード",
                    data=img_bytes,
                    file_name=f"精算書_{name}_{datetime.now().strftime('%Y%m%d')}.png",
                    mime="image/png"
                )

if __name__ == "__main__":
    main()
