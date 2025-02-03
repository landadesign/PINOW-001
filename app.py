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
    row_height = 60
    header_height = 80
    padding = 40
    height = header_height + (len(df) + 2) * row_height + padding * 2
    
    # 画像作成
    img = Image.new('RGB', (width, height), 'white')
    draw = ImageDraw.Draw(img)
    
    # フォントサイズ
    title_font_size = 32
    header_font_size = 24
    content_font_size = 24
    
    # デフォルトフォント使用
    font = ImageFont.load_default()
    
    # タイトル描画
    title = f"{name}様 1月 交通費清算書"
    draw.text((padding, padding), title, fill='black', font=font)
    
    # ヘッダー
    headers = ['日付', '経路', '距離(km)', '交通費(円)', '手当(円)', '合計(円)']
    x_positions = [padding, padding + 100, padding + 500, padding + 650, padding + 800, padding + 950]
    
    for header, x in zip(headers, x_positions):
        draw.text((x, padding + header_height), header, fill='black', font=font)
    
    # データ行
    y = padding + header_height + row_height
    for _, row in df.iterrows():
        draw.text((x_positions[0], y), str(row['date']), fill='black', font=font)
        draw.text((x_positions[1], y), str(row['route']), fill='black', font=font)
        draw.text((x_positions[2], y), f"{row['total_distance']:.1f}", fill='black', font=font)
        draw.text((x_positions[3], y), f"{int(row['transportation_fee']):,}", fill='black', font=font)
        draw.text((x_positions[4], y), f"{int(row['allowance']):,}", fill='black', font=font)
        draw.text((x_positions[5], y), f"{int(row['total']):,}", fill='black', font=font)
        y += row_height
    
    # 合計行
    y += row_height
    draw.text((x_positions[0], y), "合計", fill='black', font=font)
    draw.text((x_positions[2], y), f"{df['total_distance'].sum():.1f}", fill='black', font=font)
    draw.text((x_positions[3], y), f"{int(df['transportation_fee'].sum()):,}", fill='black', font=font)
    draw.text((x_positions[4], y), f"{int(df['allowance'].sum()):,}", fill='black', font=font)
    draw.text((x_positions[5], y), f"{int(df['total'].sum()):,}", fill='black', font=font)
    
    # 注釈
    draw.text((padding, height - row_height), "※2025年1月分給与にて清算しました。", fill='black', font=font)
    
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
        
        for line in lines:
            if '様' in line:
                current_name = line.replace('様', '').strip()
                continue
                
            parts = line.split()
            if len(parts) >= 2 and current_name:
                date = parts[0]
                route = ' '.join(parts[1:])
                route_points = route.split('→')
                distance = (len(route_points) - 1) * 5.0
                
                transportation_fee = distance * RATE_PER_KM
                allowance = DAILY_ALLOWANCE
                total = transportation_fee + allowance
                
                data.append({
                    'name': current_name,
                    'date': date,
                    'route': route,
                    'total_distance': distance,
                    'transportation_fee': transportation_fee,
                    'allowance': allowance,
                    'total': total
                })
        
        if data:
            return pd.DataFrame(data)
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
        
        # 全体サマリーの表示
        st.markdown("### 精算サマリー")
        summary_df = df.groupby('name').agg({
            'total_distance': 'sum',
            'transportation_fee': 'sum',
            'allowance': 'sum',
            'total': 'sum'
        }).reset_index()
        
        summary_df.columns = ['担当者', '総距離(km)', '総交通費(円)', '総手当(円)', '総合計(円)']
        
        # 全体の合計行を追加
        total_row = pd.DataFrame([{
            '担当者': '全体合計',
            '総距離(km)': summary_df['総距離(km)'].sum(),
            '総交通費(円)': summary_df['総交通費(円)'].sum(),
            '総手当(円)': summary_df['総手当(円)'].sum(),
            '総合計(円)': summary_df['総合計(円)'].sum()
        }])
        summary_df = pd.concat([summary_df, total_row])
        
        st.dataframe(
            summary_df.style.format({
                '総距離(km)': '{:.1f}',
                '総交通費(円)': '{:,.0f}',
                '総手当(円)': '{:,.0f}',
                '総合計(円)': '{:,.0f}'
            }),
            use_container_width=True,
            hide_index=True
        )
        
        # 個人別の詳細表示
        st.markdown("### 個人別精算詳細")
        unique_names = df['name'].unique().tolist()
        tabs = st.tabs(unique_names)
        
        for i, name in enumerate(unique_names):
            with tabs[i]:
                person_data = df[df['name'] == name].copy()
                
                # データ表示
                st.markdown(f"#### {name}様の精算データ")
                
                display_df = person_data[['date', 'route', 'total_distance', 'transportation_fee', 'allowance', 'total']]
                display_df.columns = ['日付', '経路', '距離(km)', '交通費(円)', '手当(円)', '合計(円)']
                
                # 合計行を追加
                totals = display_df.sum(numeric_only=True).to_frame().T
                totals['日付'] = '合計'
                totals['経路'] = ''
                display_df = pd.concat([display_df, totals])
                
                # データフレーム表示
                st.dataframe(
                    display_df.style.format({
                        '距離(km)': '{:.1f}',
                        '交通費(円)': '{:,.0f}',
                        '手当(円)': '{:,.0f}',
                        '合計(円)': '{:,.0f}'
                    }),
                    use_container_width=True,
                    hide_index=True
                )
                
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
