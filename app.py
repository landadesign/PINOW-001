import streamlit as st
import pandas as pd
from datetime import datetime
import re
import io
from PIL import Image, ImageDraw, ImageFont
import numpy as np
import zipfile
import os
from pathlib import Path

# ページ設定
st.set_page_config(
    page_title="PINO精算アプリケーション",
    layout="wide"
)

# 定数
RATE_PER_KM = 15
DAILY_ALLOWANCE = 200

# スタイル
st.markdown("""
    <style>
    .main {
        padding: 20px;
    }
    .stTextArea textarea {
        font-size: 16px;
    }
    .expense-table {
        font-size: 14px;
    }
    .stButton>button {
        width: 100%;
        background-color: #4CAF50;
        color: white;
        padding: 8px 16px;
        border: none;
        border-radius: 4px;
    }
    .expense-table th {
        text-align: center !important;
    }
    .expense-table td {
        text-align: right !important;
    }
    .expense-table td:nth-child(2) {
        text-align: left !important;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        padding: 8px 16px;
        background-color: #f0f2f6;
    }
    .stTabs [aria-selected="true"] {
        background-color: #4CAF50 !important;
        color: white !important;
    }
    </style>
""", unsafe_allow_html=True)

def parse_expense_data(text):
    if not text:
        return None
    
    entries = []
    
    # テキストを正規化
    text = text.replace('\r\n', '\n').strip()
    
    # 【ピノ】で始まるエントリを抽出
    for entry in text.split('【ピノ】'):
        if not entry.strip():
            continue
            
        # 基本情報を抽出
        match = re.search(r'([^　\s]+(?:[ 　]+[^　\s]+)*)\s+(\d+\/\d+)\s*\([月火水木金土日]\)', entry)
        if not match:
            continue
            
        name, date = match.groups()
        
        # 距離を抽出
        distance_match = re.search(r'(\d+\.?\d*)(?:km|㎞|ｋｍ|kｍ)', entry)
        if not distance_match:
            continue
            
        distance = float(distance_match.group(1))
        
        # 経路を抽出（括弧を含む完全な経路）
        route_start = entry.find(')') + 1
        route_end = entry.find(distance_match.group())
        route = entry[route_start:route_end].strip()
        
        entries.append({
            'name': name.strip(),
            'date': date,
            'route': route,
            'distance': distance
        })
    
    # DataFrameに変換
    df = pd.DataFrame(entries)
    
    # 日付でソート
    df['date'] = pd.to_datetime(df['date'].apply(lambda x: f"2024/{x}"))
    df = df.sort_values('date')
    
    # 日付ごとの集計を作成
    result = []
    for name in df['name'].unique():
        person_data = df[df['name'] == name]
        for date, group in person_data.groupby('date'):
            routes = group['route'].tolist()
            distances = group['distance'].tolist()
            total_distance = sum(distances)
            transportation_fee = int(total_distance * RATE_PER_KM)  # 切り捨て
            
            # 複数経路がある場合
            if len(routes) > 1:
                # 最初の行に計算式を含める
                result.append({
                    'name': name,
                    'date': date.strftime('%-m/%-d'),
                    'route': routes[0],
                    'total_distance': f"{'+'.join([str(d) for d in distances])}={total_distance}",
                    'transportation_fee': transportation_fee,
                    'allowance': DAILY_ALLOWANCE,
                    'total': transportation_fee + DAILY_ALLOWANCE
                })
                # 残りの行
                for route in routes[1:]:
                    result.append({
                        'name': name,
                        'date': date.strftime('%-m/%-d'),
                        'route': route,
                        'total_distance': None,
                        'transportation_fee': None,
                        'allowance': None,
                        'total': None
                    })
            else:
                # 単一経路の場合は通常通り
                result.append({
                    'name': name,
                    'date': date.strftime('%-m/%-d'),
                    'route': routes[0],
                    'total_distance': total_distance,
                    'transportation_fee': transportation_fee,
                    'allowance': DAILY_ALLOWANCE,
                    'total': transportation_fee + DAILY_ALLOWANCE
                })
    
    # 各担当者の最後に合計行を追加
    df_result = pd.DataFrame(result)
    for name in df_result['name'].unique():
        person_data = df_result[df_result['name'] == name]
        result.append({
            'name': name,
            'date': '合計',
            'route': '',
            'total_distance': sum(float(d.split('=')[-1]) if isinstance(d, str) and '=' in d else d 
                               for d in person_data['total_distance'] if d is not None),
            'transportation_fee': person_data['transportation_fee'].sum(),
            'allowance': person_data['allowance'].sum(),
            'total': person_data['total'].sum()
        })
    
    return pd.DataFrame(result)

def format_number(val):
    if pd.isna(val):
        return ''
    if isinstance(val, str) and '+' in val:  # 計算式の場合
        return val
    if isinstance(val, float):
        if val.is_integer():
            return f"{int(val):,}"
        return f"{val:.1f}"
    if isinstance(val, (int, str)):
        try:
            return f"{int(val):,}"
        except (ValueError, TypeError):
            return val
    return val

def create_expense_table_image(df, name, start_date):
    # フォントファイルのパスを設定
    font_path = Path(__file__).parent / "fonts" / "NotoSansJP-Regular.ttf"
    
    # フォントとサイズの設定
    title_font_size = 40
    header_font_size = 24
    content_font_size = 20
    padding = 60
    row_height = 65
    col_widths = [100, 520, 120, 140, 120, 120]
    
    # 画像サイズ設定
    width = sum(col_widths) + padding * 2
    height = (len(df) + 3) * row_height + padding * 3
    
    # 画像の作成（サイズを2倍に）
    scale_factor = 2
    img = Image.new('RGB', (int(width * scale_factor), int(height * scale_factor)), 'white')
    draw = ImageDraw.Draw(img)
    
    # フォント読み込み
    try:
        title_font = ImageFont.truetype(str(font_path), int(title_font_size * scale_factor))
        header_font = ImageFont.truetype(str(font_path), int(header_font_size * scale_factor))
        content_font = ImageFont.truetype(str(font_path), int(content_font_size * scale_factor))
    except Exception as e:
        st.error(f"フォントの読み込みに失敗しました: {e}")
        return None
    
    def scale(x): return int(x * scale_factor)
    
    # タイトルの描画
    title = f"{name}様　1月　交通費清算書"
    draw.text((scale(padding), scale(padding)), title, fill='black', font=title_font)
    
    # ヘッダーの描画
    headers = ['日付', '経路', '距離\n(km)', '交通費\n(円)', '手当\n(円)', '合計\n(円)']
    x = scale(padding)
    y = scale(padding + row_height * 1.5)
    
    # ヘッダー背景
    for header, width in zip(headers, col_widths):
        draw.rectangle([x, y, x + scale(width), y + scale(row_height)], 
                      fill='#f5f5f5', outline='#666666', width=3)
        
        lines = header.split('\n')
        for i, line in enumerate(lines):
            text_width = header_font.getlength(line)
            text_x = x + (scale(width) - text_width) / 2
            text_y = y + scale(5) + (i * scale(row_height - 10) / len(lines))
            draw.text((text_x, text_y), line, fill='black', font=header_font)
        x += scale(width)
    
    # データの描画
    y += scale(row_height)
    for i, (_, row) in enumerate(df.iterrows()):
        x = scale(padding)
        for col_idx, (value, width) in enumerate(zip(row, col_widths)):
            draw.rectangle([x, y, x + scale(width), y + scale(row_height)], 
                         outline='#666666', width=3)
            
            text = str(value) if pd.notna(value) else ''
            
            if col_idx == 1:  # 経路列
                words = text.split('→')
                if len(words) >= 3:
                    mid_point = len(words) // 2
                    line1 = '→'.join(words[:mid_point]) + '→'
                    line2 = '→'.join(words[mid_point:])
                    
                    draw.text((x + scale(10), y + scale(5)), line1, 
                             fill='black', font=content_font)
                    draw.text((x + scale(10), y + scale(row_height/2)), line2, 
                             fill='black', font=content_font)
                else:
                    text_y = y + scale(row_height/2 - content_font_size/2)
                    draw.text((x + scale(10), text_y), text, 
                             fill='black', font=content_font)
            else:
                text_width = content_font.getlength(text)
                text_x = x + scale(width) - text_width - scale(10)
                text_y = y + scale(row_height/2 - content_font_size/2)
                draw.text((text_x, text_y), text, fill='black', font=content_font)
            x += scale(width)
        y += scale(row_height)
    
    # 注釈の描画
    note = "※2025年1月分給与にて清算しました。"
    draw.text((scale(padding), scale(height - row_height)), note, 
              fill='black', font=content_font)
    
    # 画像をバイト列に変換（300dpiで保存）
    img_byte_arr = io.BytesIO()
    img.save(img_byte_arr, format='PNG', dpi=(300, 300))
    img_byte_arr = img_byte_arr.getvalue()
    
    return img_byte_arr

def create_zip_file(images_dict):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for name, img_bytes in images_dict.items():
            zip_file.writestr(
                f"精算書_{name}_{datetime.now().strftime('%Y%m%d')}.png",
                img_bytes
            )
    return zip_buffer.getvalue()

def main():
    st.title("PINO精算アプリケーション")
    
    input_text = st.text_area("精算データを貼り付けてください", height=200)
    
    if st.button("データを解析"):
        if input_text:
            df = parse_expense_data(input_text)
            if df is not None:
                st.session_state['expense_data'] = df
                st.success("データを解析しました！")
    
    if 'expense_data' in st.session_state:
        df = st.session_state['expense_data']
        unique_names = df['name'].unique().tolist()
        
        # 画像変換ボタンを2列で配置
        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("個別画像としてダウンロード", type="primary"):
                st.write("### 個別ダウンロード")
                images_dict = {}
                for name in unique_names:
                    person_data = df[df['name'] == name]
                    start_date = person_data['date'].iloc[0]
                    if start_date != '合計':
                        styled_df = person_data[['date', 'route', 'total_distance', 'transportation_fee', 'allowance', 'total']]
                        styled_df.columns = ['日付', '経路', '合計距離(km)', '交通費（距離×15P）(円)', '運転手当(円)', '合計(円)']
                        formatted_df = styled_df.copy()
                        for col in styled_df.columns:
                            if col != '経路' and col != '日付':
                                formatted_df[col] = styled_df[col].apply(format_number)
                        img_bytes = create_expense_table_image(formatted_df, name, start_date)
                        st.download_button(
                            label=f"{name}様の精算書をダウンロード",
                            data=img_bytes,
                            file_name=f"精算書_{name}_{datetime.now().strftime('%Y%m%d')}.png",
                            mime="image/png",
                            key=f"download_{name}"
                        )
        
        with col2:
            if st.button("一括ZIPでダウンロード", type="primary"):
                images_dict = {}
                for name in unique_names:
                    person_data = df[df['name'] == name]
                    start_date = person_data['date'].iloc[0]
                    if start_date != '合計':
                        styled_df = person_data[['date', 'route', 'total_distance', 'transportation_fee', 'allowance', 'total']]
                        styled_df.columns = ['日付', '経路', '合計距離(km)', '交通費（距離×15P）(円)', '運転手当(円)', '合計(円)']
                        formatted_df = styled_df.copy()
                        for col in styled_df.columns:
                            if col != '経路' and col != '日付':
                                formatted_df[col] = styled_df[col].apply(format_number)
                        images_dict[name] = create_expense_table_image(formatted_df, name, start_date)
                
                zip_bytes = create_zip_file(images_dict)
                st.download_button(
                    label="全ての精算書をZIPでダウンロード",
                    data=zip_bytes,
                    file_name=f"精算書一括_{datetime.now().strftime('%Y%m%d')}.zip",
                    mime="application/zip"
                )
        
        # タブ表示
        tabs = st.tabs(unique_names)
        for idx, name in enumerate(unique_names):
            with tabs[idx]:
                person_data = df[df['name'] == name]
                if len(person_data) > 0:
                    st.write(f"### {name}様　1月　交通費清算書")
                    styled_df = person_data[['date', 'route', 'total_distance', 'transportation_fee', 'allowance', 'total']]
                    styled_df.columns = ['日付', '経路', '合計距離(km)', '交通費（距離×15P）(円)', '運転手当(円)', '合計(円)']
                    formatted_df = styled_df.copy()
                    for col in styled_df.columns:
                        if col != '経路' and col != '日付':
                            formatted_df[col] = styled_df[col].apply(format_number)
                    st.table(formatted_df)
                    st.write("※2025年1月分給与にて清算しました。")

if __name__ == "__main__":
    main()
