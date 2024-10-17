import streamlit as st
from openai import OpenAI
import base64
from PIL import Image
import plotly.graph_objects as go
import io
import json
from collections import Counter, defaultdict
import re
import random
import zipfile
import pandas as pd
import openpyxl
from openpyxl_image_loader import SheetImageLoader
import cv2
import numpy as np
import time
import colorsys

# OpenAI API 키 설정 (Streamlit Cloud의 secrets에서 가져옴)
client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

# 프로그레스 바 및 상태 메시지를 위한 전역 변수
progress_bar = None
status_text = None

# 사용자 인증 및 사용량 추적
def authenticate_user():
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    
    if not st.session_state.authenticated:
        email = st.text_input("이메일 주소를 입력하세요")
        if st.button("인증"):
            if email.endswith("@yakjin.com"):
                st.session_state.authenticated = True
                st.session_state.email = email
                st.success("인증되었습니다.")
                return True
            else:
                st.error("허용되지 않은 이메일 주소입니다. @yakjin.com 도메인의 이메일만 사용 가능합니다.")
                return False
    return st.session_state.authenticated

# 엑셀에서 이미지 추출 및 ZIP 생성 함수
def extract_images_from_excel(uploaded_file):
    # 메모리에서 엑셀 파일 열기
    wb = openpyxl.load_workbook(io.BytesIO(uploaded_file.getvalue()))
    sheet = wb.active
    image_loader = SheetImageLoader(sheet)
    
    # 이미지 추출
    images = []
    for row in sheet.iter_rows():
        for cell in row:
            try:
                if image_loader.image_in(cell.coordinate):
                    image = image_loader.get(cell.coordinate)
                    images.append(image)
            except Exception as e:
                # 'I/O operation on closed file' 오류는 무시
                if "I/O operation on closed file" not in str(e):
                    st.warning(f"셀 {cell.coordinate}에서 이미지를 추출하는 중 오류가 발생했습니다: {str(e)}")
                continue
    
    return images

def enhance_image(image, scale_factor=2):
    # PIL Image를 OpenCV 형식으로 변환
    cv_image = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
    
    # 이미지 크기 조정
    height, width = cv_image.shape[:2]
    resized = cv2.resize(cv_image, (width*scale_factor, height*scale_factor), interpolation=cv2.INTER_CUBIC)
    
    # 언샤프 마스크 필터 적용
    gaussian = cv2.GaussianBlur(resized, (0, 0), 3.0)
    sharpened = cv2.addWeighted(resized, 1.5, gaussian, -0.5, 0, resized)
    
    # 노이즈 제거
    denoised = cv2.fastNlMeansDenoisingColored(sharpened, None, 10, 10, 7, 21)
    
    # OpenCV 형식을 다시 PIL Image로 변환
    return Image.fromarray(cv2.cvtColor(denoised, cv2.COLOR_BGR2RGB))

# 이미지 분석 관련 함수
def encode_image(image_file):
    return base64.b64encode(image_file.getvalue()).decode('utf-8')

def preprocess_response(response):
    json_match = re.search(r'```json\s*(.*?)\s*```', response, re.DOTALL)
    if json_match:
        return json_match.group(1)
    return response

# 분석 항목 정의 (확장됨)
analysis_options = {
    "Category": ["Top", "Bottom", "Dress", "Outerwear", "Accessories"],
    "Fit": ["Slim Fit", "Regular Fit", "Loose Fit", "Oversized", "Skinny", "Straight", "Bootcut", "Flare"],
    "Neckline": ["Crew Neck", "V-Neck", "Scoop Neck", "Turtleneck", "Cowl Neck", "Boat Neck", "Halter Neck", "Off-Shoulder", "Polo Collar", "Shirt Collar"],
    "Sleeves": ["Short Sleeves", "Long Sleeves", "Three-Quarter Sleeves", "Cap Sleeves", "Sleeveless", "Puff Sleeves"],
    "Length": ["Crop", "Regular", "Long", "Mini", "Midi", "Maxi"],
    "Color": ["Red", "Blue", "Green", "Yellow", "Purple", "Orange", "Pink", "Brown", "Black", "White", "Gray", "Multicolor"],
    "Pattern": ["Solid", "Striped", "Polka Dot", "Floral", "Plaid", "Checkered", "Animal Print"],
    "Material": ["Cotton", "Polyester", "Denim", "Leather", "Silk", "Wool", "Linen"],
    "Details": ["Ruffles", "Pleats", "Embroidery", "Sequins", "Beading", "Appliqué", "Fringe", "Cutouts", "Draping", "Gathering", "Buttons", "Zippers", "Pockets"]
}

def analyze_image(image, options):
    base64_image = encode_image(image)
    
    prompt = "이미지에 있는 패션 아이템을 분석하고 다음 측면에 대한 자세한 정보를 제공해주세요:\n\n"
    for option in options:
        prompt += f"{option}: 이미지에서 가장 적절한 옵션을 선택하거나, 해당되는 경우 새로운 값을 제안해주세요.\n"
    
    prompt += "\n결과를 선택된 측면을 키로 하고 감지된 옵션을 값으로 하는 JSON 객체로 제공해주세요. 여러 옵션이 감지되면 신뢰도 순으로 나열해주세요."

    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt},
                        {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{base64_image}"}}
                    ]
                }
            ],
            max_tokens=500
        )
        
        result = response.choices[0].message.content.strip()
        return preprocess_response(result)
    except Exception as e:
        st.error(f"이미지 분석 중 오류 발생: {e}")
        return ""

def generate_sophisticated_colors(n, used_colors=None):
    if used_colors is None:
        used_colors = set()
    colors = []
    hue_step = 1.0 / n
    for i in range(n):
        attempts = 0
        while attempts < 100:
            hue = (i * hue_step + random.random() * 0.5 * hue_step) % 1.0
            saturation = 0.6 + random.random() * 0.2
            lightness = 0.5 + random.random() * 0.2
            r, g, b = [int(x * 255) for x in colorsys.hls_to_rgb(hue, lightness, saturation)]
            color = f"rgb({r},{g},{b})"
            if color not in used_colors:
                colors.append(color)
                used_colors.add(color)
                break
            attempts += 1
        if attempts == 100:
            colors.append(f"rgb({random.randint(0, 255)},{random.randint(0, 255)},{random.randint(0, 255)})")
    return colors, used_colors

def create_donut_chart(data, title, used_colors=None):
    labels = list(data.keys())
    values = list(data.values())
    
    is_color_category = title.lower() == "color"
    if is_color_category:
        colors = []
        for color in labels:
            if color.lower() == "white":
                colors.append("rgb(248, 248, 248)")
            else:
                colors.append(color)
    else:
        colors, used_colors = generate_sophisticated_colors(len(labels), used_colors)

    # 텍스트 색상 결정
    text_colors = []
    for color in colors:
        if is_color_category and color == "rgb(248, 248, 248)":
            text_color = 'black'
        elif is_color_category:
            text_color = 'white' if color.lower() in ['black', 'navy', 'dark blue', 'dark green'] else 'black'
        else:
            r, g, b = map(int, color.strip('rgb()').split(','))
            brightness = (r * 299 + g * 587 + b * 114) / 1000
            text_color = 'white' if brightness < 128 else 'black'
        text_colors.append(text_color)

    fig = go.Figure(data=[go.Pie(
        labels=labels,
        values=values, 
        hole=.4,
        marker=dict(colors=colors, line=dict(color='#FFFFFF', width=0)),
        textposition='inside',
        textinfo='percent',
        hoverinfo='label+value',
        hovertemplate='%{label}<br>Count: %{value}<extra></extra>',
        textfont=dict(size=12, color=text_colors),
        insidetextorientation='radial'
    )])
    
    fig.update_layout(
        title=dict(
            text=title.capitalize(),
            font=dict(size=31),
            y=0.95,
            x=0.5,
            xanchor='center',
            yanchor='top'
        ),
        font=dict(family="Arial", size=14),
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        height=470,
        margin=dict(l=20, r=20, t=60, b=20),
        showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.2,
            xanchor="center",
            x=0.5
        )
    )
    return fig, used_colors

def process_zip_file(zip_file):
    image_files = []
    with zipfile.ZipFile(zip_file, 'r') as zip_ref:
        for file_name in zip_ref.namelist():
            if file_name.lower().endswith(('.png', '.jpg', '.jpeg')):
                with zip_ref.open(file_name) as file:
                    image_files.append((file_name, io.BytesIO(file.read())))
    return image_files

# 이미지 처리 함수 수정 (단순화)
def process_images(images):
    processed_images = []
    for image in images:
        # 필요한 이미지 처리 로직 추가
        # 예: image = some_processing_function(image)
        processed_images.append(image)
    return processed_images

def show_images_for_category(category, value, images):
    with st.expander(f"{category}: {value} (클릭하여 이미지 보기)", expanded=False):
        cols = st.columns(4)
        for i, img in enumerate(images):
            with cols[i % 4]:
                st.image(img, use_column_width=True)
                if st.button(f"전체 크기로 보기 {i+1}", key=f"{category}_{value}_{i}"):
                    st.image(img, use_column_width=True)

# 메인 앱 로직
def main():
    global progress_bar, status_text
    
    st.set_page_config(layout="wide")
    
    # CSS를 사용하여 이모지 크기 조절
    st.markdown("""
    <style>
    .emoji-title {
        font-size: 2.4em;
    }
    .emoji {
        font-size: 0.8em;
    }
    </style>
    """, unsafe_allow_html=True)
    
    st.markdown("<h1 class='emoji-title'>패션 이미지 분석기</h1>", unsafe_allow_html=True)
    
    if authenticate_user():
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        step1 = st.empty()
        step1.markdown("<h3><span class='emoji'>📁</span> 1단계: 파일 업로드</h3>", unsafe_allow_html=True)
        uploaded_file = st.file_uploader("파일 선택", type=["xlsx", "xls", "png", "jpg", "jpeg", "zip"])
        
        if uploaded_file is not None:
            step1.empty()
            
            step2 = st.empty()
            step2.markdown("<h3><span class='emoji'>🖼️</span> 2단계: 이미지 처리</h3>", unsafe_allow_html=True)
            
            images = []
            if uploaded_file.type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" or uploaded_file.type == "application/vnd.ms-excel":
                try:
                    images = extract_images_from_excel(uploaded_file)
                    if images:
                        images = images[1:]  # 첫 번째 이미지(로고) 제외
                except Exception as e:
                    st.error(f"엑셀 파일에서 이미지를 추출하는 중 오류가 발생했습니다: {str(e)}")
            elif uploaded_file.type.startswith('image/'):
                images = [Image.open(uploaded_file)]
            elif uploaded_file.type == 'application/zip':
                images = [Image.open(io.BytesIO(img_data)) for _, img_data in process_zip_file(uploaded_file)]
            
            if images:
                st.markdown(f"<p><span class='emoji'>✅</span> {len(images)}개의 이미지가 처리되었습니다.</p>", unsafe_allow_html=True)
                
                processed_images = process_images(images)
                
                st.markdown("<h3><span class='emoji'>🔍</span> 3단계: 분석 결과</h3>", unsafe_allow_html=True)
                
                selected_options = st.multiselect(
                    label="분석할 항목 선택",
                    options=list(analysis_options.keys()),
                    key="analysis_options"
                )
                
                if st.button("🚀 분석 시작"):
                    if not selected_options:
                        st.markdown("<p><span class='emoji'>⚠️</span> 분석할 항목을 하나 이상 선택해주세요.</p>", unsafe_allow_html=True)
                    else:
                        step2.empty()
                        
                        aggregated_results = {option: Counter() for option in selected_options}
                        image_categories = defaultdict(lambda: defaultdict(list))
                        
                        progress_bar = st.progress(0)
                        status_text = st.empty()
                        
                        for i, image in enumerate(processed_images):
                            progress = (i + 1) / len(processed_images)
                            progress_bar.progress(progress)
                            status_text.text(f"이미지 분석 중: {i+1}/{len(processed_images)}")
                            
                            img_byte_arr = io.BytesIO()
                            image.save(img_byte_arr, format='PNG')
                            img_byte_arr = img_byte_arr.getvalue()
                            
                            try:
                                result = analyze_image(io.BytesIO(img_byte_arr), selected_options)
                                if result:
                                    analysis_results = json.loads(result)
                                    for option, detected in analysis_results.items():
                                        if isinstance(detected, list):
                                            aggregated_results[option].update(detected)
                                            for value in detected:
                                                image_categories[option][value].append(image)
                                        elif isinstance(detected, str):
                                            aggregated_results[option][detected] += 1
                                            image_categories[option][detected].append(image)
                                        else:
                                            st.warning(f"이미지 {i+1}의 {option}에 대해 예상치 못한 결과 형식입니다: {detected}")
                            except Exception as e:
                                st.error(f"이미지 {i+1} 처리 중 오류 발생: {str(e)}")
                                continue
                        
                        progress_bar.empty()
                        status_text.empty()
                        
                        st.markdown("<h3 style='text-align: center;'><span class='emoji'>📊</span> 분석 결과</h3>", unsafe_allow_html=True)
                        col1, col2 = st.columns(2)
                        used_colors = set()
                        for i, (option, results) in enumerate(aggregated_results.items()):
                            if results:
                                fig, used_colors = create_donut_chart(results, option, used_colors)
                                if i % 2 == 0:
                                    with col1:
                                        st.plotly_chart(fig, use_container_width=True)
                                        for value, count in results.items():
                                            show_images_for_category(option, value, image_categories[option][value])
                                else:
                                    with col2:
                                        st.plotly_chart(fig, use_container_width=True)
                                        for value, count in results.items():
                                            show_images_for_category(option, value, image_categories[option][value])
                            else:
                                st.write(f"{option}에 대한 데이터가 없습니다.")
                        
                        # 사용량 추적
                        if "image_count" not in st.session_state:
                            st.session_state.image_count = 0
                        st.session_state.image_count += len(images)
                        st.write(f"총 분석된 이미지 수: {st.session_state.image_count}")
            else:
                st.markdown("<p><span class='emoji'>⚠️</span> 업로드된 파일에서 이미지를 찾을 수 없습니다.</p>", unsafe_allow_html=True)
    
    progress_bar.empty()
    status_text.empty()

if __name__ == "__main__":
    main()

# Streamlit 테마 설정을 위한 CSS
st.markdown("""
<style>
    .stMultiSelect [data-baseweb="tag"] {
        background-color: #007AFF !important;
    }
    .stMultiSelect [data-baseweb="tag"] span {
        color: white !important;
    }
    .stProgress .st-bo {
        background-color: #4CD964;
    }
    .stProgress .st-bp {
        background-color: #E5E5EA;
    }
    .stSelectbox label, .stMultiSelect label, .stFileUploader label {
        font-size: 16px !important;
        color: rgba(49, 51, 63, 0.6) !important;
    }
</style>
""", unsafe_allow_html=True)
