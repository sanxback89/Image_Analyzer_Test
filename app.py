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

# 분석 항목 정의 (수정됨)
analysis_options = {
    "Top": {
        "Fit": ["Slim Fit", "Regular Fit", "Loose Fit", "Oversized"],
        "Neckline": ["Crew Neck", "V-Neck", "Scoop Neck", "Turtleneck", "Cowl Neck", "Boat Neck", "Halter Neck", "Off-Shoulder", "Sweetheart", "Polo Collar", "Shirt Collar"],
        "Sleeves": ["Short Sleeves", "Long Sleeves", "Three-Quarter Sleeves", "Cap Sleeves", "Sleeveless", "Puff Sleeves"],
        "Length": ["Crop", "Regular", "Long"],
        "Color": ["Red", "Blue", "Green", "Yellow", "Purple", "Orange", "Pink", "Brown", "Black", "White", "Gray", "Multicolor"],
        "Pattern": ["Solid", "Striped", "Polka Dot", "Floral", "Plaid", "Checkered", "Animal Print"],
        "Material": ["Cotton", "Polyester", "Silk", "Wool", "Linen"],
        "Details": ["Ruffles", "Pleats", "Embroidery", "Sequins", "Beading", "Appliqué", "Buttons", "Zippers", "Pockets"]
    },
    "Bottom": {
        "Fit": ["Slim Fit", "Regular Fit", "Loose Fit", "Skinny", "Straight", "Bootcut", "Flare", "Wide Leg"],
        "Length": ["Short", "Knee Length", "Ankle Length", "Full Length"],
        "Rise": ["Low Rise", "Mid Rise", "High Rise"],
        "Color": ["Red", "Blue", "Green", "Yellow", "Purple", "Orange", "Pink", "Brown", "Black", "White", "Gray", "Multicolor"],
        "Pattern": ["Solid", "Striped", "Polka Dot", "Plaid", "Checkered"],
        "Material": ["Denim", "Cotton", "Polyester", "Wool", "Leather"],
        "Details": ["Distressed", "Ripped", "Embroidery", "Pockets", "Belt Loops", "Pleats"]
    },
    "Dress": {
        "Fit": ["Bodycon", "A-Line", "Shift", "Wrap", "Sheath", "Empire Waist"],
        "Neckline": ["V-Neck", "Scoop Neck", "Halter Neck", "Off-Shoulder", "Sweetheart"],
        "Sleeves": ["Sleeveless", "Short Sleeves", "Long Sleeves", "Cap Sleeves", "Puff Sleeves"],
        "Length": ["Mini", "Midi", "Maxi"],
        "Color": ["Red", "Blue", "Green", "Yellow", "Purple", "Orange", "Pink", "Brown", "Black", "White", "Gray", "Multicolor"],
        "Pattern": ["Solid", "Floral", "Polka Dot", "Striped", "Animal Print"],
        "Material": ["Cotton", "Silk", "Polyester", "Chiffon", "Lace"],
        "Details": ["Ruffles", "Pleats", "Embroidery", "Sequins", "Beading", "Belt", "Pockets"]
    },
    "Outerwear": {
        "Type": ["Jacket", "Coat", "Blazer", "Cardigan", "Vest"],
        "Fit": ["Slim Fit", "Regular Fit", "Oversized"],
        "Length": ["Cropped", "Hip Length", "Knee Length", "Long"],
        "Color": ["Red", "Blue", "Green", "Yellow", "Purple", "Orange", "Pink", "Brown", "Black", "White", "Gray", "Multicolor"],
        "Material": ["Leather", "Denim", "Wool", "Cotton", "Polyester"],
        "Details": ["Pockets", "Buttons", "Zippers", "Hood", "Fur Trim", "Quilted"]
    }
}

# 개별 이미지 분석 함수 (캐싱 적용)
@st.cache_data
def analyze_single_image(image, category, options):
    base64_image = encode_image(image)
    
    prompt = f"이미지에 있는 {category} 의류 아이템을 분석하고 다음 측면에 대한 정보를 제공해주세요. 각 옵션에 대해 가장 적합한 하나의 선택지만 선택해주세요:\n\n"
    for option in options:
        prompt += f"{option}: {', '.join(analysis_options[category][option])}\n"
    
    prompt += "\n결과를 선택된 측면을 키로 하고 감지된 옵션을 값으로 하는 JSON 객체로 제공해주세요. 각 키에 대해 하나의 값만 선택해야 합니다."

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
            max_tokens=300
        )
        
        result = response.choices[0].message.content.strip()
        processed_result = preprocess_response(result)
        
        # JSON 파싱 시도
        try:
            return json.loads(processed_result)
        except json.JSONDecodeError:
            st.error(f"JSON 파싱 오류: {processed_result}")
            return {}
    except Exception as e:
        st.error(f"이미지 분석 중 오류 발생: {e}")
        return {}

# 이미지 인코딩 함수
def encode_image(image):
    if isinstance(image, Image.Image):
        # PIL Image 객체인 경우
        buffered = io.BytesIO()
        image.save(buffered, format="PNG")
        return base64.b64encode(buffered.getvalue()).decode('utf-8')
    elif hasattr(image, 'getvalue'):
        # BytesIO 또는 파일 객체인 경우
        return base64.b64encode(image.getvalue()).decode('utf-8')
    else:
        raise ValueError("Unsupported image type")

# 응답 전처리 함수
def preprocess_response(response):
    json_match = re.search(r'```json\s*(.*?)\s*```', response, re.DOTALL)
    if json_match:
        return json_match.group(1)
    return response

# 엑셀에서 이미지 추출 함수
def extract_images_from_excel(uploaded_file):
    wb = openpyxl.load_workbook(io.BytesIO(uploaded_file.getvalue()))
    sheet = wb.active
    image_loader = SheetImageLoader(sheet)
    
    images = []
    for row in sheet.iter_rows():
        for cell in row:
            try:
                if image_loader.image_in(cell.coordinate):
                    image = image_loader.get(cell.coordinate)
                    images.append(image)
            except Exception as e:
                if "I/O operation on closed file" not in str(e):
                    st.warning(f"셀 {cell.coordinate}에서 이미지를 추출하는 중 오류가 발생했습니다: {str(e)}")
                continue
    
    return images

# ZIP 파일 처리 함수
def process_zip_file(uploaded_file):
    with zipfile.ZipFile(io.BytesIO(uploaded_file.getvalue()), 'r') as zip_ref:
        for file_name in zip_ref.namelist():
            if file_name.lower().endswith(('.png', '.jpg', '.jpeg')):
                with zip_ref.open(file_name) as file:
                    yield file_name, file.read()

# 이미지 처리 
def process_images(images):
    processed_images = []
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    for i, img in enumerate(images):
        processed_img = enhance_image(img)
        processed_images.append(processed_img)
        
        # 진행 상황 업데이트
        progress = (i + 1) / len(images)
        progress_bar.progress(progress)
        status_text.text(f"이미지 처리 중: {i+1}/{len(images)}")
    
    progress_bar.empty()
    status_text.empty()
    return processed_images

# 이미지 향상 함수
def enhance_image(image, scale_factor=2):
    cv_image = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
    height, width = cv_image.shape[:2]
    resized = cv2.resize(cv_image, (width*scale_factor, height*scale_factor), interpolation=cv2.INTER_CUBIC)
    gaussian = cv2.GaussianBlur(resized, (0, 0), 3.0)
    sharpened = cv2.addWeighted(resized, 1.5, gaussian, -0.5, 0, resized)
    denoised = cv2.fastNlMeansDenoisingColored(sharpened, None, 10, 10, 7, 21)
    return Image.fromarray(cv2.cvtColor(denoised, cv2.COLOR_BGR2RGB))

# 도넛 차 생성 함수 수정
def create_donut_chart(data, title):
    labels = list(data.keys())
    values = list(data.values())
    
    if title.lower() == 'color':
        colors = [get_color(label) for label in labels]
        # 흰색을 아주 연한 회색으로 변경
        colors = ['#F0F0F0' if color == '#FFFFFF' else color for color in colors]
    else:
        colors = generate_colors(len(labels))
    
    def get_text_color(background_color):
        if background_color == '#000000':
            return '#FFFFFF'
        else:
            return '#000000'
    
    text_colors = [get_text_color(color) for color in colors]
    
    fig = go.Figure(data=[go.Pie(
        labels=labels,
        values=values,
        hole=.3,
        marker_colors=colors,
        textinfo='percent',
        textfont=dict(size=14, color=text_colors, weight='bold'),
        hoverinfo='label+percent+text',
        text=[f'Count: {v}' for v in values],
        hovertemplate='%{label}<br>%{percent}<br>%{text}<extra></extra>'
    )])
    
    fig.update_layout(
        showlegend=True,
        legend=dict(
            orientation='h',
            yanchor='bottom',
            y=-0.2,
            xanchor='center',
            x=0.5,
            font=dict(size=13, weight='bold'),
            itemsizing='constant',
            itemwidth=30
        ),
        width=500,
        height=450,
        margin=dict(t=80, b=80, l=20, r=20),  # 상단 마진을 늘림
        annotations=[
            dict(
                text=f'<b>{title}</b>',
                x=0.5,  # x 위치를 중앙으로 설정
                y=1.05,  # y 위치를 그래프 위로 설정
                xref='paper',
                yref='paper',
                showarrow=False,
                font=dict(size=24, color='black'),  # 타이틀 색상을 검정색으로 변경
                align='center'
            )
        ]
    )
    
    return fig

# 색상 매핑 함수 수정
def get_color(label):
    color_map = {
        'Red': '#FF0000', 'Blue': '#0000FF', 'Green': '#00FF00',
        'Yellow': '#FFFF00', 'Purple': '#800080', 'Orange': '#FFA500',
        'Pink': '#FFC0CB', 'Brown': '#A52A2A', 'Black': '#000000',
        'White': '#FFFFFF', 'Gray': '#808080', 'Multicolor': '#FFFFFF'
    }
    return color_map.get(label, '#000000')

# 색상 생성 함수
def generate_colors(n):
    colors = []
    for _ in range(n):
        hue = random.random()
        saturation = 0.5 + random.random() * 0.5
        lightness = 0.4 + random.random() * 0.2
        rgb = colorsys.hls_to_rgb(hue, lightness, saturation)
        hex_color = '#{:02x}{:02x}{:02x}'.format(int(rgb[0]*255), int(rgb[1]*255), int(rgb[2]*255))
        colors.append(hex_color)
    return colors

# 메인 앱 로직 수정 (이미지 리스트 부분)
def main():
    st.set_page_config(layout="centered")
    
    st.markdown("""
    <style>
    .emoji-title { font-size: 2.4em; }
    .emoji { font-size: 0.8em; }
    </style>
    """, unsafe_allow_html=True)
    
    st.markdown("<h1 class='emoji-title'>패션 이미지 분석기</h1>", unsafe_allow_html=True)
    
    if authenticate_user():
        st.markdown("<h3><span class='emoji'>👚</span> 1단계: 의상 복종 선택</h3>", unsafe_allow_html=True)
        selected_category = st.selectbox(
            "의상 복종을 선택하세요",
            options=list(analysis_options.keys())
        )
        
        st.markdown("<h3><span class='emoji'>🔍</span> 2단계: 분석 항목 선택</h3>", unsafe_allow_html=True)
        selected_options = st.multiselect(
            label="분석할 항목 선택",
            options=list(analysis_options[selected_category].keys()),
            key="analysis_options"
        )
        
        st.markdown("<h3><span class='emoji'>📁</span> 3단계: 파일 업로드</h3>", unsafe_allow_html=True)
        uploaded_file = st.file_uploader("파일 선택", type=["xlsx", "xls", "png", "jpg", "jpeg", "zip"])
        
        if uploaded_file is not None:
            st.markdown("<h3><span class='emoji'>🖼️</span> 4계: 이미지 처리</h3>", unsafe_allow_html=True)
            
            images = []
            if uploaded_file.type in ["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "application/vnd.ms-excel"]:
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
                with st.spinner('이미지 처리 중...'):
                    processed_images = process_images(images)
                
                st.success(f"{len(processed_images)}개의 이미지가 처리되었습니다.")
                
                if st.button("🚀 5단계: 분석 시작", key="start_analysis"):
                    if not selected_options:
                        st.markdown("<p><span class='emoji'>⚠️</span> 분석할 항목을 하나 이상 선택해주세요.</p>", unsafe_allow_html=True)
                    else:
                        progress_bar = st.progress(0)
                        status_text = st.empty()
                        
                        aggregated_results = {option: Counter() for option in selected_options}
                        image_categories = defaultdict(lambda: defaultdict(list))
                        
                        for i, image in enumerate(processed_images):
                            result = analyze_single_image(image, selected_category, selected_options)
                            if result and isinstance(result, dict):  # 결과가 존재하고 딕셔너리인지 확인
                                for option, detected in result.items():
                                    if option in selected_options:  # 선택된 옵션에 대해서만 처리
                                        aggregated_results[option][detected] += 1
                                        image_categories[option][detected].append(image)
                            else:
                                st.warning(f"이미지 {i+1}의 분석 결과가 유효하지 않습니다.")
                            
                            # 진행 상황 업데이트
                            progress = (i + 1) / len(processed_images)
                            progress_bar.progress(progress)
                            status_text.text(f"이미지 분석 중: {i+1}/{len(processed_images)}")
                        
                        progress_bar.empty()
                        status_text.empty()
                        
                        st.markdown("<h3 style='text-align: center;'><span class='emoji'>📊</span> 분석 결과</h3>", unsafe_allow_html=True)
                        
                        for option, results in aggregated_results.items():
                            if results:
                                fig = create_donut_chart(results, option)
                                st.plotly_chart(fig, use_container_width=True)
                                
                                # 세부 결과를 토글 형태로 표시
                                with st.expander(f"{option}"):
                                    for value, count in results.items():
                                        st.markdown(f"**{value}** (Count: {count})", unsafe_allow_html=True)
                                        if option in image_categories and value in image_categories[option]:
                                            images = image_categories[option][value]
                                            cols = st.columns(5)
                                            for i, img in enumerate(images):
                                                with cols[i % 5]:
                                                    st.image(img, use_column_width=True)
                                                if (i + 1) % 5 == 0:
                                                    st.write("")
                                        else:
                                            st.write("해당하는 이미지가 없습니다.")
                                        st.write("---")
                            else:
                                st.write(f"{option}에 대한 데이터가 없습니다.")
            else:
                st.markdown("<p><span class='emoji'>⚠️</span> 업로드된 파일에서 이미지를 찾을 수 없습니다.</p>", unsafe_allow_html=True)

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
    .stExpander {
        border: none !important;
        box-shadow: none !important;
    }
    .stExpander > div:first-child {
        border-radius: 0 !important;
        background-color: transparent !important;
    }
    .stExpander > div:first-child > div:first-child > p {
        font-size: 25px !important;  /* 21px에서 20% 증가 */
        font-weight: bold;  /* 이미 bold로 설정되어 있지만, 확실히 하기 위해 다시 명시 */
    }
    .stButton > button {
        width: 100%;
        text-align: left;
        padding: 0.5rem;
        background-color: #f0f2f6;
        border: none;
        border-radius: 0.3rem;
        margin-bottom: 0.5rem;
        font-weight: bold;
    }
    .stButton > button:hover {
        background-color: #e0e2e6;
    }
</style>
""", unsafe_allow_html=True)
