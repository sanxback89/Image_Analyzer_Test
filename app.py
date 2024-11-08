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
import concurrent.futures
from itertools import islice

# OpenAI API key setup (fetched from Streamlit Cloud secrets)
client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

# Global variables for progress bar and status message
progress_bar = None
status_text = None

# Sleeve length guide definition
sleeve_length_guide = """
For Sleeve Length analysis, please consider these important factors:

1. Look for design intention and original garment construction:
- Check for cuffs, hem finishing, or design details that indicate the intended sleeve length
- Observe if there are buttons or tabs designed for rolling up sleeves
- Look for permanent design elements like elastic bands or fixed cuffs

2. Important: Distinguish between designed length vs. styled wearing:
- If sleeves appear rolled up or pushed up, analyze the original intended length
- Look for fabric bunching or gathering that suggests rolled-up long sleeves
- Consider the overall garment style and category to determine original design

3. Length definitions:
- Long Sleeves: Full arm length to wrist, even if currently rolled up
- Three-Quarter Sleeves: Designed to end between elbow and wrist
- Short Sleeves: Designed to end at or above elbow
- Cap Sleeves: Very short, just covering shoulder
- Sleeveless: No sleeve coverage

4. Key indicators of rolled-up long sleeves:
- Visible fabric bunching or folding
- Uneven or casual sleeve ending
- Presence of cuffs or buttons above current sleeve end
- Wrinkles or creases indicating temporary folding

Please analyze the ORIGINAL designed sleeve length, not how it's currently styled or worn.
"""

# 허용된 사용자 딕셔너리 (이메일: 비밀번호)
ALLOWED_USERS = {
    "baekdoo28@gmail.com": "Yakjin135#",
    "jenna.lee@yakjin.com": "Yakjin135#",
    "cielito@yakjin.com": "Yakjin135#",
    "jesssieyun@yakjin.com": "Yakjin135#",
    "jake@yakjin.com": "Yakjin135#",
    "zoe.choi@yakjin.com": "Yakjin135#",
    "silvia@yakjin.com": "Yakjin135#",
    "jiwoo.shin@yakjin.com": "Yakjin135#",
    "elena.lee@yakjin.com": "Yakjin135#",
    "eunh.choi@yakjin.com": "Yakjin135#"
}

# User authentication and usage tracking
def authenticate_user():
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    
    if not st.session_state.authenticated:
        email = st.text_input("Enter your email address")
        password = st.text_input("Enter your password", type="password")
        if st.button("Authentication"):
            if email in ALLOWED_USERS and ALLOWED_USERS[email] == password:
                st.session_state.authenticated = True
                st.session_state.email = email
                st.success("Authentication succeeded.")
                return True
            else:
                st.error("This is an unverified email address or incorrect password. Access denied.")
                return False
    return st.session_state.authenticated

# Analysis options definition (modified)
analysis_options = {
    "Top": {
        "Fit": ["Slim Fit", "Loose Fit", "Oversized"],
        "Neckline": ["Crew Neck", "V-Neck", "Square Neck", "Scoop Neck", "Henley Neck", "Turtleneck", "Cowl Neck", "Boat Neck", "Halter Neck", "Off-Shoulder", "Sweetheart", "Polo Collar", "Shirt Collar"],
        "Sleeves": ["Short Sleeves", "Long Sleeves", "Three-Quarter Sleeves", "Cap Sleeves", "Sleeveless", "Half Sleeves", "Puff Sleeves"],
        "Sleeves Construction": ["Set-In", "Raglan", "Dolman", "Drop Shoulder", "Unspecified"],
        "Length": ["Crop", "Regular", "Long"],
        "Color": ["Red", "Blue", "Green", "Yellow", "Purple", "Orange", "Pink", "Brown", "Black", "White", "Gray", "Multicolor"],
        "Pattern": ["Floral", "Animal print", "Tropical", "Camouflage", "Geometric Print", "Abstract Print", "Heart/Dot/Star", "Bandana/Paisley", "Conversational Print", "Logo", "Lettering", "Dyeing Effect", "Ethnic/Tribal", "Stripes", "Plaid/Checks", "Christmas", "Shine", "Unspecified"],
        "Material": ["Cotton", "Polyester", "Silk", "Wool", "Linen"],
        "Details": ["Ruffles", "Pleats", "Embroidery", "Sequins", "Beading", "Appliqué",
                   "Shirring", "Wrap", "Twist", "Knot", "mixed_media", "Seam detail", "Cut out", "Seamless", "Contrast Binding"]
    },
    "Bottom": {
        "Fit": ["Slim Fit", "Regular Fit", "Loose Fit", "Skinny", "Straight", "Bootcut", "Flare", "Wide Leg"],
        "Length": ["Short", "Knee Length", "Ankle Length", "Full Length"],
        "Rise": ["Low Rise", "Mid Rise", "High Rise"],
        "Color": ["Red", "Blue", "Green", "Yellow", "Purple", "Orange", "Pink", "Brown", "Black", "White", "Gray", "Multicolor"],
        "Pattern": ["Floral", "Animal print", "Tropical", "Camouflage", "Geometric Print", "Abstract Print", "Heart/Dot/Star", "Bandana/Paisley", "Conversational Print", "Logo", "Lettering", "Dyeing Effect", "Ethnic/Tribal", "Stripes", "Plaid/Checks", "Christmas", "Shine", "Unspecified"],
        "Material": ["Denim", "Cotton", "Polyester", "Wool", "Leather"],
        "Details": ["Distressed", "Ripped", "Embroidery", "Pockets", "Belt Loops", "Pleats"]
    },
    "Dress": {
        "Fit": ["Bodycon", "A-Line", "Fit&Flare", "Shift", "Sheath", "Empire Waist"],
        "Neckline": ["Crew Neck", "V-Neck", "Square Neck", "Scoop Neck", "Henley Neck", "Turtleneck", "Cowl Neck", "Boat Neck", "Halter Neck", "Off-Shoulder", "Sweetheart", "Polo Collar", "Shirt Collar"],
        "Sleeves": ["Short Sleeves", "Long Sleeves", "Three-Quarter Sleeves", "Cap Sleeves", "Sleeveless", "Half Sleeves", "Puff Sleeves"],
        "Sleeves Construction": ["Set-In", "Raglan", "Dolman", "Drop Shoulder", "Unspecified"],
        "Length": ["Mini", "Midi", "Maxi", "Above Knee", "Knee Length", "Below Knee"],
        "Color": ["Red", "Blue", "Green", "Yellow", "Purple", "Orange", "Pink", "Brown", "Black", "White", "Gray", "Multicolor"],
        "Pattern": ["Floral", "Animal print", "Tropical", "Camouflage", "Geometric Print", "Abstract Print", "Heart/Dot/Star", "Bandana/Paisley", "Conversational Print", "Logo", "Lettering", "Dyeing Effect", "Ethnic/Tribal", "Stripes", "Plaid/Checks", "Christmas", "Shine", "Unspecified"],
        "Material": ["Cotton", "Silk", "Polyester", "Chiffon", "Lace"],
        "Details": ["Ruffles", "Pleats", "Embroidery", "Sequins", "Beading",  
                   "Shirring", "Wrap", "Twist", "Knot", "mixed_media", "Seam detail", "Cut out", "Seamless", "Contrast Binding"]
    },
    "Outerwear": {
        "Type": ["Jacket", "Coat", "Blazer", "Cardigan", "Vest"],
        "Fit": ["Slim Fit", "Regular Fit", "Oversized"],
        "Length": ["Cropped", "Hip Length", "Knee Length", "Long"],
        "Color": ["Red", "Blue", "Green", "Yellow", "Purple", "Orange", "Pink", "Brown", "Black", "White", "Gray", "Multicolor"],
        "Material": ["Leather", "Denim", "Wool", "Cotton", "Polyester"],
        "Details": ["Pockets", "Buttons", "Zippers", "Hood", "Fur Trim", "Quilted"],
        "Pattern": ["Floral", "Animal print", "Tropical", "Camouflage", "Geometric Print", "Abstract Print", "Heart/Dot/Star", "Bandana/Paisley", "Conversational Print", "Logo", "Lettering", "Dyeing Effect", "Ethnic/Tribal", "Stripes", "Plaid/Checks", "Christmas", "Shine", "Unspecified"]
    }
}

# 배치 처리를 위한 헬퍼 함수
def batch_images(iterable, batch_size):
    iterator = iter(iterable)
    return iter(lambda: list(islice(iterator, batch_size)), [])

# 병렬 처리를 위한 분석 함수
def analyze_image_batch(batch_data):
    image, category, options = batch_data
    return analyze_single_image(image, category, options)

# 이미지 해시 함수 추가
def get_image_hash(image):
    if isinstance(image, Image.Image):
        # PIL 이미지를 numpy 배열로 변환
        img_array = np.array(image)
    else:
        # 이미 numpy 배열인 경우
        img_array = image
    
    # 이미지를 32x32로 리사이즈하고 평균 해시 계산
    resized = cv2.resize(img_array, (32, 32))
    gray = cv2.cvtColor(resized, cv2.COLOR_RGB2GRAY)
    avg = gray.mean()
    hash_str = ''.join(['1' if pixel > avg else '0' for pixel in gray.flatten()])
    return hash_str

# 수정된 분석 함수
@st.cache_data(ttl=24*3600, show_spinner=False, hash_funcs={Image.Image: get_image_hash})
def analyze_single_image(image, category, options):
    base64_image = encode_image(image)
    
    prompt = f"Analyze the {category} clothing item in the image and provide information on the following aspects:\n\n"
    
    for option in options:
        if option == "Sleeves":
            prompt += f"\n{sleeve_length_guide}\n"
        
        if option == "Details":
            prompt += f"{option}: Select ALL that apply from [{', '.join(analysis_options[category][option])}]\n"
        else:
            prompt += f"{option}: Select ONE from [{', '.join(analysis_options[category][option])}]\n"
    
    prompt += "\nProvide the result as a JSON object with the selected aspects as keys and the detected options as values. For 'Details', provide an array of all applicable options. For other aspects, provide a single value."

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
            max_tokens=300,
            temperature=0.0,
            seed=42
        )
        
        result = response.choices[0].message.content.strip()
        processed_result = preprocess_response(result)
        
        try:
            return json.loads(processed_result)
        except json.JSONDecodeError:
            st.error(f"JSON Parsing Error: {processed_result}")
            return {}
            
    except Exception as e:
        st.error(f"Error During Image Analysis: {e}")
        return {}

# Image encoding function
def encode_image(image):
    if isinstance(image, Image.Image):
        # If it's a PIL Image object
        buffered = io.BytesIO()
        image.save(buffered, format="PNG")
        return base64.b64encode(buffered.getvalue()).decode('utf-8')
    elif hasattr(image, 'getvalue'):
        # If it's a BytesIO or file object
        return base64.b64encode(image.getvalue()).decode('utf-8')
    else:
        raise ValueError("Unsupported Image Type")

# Response preprocessing function
def preprocess_response(response):
    json_match = re.search(r'```json\s*(.*?)\s*```', response, re.DOTALL)
    if json_match:
        return json_match.group(1)
    return response

# Function to extract images from Excel
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
                    st.warning(f"Error Extracting Image from Cell {cell.coordinate}: {str(e)}")
                continue
    
    return images

# ZIP file processing function
def process_zip_file(uploaded_file):
    with zipfile.ZipFile(io.BytesIO(uploaded_file.getvalue()), 'r') as zip_ref:
        for file_name in zip_ref.namelist():
            if file_name.lower().endswith(('.png', '.jpg', '.jpeg')):
                with zip_ref.open(file_name) as file:
                    yield file_name, file.read()

# Image processing
def process_images(images):
    processed_images = []
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    for i, img in enumerate(images):
        processed_img = enhance_image(img)
        processed_images.append(processed_img)
        
        # Update progress
        progress = (i + 1) / len(images)
        progress_bar.progress(progress)
        status_text.text(f"Processing Images: {i+1}/{len(images)}")
    
    progress_bar.empty()
    status_text.empty()
    return processed_images

# Image enhancement function
def enhance_image(image, scale_factor=1):
    # PIL 이미지를 OpenCV 형식으로 변환
    cv_image = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
    
    # 1. 이미지 크기 조정 (최적 크기로 조정)
    min_dimension = 800  # 최소 크기 제한
    max_dimension = 1200  # 최대 크기 제한
    height, width = cv_image.shape[:2]
    
    # 작은 이미지는 확대
    if max(height, width) < min_dimension:
        scale = min_dimension / max(height, width)
        new_width = int(width * scale)
        new_height = int(height * scale)
        cv_image = cv2.resize(cv_image, (new_width, new_height), interpolation=cv2.INTER_LINEAR)
    # 큰 이미지는 축소
    elif max(height, width) > max_dimension:
        scale = max_dimension / max(height, width)
        new_width = int(width * scale)
        new_height = int(height * scale)
        cv_image = cv2.resize(cv_image, (new_width, new_height), interpolation=cv2.INTER_AREA)
    
    # 2. 기본적인 노이즈 제거 (빠른 처리를 위해 파라미터 조정)
    denoised = cv2.fastNlMeansDenoisingColored(cv_image, None, 7, 7, 5, 12)
    
    return Image.fromarray(cv2.cvtColor(denoised, cv2.COLOR_BGR2RGB))

# 고유한 색상 세트를 생성하는 함수
def generate_unique_color_sets(num_sets, colors_per_set):
    all_colors = []
    for _ in range(num_sets):
        set_colors = []
        for _ in range(colors_per_set):
            while True:
                hue = random.random()
                saturation = 0.5 + random.random() * 0.5
                lightness = 0.4 + random.random() * 0.2
                rgb = colorsys.hls_to_rgb(hue, lightness, saturation)
                hex_color = '#{:02x}{:02x}{:02x}'.format(int(rgb[0]*255), int(rgb[1]*255), int(rgb[2]*255))
                if hex_color not in all_colors:
                    set_colors.append(hex_color)
                    all_colors.append(hex_color)
                    break
        yield set_colors

# 수정된 create_donut_chart 함수
def create_donut_chart(data, title, color_set):
    labels = list(data.keys())
    values = list(data.values())
    
    if title.lower() == 'color':
        colors = [get_color(label) for label in labels]
        colors = ['#F0F0F0' if color == '#FFFFFF' else color for color in colors]
    else:
        colors = color_set[:len(labels)]
    
    def get_text_color(background_color):
        r, g, b = int(background_color[1:3], 16), int(background_color[3:5], 16), int(background_color[5:7], 16)
        luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
        return '#000000' if luminance > 0.5 else '#FFFFFF'
    
    text_colors = [get_text_color(color) for color in colors]
    
    fig = go.Figure(data=[go.Pie(
        labels=labels,
        values=values,
        hole=.3,
        marker_colors=colors,
        textinfo='percent',
        textfont=dict(size=14, color=text_colors),
        hoverinfo='label+percent+text',
        text=[f'Count: {v}' for v in values],
        hovertemplate='%{label}<br>%{percent}<br>%{text}<extra></extra>'
    )])
    
    # 이아웃 설정 (이전과 동일)
    fig.update_layout(
        showlegend=True,
        legend=dict(
            orientation='h',
            yanchor='bottom',
            y=-0.3,
            xanchor='center',
            x=0.5,
            font=dict(size=15),
            itemsizing='constant',
            itemwidth=30
        ),
        width=500,
        height=450,
        margin=dict(t=70, b=90, l=20, r=20),
        annotations=[
            dict(
                text=f'<b>{title}</b>',
                x=0.5,
                y=1.2,
                xref='paper',
                yref='paper',
                showarrow=False,
                font=dict(size=32, color='black'),
                align='center'
            )
        ]
    )
    
    return fig

# Modified color mapping function
def get_color(label):
    color_map = {
        'Red': '#FF0000', 'Blue': '#0000FF', 'Green': '#00FF00',
        'Yellow': '#FFFF00', 'Purple': '#800080', 'Orange': '#FFA500',
        'Pink': '#FFC0CB', 'Brown': '#A52A2A', 'Black': '#000000',
        'White': '#FFFFFF', 'Gray': '#808080', 'Multicolor': '#FFFFFF'
    }
    return color_map.get(label, '#000000')

# Color generation function
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

# 세션 상태에 분석 결과 저장을 위한 초기화 함수 추가
def initialize_session_state():
    if 'analysis_results' not in st.session_state:
        st.session_state.analysis_results = {}
    if 'image_categories' not in st.session_state:
        st.session_state.image_categories = defaultdict(lambda: defaultdict(list))

# 이미지 제거 함수 추가
def remove_image(option, value, image_index):
    if option in st.session_state.analysis_results and value in st.session_state.image_categories[option]:
        # 이미지 리스트에서 제거
        st.session_state.image_categories[option][value].pop(image_index)
        
        # 카운터 업데이트
        if option == "Details":
            st.session_state.analysis_results[option][value] -= 1
            if st.session_state.analysis_results[option][value] == 0:
                del st.session_state.analysis_results[option][value]
        else:
            st.session_state.analysis_results[option][value] -= 1
        
        # 세션 상태 업데이트 트리거
        st.session_state.update_charts = True

# 이미지 이동 함수 추가
def move_image(from_option, from_value, to_value, image_index):
    if (from_option in st.session_state.image_categories and 
        from_value in st.session_state.image_categories[from_option]):
        # 이미지 가져오기
        image = st.session_state.image_categories[from_option][from_value][image_index]
        
        # 원래 카테고리에서 이미지 제거
        st.session_state.image_categories[from_option][from_value].pop(image_index)
        st.session_state.analysis_results[from_option][from_value] -= 1
        
        # 카운트가 0이 되면 카테고리 제거
        if st.session_state.analysis_results[from_option][from_value] == 0:
            del st.session_state.analysis_results[from_option][from_value]
            del st.session_state.image_categories[from_option][from_value]
        
        # 새 카테고리에 이미지 추가
        st.session_state.image_categories[from_option][to_value].append(image)
        st.session_state.analysis_results[from_option][to_value] = (
            st.session_state.analysis_results[from_option].get(to_value, 0) + 1
        )
        
        return True
    return False

# 이미지 이동을 위한 새로운 함수 추가
def move_selected_images(from_option, from_value, to_value, selected_indices):
    """
    선택된 이미지들을 한 카테고리에서 다른 카테고리로 이동
    """
    if not selected_indices:
        return False
    
    # 인덱스를 내림차순으로 정렬 (높은 인덱스부터 제거)
    selected_indices.sort(reverse=True)
    
    moved_images = []
    for idx in selected_indices:
        if (from_option in st.session_state.image_categories and 
            from_value in st.session_state.image_categories[from_option] and
            idx < len(st.session_state.image_categories[from_option][from_value])):
            
            # 이미지 가져오기
            image = st.session_state.image_categories[from_option][from_value][idx]
            moved_images.append(image)
            
            # 원래 카테고리에서 이미지 제거
            st.session_state.image_categories[from_option][from_value].pop(idx)
            st.session_state.analysis_results[from_option][from_value] -= 1
            
            # 카운트가 0이 되면 카테고리 제거
            if st.session_state.analysis_results[from_option][from_value] == 0:
                del st.session_state.analysis_results[from_option][from_value]
                del st.session_state.image_categories[from_option][from_value]
    
    # 새 카테고리에 이미지들 추가
    if moved_images:
        st.session_state.image_categories[from_option][to_value].extend(moved_images)
        st.session_state.analysis_results[from_option][to_value] = (
            st.session_state.analysis_results[from_option].get(to_value, 0) + len(moved_images)
        )
        st.session_state.needs_rerun = True
        return True
    
    return False

# 수정된 display_images_with_controls 함수
def display_images_with_controls(option, value, images, category):
    """
    체크박스와 이동 컨트롤이 있는 이미지 그리드 표시
    """
    st.markdown(f"""
        <div style="margin-bottom: 5px;">
            <strong>{value}</strong> (Count: {len(images)})
        </div>
    """, unsafe_allow_html=True)
    
    # 이미지 그리드 생성
    cols = st.columns(5)
    selected_indices = []
    
    # 이미지 크기 계산
    image_width = 150
    new_image_width = int(image_width * 1.5)
    
    # 체크박스 상태를 저장할 고유한 키 생성
    checkbox_key = f"checkbox_state_{option}_{value}"
    
    # 체크박스 상태 초기화 또는 업데이트
    if checkbox_key not in st.session_state or len(st.session_state[checkbox_key]) != len(images):
        st.session_state[checkbox_key] = [False] * len(images)
    
    for idx, img in enumerate(images):
        with cols[idx % 5]:
            with st.container():
                st.markdown(
                    """
                    <div style="position: relative; padding: 10px 0 0 10px;">
                        <div style="position: absolute; top: 10px; left: 10px; z-index: 1;">
                    """,
                    unsafe_allow_html=True
                )
                
                # 체크박스 상태 관리
                checkbox_unique_key = f"select_{option}_{value}_{idx}_{hash(str(img))}"
                if st.checkbox("", key=checkbox_unique_key,
                             value=st.session_state[checkbox_key][idx],
                             label_visibility="collapsed"):
                    selected_indices.append(idx)
                    st.session_state[checkbox_key][idx] = True
                else:
                    st.session_state[checkbox_key][idx] = False
                
                st.markdown("</div>", unsafe_allow_html=True)
                
                # 이미지 표시
                img_resized = img.resize((new_image_width, int(new_image_width * img.size[1] / img.size[0])))
                st.image(img_resized, use_column_width=True)
    
    # 컨트롤 버튼들을 하단에 배치
    st.markdown("<div style='margin-top: 15px;'>", unsafe_allow_html=True)
    
    # Move와 Remove 컨트롤을 같은 행에 배치
    col1, col2, col3 = st.columns([4, 1, 1])
    with col1:
        other_options = ["Select Category"] + [opt for opt in analysis_options[category][option] 
                                             if opt != value]
        move_to = st.selectbox(
            "Move to:",
            other_options,
            key=f"move_to_{option}_{value}_{hash(str(images))}",
            label_visibility="collapsed"
        )
    
    with col2:
        if st.button("Move", key=f"move_btn_{option}_{value}_{hash(str(images))}", use_container_width=True):
            if move_to == "Select Category":
                st.warning("Please select a category to move to")
            elif selected_indices:
                if move_selected_images(option, value, move_to, selected_indices):
                    # 체크박스 상태 초기화
                    st.session_state[checkbox_key] = [False] * len(images)
                    st.success(f"Successfully moved {len(selected_indices)} images to {move_to}")
                    st.rerun()
            else:
                st.warning("Please select images to move")
    
    with col3:
        if st.button("Remove", key=f"remove_btn_{option}_{value}_{hash(str(images))}", use_container_width=True):
            if selected_indices:
                for idx in sorted(selected_indices, reverse=True):
                    remove_image(option, value, idx)
                # 체크박스 상태 초기화
                st.session_state[checkbox_key] = [False] * len(images)
                st.success(f"Successfully removed {len(selected_indices)} images")
                st.rerun()
            else:
                st.warning("Please select images to remove")
    
    st.markdown("</div>", unsafe_allow_html=True)

# Modified main app logic (image list part)
def main():
    st.set_page_config(layout="centered")
    
    st.markdown("""
    <style>
    .emoji-title { 
        font-size: 2.4em; 
        text-align: center;
    }
    .emoji { font-size: 0.8em; }
    .results-container { display: flex; flex-wrap: wrap; justify-content: space-between; }
    .chart-container { width: 48%; margin-bottom: 20px; }
    .fullwidth { width: 100vw; position: relative; left: 50%; right: 50%; margin-left: -50vw; margin-right: -50vw; }
    </style>
    """, unsafe_allow_html=True)
    
    st.markdown("<h1 class='emoji-title'>Yakjin Fashion Image Analyzer</h1>", unsafe_allow_html=True)
    
    if authenticate_user():
        st.markdown("<h3><span class='emoji'>👚</span> Step 1: Select Clothing Category</h3>", unsafe_allow_html=True)
        selected_category = st.selectbox(
            "Choose a Clothing Category",
            options=list(analysis_options.keys())
        )
        
        st.markdown("<h3><span class='emoji'>🔍</span> Step 2: Select Analysis Items</h3>", unsafe_allow_html=True)
        selected_options = st.multiselect(
            label="Choose Analysis Items",
            options=list(analysis_options[selected_category].keys()),
            key="analysis_options"
        )
        
        st.markdown("<h3><span class='emoji'>📁</span> Step 3: Upload and Analyze</h3>", unsafe_allow_html=True)
        uploaded_files = st.file_uploader("Choose File(s)", 
                                        type=["xlsx", "xls", "png", "jpg", "jpeg", "jfif", "zip"], 
                                        accept_multiple_files=True)
        
        if uploaded_files and selected_options:  # 파일과 분석 항목이 모두 선택된 경우
            images = []
            for uploaded_file in uploaded_files:
                if uploaded_file.type in ["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "application/vnd.ms-excel"]:
                    try:
                        excel_images = extract_images_from_excel(uploaded_file)
                        if excel_images:
                            images.extend(excel_images[1:])
                    except Exception as e:
                        st.error(f"Excel 파일에서 이미지 추출 중 오류 발생: {str(e)}")
                elif uploaded_file.type.startswith('image/'):
                    try:
                        img = Image.open(uploaded_file)
                        if img.mode != 'RGB':
                            img = img.convert('RGB')
                        images.append(img)
                    except Exception as e:
                        st.error(f"이지 파일 처리 중 오류 발생: {str(e)}")
                elif uploaded_file.type == 'application/zip':
                    for _, img_data in process_zip_file(uploaded_file):
                        try:
                            img = Image.open(io.BytesIO(img_data))
                            if img.mode != 'RGB':
                                img = img.convert('RGB')
                            images.append(img)
                        except Exception as e:
                            st.error(f"ZIP 파일 내 이미지 처리 중 오류 발생: {str(e)}")
            
            if images:
                status_message = st.empty()  # 상태 메시지를 위한 컨테이너 생성
                status_message.text('이미지 처리 중...')
                
                # 이미지 처리
                processed_images = process_images(images)
                
                # 이미지 처리가 끝나면 상태 메시지 업데이트
                status_message.text('이미지 분석 중...')
                
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                aggregated_results = {option: Counter() for option in selected_options}
                image_categories = defaultdict(lambda: defaultdict(list))
                
                total_images = len(processed_images)
                batch_size = 4
                
                batch_data = [(img, selected_category, selected_options) 
                             for img in processed_images]
                
                completed_images = 0
                
                with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
                    for batch in batch_images(batch_data, batch_size):
                        future_to_image = {executor.submit(analyze_image_batch, data): data 
                                         for data in batch}
                        
                        for future in concurrent.futures.as_completed(future_to_image):
                            result = future.result()
                            if result and isinstance(result, dict):
                                image_data = future_to_image[future]
                                image = image_data[0]
                                
                                for option, detected in result.items():
                                    if option in selected_options:
                                        if option == "Details" and isinstance(detected, list):
                                            for detail in detected:
                                                aggregated_results[option][detail] += 1
                                                image_categories[option][detail].append(image)
                                        else:
                                            aggregated_results[option][detected] += 1
                                            image_categories[option][detected].append(image)
                                
                            completed_images += 1
                            progress = completed_images / total_images
                            progress_bar.progress(progress)
                            status_text.text(f"이미지 분석 중: {completed_images}/{total_images}")

                # 분석 완료 후 상태 메시지와 프로그레스 바 삭제
                progress_bar.empty()
                status_text.empty()
                status_message.empty()
                
                # 분석 결과를 세션 상태에 저장
                st.session_state.analysis_results = aggregated_results
                st.session_state.image_categories = image_categories
                
                # 결과 표시
                st.markdown("<div class='fullwidth'>", unsafe_allow_html=True)
                st.markdown("<hr>", unsafe_allow_html=True)
                st.markdown("<h2 style='text-align: center;'>📊 Analysis Results</h2>", unsafe_allow_html=True)
                st.markdown("<div class='results-container'>", unsafe_allow_html=True)
                
                # 각 분석 항목에 대한 고유한 색상 세트 생성
                color_sets = list(generate_unique_color_sets(len(selected_options), 12))  # 12는 최 카테고리 
                
                for i, (option, results) in enumerate(aggregated_results.items()):
                    if results:
                        st.markdown(f"<div class='chart-container'>", unsafe_allow_html=True)
                        fig = create_donut_chart(results, option, color_sets[i])
                        st.plotly_chart(fig, use_container_width=True)
                        
                        with st.expander(f"{option} Details"):
                            for value, count in results.items():
                                st.markdown(f"**{value}** (Count: {count})", unsafe_allow_html=True)
                                if option in image_categories and value in image_categories[option]:
                                    images = image_categories[option][value]
                                    cols = st.columns(5)
                                    for j, img in enumerate(images):
                                        with cols[j % 5]:
                                            st.image(img, use_column_width=True)
                                        if (j + 1) % 5 == 0:
                                            st.write("")
                                else:
                                    st.write("No Matching Images Found.")
                                st.write("---")
                        st.markdown("</div>", unsafe_allow_html=True)
                    else:
                        st.write(f"No Data Available for {option}.")
                    
                    # 2개의 차트마다 새 줄 시작
                    if (i + 1) % 2 == 0:
                        st.markdown("</div><div class='results-container'>", unsafe_allow_html=True)
                
                st.markdown("</div></div>", unsafe_allow_html=True)
            else:
                st.markdown("<p><span class='emoji'>⚠️</span> No Images Found in the Uploaded File.</p>", unsafe_allow_html=True)
    else:
        st.info("로그인이 필요합니다. 위의 인증 정보를 입력해주세요.")

if __name__ == "__main__":
    main()

# CSS for Streamlit theme settings
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
        font-size: 25px !important;
        font-weight: bold;
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

# CSS 스타일 업데이트
st.markdown("""
<style>
    /* 기존 스타일 유지 */
    .stMultiSelect [data-baseweb="tag"] {
        background-color: #007AFF !important;
    }
    // ... (기존 스타일 유지)

    /* 새로운 이미지 컨트롤 관련 스타일 */
    .stCheckbox {
        margin: 0;
        padding: 0;
    }
    
    .stButton > button {
        height: 38px;
        margin-top: 0 !important;
        border-radius: 4px;
        background-color: #f0f2f6 !important;
        color: #000000 !important;
    }
    
    /* Move와 Remove 버튼 스타일 */
    [data-testid="stButton"] button[key*="move_btn"] {
        background-color: #007AFF !important;
        color: white !important;
    }
    
    [data-testid="stButton"] button[key*="remove_btn"] {
        background-color: #FF3B30 !important;
        color: white !important;
    }
    
    /* 이미지 컨테이너 스타일 */
    .stImage {
        padding: 5px;
    }
    
    /* 마진 관련 스타일 */
    .stMarkdown {
        margin-bottom: 0 !important;
    }
    
    .row-widget {
        margin-top: 0 !important;
        margin-bottom: 0 !important;
    }
    
    /* View fullscreen 버튼 숨기기 */
    button[title="View fullscreen"] {
        display: none !important;
    }
</style>
""", unsafe_allow_html=True)

