import streamlit as st
from openai import OpenAI
import base64
from PIL import Image
import io
import zipfile
import openpyxl
from openpyxl_image_loader import SheetImageLoader
import cv2
import numpy as np
from collections import Counter, defaultdict

# OpenAI API 키 설정 (Streamlit Cloud의 secrets에서 가져옴)
client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

# 분석 항목 정의 (기존 코드에서 가져옴)
analysis_options = {
    "Top": {
        "Fit": ["Slim Fit", "Regular Fit", "Loose Fit", "Oversized"],
        "Neckline": ["Crew Neck", "V-Neck", "Scoop Neck", "Turtleneck", "Cowl Neck"],
        # ... 다른 옵션들 ...
    },
    # ... 다른 카테고리들 ...
}

# 이미지 추출 함수
def extract_images(uploaded_file):
    images = []
    if uploaded_file.type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
        # 엑셀 파일 처리
        wb = openpyxl.load_workbook(io.BytesIO(uploaded_file.getvalue()))
        sheet = wb.active
        image_loader = SheetImageLoader(sheet)
        for row in sheet.iter_rows():
            for cell in row:
                if image_loader.image_in(cell.coordinate):
                    image = image_loader.get(cell.coordinate)
                    images.append(image)
    elif uploaded_file.type.startswith('image/'):
        # 단일 이미지 파일 처리
        image = Image.open(io.BytesIO(uploaded_file.getvalue()))
        images.append(image)
    elif uploaded_file.type == "application/zip":
        # ZIP 파일 처리
        with zipfile.ZipFile(io.BytesIO(uploaded_file.getvalue()), 'r') as zip_ref:
            for file_name in zip_ref.namelist():
                if file_name.lower().endswith(('.png', '.jpg', '.jpeg')):
                    with zip_ref.open(file_name) as file:
                        image = Image.open(io.BytesIO(file.read()))
                        images.append(image)
    return images

# 이미지 처리 함수
def process_images(images):
    processed_images = []
    for img in images:
        # 이미지 전처리 로직 (예: 크기 조정, 품질 향상 등)
        img_array = np.array(img)
        processed_img = cv2.resize(img_array, (224, 224))  # 예시: 224x224로 크기 조정
        processed_images.append(Image.fromarray(processed_img))
    return processed_images

# 이미지 분석 함수
def analyze_image(image, category, options):
    # 이미지를 base64로 인코딩
    buffered = io.BytesIO()
    image.save(buffered, format="PNG")
    base64_image = base64.b64encode(buffered.getvalue()).decode('utf-8')
    
    # OpenAI API에 요청할 프롬프트 생성
    prompt = f"이미지에 있는 {category} 의류 아이템을 분석하고 다음 측면에 대한 정보를 제공해주세요:\n\n"
    for option in options:
        prompt += f"{option}: {', '.join(analysis_options[category][option])}\n"
    prompt += "\n결과를 JSON 형식으로 제공해주세요."

    try:
        response = client.chat.completions.create(
            model="gpt-4-vision-preview",
            messages=[
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt},
                        {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{base64_image}"}}
                    ]
                }
            ],
            max_tokens=300
        )
        return response.choices[0].message.content
    except Exception as e:
        st.error(f"이미지 분석 중 오류 발생: {e}")
        return None

def perform_analysis(images, category, options):
    results = defaultdict(Counter)
    for image in images:
        analysis_result = analyze_image(image, category, options)
        if analysis_result:
            try:
                result_dict = eval(analysis_result)  # JSON 문자열을 딕셔너리로 변환
                for key, value in result_dict.items():
                    results[key][value] += 1
            except:
                st.error("분석 결과 처리 중 오류가 발생했습니다.")
    return results

def display_results(results):
    for option, counts in results.items():
        st.subheader(f"{option} 분석 결과")
        for value, count in counts.items():
            st.write(f"{value}: {count}")

def main():
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
            st.markdown("<h3><span class='emoji'>🖼️</span> 4단계: 이미지 처리</h3>", unsafe_allow_html=True)
            
            images = extract_images(uploaded_file)
            
            if images:
                with st.spinner('이미지 처리 중...'):
                    processed_images = process_images(images)
                
                st.success(f"{len(processed_images)}개의 이미지가 처리되었습니다.")
                
                if st.button("🚀 5단계: 분석 시작"):
                    if not selected_options:
                        st.markdown("<p><span class='emoji'>⚠️</span> 분석할 항목을 하나 이상 선택해주세요.</p>", unsafe_allow_html=True)
                    else:
                        results = perform_analysis(processed_images, selected_category, selected_options)
                        display_results(results)
            else:
                st.error("업로드된 파일에서 이미지를 찾을 수 없습니다.")

# 사용자 인증 함수 (기존 코드에서 가져옴)
def authenticate_user():
    # 인증 로직 구현
    return True  # 예시로 항상 True 반환

if __name__ == "__main__":
    main()
