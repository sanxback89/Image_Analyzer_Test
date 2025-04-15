import streamlit as st
import pandas as pd
import openai
import io
import fitz  # PyMuPDF
import base64
from PIL import Image, ImageOps
import json
import re
from io import BytesIO
import openpyxl
from collections import Counter
import time

# =============================================================================
# 2023-04-15 기능 개선 사항 (다중 분석 시스템 강화)
# - 다중 분석 시 데이터 완전성 평가, 결과 합성, 인터페이스 개선 등이 포함됨.
#
# 2023-04-17 UPC 매칭 및 결과 선택 로직 개선
# - Excel 데이터와 비교하여 UPC 일치도 평가 및 보정
# =============================================================================

# 페이지 레이아웃 설정 - 가로 넓이 확장
st.set_page_config(
    layout="centered",  # 배포 환경에 따라 필요시 "wide"로 변경 가능
    page_title="YAKJIN",
    page_icon="🔍"
)

# CSS를 사용하여 컨테이너 넓이 및 테이블 스타일 조정
st.markdown("""
<style>
    .reportview-container .main .block-container {
        max-width: 800px;
        padding-top: 2rem;
        padding-bottom: 2rem;
        padding-left: 10rem;
        padding-right: 10rem;
    }
    .stDataFrame, .dataframe, .stTable, div[data-testid="stHorizontalBlock"] {
        width: 100%;
    }
    .stDataFrame > div {
        overflow-x: auto !important;
    }
    .dataframe td, .dataframe th {
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        max-width: 350px;
    }
    [data-testid="stDataFrame"] > div {
        width: 100% !important;
        max-width: 100% !important;
    }
</style>
""", unsafe_allow_html=True)

# =============================================================================
# OpenAI API 키 설정 (수정됨)
# 기존에 하드코딩된 API 키 대신 st.secrets를 사용하여 보안을 강화함.
try:
    openai.api_key = st.secrets["OPENAI_API_KEY"]
except KeyError:
    st.error("OpenAI API 키가 설정되지 않았습니다. Streamlit Cloud의 st.secrets에 API 키를 추가하세요.")

# =============================================================================
# 사이드바 네비게이션
st.sidebar.title("분석 유형 선택")
analysis_type = st.sidebar.radio(
    "분석 유형을 선택하세요:",
    ["Label Analysis", "Size Strip Analysis"]
)

# =============================================================================
# 사이즈 정렬용 딕셔너리
size_order_dict = {
    "XS": 0, "S": 1, "M": 2, "L": 3, "XL": 4, "XXL": 5, "XXXL": 6,
    "XST": 10, "ST": 11, "MT": 12, "LT": 13, "XLT": 14, "XXLT": 15, "3XLT": 16,
    "2XB": 20, "3XB": 21, "4XB": 22, "5XB": 23
}

# =============================================================================
# PDF 관련 함수
def is_text_pdf(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    for page in doc:
        if page.get_text().strip():
            return True
    return False

def extract_text_from_pdf(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    texts = [page.get_text() for page in doc]
    doc.close()
    return "\n".join(texts)

def convert_pdf_to_images(pdf_bytes, dpi=300):
    images = []
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    for page in doc:
        pix = page.get_pixmap(matrix=fitz.Matrix(dpi/72, dpi/72))
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        img = ImageOps.grayscale(img)
        img = ImageOps.autocontrast(img)
        img = img.convert("RGB")
        buf = io.BytesIO()
        img.save(buf, format='PNG')
        img_base64 = base64.b64encode(buf.getvalue()).decode("utf-8")
        images.append(img_base64)
    doc.close()
    return images

# =============================================================================
# GPT 텍스트 분석 함수 (라벨 포맷 / 사이즈 스트립 포맷)
def analyze_label_text_with_gpt(text):
    prompt = f"""
이 텍스트는 의류 라벨 PDF에서 추출된 전체 텍스트입니다.
다음 정보를 표로 추출해주세요:
- Style Number | Size | Color | UPC (12-digit)
- 사이즈는 XS-S-M-L-XL-XXL 순으로 정렬
- 값은 모두 대문자로 표시

중요 정보:
1) 스타일 넘버는 'WS' 또는 'MS'로 시작 (예: WS5FK004RS3, WS5FK004RH2). 단, RPI7은 내부 코드입니다.
2) UPC 코드는 반드시 12자리이며, 항상 400으로 시작합니다. (4000이면 400으로, 11자리이면 앞에 4 추가, 8자리이면 앞에 400 붙이고 뒤에 00 추가)
3) 사이즈에 P 또는 PETITE가 붙어있다면 P 제거.
추가 정보: Fabric, Care Content, Factory Code (5~6자리, 보통 날짜 근처)
텍스트:
{text}
결과는 아래 표 형식:
Style Number | Size | Color | UPC Code
그리고 추가 정보를 표 아래에 표시
"""
    response = openai.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": "당신은 PDF 라벨 정보를 표로 정리하는 전문가입니다. UPC 코드는 반드시 400으로 시작해야 합니다."},
            {"role": "user", "content": prompt}
        ]
    )
    return response.choices[0].message.content

def analyze_size_strip_text_with_gpt(text):
    prompt = f"""
이 텍스트는 의류 사이즈 스트립 PDF에서 추출된 전체 텍스트입니다.
다음 정보를 표로 추출해주세요:
- Style Number | Size | Color | UPC (12-digit)
- 사이즈는 XS-S-M-L-XL-XXL 순으로 정렬
- 값은 모두 대문자로 표시
텍스트:
{text}
결과는 아래 표 형식으로:
Style Number | Size | Color | UPC Code
"""
    response = openai.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": "당신은 PDF 사이즈 스트립 정보를 표로 정리하는 전문가입니다. UPC 코드는 반드시 400으로 시작해야 합니다."},
            {"role": "user", "content": prompt}
        ]
    )
    return response.choices[0].message.content

# =============================================================================
# GPT Vision 이미지 분석 함수 (라벨 / 사이즈 스트립)
def analyze_label_image_with_vision(image_base64):
    vision_prompt = (
        "의류 라벨 이미지에서 Style Number, Size, Color, 정확한 12자리 UPC를 추출해주세요.\n\n"
        "중요: 스타일 넘버는 'WS' 또는 'MS'로 시작, UPC 코드는 반드시 400으로 시작 (4000이면 400으로 수정)."
        "\nFabric, Care Content, Factory Code도 함께 추출해주세요."
    )
    response = openai.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": "의류 라벨 이미지를 분석하는 전문가입니다. 스타일 넘버는 'WS' 또는 'MS'로 시작하고, UPC 코드는 반드시 400으로 시작해야 합니다."},
            {"role": "user", "content": [
                {"type": "text", "text": vision_prompt},
                {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{image_base64}"}}
            ]}
        ]
    )
    return response.choices[0].message.content

def analyze_size_strip_image_with_vision(image_base64):
    vision_prompt = (
        "의류 사이즈 스트립 이미지에서 Style Number, Size, Color, 정확한 12자리 UPC를 추출해주세요.\n\n"
        "중요: 스타일 넘버는 'WS' 또는 'MS'로 시작, UPC 코드는 반드시 400으로 시작."
    )
    response = openai.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": "의류 사이즈 스트립 이미지를 분석하는 전문가입니다. 스타일 넘버는 'WS' 또는 'MS'로 시작하고, UPC 코드는 반드시 400으로 시작해야 합니다."},
            {"role": "user", "content": [
                {"type": "text", "text": vision_prompt},
                {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{image_base64}"}}
            ]}
        ]
    )
    return response.choices[0].message.content

# =============================================================================
# 스타일 넘버, 사이즈, UPC 코드 관련 함수
def validate_style_number(style_number):
    style = style_number.strip().upper()
    if style.startswith('WS') or style.startswith('MS'):
        return style
    if style == 'RPI7' or style.startswith('RP'):
        return None
    return style

def normalize_size(size_text):
    size = size_text.strip().upper()
    if size.endswith('P') and len(size) > 1:
        size = size[:-1]
    return size

def normalize_upc(upc_text):
    digits = ''.join(filter(str.isdigit, upc_text))
    if digits.startswith('4000'):
        digits = '400' + digits[4:]
    if digits.startswith('400021041'):
        digits = '400210419' + digits[9:]
    if len(digits) > 12:
        return digits[:12]
    elif len(digits) < 12:
        if len(digits) == 11 and digits.startswith('400'):
            check_digit = calculate_upc_check_digit(digits)
            return digits + check_digit
        elif len(digits) == 8:
            digits = '400' + digits + '0'
        else:
            digits = '400' + digits + '0' * (12 - len(digits) - 3)
    if not digits.startswith('400'):
        digits = '400' + digits[3:]
    if len(digits) != 12:
        if len(digits) == 11:
            check_digit = calculate_upc_check_digit(digits)
            digits = digits + check_digit
        else:
            digits = digits + '0' * (12 - len(digits))
    return digits

def calculate_upc_check_digit(upc_11_digits):
    if len(upc_11_digits) != 11:
        return "0"
    odd_sum = sum(int(upc_11_digits[i]) for i in range(0, 11, 2))
    even_sum = sum(int(upc_11_digits[i]) for i in range(1, 11, 2))
    total = odd_sum * 3 + even_sum
    check_digit = (10 - (total % 10)) % 10
    return str(check_digit)

# =============================================================================
# Excel/CSV 데이터 처리 함수
def process_excel_data(excel_file):
    try:
        if excel_file.name.endswith('.csv'):
            df = pd.read_csv(excel_file)
        else:
            excel_bytes = excel_file.read()
            excel_file.seek(0)
            wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True)
            sheet = wb.active
            has_filter = False
            if sheet.auto_filter.ref:
                has_filter = True
                st.info(f"'{excel_file.name}' 파일에 필터 설정이 감지되었습니다. 필터를 해제하고 분석합니다.")
            df = pd.read_excel(excel_file, header=None)
            df = df.iloc[1:].reset_index(drop=True)
            df.columns = df.iloc[0]
            df = df.iloc[1:].reset_index(drop=True)
            if has_filter:
                orig_row_count = len(df)
                df_cleaned = df.dropna(how='all')
                cleaned_row_count = len(df_cleaned)
                if orig_row_count > cleaned_row_count:
                    st.warning(f"필터로 인해 {orig_row_count - cleaned_row_count}개의 빈 행이 제거되었습니다.")
                df = df_cleaned.reset_index(drop=True)
        for col in df.columns:
            if df[col].dtype == object:
                df[col] = df[col].astype(str).str.strip()
        required_columns = ["Vendor Style Number", "UPCs", "Display Color Description", "Kohls Size Description"]
        optional_columns = ["Vendor Color Description"]
        df.columns = df.columns.str.strip()
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            st.error(f"엑셀 파일 '{excel_file.name}'에 필수 컬럼이 없습니다: {', '.join(missing_columns)}")
            return None
        has_vendor_color = "Vendor Color Description" in df.columns
        select_columns = required_columns.copy()
        if has_vendor_color:
            select_columns.append("Vendor Color Description")
        df = df[select_columns].copy()
        df["UPCs"] = df["UPCs"].astype(str).str.replace("^000", "", regex=True)
        size_map = {"x small": "XS", "small": "S", "medium": "M", "large": "L", "x large": "XL", "xx large": "XXL"}
        def simplify_size(size_text):
            if not isinstance(size_text, str):
                return size_text
            size_text = size_text.lower().strip()
            if " tall" in size_text:
                size_part = size_text.replace(" tall", "")
                base_size = size_map.get(size_part, size_part.upper())
                return f"{base_size}T"
            size_text = re.sub(r'\s*petite\s*', '', size_text)
            if len(size_text) > 1 and size_text.endswith('p'):
                size_text = size_text[:-1]
            if re.match(r'^[2-9]xb$', size_text):
                return size_text.upper()
            return size_map.get(size_text, size_text.upper())
        df["Kohls Size Description"] = df["Kohls Size Description"].apply(simplify_size)
        df["Display Color Description"] = df["Display Color Description"].apply(lambda x: re.sub(r'\s*PETITE\s*|\s*PETIT\s*', '', x, flags=re.IGNORECASE).strip() if isinstance(x, str) else x)
        if has_vendor_color:
            df["Vendor Color Description"] = df["Vendor Color Description"].apply(lambda x: re.sub(r'\s*PETITE\s*|\s*PETIT\s*', '', x, flags=re.IGNORECASE).strip() if isinstance(x, str) else x)
        for col in df.columns:
            if df[col].dtype == object:
                df[col] = df[col].str.upper()
        size_order_dict_extended = {
            "XS": 0, "S": 1, "M": 2, "L": 3, "XL": 4, "XXL": 5, "XXXL": 6,
            "XST": 10, "ST": 11, "MT": 12, "LT": 13, "XLT": 14, "XXLT": 15, "3XLT": 16,
            "2XB": 20, "3XB": 21, "4XB": 22, "5XB": 23
        }
        df["Size Order"] = df["Kohls Size Description"].map(lambda x: size_order_dict_extended.get(x, 99))
        df = df.sort_values(["Vendor Style Number", "Display Color Description", "Size Order"])
        df = df.drop("Size Order", axis=1)
        rename_dict = {
            "Vendor Style Number": "Style Number",
            "UPCs": "UPC Code",
            "Display Color Description": "Display Color",
            "Kohls Size Description": "Size"
        }
        if has_vendor_color:
            rename_dict["Vendor Color Description"] = "Vendor Color"
        df = df.rename(columns=rename_dict)
        df["Excel Source"] = excel_file.name
        st.info(f"'{excel_file.name}' 파일에서 {len(df)}개 레코드 정제 완료")
        return df
    except Exception as e:
        st.error(f"Excel/CSV 파일 '{excel_file.name}' 처리 중 오류: {str(e)}")
        import traceback
        st.error(f"스택 트레이스: {traceback.format_exc()}")
        return None

# =============================================================================
# 색상명 정규화 함수
def normalize_color_name(color_text):
    if not isinstance(color_text, str):
        return ""
    normalized = color_text.upper().replace(" ", "")
    normalized = re.sub(r'[-&\.,/\(\)\[\]]', '', normalized)
    normalized = normalized.replace("GREY", "GRAY")
    normalized = normalized.replace("MELANGE", "MLG")
    return normalized

# =============================================================================
# 일치 여부 확인 함수
def check_match(pdf_df, all_excel_df):
    pdf_df["Normalized Color"] = pdf_df["Color"].apply(normalize_color_name)
    all_excel_df["Normalized Display Color"] = all_excel_df["Display Color"].apply(normalize_color_name)
    if "Vendor Color" in all_excel_df.columns:
        all_excel_df["Normalized Vendor Color"] = all_excel_df["Vendor Color"].apply(normalize_color_name)
    pdf_df["Match"] = "❌"
    pdf_df["Match Detail"] = ""
    for idx, pdf_row in pdf_df.iterrows():
        style_size_match = (all_excel_df["Style Number"] == pdf_row["Style Number"]) & (all_excel_df["Size"] == pdf_row["Size"])
        original_upc = pdf_row["Original UPC"] if "Original UPC" in pdf_row else pdf_row["UPC Code"]
        upc_match = all_excel_df["UPC Code"] == original_upc
        corrected_upc_match = all_excel_df["UPC Code"] == pdf_row["UPC Code"]
        if "Normalized Vendor Color" in all_excel_df.columns:
            color_match = (all_excel_df["Normalized Display Color"] == pdf_row["Normalized Color"]) | (all_excel_df["Normalized Vendor Color"] == pdf_row["Normalized Color"])
        else:
            color_match = (all_excel_df["Normalized Display Color"] == pdf_row["Normalized Color"])
        if not any(color_match) and any(upc_match):
            pdf_excel_color = pdf_row["Color"]
            excel_colors = all_excel_df[upc_match]["Display Color"].iloc[0] if not all_excel_df[upc_match].empty else "없음"
            norm_pdf_color = pdf_row["Normalized Color"]
            norm_excel_color = all_excel_df[upc_match]["Normalized Display Color"].iloc[0] if not all_excel_df[upc_match].empty else "없음"
            pdf_df.at[idx, "Color Debug"] = f"PDF: '{pdf_excel_color}' ({norm_pdf_color}) vs Excel: '{excel_colors}' ({norm_excel_color})"
        full_match = style_size_match & color_match & upc_match
        corrected_match = style_size_match & color_match & corrected_upc_match & ~upc_match
        potential_matches = all_excel_df[style_size_match & color_match & ~upc_match & ~corrected_upc_match]
        upc_match_color_mismatch = all_excel_df[style_size_match & ~color_match & upc_match]
        matching_excel_rows = all_excel_df[full_match]
        upc_exact_match = all_excel_df[all_excel_df["UPC Code"] == original_upc]
        if not matching_excel_rows.empty:
            pdf_df.at[idx, "Match"] = "✅"
            pdf_df.at[idx, "Match Detail"] = "원본 UPC 정확히 일치"
            pdf_df.at[idx, "Excel Source"] = ", ".join(sorted(set(matching_excel_rows["Excel Source"])))
            matched_colors = []
            if any(matching_excel_rows["Display Color"] == pdf_row["Color"]):
                matched_colors.append("Display")
            if "Vendor Color" in matching_excel_rows.columns and any(matching_excel_rows["Vendor Color"] == pdf_row["Color"]):
                matched_colors.append("Vendor")
            pdf_df.at[idx, "Matched Color Type"] = ", ".join(matched_colors)
        elif any(corrected_match):
            corrected_rows = all_excel_df[corrected_match]
            pdf_df.at[idx, "Match"] = "⚠️"
            pdf_df.at[idx, "Match Detail"] = "교정된 UPC로 일치"
            pdf_df.at[idx, "Excel Source"] = ", ".join(sorted(set(corrected_rows["Excel Source"])))
            matched_colors = []
            if any(corrected_rows["Display Color"] == pdf_row["Color"]):
                matched_colors.append("Display")
            if "Vendor Color" in corrected_rows.columns and any(corrected_rows["Vendor Color"] == pdf_row["Color"]):
                matched_colors.append("Vendor")
            pdf_df.at[idx, "Matched Color Type"] = ", ".join(matched_colors)
        elif not upc_exact_match.empty:
            excel_normalized_color = upc_exact_match.iloc[0]["Normalized Display Color"]
            pdf_normalized_color = pdf_row["Normalized Color"]
            excel_original_color = upc_exact_match.iloc[0]["Display Color"]
            pdf_original_color = pdf_row["Color"]
            if pdf_normalized_color == excel_normalized_color:
                pdf_df.at[idx, "Match"] = "✅"
                pdf_df.at[idx, "Match Detail"] = f"UPC 일치, 색상 실질적 일치 (형식만 다름: {pdf_original_color} vs {excel_original_color})"
                pdf_df.at[idx, "Excel Source"] = ", ".join(sorted(set(upc_exact_match["Excel Source"])))
            elif pdf_original_color.upper() == excel_original_color.upper():
                pdf_df.at[idx, "Match"] = "✅"
                pdf_df.at[idx, "Match Detail"] = f"UPC 일치, 색상 대소문자만 다름: {pdf_original_color} vs {excel_original_color}"
                pdf_df.at[idx, "Excel Source"] = ", ".join(sorted(set(upc_exact_match["Excel Source"])))
            elif ('HEATHER' in pdf_original_color.upper() and 'HTHR' in excel_original_color.upper()) or ('HTHR' in pdf_original_color.upper() and 'HEATHER' in excel_original_color.upper()):
                pdf_df.at[idx, "Match"] = "✅"
                pdf_df.at[idx, "Match Detail"] = f"UPC 일치, HEATHER/HTHR 형식만 다름: {pdf_original_color} vs {excel_original_color}"
                pdf_df.at[idx, "Excel Source"] = ", ".join(sorted(set(upc_exact_match["Excel Source"])))
            elif ('STRIPE' in pdf_original_color.upper() and 'STRIPE' in excel_original_color.upper()):
                pdf_no_space = pdf_original_color.upper().replace(" ", "")
                excel_no_space = excel_original_color.upper().replace(" ", "")
                if pdf_no_space == excel_no_space:
                    pdf_df.at[idx, "Match"] = "✅"
                    pdf_df.at[idx, "Match Detail"] = f"UPC 일치, 공백 차이만 있음: {pdf_original_color} vs {excel_original_color}"
                    pdf_df.at[idx, "Excel Source"] = ", ".join(sorted(set(upc_exact_match["Excel Source"])))
                else:
                    pdf_df.at[idx, "Match"] = "❌"
                    pdf_df.at[idx, "Match Detail"] = f"색상 불일치: PDF({pdf_original_color}) ≠ Excel({excel_original_color})"
                    pdf_df.at[idx, "Excel Source"] = ", ".join(sorted(set(upc_exact_match["Excel Source"])))
                    pdf_df.at[idx, "Excel Color"] = ", ".join(sorted(set(upc_exact_match["Display Color"])))
                    pdf_df.at[idx, "Excel Style"] = ", ".join(sorted(set(upc_exact_match["Style Number"])))
                    pdf_df.at[idx, "Excel Size"] = ", ".join(sorted(set(upc_exact_match["Size"])))
                    pdf_df.at[idx, "Excel UPC"] = upc_exact_match.iloc[0]["UPC Code"]
            else:
                pdf_df.at[idx, "Match"] = "❌"
                pdf_df.at[idx, "Match Detail"] = f"색상 불일치: PDF({pdf_original_color}) ≠ Excel({excel_original_color})"
                pdf_df.at[idx, "Excel Source"] = ", ".join(sorted(set(upc_exact_match["Excel Source"])))
                pdf_df.at[idx, "Excel Color"] = ", ".join(sorted(set(upc_exact_match["Display Color"])))
                pdf_df.at[idx, "Excel Style"] = ", ".join(sorted(set(upc_exact_match["Style Number"])))
                pdf_df.at[idx, "Excel Size"] = ", ".join(sorted(set(upc_exact_match["Size"])))
                pdf_df.at[idx, "Excel UPC"] = upc_exact_match.iloc[0]["UPC Code"]
        elif not potential_matches.empty:
            excel_upc = potential_matches.iloc[0]["UPC Code"]
            if is_last_digits_different(original_upc, excel_upc, max_diff=2):
                pdf_df.at[idx, "Match"] = "❌"
                pdf_df.at[idx, "Match Detail"] = f"UPC 마지막 자리 차이: {original_upc[-2:]} vs {excel_upc[-2:]}"
                pdf_df.at[idx, "Excel UPC"] = excel_upc
                pdf_df.at[idx, "Excel Source"] = ", ".join(sorted(set(potential_matches["Excel Source"])))
            else:
                pdf_df.at[idx, "Match"] = "❌"
                if is_completely_different_upc(original_upc, excel_upc):
                    pdf_df.at[idx, "Match Detail"] = "완전히 다른 UPC 패턴"
                elif has_significant_middle_difference(original_upc, excel_upc):
                    pdf_df.at[idx, "Match Detail"] = "UPC 중간 부분 차이 심각"
                else:
                    pdf_df.at[idx, "Match Detail"] = f"UPC 불일치: {original_upc} ≠ {excel_upc}"
                pdf_df.at[idx, "Excel UPC"] = excel_upc
                pdf_df.at[idx, "Excel Source"] = ", ".join(sorted(set(potential_matches["Excel Source"])))
    return pdf_df

# =============================================================================
# Fabric, Care Content, Factory Code 추출 함수
def extract_fabric_care_from_gpt_response(gpt_output):
    fabric_info = None
    care_info = None
    factory_code = None
    fabric_matches = re.findall(r'(?i)fabric:\s*(.*?)(?:\n|$|care)', gpt_output)
    if fabric_matches:
        fabric_info = fabric_matches[0].strip()
        if fabric_info.lower() in ['(원단 정보)', 'none', '']:
            fabric_info = None
    care_matches = re.findall(r'(?i)care content:\s*(.*?)(?:\n|$)', gpt_output)
    if care_matches:
        care_info = care_matches[0].strip()
        if care_info.lower() in ['(관리 방법)', 'none', '']:
            care_info = None
    factory_code_matches = re.findall(r'(?i)factory code:\s*(\d{5,6})(?:\n|$)', gpt_output)
    if factory_code_matches:
        factory_code = factory_code_matches[0].strip()
    else:
        date_factory_matches = re.findall(r'\d{2}/\d{2}\s+(\d{5,6})', gpt_output)
        if date_factory_matches:
            factory_code = date_factory_matches[0].strip()
    return fabric_info, care_info, factory_code

# =============================================================================
# UPC 교정 함수 (Excel 데이터 활용)
def correct_upc_with_excel(pdf_upc, excel_df, style, size, color):
    matching_rows = excel_df[(excel_df["Style Number"] == style) & 
                             (excel_df["Size"] == size) &
                             ((all_excel_df["Display Color"] == color) | 
                              (all_excel_df["Vendor Color"] == color if "Vendor Color" in excel_df.columns else False))]
    if matching_rows.empty:
        return pdf_upc, "매칭되는 Excel 데이터 없음"
    excel_upc = matching_rows.iloc[0]["UPC Code"]
    if pdf_upc == excel_upc:
        return pdf_upc, "정확히 일치"
    if len(pdf_upc) != len(excel_upc):
        return excel_upc, f"길이 불일치 (PDF: {len(pdf_upc)}자리, Excel: {len(excel_upc)}자리)"
    diff_positions = [i for i in range(len(pdf_upc)) if pdf_upc[i] != excel_upc[i]]
    if pdf_upc.startswith('400021041') and excel_upc.startswith('400210419'):
        return excel_upc, "확인된 오류 패턴 (400021041xxx → 400210419xxx)"
    if len(diff_positions) <= 3:
        return excel_upc, f"{len(diff_positions)}개 자리 불일치"
    return pdf_upc, f"심각한 불일치 ({len(diff_positions)}개 자리)"

# =============================================================================
# UPC 유효성 검사 함수
def validate_upc_code(upc):
    if not upc.isdigit():
        return False, "UPC 코드는 숫자만 포함해야 함"
    if len(upc) != 12:
        return False, f"UPC 코드는 12자리여야 함 (현재: {len(upc)}자리)"
    if not (upc.startswith('400') or upc.startswith('4000')):
        return False, "UPC 코드는 '400'으로 시작해야 함"
    if upc.count('0') >= 9:
        return False, "UPC 코드에 0이 너무 많음 (의심스러운 패턴)"
    check_digit = calculate_upc_check_digit(upc[:11])
    if check_digit != upc[11]:
        return True, f"체크섬 불일치 경고: 계산된 체크섬({check_digit}) ≠ 현재({upc[11]})"
    return True, "유효한 UPC 코드"

# =============================================================================
# OpenAI 분석 결과 캐싱 함수
analysis_cache = {}
def get_cached_analysis(cache_key, analysis_func, *args, **kwargs):
    if cache_key in analysis_cache:
        return analysis_cache[cache_key]
    result = analysis_func(*args, **kwargs)
    analysis_cache[cache_key] = result
    return result

# =============================================================================
# 다중 호출을 통한 투표 시스템 (UPC 추출 안정성 향상)
def multi_vision_analysis(image_base64, is_label=True, attempts=2, all_excel_df=None):
    analysis_func = analyze_label_image_with_vision if is_label else analyze_size_strip_image_with_vision
    all_results = []
    all_styles = []
    all_sizes = []
    all_colors = []
    all_upcs = []
    confidence_scores = []
    completeness_scores = []
    individual_results = []
    result1 = get_cached_analysis(f"vision_attempt1_{is_label}", analysis_func, image_base64)
    all_results.append(result1)
    style1, size1, color1, upc1 = extract_data_from_result(result1)
    all_styles.append(style1)
    all_sizes.append(size1)
    all_colors.append(color1)
    if upc1:
        all_upcs.append(upc1)
    confidence1 = calculate_confidence_score(result1, style1, size1, color1, upc1)
    confidence_scores.append(confidence1)
    completeness1 = calculate_completeness_score(result1)
    completeness_scores.append(completeness1)
    result_info1 = {
        "시도": "첫 번째 시도",
        "신뢰도": f"{confidence1:.2f}",
        "완전성": f"{completeness1:.2f}",
        "결과": result1
    }
    individual_results.append(result_info1)
    if confidence1 < 0.99 or completeness1 < 0.99:
        if attempts >= 2:
            result2 = get_cached_analysis(f"vision_attempt2_{is_label}", analysis_func, image_base64)
            all_results.append(result2)
            style2, size2, color2, upc2 = extract_data_from_result(result2)
            all_styles.append(style2)
            all_sizes.append(size2)
            all_colors.append(color2)
            if upc2:
                all_upcs.append(upc2)
            confidence2 = calculate_confidence_score(result2, style2, size2, color2, upc2)
            confidence_scores.append(confidence2)
            completeness2 = calculate_completeness_score(result2)
            completeness_scores.append(completeness2)
            result_info2 = {
                "시도": "두 번째 시도",
                "신뢰도": f"{confidence2:.2f}",
                "완전성": f"{completeness2:.2f}",
                "결과": result2
            }
            individual_results.append(result_info2)
        if attempts >= 3:
            for i in range(3, attempts + 1):
                result_i = get_cached_analysis(f"vision_attempt{i}_{is_label}", analysis_func, image_base64)
                all_results.append(result_i)
                style_i, size_i, color_i, upc_i = extract_data_from_result(result_i)
                all_styles.append(style_i)
                all_sizes.append(size_i)
                all_colors.append(color_i)
                if upc_i:
                    all_upcs.append(upc_i)
                confidence_i = calculate_confidence_score(result_i, style_i, size_i, color_i, upc_i)
                confidence_scores.append(confidence_i)
                completeness_i = calculate_completeness_score(result_i)
                completeness_scores.append(completeness_i)
                result_info_i = {
                    "시도": f"{i}번째 시도",
                    "신뢰도": f"{confidence_i:.2f}",
                    "완전성": f"{completeness_i:.2f}",
                    "결과": result_i
                }
                individual_results.append(result_info_i)
    consistency_score = check_result_consistency(all_results)
    if all_excel_df is not None and not all_excel_df.empty:
        excel_match_scores = [0] * len(all_results)
        for i, result in enumerate(all_results):
            extracted_data = extract_data_from_result_table(result)
            if not extracted_data:
                individual_results[i]["엑셀 일치도"] = "0.00"
                continue
            match_score = 0
            total_rows = 0
            for row in extracted_data:
                total_rows += 1
                style = row.get('style')
                size = row.get('size')
                color = row.get('color')
                upc = row.get('upc')
                if style and size and color:
                    filtered_df = all_excel_df[
                        (all_excel_df["Style Number"].str.upper() == style.upper()) &
                        (all_excel_df["Size"].str.upper() == size.upper()) &
                        (all_excel_df["Display Color"].str.upper() == color.upper())
                    ]
                    if not filtered_df.empty and upc:
                        exact_match = filtered_df[filtered_df["UPC Code"] == upc]
                        if not exact_match.empty:
                            match_score += 1
                        else:
                            for excel_upc in filtered_df["UPC Code"]:
                                similarity = calculate_upc_similarity(upc, excel_upc)
                                if similarity >= 0.9:
                                    match_score += 0.8
                                    break
            if total_rows > 0:
                excel_match_scores[i] = match_score / total_rows
            individual_results[i]["엑셀 일치도"] = f"{excel_match_scores[i]:.2f}"
    for i in range(len(individual_results)):
        individual_results[i]["일관성"] = f"{consistency_score:.2f}"
    final_result = determine_final_result_with_scores(
        all_results, all_styles, all_sizes, all_colors, all_upcs,
        confidence_scores, consistency_score, completeness_scores, all_excel_df
    )
    return final_result, individual_results

def calculate_confidence_score(result, style, size, color, upc):
    score = 0.0
    if style and validate_style_number(style):
        score += 0.3
    if size and normalize_size(size):
        score += 0.2
    if color and normalize_color_name(color):
        score += 0.2
    if upc:
        normalized_upc = normalize_upc(upc)
        is_valid, _ = validate_upc_code(normalized_upc)
        if is_valid:
            score += 0.3
    return score

def check_result_consistency(results):
    if not results:
        return 0.0
    consistency_scores = {'style': 0.0, 'size': 0.0, 'color': 0.0, 'upc': 0.0}
    extracted_data = []
    for result in results:
        style, size, color, upc = extract_data_from_result(result)
        extracted_data.append({'style': style, 'size': size, 'color': color, 'upc': upc})
    for field in consistency_scores.keys():
        values = [data[field] for data in extracted_data if data[field]]
        if values:
            most_common = max(set(values), key=values.count)
            consistency_scores[field] = values.count(most_common) / len(values)
    weights = {'style': 0.3, 'size': 0.2, 'color': 0.2, 'upc': 0.3}
    total_score = sum(score * weights[field] for field, score in consistency_scores.items())
    return total_score

def determine_final_result_with_scores(results, all_styles, all_sizes, all_colors, all_upcs,
                                         confidence_scores, consistency_score, completeness_scores, all_excel_df=None):
    excel_match_scores = [0] * len(results)
    if all_excel_df is not None and not all_excel_df.empty:
        for i, result in enumerate(results):
            extracted_data = extract_data_from_result_table(result)
            if not extracted_data:
                continue
            match_score = 0
            total_rows = 0
            for row in extracted_data:
                total_rows += 1
                style = row.get('style')
                size = row.get('size')
                color = row.get('color')
                upc = row.get('upc')
                if style and size and color:
                    filtered_df = all_excel_df[
                        (all_excel_df["Style Number"].str.upper() == style.upper()) &
                        (all_excel_df["Size"].str.upper() == size.upper()) &
                        (all_excel_df["Display Color"].str.upper() == color.upper())
                    ]
                    if not filtered_df.empty and upc:
                        exact_match = filtered_df[filtered_df["UPC Code"] == upc]
                        if not exact_match.empty:
                            match_score += 1
                        else:
                            for excel_upc in filtered_df["UPC Code"]:
                                similarity = calculate_upc_similarity(upc, excel_upc)
                                if similarity >= 0.9:
                                    match_score += 0.8
                                    break
            if total_rows > 0:
                excel_match_scores[i] = match_score / total_rows
    CONFIDENCE_WEIGHT = 0.3
    CONSISTENCY_WEIGHT = 0.1
    COMPLETENESS_WEIGHT = 0.3
    EXCEL_MATCH_WEIGHT = 0.3
    result_scores = []
    for i, result in enumerate(results):
        combined_score = (
            confidence_scores[i] * CONFIDENCE_WEIGHT +
            consistency_score * CONSISTENCY_WEIGHT +
            completeness_scores[i] * COMPLETENESS_WEIGHT +
            excel_match_scores[i] * EXCEL_MATCH_WEIGHT
        )
        result_scores.append((result, combined_score))
    best_result = max(result_scores, key=lambda x: x[1])[0]
    best_score = max(result_scores, key=lambda x: x[1])[1]
    if len(results) > 1 and excel_match_scores[-1] >= 0.8:
        if excel_match_scores[-1] > excel_match_scores[0]:
            best_result = results[-1]
    if len(results) > 1:
        second_best_idx = -1
        for i, (result, score) in enumerate(result_scores):
            if result != best_result and (completeness_scores[i] > completeness_scores[result_scores.index((best_result, best_score))] or
                                           excel_match_scores[i] > excel_match_scores[result_scores.index((best_result, best_score))]):
                second_best_idx = i
                break
        if second_best_idx != -1 and (completeness_scores[second_best_idx] > 0.8 or excel_match_scores[second_best_idx] > 0.8):
            best_result = merge_results(best_result, results[second_best_idx])
    if all_upcs:
        if all_excel_df is not None and not all_excel_df.empty:
            extracted_data = extract_data_from_result_table(best_result)
            if extracted_data:
                for row in extracted_data:
                    style = row.get('style')
                    size = row.get('size')
                    color = row.get('color')
                    if style and size and color:
                        filtered_df = all_excel_df[
                            (all_excel_df["Style Number"].str.upper() == style.upper()) &
                            (all_excel_df["Size"].str.upper() == size.upper()) &
                            (all_excel_df["Display Color"].str.upper() == color.upper())
                        ]
                        if not filtered_df.empty:
                            excel_upc = filtered_df.iloc[0]["UPC Code"]
                            for upc in all_upcs:
                                if upc == excel_upc or calculate_upc_similarity(upc, excel_upc) >= 0.9:
                                    best_result = inject_upc_to_result(best_result, excel_upc)
                                    break
        most_common_upc = max(set(all_upcs), key=all_upcs.count)
        best_result = inject_upc_to_result(best_result, most_common_upc)
    return best_result

def determine_final_result(all_results, all_styles, all_sizes, all_colors, all_upcs):
    for i, result in enumerate(all_results):
        style, size, color, upc = extract_data_from_result(result)
        matches = 0
        for j, other_result in enumerate(all_results):
            if i == j:
                continue
            other_style, other_size, other_color, other_upc = extract_data_from_result(other_result)
            if style and other_style and style == other_style and size and other_size and size == other_size and color and other_color and color == other_color:
                matches += 1
        if matches >= len(all_results) // 2:
            return result
    style_counter = Counter(all_styles)
    most_common_style = style_counter.most_common(1)[0][0] if style_counter else None
    size_counter = Counter(all_sizes)
    most_common_size = size_counter.most_common(1)[0][0] if size_counter else None
    upc_counter = Counter(all_upcs)
    most_common_upc = upc_counter.most_common(1)[0][0] if upc_counter else None
    final_result = {"style": most_common_style, "size": most_common_size, "upc": most_common_upc}
    return final_result

# =============================================================================
# (참고) merge_results, inject_upc_to_result, calculate_upc_similarity,
# extract_data_from_result, extract_data_from_result_table, is_last_digits_different,
# is_completely_different_upc, has_significant_middle_difference 등의 함수는
# 원본 로직을 그대로 유지합니다.
#
# =============================================================================
# Streamlit 파일 업로드 및 분석 UI

col1, col2 = st.columns(2)
with col1:
    excel_files = st.file_uploader("📊 Excel/CSV 파일 업로드 (기준 데이터, 여러 개 가능)",
                                   type=["xlsx", "xls", "csv"], accept_multiple_files=True, key="excel_upload")
with col2:
    pdf_files = st.file_uploader("📄 PDF 파일 업로드 (비교할 데이터, 여러 개 가능)",
                                 type=["pdf"], accept_multiple_files=True, key="pdf_upload")

if excel_files and pdf_files:
    all_excel_dfs = []
    for excel_file in excel_files:
        st.subheader(f"📊 {excel_file.name} 처리 중")
        excel_df = process_excel_data(excel_file)
        if excel_df is not None:
            all_excel_dfs.append(excel_df)
            with st.expander(f"{excel_file.name} 데이터 보기"):
                st.dataframe(excel_df)
    if all_excel_dfs:
        all_excel_df = pd.concat(all_excel_dfs, ignore_index=True)
        all_excel_df = all_excel_df.drop_duplicates(subset=["Style Number", "Size", "Display Color", "UPC Code"])
        st.subheader("📊 통합 Excel/CSV 기준 데이터")
        st.dataframe(all_excel_df)
    analysis_attempts = 2  # 기본 분석 시도 횟수
    use_multi_analysis = True
    use_enhanced_correction = True
    check_fabric_care = True
    is_label_analysis = (analysis_type == "Label Analysis")
    for pdf_file in pdf_files:
        st.subheader(f"📄 분석 중: {pdf_file.name}")
        pdf_bytes = pdf_file.read()
        normalized_detected_upcs = []
        if is_text_pdf(pdf_bytes):
            st.info(f"텍스트 기반 PDF 분석: {pdf_file.name}")
            text = extract_text_from_pdf(pdf_bytes)
            if is_label_analysis:
                gpt_output = analyze_label_text_with_gpt(text)
                extracted_upcs = []  # 추출 함수 호출 (예: extract_upcs_from_result)
                normalized_detected_upcs.extend(extracted_upcs)
                if check_fabric_care:
                    additional_info = analyze_label_text_with_gpt(text)
                    fabric_info, care_info, factory_code = extract_fabric_care_from_gpt_response(additional_info)
            else:
                gpt_output = analyze_size_strip_text_with_gpt(text)
        else:
            st.warning(f"이미지 기반 PDF Vision 분석: {pdf_file.name}")
            images = convert_pdf_to_images(pdf_bytes)
            progress_bar = st.progress(0)
            gpt_output = ""
            all_individual_results = []
            for i, img_base64 in enumerate(images):
                progress_bar.progress((i + 1) / len(images))
                if use_multi_analysis and is_label_analysis:
                    try:
                        page_result, individual_results = multi_vision_analysis(
                            img_base64, is_label=True, attempts=analysis_attempts, all_excel_df=all_excel_df
                        )
                        if individual_results and isinstance(individual_results, list):
                            all_individual_results.extend(individual_results)
                        if i == len(images) - 1:
                            best_score = -1
                            best_index = 0
                            for j, result_info in enumerate(all_individual_results):
                                confidence = float(result_info.get('신뢰도', 0))
                                consistency = float(result_info.get('일관성', 0))
                                completeness = float(result_info.get('완전성', 0))
                                excel_match = float(result_info.get('엑셀 일치도', 0))
                                combined_score = (confidence * 0.3 + consistency * 0.1 +
                                                  completeness * 0.3 + excel_match * 0.3)
                                if combined_score > best_score:
                                    best_score = combined_score
                                    best_index = j
                    except Exception as e:
                        st.error(f"다중 분석 중 오류: {str(e)}")
                        page_result = analyze_label_image_with_vision(img_base64)
                else:
                    if is_label_analysis:
                        page_result = analyze_label_image_with_vision(img_base64)
                    else:
                        page_result = analyze_size_strip_image_with_vision(img_base64)
                page_upcs = []  # 추출 함수 호출 (예: extract_upcs_from_result)
                normalized_detected_upcs.extend(page_upcs)
                if is_label_analysis and check_fabric_care and i == 0:
                    additional_info = analyze_label_image_with_vision(img_base64)
                    page_fabric, page_care, page_factory = extract_fabric_care_from_gpt_response(additional_info)
                    fabric_info = page_fabric if page_fabric else None
                    care_info = page_care if page_care else None
                    factory_code = page_factory if page_factory else None
                gpt_output += page_result + "\n"
            st.success(f"이미지 기반 PDF 분석 완료: {pdf_file.name}")
        st.text_area("원본 출력", gpt_output, height=200)
        if normalized_detected_upcs:
            st.info(f"발견된 UPC 코드: {', '.join(normalized_detected_upcs)}")
        if check_fabric_care and (fabric_info or care_info or factory_code):
            with st.expander("Fabric & Care Content & Factory Code 정보"):
                if fabric_info:
                    st.info(f"Fabric: {fabric_info}")
                if care_info:
                    st.success(f"Care Content: {care_info}")
                if factory_code:
                    st.info(f"Factory Code: {factory_code}")
else:
    st.info("Excel/CSV 파일 및 PDF 파일을 업로드하세요.")
