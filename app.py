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

# OpenAI API key setup (fetched from Streamlit Cloud secrets)
client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

# Global variables for progress bar and status message
progress_bar = None
status_text = None

# User authentication and usage tracking
def authenticate_user():
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    
    if not st.session_state.authenticated:
        email = st.text_input("Enter Your Email Address")
        if st.button("Authenticate"):
            if email.endswith("@yakjin.com"):
                st.session_state.authenticated = True
                st.session_state.email = email
                st.success("Authentication Successful.")
                return True
            else:
                st.error("Unauthorized email address. Only @yakjin.com domain emails are allowed.")
                return False
    return st.session_state.authenticated

# Analysis options definition (modified)
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

# Individual image analysis function (with caching)
@st.cache_data
def analyze_single_image(image, category, options):
    base64_image = encode_image(image)
    
    prompt = f"Analyze the {category} clothing item in the image and provide information on the following aspects. Choose only the most appropriate option for each:\n\n"
    for option in options:
        prompt += f"{option}: {', '.join(analysis_options[category][option])}\n"
    
    prompt += "\nProvide the result as a JSON object with the selected aspects as keys and the detected options as values. Choose only one value for each key."

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
        
        # Attempt JSON parsing
        try:
            return json.loads(processed_result)
        except json.JSONDecodeError:
            st.error(f"JSON Parsing Error: {processed_result}")
            return {}
    except Exception as e:
        st.error(f"Error Occurred During Image Analysis: {e}")
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
def enhance_image(image, scale_factor=2):
    cv_image = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
    height, width = cv_image.shape[:2]
    resized = cv2.resize(cv_image, (width*scale_factor, height*scale_factor), interpolation=cv2.INTER_CUBIC)
    gaussian = cv2.GaussianBlur(resized, (0, 0), 3.0)
    sharpened = cv2.addWeighted(resized, 1.5, gaussian, -0.5, 0, resized)
    denoised = cv2.fastNlMeansDenoisingColored(sharpened, None, 10, 10, 7, 21)
    return Image.fromarray(cv2.cvtColor(denoised, cv2.COLOR_BGR2RGB))

# Modified donut chart creation function
def create_donut_chart(data, title):
    labels = list(data.keys())
    values = list(data.values())
    
    if title.lower() == 'color':
        colors = [get_color(label) for label in labels]
        # Change white to very light gray
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
        textfont=dict(size=14, color=text_colors),
        hoverinfo='label+percent+text',
        text=[f'Count: {v}' for v in values],
        hovertemplate='%{label}<br>%{percent}<br>%{text}<extra></extra>'
    )])
    
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
        margin=dict(t=80, b=80, l=20, r=20),  # Increased top margin
        annotations=[
            dict(
                text=f'<b>{title}</b>',
                x=0.5,  # Set x position to center
                y=1.15,  # Set y position above the chart
                xref='paper',
                yref='paper',
                showarrow=False,
                font=dict(size=32, color='black'),  # Changed title color to black
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

# Modified main app logic (image list part)
def main():
    st.set_page_config(layout="centered")
    
    st.markdown("""
    <style>
    .emoji-title { font-size: 2.4em; }
    .emoji { font-size: 0.8em; }
    .results-container { display: flex; flex-wrap: wrap; justify-content: space-between; }
    .chart-container { width: 48%; margin-bottom: 20px; }
    </style>
    """, unsafe_allow_html=True)
    
    st.markdown("<h1 class='emoji-title'>📊 Fashion Image Analyzer</h1>", unsafe_allow_html=True)
    
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
        
        st.markdown("<h3><span class='emoji'>📁</span> Step 3: Upload File</h3>", unsafe_allow_html=True)
        uploaded_file = st.file_uploader("Choose File", type=["xlsx", "xls", "png", "jpg", "jpeg", "zip"])
        
        if uploaded_file is not None:
            st.markdown("<h3><span class='emoji'>🖼️</span> Step 4: Image Processing</h3>", unsafe_allow_html=True)
            
            images = []
            if uploaded_file.type in ["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "application/vnd.ms-excel"]:
                try:
                    images = extract_images_from_excel(uploaded_file)
                    if images:
                        images = images[1:]  # Exclude the first image (logo)
                except Exception as e:
                    st.error(f"Error Occurred While Extracting Images from Excel File: {str(e)}")
            elif uploaded_file.type.startswith('image/'):
                images = [Image.open(uploaded_file)]
            elif uploaded_file.type == 'application/zip':
                images = [Image.open(io.BytesIO(img_data)) for _, img_data in process_zip_file(uploaded_file)]
            
            if images:
                with st.spinner('Processing Images...'):
                    processed_images = process_images(images)
                
                st.success(f"{len(processed_images)} Images Processed Successfully.")
                
                if st.button("🚀 Step 5: Start Analysis", key="start_analysis"):
                    if not selected_options:
                        st.markdown("<p><span class='emoji'>⚠️</span> Please Select at Least One Analysis Item.</p>", unsafe_allow_html=True)
                    else:
                        # 분석 로직 (기존 코드 유지)
                        progress_bar = st.progress(0)
                        status_text = st.empty()
                        
                        aggregated_results = {option: Counter() for option in selected_options}
                        image_categories = defaultdict(lambda: defaultdict(list))
                        
                        for i, image in enumerate(processed_images):
                            result = analyze_single_image(image, selected_category, selected_options)
                            if result and isinstance(result, dict):
                                for option, detected in result.items():
                                    if option in selected_options:
                                        aggregated_results[option][detected] += 1
                                        image_categories[option][detected].append(image)
                            else:
                                st.warning(f"Invalid Analysis Result for Image {i+1}.")
                            
                            progress = (i + 1) / len(processed_images)
                            progress_bar.progress(progress)
                            status_text.text(f"Analyzing Images: {i+1}/{len(processed_images)}")
                        
                        progress_bar.empty()
                        status_text.empty()
                        
                        # 분석 결과 표시 부분을 전체 화면으로 전환
                        st.markdown("<hr>", unsafe_allow_html=True)
                        st.markdown("<h2 style='text-align: center;'>📊 Analysis Results</h2>", unsafe_allow_html=True)
                        st.markdown("<div class='results-container'>", unsafe_allow_html=True)
                        
                        # 결과 표시 (2열 레이아웃)
                        for i, (option, results) in enumerate(aggregated_results.items()):
                            if results:
                                st.markdown(f"<div class='chart-container'>", unsafe_allow_html=True)
                                fig = create_donut_chart(results, option)
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
                        
                        st.markdown("</div>", unsafe_allow_html=True)
            else:
                st.markdown("<p><span class='emoji'>⚠️</span> No Images Found in the Uploaded File.</p>", unsafe_allow_html=True)

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
