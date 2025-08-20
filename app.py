import cv2
import os
import numpy as np
import tempfile
import zipfile
import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from PIL import Image

# Thresholds
BLUR_THRESHOLD = 100
BRIGHT_LOW = 50
BRIGHT_HIGH = 200
CONTRAST_LOW = 15
SMALL_PRODUCT_RATIO = 0.1

# ================== DETECTION FUNCTIONS ==================
def detect_blur(image):
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    fm = cv2.Laplacian(gray, cv2.CV_64F).var()
    return fm < BLUR_THRESHOLD, f"Blur score = {fm:.2f} (Threshold {BLUR_THRESHOLD})"

def detect_brightness_contrast(image):
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    mean = np.mean(gray)
    contrast = gray.std()
    if mean < BRIGHT_LOW:
        return "Dark", f"Brightness too low = {mean:.2f} (Threshold {BRIGHT_LOW})"
    elif mean > BRIGHT_HIGH:
        return "Bright", f"Brightness too high = {mean:.2f} (Threshold {BRIGHT_HIGH})"
    elif contrast < CONTRAST_LOW:
        return "Low Contrast", f"Contrast too low = {contrast:.2f} (Threshold {CONTRAST_LOW})"
    return None, None

def detect_rotation(image):
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    _, thresh = cv2.threshold(gray, 128, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    contours, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not contours:
        return False, None
    c = max(contours, key=cv2.contourArea)
    rect = cv2.minAreaRect(c)
    angle = rect[-1]
    if angle < -45:
        angle = 90 + angle
    tilted = abs(angle) > 15
    return tilted, f"Tilt angle = {angle:.2f}° (Threshold 15°)" if tilted else None

def detect_small_object(image):
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    _, thresh = cv2.threshold(gray, 128, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    contours, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not contours:
        return False, None
    c = max(contours, key=cv2.contourArea)
    area = cv2.contourArea(c)
    img_area = image.shape[0] * image.shape[1]
    ratio = area / img_area
    small = ratio < SMALL_PRODUCT_RATIO
    return small, f"Object covers {ratio*100:.2f}% of image (Threshold {SMALL_PRODUCT_RATIO*100:.0f}%)" if small else None

# ================== PPT CREATION ==================
def create_ppt(edge_cases, output_path):
    prs = Presentation()
    # Title slide
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = "Edge Case Analysis Report"
    slide.placeholders[1].text = "Automatically generated from product images"

    for case, img_path, reason in edge_cases:
        slide_layout = prs.slide_layouts[5]  # Title + Content
        slide = prs.slides.add_slide(slide_layout)
        slide.shapes.title.text = f"Edge Case: {case}"

        # Insert image
        slide.shapes.add_picture(img_path, Inches(1), Inches(1.5), height=Inches(3))

        # Insert reason text
        tx_box = slide.shapes.add_textbox(Inches(1), Inches(4.8), Inches(7), Inches(1.5))
        tf = tx_box.text_frame
        p = tf.add_paragraph()
        p.text = reason
        p.font.size = Pt(14)

    prs.save(output_path)

# ================== STREAMLIT APP ==================
st.title("📊 Edge Case Detector for Product Images")

uploaded_zip = st.file_uploader("Upload a ZIP file containing product images", type=["zip"])

if uploaded_zip:
    with tempfile.TemporaryDirectory() as tmpdir:
        # Extract ZIP
        zip_path = os.path.join(tmpdir, "images.zip")
        with open(zip_path, "wb") as f:
            f.write(uploaded_zip.read())
        with zipfile.ZipFile(zip_path, "r") as zip_ref:
            zip_ref.extractall(tmpdir)

        # Process images
        edge_cases = []
        for root, _, files in os.walk(tmpdir):
            for file in files:
                if file.lower().endswith((".jpg", ".jpeg", ".png")):
                    img_path = os.path.join(root, file)
                    image = cv2.imread(img_path)
                    if image is None:
                        continue

                    blur_flag, blur_reason = detect_blur(image)
                    if blur_flag:
                        edge_cases.append(("Blurry", img_path, blur_reason))

                    bc, bc_reason = detect_brightness_contrast(image)
                    if bc:
                        edge_cases.append((bc, img_path, bc_reason))

                    tilt_flag, tilt_reason = detect_rotation(image)
                    if tilt_flag:
                        edge_cases.append(("Tilted", img_path, tilt_reason))

                    small_flag, small_reason = detect_small_object(image)
                    if small_flag:
                        edge_cases.append(("Small Product", img_path, small_reason))

        if edge_cases:
            st.subheader("🔎 Detected Edge Cases")
            for case, img_path, reason in edge_cases:
                col1, col2 = st.columns([1,2])
                with col1:
                    img = Image.open(img_path)
                    st.image(img, caption=f"{case}", use_column_width=True)
                with col2:
                    st.write(f"**Reason:** {reason}")

            ppt_path = os.path.join(tmpdir, "edge_cases_report.pptx")
            create_ppt(edge_cases, ppt_path)

            with open(ppt_path, "rb") as f:
                st.download_button(
                    label="📥 Download Edge Case Report (PPTX)",
                    data=f,
                    file_name="edge_cases_report.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
        else:
            st.info("✅ No edge cases detected in the uploaded images.")
