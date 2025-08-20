import streamlit as st
import cv2
import numpy as np
import os
import tempfile
from pptx import Presentation
from pptx.util import Inches, Pt
from PIL import Image, ExifTags

# --------- Helper: Ensure PPT-compatible format ----------
def ensure_supported_format(image_path):
    supported_formats = ["BMP", "GIF", "JPEG", "PNG", "TIFF", "WMF"]
    try:
        img = Image.open(image_path)
        fmt = img.format.upper() if img.format else None

        if fmt not in supported_formats:
            temp_dir = tempfile.gettempdir()
            new_path = os.path.join(temp_dir, os.path.basename(image_path) + ".png")
            img.convert("RGB").save(new_path, "PNG")
            return new_path
        return image_path
    except Exception as e:
        print(f"Error converting {image_path}: {e}")
        return None

# --------- Helper: Auto rotate if EXIF orientation exists ----------
def auto_rotate(image_path):
    try:
        img = Image.open(image_path)
        for orientation in ExifTags.TAGS.keys():
            if ExifTags.TAGS[orientation] == 'Orientation':
                break
        exif = dict(img._getexif().items())

        if exif.get(orientation) == 3:
            img = img.rotate(180, expand=True)
        elif exif.get(orientation) == 6:
            img = img.rotate(270, expand=True)
        elif exif.get(orientation) == 8:
            img = img.rotate(90, expand=True)

        temp_path = os.path.join(tempfile.gettempdir(), "rotated_" + os.path.basename(image_path))
        img.save(temp_path)
        return temp_path
    except Exception:
        return image_path

# --------- Edge Case Detection ----------
def detect_edge_cases(image_path):
    reasons = []
    try:
        img = cv2.imread(image_path)
        if img is None:
            return ["Unreadable image"]

        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

        # Blur check
        blur_val = cv2.Laplacian(gray, cv2.CV_64F).var()
        if blur_val < 100:
            reasons.append(f"Blurry image detected (Blur score = {blur_val:.2f}, threshold=100)")

        # Brightness check
        brightness = np.mean(gray)
        if brightness < 50:
            reasons.append(f"Low brightness (Brightness = {brightness:.2f}, threshold=50)")

        # Contrast check
        contrast = gray.std()
        if contrast < 20:
            reasons.append(f"Low contrast (Contrast = {contrast:.2f}, threshold=20)")

        # Tilt detection using Hough lines
        edges = cv2.Canny(gray, 50, 150, apertureSize=3)
        lines = cv2.HoughLines(edges, 1, np.pi/180, 200)
        if lines is not None:
            angles = []
            for rho, theta in lines[:, 0]:
                angle = (theta * 180 / np.pi) - 90
                if -45 < angle < 45:
                    angles.append(angle)
            if angles:
                avg_angle = np.mean(angles)
                if abs(avg_angle) > 5:  # tilted
                    reasons.append(f"Tilted image auto-corrected (rotation={avg_angle:.2f}°)")
    except Exception as e:
        reasons.append(f"Error processing: {e}")

    return reasons

# --------- PPT Generator ----------
def create_ppt(edge_cases, output_path="Edge_Cases_Presentation.pptx"):
    prs = Presentation()
    blank_slide_layout = prs.slide_layouts[6]  # empty slide

    for case in edge_cases:
        image_path, reasons = case
        if not reasons:
            continue

        slide = prs.slides.add_slide(blank_slide_layout)

        # Add title
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(1))
        tf = title_box.text_frame
        tf.text = "Edge Case Detected"
        tf.paragraphs[0].font.size = Pt(24)
        tf.paragraphs[0].font.bold = True

        # Add image
        img_path = ensure_supported_format(image_path)
        if img_path:
            slide.shapes.add_picture(img_path, Inches(1), Inches(1.5), width=Inches(6))

        # Add reason text
        left = Inches(0.5)
        top = Inches(5.5)
        width = Inches(9)
        height = Inches(2)
        text_box = slide.shapes.add_textbox(left, top, width, height)
        tf = text_box.text_frame
        for reason in reasons:
            p = tf.add_paragraph()
            p.text = reason
            p.font.size = Pt(14)

    prs.save(output_path)
    return output_path

# --------- Streamlit UI ----------
st.title("🛠️ Product Image Edge Case Analyzer")

uploaded_files = st.file_uploader("Upload multiple images", type=["jpg", "jpeg", "png", "bmp", "tiff", "gif", "mpo"], accept_multiple_files=True)

if uploaded_files:
    edge_cases = []
    for uploaded_file in uploaded_files:
        # Save temp file
        temp_path = os.path.join(tempfile.gettempdir(), uploaded_file.name)
        with open(temp_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        # Auto rotate before analysis
        rotated_path = auto_rotate(temp_path)

        # Detect edge cases
        reasons = detect_edge_cases(rotated_path)

        if reasons:
            st.image(rotated_path, caption=f"Edge Cases: {', '.join(reasons)}", use_container_width=True)
            edge_cases.append((rotated_path, reasons))

    if edge_cases:
        ppt_path = create_ppt(edge_cases)
        with open(ppt_path, "rb") as f:
            st.download_button("📥 Download Edge Cases PPT", f, file_name="Edge_Cases_Presentation.pptx")
