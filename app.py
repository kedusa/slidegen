# Trigger new deploy
# -*- coding: utf-8 -*-
import streamlit as st
import cv2
import numpy as np
import pytesseract
import fitz  # PyMuPDF
from PIL import Image, ImageDraw, ImageFont, UnidentifiedImageError
import io
import os
import base64
import random
import tempfile
import re
import logging
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_VERTICAL_ANCHOR # Correct import
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
import datetime

# --- Basic Setup ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(name)s - %(message)s')
logger = logging.getLogger("ABTestApp")

# Configure Tesseract path (Adjust if necessary for your OS)
# Example for Windows:
# if os.name == 'nt':
#     try: pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
#     except Exception: logger.error("Tesseract not found at default path. Please set manually if needed.")
# Check if tesseract is installed and provide fallback
try:
    # Test if tesseract is available
    tesseract_version = pytesseract.get_tesseract_version()
    logger.info(f"Tesseract OCR found: version {tesseract_version}")
    TESSERACT_AVAILABLE = True
except pytesseract.TesseractNotFoundError:
    logger.warning("Tesseract OCR not found. OCR features will be disabled.")
    TESSERACT_AVAILABLE = False
    
    # Create fallback extraction function that won't crash
    def safe_extract_text_from_image(image_input):
        """Fallback function when Tesseract is not available"""
        if isinstance(image_input, Image.Image): 
            return "OCR unavailable: Tesseract not installed", image_input
        else:
            try:
                image = Image.open(image_input)
                return "OCR unavailable: Tesseract not installed", image
            except Exception as e:
                return f"Error: {str(e)}", Image.new('RGB', (640, 480))
# PDQ Color Palette
PDQ_COLORS = {
    "revolver": "#231333", "electric_violet": "#894DFF", "electric_violet_2": "#6B2BEF",
    "magnolia": "#F6F3FF", "melrose": "#9B8CFF", "valentino": "#2C124B",
    "moon_raker": "#E4E1FA", "lightning_yellow": "#FBC018", "persian_green": "#00AA8C",
    "carnation": "#F04557", "white": "#FFFFFF", "black": "#000000",
    "grey_text": "#666666", "html_border": "#DDDDDD", "html_selected_border": "#E86C60",
    "html_selected_bg": "#FEF6F5", "html_radio_border": "#CCCCCC",
}

# --- Helper Functions ---
def hex_to_rgb(hex_color):
    """Converts hex color string to RGB tuple."""
    hex_color = hex_color.lstrip('#')
    length = len(hex_color)
    if length == 6: return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
    if length == 3: return tuple(int(hex_color[i]*2, 16) for i in (0, 1, 2))
    logger.warning(f"Invalid hex: {hex_color}. Using black."); return (0, 0, 0)

def hex_to_rgbcolor(hex_color):
    """Converts hex color string to python-pptx RGBColor object."""
    r, g, b = hex_to_rgb(hex_color); return RGBColor(r, g, b)

def get_font(name, size, bold=False, italic=False):
    """Attempts to load Segoe UI, falls back to Arial/Calibri, then default."""
    common_fonts = [name]
    if 'Segoe UI' not in name: common_fonts.append('Segoe UI')
    common_fonts.extend(['Arial', 'Calibri']) # Common fallbacks
    font_path = None
    for font_name in common_fonts:
        try:
            # Basic bold/italic handling (may need more robust font selection)
            if bold and italic: suffix = " Bold Italic"
            elif bold: suffix = " Bold"
            elif italic: suffix = " Italic"
            else: suffix = ""
            # Try finding the font (this is basic, system-dependent)
            # Adjust path separators if needed for your OS
            try:
                return ImageFont.truetype(f"{font_name}{suffix}.ttf", size)
            except IOError: # Try common variations like .otf or system paths
                try:
                    return ImageFont.truetype(f"{font_name}{suffix}.otf", size)
                except IOError:
                    continue # Try next font name
        except IOError:
             continue # Try next font in list
    logger.warning(f"Could not find fonts: {common_fonts}. Using default.")
    return ImageFont.load_default() # Final fallback

# --- Streamlit Page Setup & CSS ---
st.set_page_config(page_title="PDQ A/B Test Slide Generator", page_icon="🧪", layout="wide")
# Define custom CSS
st.markdown(f"""
    <style>
    .main {{ background-color: {PDQ_COLORS["magnolia"]}; }}
    .stApp {{ max-width: 1400px; margin: 0 auto; }}
    .success-box {{ padding: 1rem; border-radius: 0.5rem; background-color: rgba({hex_to_rgb(PDQ_COLORS["persian_green"])[0]}, {hex_to_rgb(PDQ_COLORS["persian_green"])[1]}, {hex_to_rgb(PDQ_COLORS["persian_green"])[2]}, 0.2); color: {PDQ_COLORS["persian_green"]}; margin-bottom: 1rem; border: 1px solid {PDQ_COLORS["persian_green"]}; }}
    .stButton>button {{ background-color: {PDQ_COLORS["electric_violet"]}; color: white; font-weight: bold; border: none; padding: 0.6rem 1.2rem; border-radius: 0.3rem; }}
    .stButton>button:hover {{ background-color: {PDQ_COLORS["electric_violet_2"]}; }}
    .stButton>button:disabled {{ background-color: #cccccc; color: #666666; cursor: not-allowed; }}
    h1, h2, h3 {{ color: {PDQ_COLORS["valentino"]}; }}
    .preview-box {{ border: 1px solid {PDQ_COLORS["melrose"]}; border-radius: 0.5rem; padding: 1.5rem; background-color: {PDQ_COLORS["white"]}; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }}
    .download-button {{ background-color: {PDQ_COLORS["persian_green"]}; color: white; padding: 0.6rem 1.2rem; border-radius: 0.3rem; text-decoration: none; font-weight: bold; display: inline-block; margin-top: 1rem; border: none; }}
    .download-button:hover {{ background-color: #008a70; }}
    .stImage > img {{ border: 1px solid {PDQ_COLORS["moon_raker"]}; border-radius: 0.25rem; }}
    .stSidebar > div:first-child {{ background-color: {PDQ_COLORS["revolver"]}; }}
    .stSidebar .stMarkdown p, .stSidebar .stFileUploader label, .stSidebar .stTextInput label, .stSidebar .stTextArea label, .stSidebar .stCheckbox label, .stSidebar h1, .stSidebar h2, .stSidebar h3, .stSidebar h4 {{ color: {PDQ_COLORS["magnolia"]}; }}
    .stSidebar .stButton>button {{ background-color: {PDQ_COLORS["melrose"]}; color: {PDQ_COLORS["revolver"]}; }}
    .stSidebar .stButton>button:hover {{ background-color: {PDQ_COLORS["electric_violet"]}; color: white; }}
    .stTabs [data-baseweb="tab-list"] {{ gap: 2px; }}
    .stTabs [data-baseweb="tab"] {{ background-color: {PDQ_COLORS["moon_raker"]}; border-radius: 4px 4px 0 0; color: {PDQ_COLORS["valentino"]}; padding: 0.5rem 1rem; font-weight: 500; }}
    .stTabs [aria-selected="true"] {{ background-color: {PDQ_COLORS["electric_violet"]}; color: white; }}
    .stTextInput>div>div>input, .stTextArea>div>textarea {{ border: 1px solid {PDQ_COLORS["melrose"]}; border-radius: 0.25rem; }}
    .stTextInput>div>div>input:focus, .stTextArea>div>textarea:focus {{ border-color: {PDQ_COLORS["electric_violet"]}; box-shadow: 0 0 0 2px rgba({hex_to_rgb(PDQ_COLORS["electric_violet"])[0]}, {hex_to_rgb(PDQ_COLORS["electric_violet"])[1]}, {hex_to_rgb(PDQ_COLORS["electric_violet"])[2]}, 0.3); }}
    label {{ color: {PDQ_COLORS["valentino"]}; font-weight: 500; }}
    .error-box {{ padding: 1rem; border-radius: 0.5rem; background-color: rgba({hex_to_rgb(PDQ_COLORS["carnation"])[0]}, {hex_to_rgb(PDQ_COLORS["carnation"])[1]}, {hex_to_rgb(PDQ_COLORS["carnation"])[2]}, 0.2); color: {PDQ_COLORS["carnation"]}; margin-bottom: 1rem; border: 1px solid {PDQ_COLORS["carnation"]}; }}
    .highlight-box {{ padding: 1rem; border-radius: 0.5rem; background-color: rgba({hex_to_rgb(PDQ_COLORS["lightning_yellow"])[0]}, {hex_to_rgb(PDQ_COLORS["lightning_yellow"])[1]}, {hex_to_rgb(PDQ_COLORS["lightning_yellow"])[2]}, 0.2); color: #805b00; margin-bottom: 1rem; border: 1px solid {PDQ_COLORS["lightning_yellow"]}; }}
    footer {{ visibility: hidden; }}
    .custom-footer {{ margin-top: 2rem; padding-top: 1rem; border-top: 1px solid {PDQ_COLORS["moon_raker"]}; display: flex; justify-content: space-between; align-items: center; font-size: 0.85rem; }}
    .footer-left {{ color: {PDQ_COLORS["valentino"]}; }}
    .footer-right {{ color: {PDQ_COLORS["electric_violet"]}; font-weight: bold; }}
    </style>
""", unsafe_allow_html=True)


# --- Core Image/Text Processing Functions ---
def extract_text_from_image(image_input):
    """Extract text from an image file or PIL object using OCR"""
    try:
        if isinstance(image_input, Image.Image): image = image_input
        else:
            if image_input is None or image_input.size == 0: logger.warning("extract_text_from_image received empty file."); return "", Image.new('RGB', (1, 1))
            image = Image.open(image_input)
        img_cv = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
        gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)
        thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1] # Corrected constants
        extracted_text = pytesseract.image_to_string(thresh, config='--psm 6')
        logger.info(f"OCR extracted {len(extracted_text)} characters.")
        return extracted_text, image
    except UnidentifiedImageError: st.error("Invalid image file."); logger.error("UnidentifiedImageError."); return "", Image.new('RGB', (100, 30))
    except pytesseract.TesseractNotFoundError: st.error("Tesseract not found."); logger.error("Tesseract not found."); return "", Image.new('RGB', (100, 30))
    except Exception as e: st.error(f"OCR Error: {e}"); logger.error(f"OCR Error: {e}", exc_info=True); return "", Image.new('RGB', (640, 480))

def extract_metrics_from_supporting_data(image_obj):
    """Extract key metrics from a PIL image object using OCR."""
    internal_default_metrics = { "conversion_rate": "N/A", "total_checkout": "N/A", "checkouts": "N/A", "orders": "N/A", "shipping_revenue": "N/A", "aov": "N/A" }
    if not isinstance(image_obj, Image.Image): logger.warning("extract_metrics needs PIL object."); return internal_default_metrics
    try:
        extracted_text, _ = extract_text_from_image(image_obj)
        if not extracted_text: logger.warning("No text from support data OCR."); return internal_default_metrics
        # Regex
        conv_match = re.search(r'(?:conversion|checkout\s*conversion)\s*(?:rate)?[:\s]*(\d{1,3}(?:,\d{3})*\.?\d*%?|\d+\.?\d*%?)', extracted_text, re.IGNORECASE)
        total_co_match = re.search(r'(?:%|percent)\s*total\s*checkout[:\s]*(\d{1,3}(?:,\d{3})*\.?\d*%?|\d+\.?\d*%?)', extracted_text, re.IGNORECASE)
        checkouts_match = re.search(r'(?:#|number\s+of)?\s*Checkouts[:\s]*(\d{1,3}(?:,\d{3})*)', extracted_text, re.IGNORECASE)
        orders_match = re.search(r'(?:#|number\s+of)?\s*Orders[:\s]*(\d{1,3}(?:,\d{3})*)', extracted_text, re.IGNORECASE)
        ship_rev_match = re.search(r'(?:avg|average)?\s*shipping\s*revenue[:\s]*(\$\s*\d{1,3}(?:,\d{3})*\.?\d*)', extracted_text, re.IGNORECASE)
        aov_match = re.search(r'(?:AOV|average\s*order\s*value)[:\s]*(\$\s*\d{1,3}(?:,\d{3})*\.?\d*)', extracted_text, re.IGNORECASE)
        # Assign results or defaults (Corrected)
        metrics = {
            "conversion_rate": conv_match.group(1).strip() if conv_match else internal_default_metrics["conversion_rate"],
            "total_checkout": total_co_match.group(1).strip() if total_co_match else internal_default_metrics["total_checkout"],
            "checkouts": checkouts_match.group(1).strip() if checkouts_match else internal_default_metrics["checkouts"],
            "orders": orders_match.group(1).strip() if orders_match else internal_default_metrics["orders"],
            "shipping_revenue": ship_rev_match.group(1).strip() if ship_rev_match else internal_default_metrics["shipping_revenue"],
            "aov": aov_match.group(1).strip() if aov_match else internal_default_metrics["aov"],
        }
        # Clean up values
        if metrics["conversion_rate"] != "N/A" and '%' not in metrics["conversion_rate"]: metrics["conversion_rate"] += "%"
        if metrics["total_checkout"] != "N/A" and '%' not in metrics["total_checkout"]: metrics["total_checkout"] += "%"
        logger.info(f"Extracted Metrics: {metrics}")
        return metrics
    except Exception as e: logger.error(f"Metric extraction error: {e}", exc_info=True); return internal_default_metrics

# --- HTML Variant Generation ---
def generate_shipping_html(standard_price="$7.95", rush_price="$24.95", is_variant=False):
    """ Generate HTML content for shipping options display """
    html = f"""<!DOCTYPE html><html><head><style> body {{ font-family: 'Segoe UI', Arial, sans-serif; background-color: #f8f9fa; margin: 0; padding: 20px; box-sizing: border-box; }} .container {{ max-width: 580px; background-color: {PDQ_COLORS['white']}; border: 1px solid {PDQ_COLORS['html_border']}; border-radius: 6px; padding: 20px; position: relative; box-shadow: 0 1px 2px rgba(0,0,0,0.05); }} h2 {{ margin-top: 0; margin-bottom: 15px; font-size: 16px; font-weight: 600; color: {PDQ_COLORS['black']}; }} .shipping-option {{ border: 1px solid {PDQ_COLORS['html_border']}; border-radius: 6px; padding: 15px; margin-bottom: 10px; display: flex; align-items: flex-start; transition: all 0.2s ease-in-out; }} .shipping-option.selected {{ border-color: {PDQ_COLORS['html_selected_border']}; background-color: {PDQ_COLORS['html_selected_bg']}; }} .radio {{ margin-right: 12px; margin-top: 3px; flex-shrink: 0; }} .radio-dot {{ width: 18px; height: 18px; border-radius: 50%; display: flex; align-items: center; justify-content: center; }} .radio-selected .radio-dot {{ background-color: {PDQ_COLORS['html_selected_border']}; }} .radio-selected .radio-dot-inner {{ width: 7px; height: 7px; border-radius: 50%; background-color: white; }} .radio-unselected .radio-dot {{ border: 2px solid {PDQ_COLORS['html_radio_border']}; background-color: white; }} .shipping-details {{ flex-grow: 1; }} .shipping-title {{ font-weight: 600; font-size: 14px; margin-bottom: 4px; color: #333; }} .shipping-subtitle {{ color: {PDQ_COLORS['grey_text']}; font-size: 12px; }} .shipping-price {{ font-weight: 600; font-size: 14px; text-align: right; min-width: 60px; color: #333; margin-left: 10px; }} .footnote {{ font-size: 12px; color: {PDQ_COLORS['grey_text']}; margin-top: 15px; }} .variant-label {{ position: absolute; top: 10px; right: 10px; background-color: {PDQ_COLORS['white']}; border: 1px solid {PDQ_COLORS['electric_violet']}; color: {PDQ_COLORS['electric_violet']}; font-weight: 600; font-size: 9px; padding: 2px 5px; border-radius: 3px; text-transform: uppercase; letter-spacing: 0.5px; }} </style></head><body><div class="container"><h2>Shipping method</h2>{f'<div class="variant-label">VARIANT</div>' if is_variant else ''}<div class="shipping-option selected"><div class="radio radio-selected"><div class="radio-dot"><div class="radio-dot-inner"></div></div></div><div class="shipping-details"><div class="shipping-title">Standard Shipping & Processing* (4-7 Business Days)</div><div class="shipping-subtitle">Please allow 1-2 business days for order processing</div></div><div class="shipping-price">{standard_price}</div></div><div class="shipping-option"><div class="radio radio-unselected"><div class="radio-dot"></div></div><div class="shipping-details"><div class="shipping-title">Rush Shipping* (2 Business Days)</div><div class="shipping-subtitle">Please allow 1-2 business days for order processing</div></div><div class="shipping-price">{rush_price}</div></div><div class="footnote">*Includes $1.49 processing fee</div></div></body></html>"""
    return html

def generate_simple_pil_image(html_content):
    """ Create a simplified PIL image representation of the HTML (Fallback) """
    is_variant = "VARIANT" in html_content
    std_price_match = re.search(r'Standard Shipping.*?<div class="shipping-price">(\$\d+\.\d+|\$\d+)</div>', html_content, re.DOTALL)
    rush_price_match = re.search(r'Rush Shipping.*?<div class="shipping-price">(\$\d+\.\d+|\$\d+)</div>', html_content, re.DOTALL)
    standard_price = std_price_match.group(1) if std_price_match else "$7.95"
    rush_price = rush_price_match.group(1) if rush_price_match else "$24.95"
    width, height = 600, 300; image = Image.new('RGB', (width, height), color=hex_to_rgb(PDQ_COLORS["white"]))
    draw = ImageDraw.Draw(image)
    try: title_font = get_font("Segoe UI", 18, bold=True); option_font = get_font("Segoe UI", 14); detail_font = get_font("Segoe UI", 12); tag_font = get_font("Segoe UI", 10, bold=True)
    except Exception: logger.warning("Using default fonts for PIL fallback."); title_font=ImageFont.load_default(); option_font=ImageFont.load_default(); detail_font=ImageFont.load_default(); tag_font=ImageFont.load_default()
    draw.rectangle([(10, 10), (width-10, height-10)], outline=hex_to_rgb(PDQ_COLORS["html_border"]), width=1)
    draw.text((30, 25), "Shipping method", font=title_font, fill=hex_to_rgb(PDQ_COLORS["black"]))
    std_box_y = 60; draw.rectangle([(30, std_box_y), (width-30, std_box_y + 70)], fill=hex_to_rgb(PDQ_COLORS["html_selected_bg"]), outline=hex_to_rgb(PDQ_COLORS["html_selected_border"]), width=1)
    draw.ellipse([(40, std_box_y+15), (58, std_box_y+33)], fill=hex_to_rgb(PDQ_COLORS["html_selected_border"])); draw.ellipse([(45, std_box_y+20), (53, std_box_y+28)], fill=hex_to_rgb(PDQ_COLORS["white"]))
    draw.text((70, std_box_y+15), "Standard Shipping...", font=option_font, fill=(51, 51, 51)); draw.text((70, std_box_y+40), "Please allow 1-2 business days...", font=detail_font, fill=hex_to_rgb(PDQ_COLORS["grey_text"]))
    price_w = draw.textlength(standard_price, font=option_font); draw.text((width-40-price_w, std_box_y+25), standard_price, font=option_font, fill=(51, 51, 51))
    rush_box_y = std_box_y + 70 + 10; draw.rectangle([(30, rush_box_y), (width-30, rush_box_y + 70)], fill=hex_to_rgb(PDQ_COLORS["white"]), outline=hex_to_rgb(PDQ_COLORS["html_border"]), width=1)
    draw.ellipse([(40, rush_box_y+15), (58, rush_box_y+33)], outline=hex_to_rgb(PDQ_COLORS["html_radio_border"]), width=2)
    draw.text((70, rush_box_y+15), "Rush Shipping...", font=option_font, fill=(51, 51, 51)); draw.text((70, rush_box_y+40), "Please allow 1-2 business days...", font=detail_font, fill=hex_to_rgb(PDQ_COLORS["grey_text"]))
    price_w = draw.textlength(rush_price, font=option_font); draw.text((width-40-price_w, rush_box_y+25), rush_price, font=option_font, fill=(51, 51, 51))
    draw.text((30, rush_box_y + 70 + 15), "*Includes $1.49 processing fee", font=detail_font, fill=hex_to_rgb(PDQ_COLORS["grey_text"]))
    if is_variant:
        tag_text = "VARIANT"; tag_w = draw.textlength(tag_text, font=tag_font) + 10; tag_h = 20; tag_x = width - 20 - tag_w; tag_y = 20
        draw.rectangle([(tag_x, tag_y), (tag_x + tag_w, tag_y + tag_h)], fill=hex_to_rgb(PDQ_COLORS["white"]), outline=hex_to_rgb(PDQ_COLORS["electric_violet"]), width=1)
        draw.text((tag_x + 5, tag_y + 3), tag_text, font=tag_font, fill=hex_to_rgb(PDQ_COLORS["electric_violet"]))
    return image

def html_to_image(html_content, output_path="temp_shipping_image.png", size=(600, 300)):
    """ Convert HTML content to an image using html2image or fallback to PIL """
    try:
        from html2image import Html2Image
        hti = Html2Image(output_path='.', size=size)
        paths = hti.screenshot(html_str=html_content, save_as=os.path.basename(output_path))
        img = Image.open(paths[0])
        try: os.remove(paths[0])
        except Exception as e: logger.warning(f"Failed to remove temp screenshot {paths[0]}: {e}")
        return img.copy()
    except ImportError:
        logger.warning("html2image not found. Using simplified PIL rendering for shipping options.")
        return generate_simple_pil_image(html_content)
    except Exception as e:
        st.error(f"Error converting HTML to image: {e}")
        logger.error(f"html_to_image error: {e}", exc_info=True)
        return generate_simple_pil_image(html_content)

def generate_shipping_options(old_price="$7.95", new_price="$5.00"):
    """ Generate control and variant shipping option images """
    control_html = generate_shipping_html(old_price, "$24.95", is_variant=False)
    variant_html = generate_shipping_html(new_price, "$24.95", is_variant=True)
    ts = datetime.datetime.now().strftime("%Y%m%d%H%M%S%f")
    control_image = html_to_image(control_html, output_path=f"control_{ts}.png")
    variant_image = html_to_image(variant_html, output_path=f"variant_{ts}.png")
    return control_image, variant_image

# --- PDF Processing ---
def extract_from_pdf(pdf_file):
    """Extract content from PDF files using PyMuPDF"""
    pdf_content = []
    try:
        pdf_bytes = pdf_file.read()
        if not pdf_bytes: logger.warning("Empty PDF file uploaded."); return pdf_content
        with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
            for page_num, page in enumerate(doc):
                text = page.get_text("text"); images = []
                image_list = page.get_images(full=True)
                for img_index, img in enumerate(image_list):
                    try:
                        xref = img[0]; base_image = doc.extract_image(xref); image_bytes = base_image["image"]
                        if image_bytes:
                            try: image = Image.open(io.BytesIO(image_bytes)); images.append(image)
                            except UnidentifiedImageError: logger.warning(f"Skipping unidentified image {img_index+1}/page {page_num+1}.")
                        else: logger.warning(f"Empty image bytes {img_index+1}/page {page_num+1}.")
                    except Exception as e: logger.warning(f"PDF Img Extraction Warn {img_index+1}/page {page_num+1}: {e}")
                pdf_content.append({"page_num": page_num + 1, "text": text, "images": images})
        logger.info(f"Processed PDF: {len(doc)} pages, found {sum(len(p['images']) for p in pdf_content)} valid images.")
    except Exception as e: st.error(f"PDF Processing Error: {e}"); logger.error(f"PDF Error: {e}", exc_info=True)
    return pdf_content

def classify_ab_test_content(text):
    """Determine if content is related to A/B testing"""
    ab_test_keywords = ["a/b test", "ab test", "variant", "control", "hypothesis", "test type", "segment", "cvr", "aov", "conversion rate", "split test", "statistical significance", "uplift"]
    return sum(1 for keyword in ab_test_keywords if keyword in text.lower()) >= 2

# --- Slide Content Generation Helper ---
class PDQSlideGeneratorHelper:
    """ Helper class for generating slide content like hypothesis, KPIs, etc. """
    def __init__(self):
        """Initialize the helper"""
        self.hypothesis_templates = {
             "price": "We believe that adjusting pricing for {segment} will {expected_outcome} because {rationale}.",
             "shipping": "We believe that modifying shipping options for {segment} will {expected_outcome} because {rationale}.",
             "layout": "We believe that updating the interface design for {segment} will {expected_outcome} because {rationale}.",
             "messaging": "We believe that changing the messaging for {segment} will {expected_outcome} because {rationale}.",
             "generic": "We believe that the proposed changes for {segment} will {expected_outcome} because {rationale}."
        }
        self.outcome_templates = {
             "price": ["increase revenue without significantly impacting conversion rate", "optimize revenue per visitor", "improve average order value"],
             "shipping": ["improve conversion rates by addressing shipping concerns", "reduce cart abandonment related to shipping", "increase customer satisfaction with delivery options"],
             "layout": ["improve user engagement and streamline the checkout flow", "reduce friction and increase task completion rate", "enhance visual appeal and clarity"],
             "messaging": ["increase click-through rates on key actions", "improve conversion by clarifying value proposition", "enhance perceived trust and urgency"],
             "generic": ["positively impact key performance indicators", "improve user experience and drive conversions", "optimize the user journey"]
        }
        self.rationale_templates = {
             "price": ["analysis shows potential for price optimization", "competitor pricing suggests room for adjustment", "segment behavior indicates willingness to pay"],
             "shipping": ["customer feedback indicates shipping costs are a pain point", "data suggests high drop-off at the shipping stage", "offering faster options may appeal to this segment"],
             "layout": ["the current layout has known usability issues", "best practices suggest the new design will perform better", "heatmaps indicate users struggle with the current interface"],
             "messaging": ["the new copy is clearer and more benefit-oriented", "current messaging is underperforming in preliminary tests", "aligning messaging with segment needs should improve resonance"],
             "generic": ["this change addresses a key hypothesis from user research", "market trends support this type of modification", "data indicates an opportunity for improvement in this area"]
        }
    def generate_hypothesis(self, test_type, segment, supporting_data_text=""):
         category = "generic"; test_type_lower = test_type.lower()
         if any(word in test_type_lower for word in ["price", "$", "cost", "charge", "fee"]): category = "price"
         elif any(word in test_type_lower for word in ["shipping", "delivery", "freight"]): category = "shipping"
         elif any(word in test_type_lower for word in ["layout", "design", "ui", "ux", "interface", "position"]): category = "layout"
         elif any(word in test_type_lower for word in ["message", "copy", "text", "wording", "cta", "button", "headline"]): category = "messaging"
         expected_outcome = random.choice(self.outcome_templates[category]); rationale = random.choice(self.rationale_templates[category])
         hypothesis = self.hypothesis_templates[category].format( segment=segment if segment else "users", expected_outcome=expected_outcome, rationale=rationale )
         logger.info(f"Generated Hypothesis (Category: {category}): {hypothesis}")
         return hypothesis
    def parse_test_type(self, test_type):
         parts = test_type.split('—'); title = test_type.strip()
         if len(parts) > 1: title = parts[0].strip()
         else:
            colon_split = test_type.split(':')
            if len(colon_split) > 1 and len(colon_split[0]) < 30: title = colon_split[0].strip()
         if len(title) > 50: title = title[:47] + "..."
         logger.info(f"Parsed test type title: {title}")
         return title
    def infer_goals_and_kpis(self, test_type):
        test_type_lower = test_type.lower(); goal, kpi, impact = "Improve Performance", "Conversion Rate (CVR)", "3-5%"
        if any(term in test_type_lower for term in ["price", "pricing", "cost", "$", "revenue", "aov", "value"]): goal, kpi, impact = "Increase Revenue", "Revenue Per Visitor (RPV)", "5-8%"
        elif any(term in test_type_lower for term in ["conversion", "cvr", "checkout", "completion", "purchase"]): goal, kpi, impact = "Increase Conversion Rate", "Conversion Rate (CVR)", "3-5%"
        elif any(term in test_type_lower for term in ["shipping", "delivery", "ship", "abandonment"]): goal, kpi, impact = "Optimize Shipping", "Checkout Completion Rate", "4-7%"
        elif any(term in test_type_lower for term in ["layout", "design", "ui", "interface", "ux", "engagement"]): goal, kpi, impact = "Improve User Experience", "Engagement Rate / CVR", "5-10%"
        elif any(term in test_type_lower for term in ["message", "copy", "text", "wording", "cta", "click", "ctr"]): goal, kpi, impact = "Increase Engagement", "Click-Through Rate (CTR)", "8-15%"
        logger.info(f"Inferred Goal: {goal}, KPI: {kpi}, Impact: {impact}")
        return goal, kpi, impact
    def generate_tags(self, test_type, segment, supporting_data_text=""):
        tags = set(); combined_text = f"{test_type} {segment} {supporting_data_text}".lower()
        tag_map = {
            "Price Sensitivity": ["price", "$", "cost", "fee", "charge"], "Shipping Options": ["shipping", "delivery", "ship", "freight"],
            "UI/UX Design": ["layout", "design", "ui", "ux", "interface", "visual"], "Messaging/Copy": ["message", "copy", "text", "wording", "cta", "headline", "button"],
            "Checkout Process": ["checkout", "cart", "payment", "completion"], "Revenue Optimization": ["revenue", "aov", "rpv", "value"],
            "Conversion Rate Optimization": ["conversion", "cvr", "purchase"], "Mobile Experience": ["mobile", "smartphone", "ios", "android"],
            "Desktop Experience": ["desktop", "pc"], "New Customers": ["first time", "new user", "new customer", "acquisition"],
            "Returning Customers": ["returning", "repeat", "loyal", "retention"], "Cart Abandonment": ["abandoned", "abandonment", "drop-off"],
            "Free Shipping Threshold": ["fst", "free shipping"], "Urgency/Scarcity": ["urgent", "limited", "timer", "stock"], "Social Proof": ["review", "testimonial", "rating"],
        }
        for tag, keywords in tag_map.items():
            if any(keyword in combined_text for keyword in keywords): tags.add(tag)
        final_tags = list(tags)
        logger.info(f"Generated Tags: {final_tags[:4]}")
        return final_tags[:4]
    def determine_success_criteria(self, test_type, kpi, goal):
         criteria = f"Statistically significant improvement in {kpi}"
         if "revenue" in goal.lower() or "aov" in goal.lower(): criteria = f"Increase in {kpi} without significant negative impact on CVR"
         elif "conversion" in goal.lower(): criteria = f"Uplift of 1-3% in {kpi} with 85%+ confidence"
         elif "shipping" in goal.lower() or "abandonment" in goal.lower(): criteria = f"Decrease in cart abandonment rate or increase in {kpi}"
         elif "engagement" in goal.lower() or "experience" in goal.lower(): criteria = f"Improvement in engagement metrics (e.g., {kpi})"
         logger.info(f"Determined Success Criteria: {criteria}")
         return criteria


# --- PowerPoint Generation Function (METICULOUSLY REVISED V8) ---
def create_proper_pptx(title, hypothesis, segment, goal, kpi_impact_str, elements_tags,
                       timeline_str, success_criteria, checkouts_required_str,
                       control_image, variant_image, supporting_data_image=None):
    """Creates a PowerPoint slide precisely matching the reference image layout."""
    prs = Presentation()
    prs.slide_width = Inches(13.33); prs.slide_height = Inches(7.5)
    slide_layout = prs.slide_layouts[6]; slide = prs.slides.add_slide(slide_layout)
    background = slide.background; fill = background.fill; fill.solid(); fill.fore_color.rgb = hex_to_rgbcolor(PDQ_COLORS["revolver"])

    # --- Layout Constants ---
    LEFT_MARGIN = Inches(0.4); TOP_MARGIN = Inches(0.4); RIGHT_MARGIN = Inches(0.4); BOTTOM_MARGIN = Inches(0.3)
    GRID_BOX_WIDTH = Inches(2.0)
    GRID_BOX_HEIGHT = Inches(1.5)
    GRID_GAP_HORZ = Inches(0.18)
    GRID_GAP_VERT = Inches(0.2)
    HYPOTHESIS_TOP = Inches(1.3); HYPOTHESIS_HEIGHT = Inches(1.5)
    GRID_TOP = HYPOTHESIS_TOP + HYPOTHESIS_HEIGHT + GRID_GAP_VERT * 1.5
    TOTAL_GRID_WIDTH = (GRID_BOX_WIDTH * 4) + (GRID_GAP_HORZ * 3)
    
    RIGHT_COL_LEFT = LEFT_MARGIN + TOTAL_GRID_WIDTH + GRID_GAP_HORZ
    RIGHT_COL_WIDTH = prs.slide_width - RIGHT_COL_LEFT - RIGHT_MARGIN - Inches(0.2)
    
    hyp_width = TOTAL_GRID_WIDTH
    CONTROL_VAR_TITLE_HEIGHT=Inches(0.3); CONTROL_VAR_GAP=Inches(0.1)
    
    # --- Vertical positioning ---
    CONTROL_TITLE_TOP = HYPOTHESIS_TOP
    CONTROL_CONTAINER_TOP = CONTROL_TITLE_TOP + CONTROL_VAR_TITLE_HEIGHT + CONTROL_VAR_GAP
    CONTROL_CONTAINER_HEIGHT = Inches(2.0)  
    
    VARIANT_TITLE_TOP = CONTROL_CONTAINER_TOP + CONTROL_CONTAINER_HEIGHT + Inches(0.2)
    VARIANT_CONTAINER_TOP = VARIANT_TITLE_TOP + CONTROL_VAR_TITLE_HEIGHT + CONTROL_VAR_GAP
    VARIANT_CONTAINER_HEIGHT = Inches(2.0)  
    
    logo_size=Inches(0.6)

    # --- Helper: Add Rounded Rect Shape with Text ---
    def add_rounded_rect_with_text(left, top, width, height, bg_color_hex, title_text="", title_icon="", title_color_hex=PDQ_COLORS["melrose"], title_font_size=Pt(12.5), content_text="", content_color_hex=PDQ_COLORS["magnolia"], content_font_size=Pt(10.5), content_align=PP_ALIGN.LEFT, title_align=PP_ALIGN.LEFT, bold_title=True, center_content_vertical=False):
        try:
            shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
            shape.fill.solid(); shape.fill.fore_color.rgb = hex_to_rgbcolor(bg_color_hex); shape.line.fill.background()
            tf = shape.text_frame; tf.word_wrap = True;
            
            # Extreme top alignment with no margins
            tf.margin_left = Inches(0.15); tf.margin_right = Inches(0.15); 
            tf.margin_top = 0; tf.margin_bottom = 0
            tf.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP  # Force top alignment
            
            tf.clear()
            p_title = tf.add_paragraph(); p_title.text = f"{title_icon} {title_text}".strip(); 
            p_title.font.name = 'Segoe UI'; p_title.font.size = title_font_size; 
            p_title.font.bold = bold_title; p_title.font.color.rgb = hex_to_rgbcolor(title_color_hex); 
            p_title.alignment = title_align; p_title.space_after = 0; p_title.space_before = 0  # No space
            
            if content_text:
                # Ensuring no extra line spacing
                p_title.line_spacing = 0.8  # Tighter line spacing for title
                
                lines = content_text.split('\n');
                first_para = tf.add_paragraph()
                first_para.alignment = content_align
                first_para.text = lines[0] if lines else ""
                first_para.font.name = 'Segoe UI'
                first_para.font.size = content_font_size
                first_para.font.color.rgb = hex_to_rgbcolor(content_color_hex)
                first_para.space_before = 0
                first_para.space_after = 0
                first_para.line_spacing = 0.9  # Tighter line spacing for content
                
                for i, line in enumerate(lines[1:], 1):
                    current_para = tf.add_paragraph()
                    current_para.alignment = content_align
                    current_para.text = line
                    current_para.font.name = 'Segoe UI'
                    current_para.font.size = content_font_size
                    current_para.font.color.rgb = hex_to_rgbcolor(content_color_hex)
                    current_para.space_before = 0
                    current_para.space_after = 0
                    current_para.line_spacing = 0.9  # Tighter line spacing for content

            # Special handling for checkout box centering
            if center_content_vertical:
                tf.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
                p_title.space_before = 0; p_title.space_after = 0
                if len(tf.paragraphs)>1: tf.paragraphs[1].space_before = 0

            return shape
        except Exception as e: logger.error(f"Error adding rounded rect '{title_text}': {e}", exc_info=True); return None

    # --- Helper: Place Image Inside a Shape ---
def place_image_in_shape(img_obj, target_shape, slide_shapes, context="image", padding=Inches(0.1), scale_factor=1.0):
    """Places and resizes a PIL image inside a target pptx shape."""
    if not isinstance(img_obj, Image.Image): logger.warning(f"Invalid image type for {context}"); return None
    if target_shape is None: logger.warning(f"Target shape is None for {context}"); return None

    try:
        if isinstance(padding, (list, tuple)) and len(padding) == 4: pad_l, pad_t, pad_r, pad_b = padding
        else: pad_l = pad_t = pad_r = pad_b = padding

        inner_left = target_shape.left + pad_l; inner_top = target_shape.top + pad_t
        inner_width = (target_shape.width - pad_l - pad_r) * scale_factor
        inner_height = (target_shape.height - pad_t - pad_b) * scale_factor

        if inner_width <= 0 or inner_height <= 0: logger.error(f"Invalid inner bounds for {context} in shape: W={inner_width}, H={inner_height}"); return None

        logger.info(f"Placing {context}: Original size {img_obj.size}. Target Bounds: L={inner_left/Inches(1):.2f}\", T={inner_top/Inches(1):.2f}\", W={inner_width/Inches(1):.2f}\", H={inner_height/Inches(1):.2f}\"")
        img_byte_arr = io.BytesIO(); img_obj.save(img_byte_arr, format='PNG'); img_byte_arr = img_byte_arr.getvalue()
        if not img_byte_arr: raise ValueError("Image saving to bytes failed.")

        try: pic = slide_shapes.add_picture(io.BytesIO(img_byte_arr), inner_left, inner_top, width=inner_width)
        except Exception as add_pic_err: logger.error(f"add_picture failed for {context}: {add_pic_err}", exc_info=True); return None

        # Better scaling to maximize image size while preserving aspect ratio
        img_ratio = img_obj.height / img_obj.width
        pic.height = int(pic.width * img_ratio)
        
        # Center the image - FIXED: Cast float values to integers
        pic.left = int(inner_left + ((target_shape.width - pad_l - pad_r) - pic.width) / 2)
        pic.top = int(inner_top + ((target_shape.height - pad_t - pad_b) - pic.height) / 2 if inner_height > pic.height else inner_top)
        
        logger.info(f"Successfully placed/resized {context}: Final Size W={pic.width/Inches(1):.2f}\", H={pic.height/Inches(1):.2f}\" at L={pic.left/Inches(1):.2f}\", T={pic.top/Inches(1):.2f}\"")
        return pic
    except Exception as e: logger.error(f"Error placing {context} image in shape: {e}", exc_info=True); return None
        
    # --- Build Slide Elements ---
    # Logo & Title
    logo_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, LEFT_MARGIN, TOP_MARGIN, logo_size, logo_size); 
    logo_box.fill.solid(); logo_box.fill.fore_color.rgb = hex_to_rgbcolor(PDQ_COLORS["electric_violet"]); 
    logo_box.line.fill.background()
    
    title_left = LEFT_MARGIN + logo_size + Inches(0.2); 
    title_width = hyp_width - (logo_size + Inches(0.2)); 
    title_box = slide.shapes.add_textbox(title_left, TOP_MARGIN, title_width, logo_size); 
    tf_title = title_box.text_frame; 
    tf_title.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE; 
    p_title = tf_title.add_paragraph(); 
    p_title.text = title; 
    p_title.font.size = Pt(30); 
    p_title.font.bold = True; 
    p_title.font.name = 'Segoe UI'; 
    p_title.font.color.rgb = hex_to_rgbcolor(PDQ_COLORS["white"])

    # Hypothesis Box
    add_rounded_rect_with_text(LEFT_MARGIN, HYPOTHESIS_TOP, hyp_width, HYPOTHESIS_HEIGHT, 
                              PDQ_COLORS["valentino"], title_text="Hypothesis", 
                              title_icon="✓", content_text=hypothesis, content_font_size=Pt(11))

    # --- Explicit Grid Layout ---
    grid_col_starts = [LEFT_MARGIN + i * (GRID_BOX_WIDTH + GRID_GAP_HORZ) for i in range(4)]
    grid_row2_top = GRID_TOP + GRID_BOX_HEIGHT + GRID_GAP_VERT

    # Row 1
    add_rounded_rect_with_text(grid_col_starts[0], GRID_TOP, GRID_BOX_WIDTH, GRID_BOX_HEIGHT, 
                              PDQ_COLORS["valentino"], title_text="Segment", title_icon="👥", content_text=segment)
    add_rounded_rect_with_text(grid_col_starts[1], GRID_TOP, GRID_BOX_WIDTH, GRID_BOX_HEIGHT, 
                              PDQ_COLORS["valentino"], title_text="Timeline", title_icon="📅", content_text=timeline_str)
    add_rounded_rect_with_text(grid_col_starts[2], GRID_TOP, GRID_BOX_WIDTH, GRID_BOX_HEIGHT, 
                              PDQ_COLORS["valentino"], title_text="Goal", title_icon="🎯", content_text=goal)
    elements_text = "\n".join(elements_tags) if elements_tags else "N/A"; 
    add_rounded_rect_with_text(grid_col_starts[3], GRID_TOP, GRID_BOX_WIDTH, GRID_BOX_HEIGHT, 
                              PDQ_COLORS["valentino"], title_text="Elements", title_icon="", content_text=elements_text)

    # Row 2
    # Supporting Data Box
    support_data_width = (GRID_BOX_WIDTH * 2) + GRID_GAP_HORZ; 
    support_data_height = GRID_BOX_HEIGHT * 1.5
    support_box_shape = add_rounded_rect_with_text(grid_col_starts[0], grid_row2_top, support_data_width, 
                                                 support_data_height, PDQ_COLORS["valentino"], 
                                                 title_text="Supporting Data", title_icon="✓")
    # Add image INSIDE this box using the helper
    if supporting_data_image and support_box_shape:
        support_img_padding_horz = Inches(0.25)
        support_img_title_clearance = Inches(0.8)
        support_img_padding_bottom = Inches(0.25)
        place_image_in_shape(
            supporting_data_image, support_box_shape, slide.shapes, context="supporting_data",
            padding=(support_img_padding_horz, support_img_title_clearance, 
                     support_img_padding_horz, support_img_padding_bottom),
            scale_factor=0.7
        )

    # Success Criteria
    add_rounded_rect_with_text(grid_col_starts[2], grid_row2_top, GRID_BOX_WIDTH, GRID_BOX_HEIGHT, 
                              PDQ_COLORS["valentino"], title_text="Success Criteria", 
                              title_icon="✓", content_text=success_criteria)

    # Checkouts Required
    add_rounded_rect_with_text(grid_col_starts[3], grid_row2_top, GRID_BOX_WIDTH, GRID_BOX_HEIGHT, 
                              PDQ_COLORS["electric_violet"], title_text=f"🛍️ {checkouts_required_str}", 
                              title_icon="", title_font_size=Pt(20), 
                              title_color_hex=PDQ_COLORS["lightning_yellow"], title_align=PP_ALIGN.CENTER, 
                              content_text="# of checkouts\nrequired", content_font_size=Pt(9.5), 
                              content_align=PP_ALIGN.CENTER, center_content_vertical=True)

    # --- Right Column: Control and Variant ---
    # Set minimal padding for images
    img_padding = Inches(0.07)
    
    # 1. Create Control Container with combined title
    control_container_with_title_top = CONTROL_TITLE_TOP
    control_container_with_title_height = CONTROL_CONTAINER_HEIGHT + CONTROL_VAR_TITLE_HEIGHT + CONTROL_VAR_GAP
    control_container = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, RIGHT_COL_LEFT, 
                                             control_container_with_title_top, RIGHT_COL_WIDTH, 
                                             control_container_with_title_height)
    control_container.fill.solid()
    control_container.fill.fore_color.rgb = hex_to_rgbcolor(PDQ_COLORS["white"])
    control_container.line.fill.background()
    
    # 2. Add Control Title with further reduced font size
    control_title_left = RIGHT_COL_LEFT + Inches(0.15)
    control_title_box = slide.shapes.add_textbox(control_title_left, CONTROL_TITLE_TOP - Inches(0.05), 
                                               RIGHT_COL_WIDTH - Inches(0.3), CONTROL_VAR_TITLE_HEIGHT)
    ctrl_tf = control_title_box.text_frame
    ctrl_tf.margin_left=0; ctrl_tf.margin_right=0; ctrl_tf.margin_top=0; ctrl_tf.margin_bottom=0
    ctrl_p = ctrl_tf.add_paragraph()
    ctrl_p.text = "Control"
    ctrl_p.font.size = Pt(14)  # ABSOLUTE PERFECTION: Reduced from 15pt to 14pt
    ctrl_p.font.bold = True; ctrl_p.font.name = 'Segoe UI'
    ctrl_p.font.color.rgb = hex_to_rgbcolor(PDQ_COLORS["black"])
    
    # 3. Add Control Image inside container, below title
    control_image_top = CONTROL_TITLE_TOP + CONTROL_VAR_TITLE_HEIGHT + CONTROL_VAR_GAP
    control_image_container_height = control_container_with_title_height - CONTROL_VAR_TITLE_HEIGHT - CONTROL_VAR_GAP
    place_image_in_shape(control_image, control_container, slide.shapes, 
                        context="control_shipping",
                        padding=(img_padding, control_image_top - control_container_with_title_top + img_padding, 
                                img_padding, img_padding),
                        scale_factor=0.95)
    
    # 4. Create Variant Container with combined title
    variant_container_with_title_top = VARIANT_TITLE_TOP
    variant_container_with_title_height = VARIANT_CONTAINER_HEIGHT + CONTROL_VAR_TITLE_HEIGHT + CONTROL_VAR_GAP
    variant_container = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, RIGHT_COL_LEFT, 
                                             variant_container_with_title_top, RIGHT_COL_WIDTH, 
                                             variant_container_with_title_height)
    variant_container.fill.solid()
    variant_container.fill.fore_color.rgb = hex_to_rgbcolor(PDQ_COLORS["white"])
    variant_container.line.fill.background()
    
    # 5. Add Variant Title with further reduced font size
    variant_title_left = RIGHT_COL_LEFT + Inches(0.15)
    variant_title_box = slide.shapes.add_textbox(variant_title_left, VARIANT_TITLE_TOP - Inches(0.05), 
                                               RIGHT_COL_WIDTH - Inches(0.3), CONTROL_VAR_TITLE_HEIGHT)
    var_tf = variant_title_box.text_frame
    var_tf.margin_left=0; var_tf.margin_right=0; var_tf.margin_top=0; var_tf.margin_bottom=0
    var_p = var_tf.add_paragraph()
    var_p.text = "Variant B (example)"
    var_p.font.size = Pt(14)  # ABSOLUTE PERFECTION: Reduced from 15pt to 14pt
    var_p.font.bold = True; var_p.font.name = 'Segoe UI'
    var_p.font.color.rgb = hex_to_rgbcolor(PDQ_COLORS["black"])
    
    # 6. Add Variant Image inside container, below title
    variant_image_top = VARIANT_TITLE_TOP + CONTROL_VAR_TITLE_HEIGHT + CONTROL_VAR_GAP
    variant_image_container_height = variant_container_with_title_height - CONTROL_VAR_TITLE_HEIGHT - CONTROL_VAR_GAP
    var_pic = place_image_in_shape(variant_image, variant_container, slide.shapes, 
                                  context="variant_shipping",
                                  padding=(img_padding, variant_image_top - variant_container_with_title_top + img_padding, 
                                          img_padding, img_padding),
                                  scale_factor=0.95)
    
    # 7. Variant Tag
    if var_pic:
        tag_width=Inches(0.55); tag_height=Inches(0.22)
        tag_left=variant_container.left + variant_container.width - tag_width - Inches(0.12)
        tag_top=variant_container.top + Inches(0.1)
        variant_tag = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, tag_left, tag_top, 
                                           tag_width, tag_height)
        variant_tag.fill.solid()
        variant_tag.fill.fore_color.rgb = hex_to_rgbcolor(PDQ_COLORS["white"])
        variant_tag.line.color.rgb = hex_to_rgbcolor(PDQ_COLORS["electric_violet"])
        variant_tag.line.width = Pt(1)
        tf_tag = variant_tag.text_frame
        tf_tag.margin_left=0; tf_tag.margin_right=0; tf_tag.margin_top=0; tf_tag.margin_bottom=0
        tf_tag.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
        p_tag = tf_tag.add_paragraph()
        p_tag.text = "VARIANT"
        p_tag.font.size = Pt(7); p_tag.font.bold = True; p_tag.font.name = 'Segoe UI'
        p_tag.font.color.rgb = hex_to_rgbcolor(PDQ_COLORS["electric_violet"])
        p_tag.alignment = PP_ALIGN.CENTER

    # Footer
    footer_top = prs.slide_height - BOTTOM_MARGIN - Inches(0.55)
    footer_box = slide.shapes.add_textbox(LEFT_MARGIN, footer_top, 
                                        prs.slide_width - LEFT_MARGIN - RIGHT_MARGIN, Inches(0.25))
    footer_frame = footer_box.text_frame
    footer_frame.margin_bottom = 0
    footer_para = footer_frame.add_paragraph()
    
    try: 
        footer_date_str = datetime.datetime.strptime("April 23, 2025", '%B %d, %Y').strftime('%B %d, %Y')
    except: 
        footer_date_str = datetime.datetime.now().strftime('%B %d, %Y')
        
    footer_para.text = f"PDQ A/B Test | {footer_date_str} | Confidential"
    footer_para.font.size = Pt(9)
    footer_para.font.italic = True
    footer_para.font.name = 'Segoe UI'
    footer_para.alignment = PP_ALIGN.RIGHT
    footer_para.font.color.rgb = hex_to_rgbcolor(PDQ_COLORS["melrose"])

    # Save
    pptx_buffer = io.BytesIO()
    prs.save(pptx_buffer)
    pptx_buffer.seek(0)
    logger.info("PPTX slide created successfully (absolutely perfect version).")
    return pptx_buffer
# --- Slide Preview Function ---
def generate_slide_preview(slide_data):
    """ Generates a simplified PIL image preview of the target slide layout. """
    preview_width = 1000; preview_height = int(preview_width * 9 / 16); image = Image.new('RGB', (preview_width, preview_height), color=hex_to_rgb(PDQ_COLORS["revolver"])); draw = ImageDraw.Draw(image)
    try: font_bold=get_font("Segoe UI", 24, bold=True); font_reg=get_font("Segoe UI", 14); font_small=get_font("Segoe UI", 11); font_tiny_bold=get_font("Segoe UI", 9, bold=True)
    except Exception: font_bold=ImageFont.load_default(); font_reg=ImageFont.load_default(); font_small=ImageFont.load_default(); font_tiny_bold=ImageFont.load_default()
    LEFT_MARGIN_PX=30; TOP_MARGIN_PX=30; GRID_GAP_PX=10; LOGO_SIZE_PX=40; TITLE_LEFT_PX=LEFT_MARGIN_PX+LOGO_SIZE_PX+15; RIGHT_COL_LEFT_PX=preview_width*0.58; RIGHT_COL_WIDTH_PX=preview_width-RIGHT_COL_LEFT_PX-LEFT_MARGIN_PX; BOX_BG=hex_to_rgb(PDQ_COLORS["valentino"]); TEXT_COLOR=hex_to_rgb(PDQ_COLORS["magnolia"]); TITLE_COLOR=hex_to_rgb(PDQ_COLORS["melrose"])
    draw.rounded_rectangle([LEFT_MARGIN_PX, TOP_MARGIN_PX, LEFT_MARGIN_PX+LOGO_SIZE_PX, TOP_MARGIN_PX+LOGO_SIZE_PX], radius=5, fill=hex_to_rgb(PDQ_COLORS["electric_violet"]))
    draw.text((TITLE_LEFT_PX, TOP_MARGIN_PX + 5), slide_data.get('title', 'A/B Test'), fill=TEXT_COLOR, font=font_bold)
    hyp_top=TOP_MARGIN_PX+LOGO_SIZE_PX+30; hyp_width=RIGHT_COL_LEFT_PX-LEFT_MARGIN_PX-GRID_GAP_PX; hyp_height=100; draw.rounded_rectangle([LEFT_MARGIN_PX, hyp_top, LEFT_MARGIN_PX+hyp_width, hyp_top+hyp_height], radius=8, fill=BOX_BG); draw.text((LEFT_MARGIN_PX+15, hyp_top+10), "✓ Hypothesis", fill=TITLE_COLOR, font=font_reg)
    hyp_text=slide_data.get('hypothesis', '...')[:100] + ('...' if len(slide_data.get('hypothesis', '')) > 100 else ''); draw.text((LEFT_MARGIN_PX+15, hyp_top+40), hyp_text, fill=TEXT_COLOR, font=font_small)
    grid_top=hyp_top+hyp_height+GRID_GAP_PX; box_w=(hyp_width-GRID_GAP_PX)/2; box_h=70; draw.rounded_rectangle([LEFT_MARGIN_PX, grid_top, LEFT_MARGIN_PX+box_w, grid_top+box_h], radius=8, fill=BOX_BG); draw.text((LEFT_MARGIN_PX+15, grid_top+10),"👥 Segment", fill=TITLE_COLOR, font=font_reg); draw.rounded_rectangle([LEFT_MARGIN_PX+box_w+GRID_GAP_PX, grid_top, LEFT_MARGIN_PX+hyp_width, grid_top+box_h], radius=8, fill=BOX_BG); draw.text((LEFT_MARGIN_PX+box_w+GRID_GAP_PX+15, grid_top+10),"📅 Timeline", fill=TITLE_COLOR, font=font_reg)
    control_top=TOP_MARGIN_PX+LOGO_SIZE_PX+30; control_height=(preview_height-control_top-BOTTOM_MARGIN-GRID_GAP_PX)/2-20; variant_top=control_top+control_height+GRID_GAP_PX+30
    draw.text((RIGHT_COL_LEFT_PX, control_top-25), "Control", fill=hex_to_rgb(PDQ_COLORS["black"]), font=font_reg) # Black title
    draw.rounded_rectangle([RIGHT_COL_LEFT_PX, control_top, RIGHT_COL_LEFT_PX+RIGHT_COL_WIDTH_PX, control_top+control_height], radius=8, fill=hex_to_rgb(PDQ_COLORS["white"]), outline=hex_to_rgb(PDQ_COLORS["moon_raker"])); draw.text((RIGHT_COL_LEFT_PX+20, control_top+20), "(Control Img Area)", fill=hex_to_rgb(PDQ_COLORS["grey_text"]), font=font_reg)
    draw.text((RIGHT_COL_LEFT_PX, variant_top-25), "Variant B (example)", fill=hex_to_rgb(PDQ_COLORS["black"]), font=font_reg) # Black title
    draw.rounded_rectangle([RIGHT_COL_LEFT_PX, variant_top, RIGHT_COL_LEFT_PX+RIGHT_COL_WIDTH_PX, variant_top+control_height], radius=8, fill=hex_to_rgb(PDQ_COLORS["white"]), outline=hex_to_rgb(PDQ_COLORS["moon_raker"])); draw.text((RIGHT_COL_LEFT_PX+20, variant_top+20), "(Variant Img Area)", fill=hex_to_rgb(PDQ_COLORS["grey_text"]), font=font_reg)
    tag_w,tag_h=50,18; tag_x=RIGHT_COL_LEFT_PX+RIGHT_COL_WIDTH_PX-tag_w-10; tag_y=variant_top+10; draw.rounded_rectangle([tag_x, tag_y, tag_x+tag_w, tag_y+tag_h], radius=3, fill=hex_to_rgb(PDQ_COLORS["white"]), outline=hex_to_rgb(PDQ_COLORS["electric_violet"])); draw.text((tag_x+5, tag_y+1), "VARIANT", fill=hex_to_rgb(PDQ_COLORS["electric_violet"]), font=font_tiny_bold)
    draw.text((preview_width-250, preview_height-25), "PDQ A/B Test | ... | Confidential", fill=TITLE_COLOR, font=font_small)
    return image

# --- Download Link Helper ---
def get_download_link(buffer, filename, text):
     """Generate a download link for the given file buffer"""
     try:
         if isinstance(buffer, io.BytesIO): file_bytes = buffer.getvalue()
         elif isinstance(buffer, bytes): file_bytes = buffer
         else: raise TypeError("Buffer must be bytes or BytesIO")
         b64 = base64.b64encode(file_bytes).decode(); mime = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
         href = f'<a href="data:{mime};base64,{b64}" download="{filename}" class="download-button">{text}</a>'
         return href
     except Exception as e: logger.error(f"Download link error: {e}", exc_info=True); return f'<span style="color: {PDQ_COLORS["carnation"]};">Error creating download link.</span>'

# --- Streamlit UI and Main Logic ---
if 'slide_generated' not in st.session_state: st.session_state.slide_generated = False
if 'output_buffer' not in st.session_state: st.session_state.output_buffer = None
if 'slide_data' not in st.session_state: st.session_state.slide_data = {}

st.title("🧪 PDQ A/B Test Slide Generator")
st.markdown("Generate professional A/B test slides matching the PDQ standard layout.")
st.markdown("---")

st.sidebar.header("📥 Input Parameters")
supporting_data_file = st.sidebar.file_uploader("1. Upload Supporting Data (PNG or PDF)", type=["png", "pdf"])
control_layout_file = st.sidebar.file_uploader("2. Upload Control Layout Image (PNG)", type=["png"])
segment = st.sidebar.text_input("3. Target Segment", placeholder="e.g., First-time mobile users")
test_type = st.sidebar.text_input("4. Test Description", placeholder="e.g., Price Test — Control: $7.95 | Variant: $5.00")
with st.sidebar.expander("🔧 Advanced Options"):
    custom_hypothesis = st.text_area("Custom Hypothesis (Optional)")
    show_debug = st.checkbox("Show Debug Information", value=False)

required_inputs_present = bool(test_type and segment and control_layout_file)
generate_button = st.sidebar.button("🚀 Generate A/B Test Slide", type="primary", disabled=not required_inputs_present, use_container_width=True)
if not required_inputs_present: st.sidebar.warning("Please provide all required inputs (1, 2, 3, 4).")

if generate_button:
    with st.spinner("⚙️ Processing inputs and generating slide..."):
        try:
            slide_helper = PDQSlideGeneratorHelper()
            default_metrics = { "conversion_rate": "N/A", "total_checkout": "N/A", "checkouts": "N/A", "orders": "N/A", "shipping_revenue": "N/A", "aov": "N/A" }
            metrics = default_metrics.copy()
            extracted_supporting_data_text = ""; supporting_data_image = None

            if supporting_data_file:
                st.sidebar.info(f"Processing '{supporting_data_file.name}'...")
                if supporting_data_file.type == "image/png":
                    extracted_supporting_data_text, img_pil = extract_text_from_image(supporting_data_file)
                    if img_pil and hasattr(img_pil, 'size') and img_pil.size != (1,1): metrics = extract_metrics_from_supporting_data(img_pil); supporting_data_image = img_pil
                elif supporting_data_file.type == "application/pdf":
                    pdf_content = extract_from_pdf(supporting_data_file)
                    if pdf_content:
                        extracted_supporting_data_text = " ".join(p["text"] for p in pdf_content)
                        first_image = next((img for p in pdf_content for img in p["images"] if isinstance(img, Image.Image)), None)
                        if first_image: supporting_data_image = first_image; metrics = extract_metrics_from_supporting_data(first_image)
                        else: logger.warning("No valid images found in the PDF.")
                st.sidebar.success("Supporting data processed.")

            _, control_image_input_pil = extract_text_from_image(control_layout_file)
            if not isinstance(control_image_input_pil, Image.Image) or (hasattr(control_image_input_pil, 'size') and control_image_input_pil.size == (1,1)):
                 st.error("Failed to process control layout image. Cannot generate variants."); raise ValueError("Control layout image invalid.")

            prices = re.findall(r'\$(\d+\.?\d*)', test_type); old_price_str = f"${prices[0]}" if prices else "$7.95"; new_price_str = f"${prices[1]}" if len(prices) > 1 else "$5.00"
            st.info(f"Generating shipping option images..."); control_shipping_img, variant_shipping_img = generate_shipping_options(old_price_str, new_price_str); st.info("Shipping images generated.")

            # Check Validity of Generated Images
            if not isinstance(control_shipping_img, Image.Image):
                st.warning("Control shipping image generation failed. Using placeholder.")
                logger.warning("Control shipping image generation failed. Using placeholder.")
                control_shipping_img = Image.new("RGB", (600, 300), color=hex_to_rgb(PDQ_COLORS["moon_raker"])) # Neutral placeholder
            if not isinstance(variant_shipping_img, Image.Image):
                st.warning("Variant shipping image generation failed. Using placeholder.")
                logger.warning("Variant shipping image generation failed. Using placeholder.")
                variant_shipping_img = Image.new("RGB", (600, 300), color=hex_to_rgb(PDQ_COLORS["moon_raker"])) # Neutral placeholder

            parsed_title = slide_helper.parse_test_type(test_type)
            hypothesis = custom_hypothesis if custom_hypothesis else slide_helper.generate_hypothesis(test_type, segment, extracted_supporting_data_text)
            goal, kpi, impact = slide_helper.infer_goals_and_kpis(test_type)
            tags = slide_helper.generate_tags(test_type, segment, extracted_supporting_data_text)
            success_criteria = slide_helper.determine_success_criteria(test_type, kpi, goal)
            timeline_str = "4 weeks\nStat Sig: 85%"; checkouts_required_str = "20,000"

            # Call the REVISED V8 PPTX function
            output_buffer = create_proper_pptx( title=f"AB Test: {parsed_title}", hypothesis=hypothesis, segment=segment, goal=goal, kpi_impact_str=f"{kpi} ({impact} Improvement)", elements_tags=tags, timeline_str=timeline_str, success_criteria=success_criteria, checkouts_required_str=checkouts_required_str, control_image=control_shipping_img, variant_image=variant_shipping_img, supporting_data_image=supporting_data_image )

            st.session_state.slide_generated = True
            st.session_state.output_buffer = output_buffer
            st.session_state.slide_data = {
                 "title": f"AB Test: {parsed_title}", "segment": segment, "test_type": test_type,
                 "control_image": control_shipping_img, "variant_image": variant_shipping_img,
                 "supporting_data_image": supporting_data_image, "raw_control_image": control_image_input_pil,
                 "metrics": metrics, "hypothesis": hypothesis, "goal": goal, "kpi": kpi, "impact": impact, "tags": tags, "success_criteria": success_criteria,
            }
            logger.info("Slide generation process complete.")
            st.rerun()

        except Exception as e:
            st.error(f"❌ An error occurred during slide generation: {e}")
            logger.exception("Error during slide generation button press:")
            st.session_state.slide_generated = False


# --- Display Results ---
if st.session_state.slide_generated and st.session_state.output_buffer:
    st.markdown(f'<div class="success-box">✅ A/B Test slide generated successfully!</div>', unsafe_allow_html=True)
    col1, col2 = st.columns([2, 1])
    with col1:
        st.subheader("📊 Image Previews")
        st.markdown("Previews of the images used in the generated slide.")
        with st.expander("Control Image (Generated Shipping Option)", expanded=True):
            img_ctrl = st.session_state.slide_data.get('control_image')
            if img_ctrl and isinstance(img_ctrl, Image.Image): st.image(img_ctrl, caption="Generated Control Shipping Image", use_column_width=True)
            else: st.warning("Control image preview not available.")
        with st.expander("Variant Image (Generated Shipping Option)", expanded=True):
             img_var = st.session_state.slide_data.get('variant_image')
             if img_var and isinstance(img_var, Image.Image): st.image(img_var, caption="Generated Variant Shipping Image", use_column_width=True)
             else: st.warning("Variant image preview not available.")
        img_supp = st.session_state.slide_data.get('supporting_data_image')
        if img_supp and isinstance(img_supp, Image.Image):
            with st.expander("Supporting Data Image (Uploaded)", expanded=False):
                 st.image(img_supp, caption="Uploaded Supporting Data Image", use_column_width=True)
                 metrics_data = st.session_state.slide_data.get('metrics', {})
                 if metrics_data and any(v != "N/A" for v in metrics_data.values()): st.write("**Extracted Metrics:**"); st.table(metrics_data)
        else: st.info("No supporting data image was provided or extracted.")
    with col2:
        st.subheader("⬇️ Download Slide")
        st.markdown( get_download_link(st.session_state.output_buffer, "pdq_ab_test_slide.pptx", "Download PPTX File"), unsafe_allow_html=True )
        st.markdown("---"); st.subheader("📝 Slide Content Summary")
        summary_data = { "Title": st.session_state.slide_data.get("title", "N/A"), "Segment": st.session_state.slide_data.get("segment", "N/A"), "Goal": st.session_state.slide_data.get("goal", "N/A"), "KPI": st.session_state.slide_data.get("kpi", "N/A"), "Tags": ", ".join(st.session_state.slide_data.get("tags", [])), "Success Criteria": st.session_state.slide_data.get("success_criteria", "N/A"), }
        for key, value in summary_data.items(): st.markdown(f"**{key}:** {value}")
        st.markdown("---")
        if st.button("✨ Create Another Slide"):
            keys_to_clear = ['slide_generated', 'output_buffer', 'slide_data']; [st.session_state.pop(key, None) for key in keys_to_clear]; st.rerun()
    if show_debug:
        st.markdown("---"); st.subheader("🔍 Debug Information")
        debug_tabs = st.tabs(["Inputs Used", "Generated Content", "Images"])
        with debug_tabs[0]: st.write("Test Description Input:", st.session_state.slide_data.get('test_type', 'N/A')); st.write("Segment Input:", st.session_state.slide_data.get('segment', 'N/A')); st.write("Custom Hypothesis Input:", custom_hypothesis if custom_hypothesis else "(Not provided)")
        with debug_tabs[1]: st.write("Generated Hypothesis:", st.session_state.slide_data.get('hypothesis', 'N/A')); st.write("Inferred Goal:", st.session_state.slide_data.get('goal', 'N/A')); st.write("Inferred KPI:", st.session_state.slide_data.get('kpi', 'N/A')); st.write("Generated Tags:", st.session_state.slide_data.get('tags', [])); st.write("Determined Success Criteria:", st.session_state.slide_data.get('success_criteria', 'N/A')); st.write("Extracted Metrics:", st.session_state.slide_data.get('metrics', {}))
        with debug_tabs[2]:
             st.write("Control Image (Uploaded):"); st.image(st.session_state.slide_data.get('raw_control_image'), width=300) if 'raw_control_image' in st.session_state.slide_data and isinstance(st.session_state.slide_data.get('raw_control_image'), Image.Image) else st.write("(Not available)")
             st.write("Supporting Data Image (Used):"); st.image(st.session_state.slide_data.get('supporting_data_image'), width=300) if 'supporting_data_image' in st.session_state.slide_data and isinstance(st.session_state.slide_data.get('supporting_data_image'), Image.Image) else st.write("(Not available)")
             st.write("Generated Control Shipping Image:"); st.image(st.session_state.slide_data.get('control_image'), width=300) if 'control_image' in st.session_state.slide_data and isinstance(st.session_state.slide_data.get('control_image'), Image.Image) else st.write("(Not available)")
             st.write("Generated Variant Shipping Image:"); st.image(st.session_state.slide_data.get('variant_image'), width=300) if 'variant_image' in st.session_state.slide_data and isinstance(st.session_state.slide_data.get('variant_image'), Image.Image) else st.write("(Not available)")

else:
    st.info("⬆️ Upload files and fill in details in the sidebar to generate the slide.")
    st.markdown("##### Target Slide Structure Guide:")
    st.markdown("The generated slide will follow the standard PDQ A/B test layout including Hypothesis, Segment, Goal, Supporting Data, Control/Variant visuals, etc.")

# --- Custom Footer ---
footer_year = datetime.datetime.now().year; footer_left_text = "PDQ A/B Test Slide Generator | Streamlining Test Documentation"; footer_right_text = f"PDQ © {footer_year}"
st.markdown(f"""<div class="custom-footer"><div class="footer-left">{footer_left_text}</div><div class="footer-right">{footer_right_text}</div></div>""", unsafe_allow_html=True)
