# -*- coding: utf-8 -*-
import streamlit as st
import cv2
import numpy as np
import pytesseract # Keep this import for the python wrapper
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
from html2image import Html2Image # Import html2image

# --- Basic Setup ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(name)s - %(message)s')
logger = logging.getLogger("ABTestApp")

# REMOVE or COMMENT OUT the manual Tesseract path setting - not needed in Docker
# if os.name == 'nt':
#     try: pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
#     except Exception: logger.error("Tesseract not found at default path. Please set manually if needed.")

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
    # Prioritize installed fonts now
    common_fonts = [name, 'Segoe UI', 'Arial', 'Calibri', 'Liberation Sans'] # Added Liberation Sans
    font_path = None
    for font_name in common_fonts:
        try:
            # Basic bold/italic handling might still be tricky across systems
            # This part remains heuristic
            style_suffix = ""
            if bold and italic: style_suffix = " Bold Italic"
            elif bold: style_suffix = " Bold"
            elif italic: style_suffix = " Italic"

            # Try common file extensions and naming conventions
            for ext in ['.ttf', '.otf']:
                try:
                    # Try exact name + style
                    return ImageFont.truetype(f"{font_name}{style_suffix}{ext}", size)
                except IOError:
                    try:
                        # Try just name (for fonts where style is part of the base name)
                        return ImageFont.truetype(f"{font_name}{ext}", size)
                    except IOError:
                        continue # Try next extension or font name
        except IOError:
             continue # Try next font in list

    logger.warning(f"Could not find specified fonts: {common_fonts}. Using default PIL font.")
    # Use PIL's default if nothing else is found
    try:
        return ImageFont.load_default(size=size) # Try specifying size for default
    except AttributeError: # Older PIL might not support size argument for load_default
        return ImageFont.load_default()


# --- Streamlit Page Setup & CSS ---
st.set_page_config(page_title="PDQ A/B Test Slide Generator", page_icon="🧪", layout="wide")
# Define custom CSS
st.markdown(f"""
    <style>
    /* ... Keep your existing CSS ... */
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
        if isinstance(image_input, Image.Image):
            image = image_input
            logger.info("Processing PIL image for OCR.")
        elif hasattr(image_input, 'getvalue'): # Check if it's an UploadedFile object
            logger.info(f"Processing uploaded file '{getattr(image_input, 'name', 'N/A')}' for OCR.")
            image_bytes = image_input.getvalue()
            if not image_bytes:
                logger.warning("extract_text_from_image received empty file bytes.")
                return "", None # Return None for image if invalid
            image = Image.open(io.BytesIO(image_bytes))
        else:
            logger.error(f"Unsupported input type for OCR: {type(image_input)}")
            return "", None

        # Convert PIL Image to OpenCV format
        img_cv = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
        gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)
        # Apply thresholding
        thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]

        # Perform OCR
        extracted_text = pytesseract.image_to_string(thresh, config='--psm 6')
        logger.info(f"OCR extracted {len(extracted_text)} characters.")
        return extracted_text, image # Return the original PIL image

    except UnidentifiedImageError:
        st.error("Invalid or corrupted image file provided.")
        logger.error("UnidentifiedImageError during OCR processing.")
        return "", None
    except pytesseract.TesseractNotFoundError:
        # This error should theoretically not happen anymore with Docker setup
        st.error("Tesseract installation issue detected. Please check deployment logs.")
        logger.error("TesseractNotFoundError - This should not happen in the Docker env!")
        return "", None
    except Exception as e:
        st.error(f"An unexpected error occurred during OCR: {e}")
        logger.error(f"Unexpected OCR Error: {e}", exc_info=True)
        return "", None

def extract_metrics_from_supporting_data(image_obj):
    """Extract key metrics from a PIL image object using OCR."""
    internal_default_metrics = { "conversion_rate": "N/A", "total_checkout": "N/A", "checkouts": "N/A", "orders": "N/A", "shipping_revenue": "N/A", "aov": "N/A" }
    if not isinstance(image_obj, Image.Image):
        logger.warning("extract_metrics needs PIL object. Received non-image or None.")
        return internal_default_metrics
    try:
        # Reuse the OCR function
        extracted_text, _ = extract_text_from_image(image_obj) # We already have the image, just need text
        if not extracted_text:
            logger.warning("No text extracted from supporting data image for metrics.")
            return internal_default_metrics

        # Regex patterns (keep as they are)
        conv_match = re.search(r'(?:conversion|checkout\s*conversion)\s*(?:rate)?[:\s]*(\d{1,3}(?:,\d{3})*\.?\d*%?|\d+\.?\d*%?)', extracted_text, re.IGNORECASE)
        total_co_match = re.search(r'(?:%|percent)\s*total\s*checkout[:\s]*(\d{1,3}(?:,\d{3})*\.?\d*%?|\d+\.?\d*%?)', extracted_text, re.IGNORECASE)
        checkouts_match = re.search(r'(?:#|number\s+of)?\s*Checkouts[:\s]*(\d{1,3}(?:,\d{3})*)', extracted_text, re.IGNORECASE)
        orders_match = re.search(r'(?:#|number\s+of)?\s*Orders[:\s]*(\d{1,3}(?:,\d{3})*)', extracted_text, re.IGNORECASE)
        ship_rev_match = re.search(r'(?:avg|average)?\s*shipping\s*revenue[:\s]*(\$\s*\d{1,3}(?:,\d{3})*\.?\d*)', extracted_text, re.IGNORECASE)
        aov_match = re.search(r'(?:AOV|average\s*order\s*value)[:\s]*(\$\s*\d{1,3}(?:,\d{3})*\.?\d*)', extracted_text, re.IGNORECASE)

        # Assign results or defaults
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
    except Exception as e:
        logger.error(f"Metric extraction error: {e}", exc_info=True)
        # Don't show streamlit error here, just return defaults
        return internal_default_metrics

# --- HTML Variant Generation ---
def generate_shipping_html(standard_price="$7.95", rush_price="$24.95", is_variant=False):
    """ Generate HTML content for shipping options display """
    # Using Arial as a safe fallback now that fonts are installed
    html = f"""<!DOCTYPE html><html><head><style> body {{ font-family: Arial, sans-serif; background-color: #f8f9fa; margin: 0; padding: 20px; box-sizing: border-box; }} .container {{ max-width: 580px; background-color: {PDQ_COLORS['white']}; border: 1px solid {PDQ_COLORS['html_border']}; border-radius: 6px; padding: 20px; position: relative; box-shadow: 0 1px 2px rgba(0,0,0,0.05); }} h2 {{ margin-top: 0; margin-bottom: 15px; font-size: 16px; font-weight: 600; color: {PDQ_COLORS['black']}; }} .shipping-option {{ border: 1px solid {PDQ_COLORS['html_border']}; border-radius: 6px; padding: 15px; margin-bottom: 10px; display: flex; align-items: flex-start; transition: all 0.2s ease-in-out; }} .shipping-option.selected {{ border-color: {PDQ_COLORS['html_selected_border']}; background-color: {PDQ_COLORS['html_selected_bg']}; }} .radio {{ margin-right: 12px; margin-top: 3px; flex-shrink: 0; }} .radio-dot {{ width: 18px; height: 18px; border-radius: 50%; display: flex; align-items: center; justify-content: center; }} .radio-selected .radio-dot {{ background-color: {PDQ_COLORS['html_selected_border']}; }} .radio-selected .radio-dot-inner {{ width: 7px; height: 7px; border-radius: 50%; background-color: white; }} .radio-unselected .radio-dot {{ border: 2px solid {PDQ_COLORS['html_radio_border']}; background-color: white; }} .shipping-details {{ flex-grow: 1; }} .shipping-title {{ font-weight: 600; font-size: 14px; margin-bottom: 4px; color: #333; }} .shipping-subtitle {{ color: {PDQ_COLORS['grey_text']}; font-size: 12px; }} .shipping-price {{ font-weight: 600; font-size: 14px; text-align: right; min-width: 60px; color: #333; margin-left: 10px; }} .footnote {{ font-size: 12px; color: {PDQ_COLORS['grey_text']}; margin-top: 15px; }} .variant-label {{ position: absolute; top: 10px; right: 10px; background-color: {PDQ_COLORS['white']}; border: 1px solid {PDQ_COLORS['electric_violet']}; color: {PDQ_COLORS['electric_violet']}; font-weight: 600; font-size: 9px; padding: 2px 5px; border-radius: 3px; text-transform: uppercase; letter-spacing: 0.5px; }} </style></head><body><div class="container"><h2>Shipping method</h2>{f'<div class="variant-label">VARIANT</div>' if is_variant else ''}<div class="shipping-option selected"><div class="radio radio-selected"><div class="radio-dot"><div class="radio-dot-inner"></div></div></div><div class="shipping-details"><div class="shipping-title">Standard Shipping & Processing* (4-7 Business Days)</div><div class="shipping-subtitle">Please allow 1-2 business days for order processing</div></div><div class="shipping-price">{standard_price}</div></div><div class="shipping-option"><div class="radio radio-unselected"><div class="radio-dot"></div></div><div class="shipping-details"><div class="shipping-title">Rush Shipping* (2 Business Days)</div><div class="shipping-subtitle">Please allow 1-2 business days for order processing</div></div><div class="shipping-price">{rush_price}</div></div><div class="footnote">*Includes $1.49 processing fee</div></div></body></html>"""
    return html

# --- Removed generate_simple_pil_image as html2image should now work ---

def html_to_image(html_content, output_path="temp_shipping_image.png", size=(600, 300)):
    """ Convert HTML content to an image using html2image """
    # No fallback needed now, rely on html2image
    try:
        # Ensure output directory exists if needed (using temp dir is better)
        temp_dir = tempfile.gettempdir()
        hti = Html2Image(output_path=temp_dir, size=size, custom_flags=['--no-sandbox', '--disable-gpu']) # Added flags often needed in containers
        # Generate a unique filename
        unique_filename = f"{os.path.splitext(os.path.basename(output_path))[0]}_{datetime.datetime.now().strftime('%Y%m%d%H%M%S%f')}.png"
        full_output_path = os.path.join(temp_dir, unique_filename)

        logger.info(f"Attempting HTML screenshot to: {full_output_path}")
        paths = hti.screenshot(html_str=html_content, save_as=unique_filename)

        if not paths or not os.path.exists(paths[0]):
             raise RuntimeError(f"html2image failed to create the screenshot file at {paths[0] if paths else 'unknown path'}")

        logger.info(f"Screenshot successful: {paths[0]}")
        img = Image.open(paths[0])
        # It's crucial to copy the image data before removing the file
        img_copy = img.copy()
        img.close() # Close the file handle

        try:
            os.remove(paths[0])
            logger.info(f"Removed temporary screenshot: {paths[0]}")
        except Exception as e:
            logger.warning(f"Failed to remove temp screenshot {paths[0]}: {e}")

        return img_copy # Return the copy

    except Exception as e:
        st.error(f"Error converting HTML to image using html2image: {e}")
        logger.error(f"html_to_image error: {e}", exc_info=True)
        # Provide a basic placeholder if html2image fails catastrophically
        placeholder = Image.new('RGB', size, color=hex_to_rgb(PDQ_COLORS["moon_raker"]))
        draw = ImageDraw.Draw(placeholder)
        try:
            font = get_font("Arial", 14) # Use a reliable font
            draw.text((10, 10), "Error generating image preview", fill=(0,0,0), font=font)
        except: pass # Ignore font errors in placeholder
        return placeholder


def generate_shipping_options(old_price="$7.95", new_price="$5.00"):
    """ Generate control and variant shipping option images """
    logger.info("Generating shipping HTML for control and variant...")
    control_html = generate_shipping_html(old_price, "$24.95", is_variant=False)
    variant_html = generate_shipping_html(new_price, "$24.95", is_variant=True)
    logger.info("Converting HTML to images...")
    control_image = html_to_image(control_html, output_path="control_shipping.png")
    variant_image = html_to_image(variant_html, output_path="variant_shipping.png")
    logger.info("Shipping option image generation complete.")
    return control_image, variant_image

# --- PDF Processing ---
def extract_from_pdf(pdf_file):
    """Extract content from PDF files using PyMuPDF"""
    pdf_content = []
    if not pdf_file:
        logger.warning("extract_from_pdf received None.")
        return pdf_content
    try:
        logger.info(f"Processing PDF: {getattr(pdf_file, 'name', 'N/A')}")
        pdf_bytes = pdf_file.read()
        if not pdf_bytes:
            logger.warning("Empty PDF file uploaded.")
            return pdf_content

        with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
            logger.info(f"Opened PDF with {len(doc)} pages.")
            for page_num, page in enumerate(doc):
                page_text = page.get_text("text")
                page_images = []
                image_list = page.get_images(full=True)
                logger.debug(f"Page {page_num+1}: Found {len(image_list)} raw image entries.")
                for img_index, img in enumerate(image_list):
                    xref = img[0]
                    if xref == 0: continue # Skip invalid xref=0 entries sometimes found
                    try:
                        base_image = doc.extract_image(xref)
                        if not base_image:
                            logger.warning(f"extract_image returned None for xref {xref} on page {page_num+1}.")
                            continue
                        image_bytes = base_image["image"]
                        if image_bytes:
                            try:
                                image = Image.open(io.BytesIO(image_bytes))
                                page_images.append(image)
                                logger.debug(f"Successfully extracted image {img_index+1} (xref: {xref}) from page {page_num+1}.")
                            except UnidentifiedImageError:
                                logger.warning(f"Skipping unidentified image {img_index+1} (xref: {xref}) on page {page_num+1}.")
                            except Exception as img_open_e:
                                logger.warning(f"Error opening image {img_index+1} (xref: {xref}) from page {page_num+1}: {img_open_e}")
                        else:
                            logger.warning(f"Empty image bytes for xref {xref} on page {page_num+1}.")
                    except Exception as e:
                        logger.warning(f"PDF Image Extraction Warning - Page {page_num+1}, Img Index {img_index+1} (xref: {xref}): {e}")

                pdf_content.append({"page_num": page_num + 1, "text": page_text, "images": page_images})
                logger.debug(f"Processed Page {page_num+1}: Text length {len(page_text)}, Images found: {len(page_images)}")

        logger.info(f"Finished processing PDF: Found {sum(len(p['images']) for p in pdf_content)} valid images total.")
    except Exception as e:
        st.error(f"PDF Processing Error: {e}")
        logger.error(f"PDF Error: {e}", exc_info=True)
    return pdf_content

# --- Slide Content Generation Helper (Keep as is) ---
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

# --- PowerPoint Generation Function (Keep as is) ---
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
            img_ratio = img_obj.height / img_obj.width if img_obj.width > 0 else 1 # Avoid division by zero
            pic.height = int(pic.width * img_ratio)

            # Center the image
            pic.left = inner_left + ((target_shape.width - pad_l - pad_r) - pic.width) / 2
            pic.top = inner_top + ((target_shape.height - pad_t - pad_b) - pic.height) / 2 if inner_height > pic.height else inner_top

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
    p_title.font.name = 'Segoe UI'; # Keep trying Segoe UI, relies on system fonts
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
        support_img_title_clearance = Inches(0.8) # Space for title
        support_img_padding_bottom = Inches(0.25)
        place_image_in_shape(
            supporting_data_image, support_box_shape, slide.shapes, context="supporting_data",
            padding=(support_img_padding_horz, support_img_title_clearance,
                     support_img_padding_horz, support_img_padding_bottom),
            scale_factor=0.7 # Adjust scale factor as needed
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

    # 2. Add Control Title
    control_title_left = RIGHT_COL_LEFT + Inches(0.15)
    control_title_box = slide.shapes.add_textbox(control_title_left, CONTROL_TITLE_TOP - Inches(0.05),
                                               RIGHT_COL_WIDTH - Inches(0.3), CONTROL_VAR_TITLE_HEIGHT)
    ctrl_tf = control_title_box.text_frame
    ctrl_tf.margin_left=0; ctrl_tf.margin_right=0; ctrl_tf.margin_top=0; ctrl_tf.margin_bottom=0
    ctrl_p = ctrl_tf.add_paragraph()
    ctrl_p.text = "Control"
    ctrl_p.font.size = Pt(14)
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

    # 5. Add Variant Title
    variant_title_left = RIGHT_COL_LEFT + Inches(0.15)
    variant_title_box = slide.shapes.add_textbox(variant_title_left, VARIANT_TITLE_TOP - Inches(0.05),
                                               RIGHT_COL_WIDTH - Inches(0.3), CONTROL_VAR_TITLE_HEIGHT)
    var_tf = variant_title_box.text_frame
    var_tf.margin_left=0; var_tf.margin_right=0; var_tf.margin_top=0; var_tf.margin_bottom=0
    var_p = var_tf.add_paragraph()
    var_p.text = "Variant B (example)"
    var_p.font.size = Pt(14)
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
        # Attempt to parse a specific date, fallback to current date
        footer_date_str = datetime.datetime.strptime("April 23, 2025", '%B %d, %Y').strftime('%B %d, %Y')
    except ValueError:
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
    logger.info("PPTX slide created successfully.")
    return pptx_buffer

# --- Download Link Helper ---
def get_download_link(buffer, filename, text):
     """Generate a download link for the given file buffer"""
     try:
         if isinstance(buffer, io.BytesIO): file_bytes = buffer.getvalue()
         elif isinstance(buffer, bytes): file_bytes = buffer
         else: raise TypeError("Buffer must be bytes or BytesIO")
         b64 = base64.b64encode(file_bytes).decode(); mime = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
         # Use a unique ID for the download button/link if needed, or keep simple
         href = f'<a href="data:{mime};base64,{b64}" download="{filename}" class="download-button">{text}</a>'
         return href
     except Exception as e:
         logger.error(f"Download link error: {e}", exc_info=True)
         # Return an error message within the HTML structure if possible
         return f'<span class="error-box">Error creating download link.</span>'


# --- Streamlit UI and Main Logic ---
# Initialize session state keys if they don't exist
if 'slide_generated' not in st.session_state: st.session_state.slide_generated = False
if 'output_buffer' not in st.session_state: st.session_state.output_buffer = None
if 'slide_data' not in st.session_state: st.session_state.slide_data = {}
if 'error_message' not in st.session_state: st.session_state.error_message = None

st.title("🧪 PDQ A/B Test Slide Generator")
st.markdown("Generate professional A/B test slides matching the PDQ standard layout.")
st.markdown("---")

# Display general errors if any occurred during the last run
if st.session_state.error_message:
    st.error(f"❌ An error occurred: {st.session_state.error_message}")
    # Optionally clear the error after displaying it
    # st.session_state.error_message = None

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
if not required_inputs_present: st.sidebar.warning("Please provide all required inputs (2, 3, 4). Supporting data (1) is optional.")

# --- Generation Logic ---
if generate_button:
    st.session_state.slide_generated = False # Reset state before generation
    st.session_state.output_buffer = None
    st.session_state.slide_data = {}
    st.session_state.error_message = None   # Clear previous errors
    logger.info("Generate button clicked. Starting process...")

    with st.spinner("⚙️ Processing inputs and generating slide..."):
        try: # Wrap the entire generation process
            logger.info("Initializing SlideGeneratorHelper...")
            slide_helper = PDQSlideGeneratorHelper()
            default_metrics = { "conversion_rate": "N/A", "total_checkout": "N/A", "checkouts": "N/A", "orders": "N/A", "shipping_revenue": "N/A", "aov": "N/A" }
            metrics = default_metrics.copy()
            extracted_supporting_data_text = ""; supporting_data_image = None

            # --- Process Supporting Data (Optional) ---
            if supporting_data_file:
                logger.info(f"Processing supporting data file: {supporting_data_file.name}")
                st.sidebar.info(f"Processing '{supporting_data_file.name}'...")
                if supporting_data_file.type == "image/png":
                    extracted_supporting_data_text, img_pil = extract_text_from_image(supporting_data_file)
                    if img_pil:
                        metrics = extract_metrics_from_supporting_data(img_pil)
                        supporting_data_image = img_pil
                        logger.info("Processed PNG supporting data.")
                    else:
                        logger.warning("Failed to get PIL image from supporting PNG.")
                elif supporting_data_file.type == "application/pdf":
                    pdf_content = extract_from_pdf(supporting_data_file)
                    if pdf_content:
                        extracted_supporting_data_text = " ".join(p["text"] for p in pdf_content if p.get("text"))
                        # Find the first valid image from any page
                        first_image = next((img for p in pdf_content for img in p.get("images", []) if isinstance(img, Image.Image)), None)
                        if first_image:
                            supporting_data_image = first_image
                            metrics = extract_metrics_from_supporting_data(first_image)
                            logger.info("Processed PDF supporting data, found image.")
                        else:
                            logger.warning("No valid images found in the PDF for metrics/display.")
                    else:
                         logger.warning("extract_from_pdf returned empty content.")
                st.sidebar.success("Supporting data processed.")
            else:
                logger.info("No supporting data file provided.")

            # --- Process Control Layout Image (Required) ---
            logger.info("Processing control layout file...")
            _, control_image_input_pil = extract_text_from_image(control_layout_file)
            if not isinstance(control_image_input_pil, Image.Image):
                 logger.error("Failed to process control layout image. Cannot generate.")
                 st.error("Failed to read the Control Layout Image. Please check the file and try again.")
                 # Set error state and stop execution within this block
                 st.session_state.error_message = "Failed to read Control Layout Image."
                 st.stop() # Stop script execution for this run

            # --- Generate Shipping Option Images ---
            logger.info("Extracting prices for shipping options...")
            prices = re.findall(r'\$(\d+\.?\d*)', test_type)
            old_price_str = f"${prices[0]}" if prices else "$7.95"
            new_price_str = f"${prices[1]}" if len(prices) > 1 else "$5.00"
            logger.info(f"Generating shipping option images with prices: Control={old_price_str}, Variant={new_price_str}")
            # This step now uses html2image
            control_shipping_img, variant_shipping_img = generate_shipping_options(old_price_str, new_price_str)
            logger.info("Shipping images generated (using html2image).")

            # Check Validity of Generated Images
            if not isinstance(control_shipping_img, Image.Image):
                logger.warning("Control shipping image generation failed. Using placeholder.")
                control_shipping_img = Image.new("RGB", (600, 300), color=hex_to_rgb(PDQ_COLORS["moon_raker"])) # Neutral placeholder
            if not isinstance(variant_shipping_img, Image.Image):
                logger.warning("Variant shipping image generation failed. Using placeholder.")
                variant_shipping_img = Image.new("RGB", (600, 300), color=hex_to_rgb(PDQ_COLORS["moon_raker"])) # Neutral placeholder

            # --- Generate Slide Content ---
            logger.info("Generating slide text content (hypothesis, kpis, etc.)...")
            parsed_title = slide_helper.parse_test_type(test_type)
            hypothesis = custom_hypothesis if custom_hypothesis else slide_helper.generate_hypothesis(test_type, segment, extracted_supporting_data_text)
            goal, kpi, impact = slide_helper.infer_goals_and_kpis(test_type)
            tags = slide_helper.generate_tags(test_type, segment, extracted_supporting_data_text)
            success_criteria = slide_helper.determine_success_criteria(test_type, kpi, goal)
            timeline_str = "4 weeks\nStat Sig: 85%" # Example placeholder
            checkouts_required_str = "20,000" # Example placeholder

            # --- Create PowerPoint ---
            logger.info("Calling create_proper_pptx function...")
            output_buffer = create_proper_pptx(
                 title=f"AB Test: {parsed_title}",
                 hypothesis=hypothesis,
                 segment=segment,
                 goal=goal,
                 kpi_impact_str=f"{kpi} ({impact} Improvement)", # Combine KPI and impact
                 elements_tags=tags,
                 timeline_str=timeline_str,
                 success_criteria=success_criteria,
                 checkouts_required_str=checkouts_required_str,
                 control_image=control_shipping_img,
                 variant_image=variant_shipping_img,
                 supporting_data_image=supporting_data_image
            )
            logger.info("PPTX buffer created.")

            # --- Update Session State ---
            st.session_state.slide_generated = True
            st.session_state.output_buffer = output_buffer
            st.session_state.slide_data = {
                 "title": f"AB Test: {parsed_title}", "segment": segment, "test_type": test_type,
                 "control_image": control_shipping_img, "variant_image": variant_shipping_img,
                 "supporting_data_image": supporting_data_image, "raw_control_image": control_image_input_pil,
                 "metrics": metrics, "hypothesis": hypothesis, "goal": goal, "kpi": kpi, "impact": impact, "tags": tags, "success_criteria": success_criteria,
            }
            logger.info("Session state updated successfully.")

        # --- Catch ALL exceptions during generation ---
        except Exception as e:
            st.error(f"❌ An error occurred during slide generation: {e}")
            logger.exception("Error during slide generation button press:")
            st.session_state.slide_generated = False
            st.session_state.error_message = str(e) # Store error message

        # --- Trigger UI Update ---
        # This needs to happen OUTSIDE the spinner context maybe?
        # Let's try running it regardless of success/failure to see results/errors
        logger.info("Attempting st.rerun() to update UI.")
        st.rerun() # Rerun to display results or the error message at the top


# --- Display Results Section ---
# This block runs on subsequent reruns after the button is pressed and state is set
if st.session_state.slide_generated and st.session_state.output_buffer:
    logger.info("Displaying results section...")
    try: # Wrap the results display in a try-except as well
        st.markdown(f'<div class="success-box">✅ A/B Test slide generated successfully!</div>', unsafe_allow_html=True)
        col1, col2 = st.columns([2, 1])

        with col1:
            st.subheader("📊 Image Previews")
            st.markdown("Previews of the images used in the generated slide.")

            with st.expander("Control Image (Generated Shipping Option)", expanded=True):
                img_ctrl = st.session_state.slide_data.get('control_image')
                if img_ctrl and isinstance(img_ctrl, Image.Image):
                    st.image(img_ctrl, caption="Generated Control Shipping Image", use_column_width=True)
                else:
                    st.warning("Control image preview not available or invalid.")
                    logger.warning("Control image not available or invalid in session state for display.")

            with st.expander("Variant Image (Generated Shipping Option)", expanded=True):
                 img_var = st.session_state.slide_data.get('variant_image')
                 if img_var and isinstance(img_var, Image.Image):
                     st.image(img_var, caption="Generated Variant Shipping Image", use_column_width=True)
                 else:
                     st.warning("Variant image preview not available or invalid.")
                     logger.warning("Variant image not available or invalid in session state for display.")

            img_supp = st.session_state.slide_data.get('supporting_data_image')
            if img_supp and isinstance(img_supp, Image.Image):
                with st.expander("Supporting Data Image (Uploaded)", expanded=False):
                     st.image(img_supp, caption="Uploaded Supporting Data Image", use_column_width=True)
                     metrics_data = st.session_state.slide_data.get('metrics', {})
                     # Check if metrics_data is not None and has valid entries before displaying
                     if metrics_data and isinstance(metrics_data, dict) and any(v != "N/A" for v in metrics_data.values()):
                         st.write("**Extracted Metrics:**")
                         st.table(metrics_data)
            # Only show info message if no supporting data file was uploaded initially
            elif not supporting_data_file:
                 st.info("No supporting data image was provided.")
            else:
                 st.warning("Supporting data image preview not available (check processing logs).")
                 logger.warning("Supporting data image not available or invalid in session state for display.")


        with col2:
            st.subheader("⬇️ Download Slide")
            download_link_html = get_download_link(
                st.session_state.output_buffer,
                "pdq_ab_test_slide.pptx",
                "Download PPTX File"
            )
            st.markdown(download_link_html, unsafe_allow_html=True)

            st.markdown("---"); st.subheader("📝 Slide Content Summary")
            summary_data = {
                "Title": st.session_state.slide_data.get("title", "N/A"),
                "Segment": st.session_state.slide_data.get("segment", "N/A"),
                "Goal": st.session_state.slide_data.get("goal", "N/A"),
                "KPI": st.session_state.slide_data.get("kpi", "N/A"),
                "Tags": ", ".join(st.session_state.slide_data.get("tags", [])),
                "Success Criteria": st.session_state.slide_data.get("success_criteria", "N/A"),
            }
            for key, value in summary_data.items():
                st.markdown(f"**{key}:** {value}")

            st.markdown("---")
            if st.button("✨ Create Another Slide"):
                logger.info("Clearing session state for new slide.")
                # Clear specific keys related to the last generation
                keys_to_clear = ['slide_generated', 'output_buffer', 'slide_data', 'error_message']
                for key in keys_to_clear:
                    st.session_state.pop(key, None)
                st.rerun() # Rerun to reset the UI

        # --- Debug Information Display ---
        if show_debug:
            logger.info("Displaying debug information.")
            st.markdown("---"); st.subheader("🔍 Debug Information")
            debug_tabs = st.tabs(["Inputs Used", "Generated Content", "Images"])
            with debug_tabs[0]:
                st.write("Test Description Input:", st.session_state.slide_data.get('test_type', 'N/A'))
                st.write("Segment Input:", st.session_state.slide_data.get('segment', 'N/A'))
                st.write("Custom Hypothesis Input:", custom_hypothesis if custom_hypothesis else "(Not provided)")
            with debug_tabs[1]:
                st.write("Generated Hypothesis:", st.session_state.slide_data.get('hypothesis', 'N/A'))
                st.write("Inferred Goal:", st.session_state.slide_data.get('goal', 'N/A'))
                st.write("Inferred KPI:", st.session_state.slide_data.get('kpi', 'N/A'))
                st.write("Generated Tags:", st.session_state.slide_data.get('tags', []))
                st.write("Determined Success Criteria:", st.session_state.slide_data.get('success_criteria', 'N/A'))
                st.write("Extracted Metrics:", st.session_state.slide_data.get('metrics', {}))
            with debug_tabs[2]:
                 st.write("Control Image (Uploaded - Raw Input):")
                 raw_ctrl_img = st.session_state.slide_data.get('raw_control_image')
                 if raw_ctrl_img and isinstance(raw_ctrl_img, Image.Image): st.image(raw_ctrl_img, width=300)
                 else: st.write("(Not available or invalid)")

                 st.write("Supporting Data Image (Used in PPTX):")
                 supp_img = st.session_state.slide_data.get('supporting_data_image')
                 if supp_img and isinstance(supp_img, Image.Image): st.image(supp_img, width=300)
                 else: st.write("(Not available or invalid)")

                 st.write("Generated Control Shipping Image (Used in PPTX):")
                 gen_ctrl_img = st.session_state.slide_data.get('control_image')
                 if gen_ctrl_img and isinstance(gen_ctrl_img, Image.Image): st.image(gen_ctrl_img, width=300)
                 else: st.write("(Not available or invalid)")

                 st.write("Generated Variant Shipping Image (Used in PPTX):")
                 gen_var_img = st.session_state.slide_data.get('variant_image')
                 if gen_var_img and isinstance(gen_var_img, Image.Image): st.image(gen_var_img, width=300)
                 else: st.write("(Not available or invalid)")

    # --- Catch errors during results display ---
    except Exception as display_e:
        st.error(f"❌ An error occurred while displaying the results: {display_e}")
        logger.exception("Error during results display section:")
        st.session_state.error_message = f"Error displaying results: {display_e}"
        # Avoid rerun loop if display fails
        st.session_state.slide_generated = False # Prevent trying to display again immediately

# --- Fallback message if not generated ---
elif not st.session_state.error_message and not generate_button: # Avoid showing this if an error is already displayed or button just pressed
    st.info("⬆️ Upload files and fill in details in the sidebar to generate the slide.")
    # Keep the structure guide
    # st.markdown("##### Target Slide Structure Guide:")
    # st.markdown("The generated slide will follow the standard PDQ A/B test layout including Hypothesis, Segment, Goal, Supporting Data, Control/Variant visuals, etc.")


# --- Custom Footer ---
footer_year = datetime.datetime.now().year; footer_left_text = "PDQ A/B Test Slide Generator | Streamlining Test Documentation"; footer_right_text = f"PDQ © {footer_year}"
st.markdown(f"""<div class="custom-footer"><div class="footer-left">{footer_left_text}</div><div class="footer-right">{footer_right_text}</div></div>""", unsafe_allow_html=True)

