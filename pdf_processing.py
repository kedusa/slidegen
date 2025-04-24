import io
import fitz  # PyMuPDF
import pdfplumber
from PIL import Image
import pytesseract
import re

def extract_text_from_pdf(pdf_file):
    """
    Extract text from PDF using PyMuPDF
    
    Args:
        pdf_file: File object containing PDF data
    
    Returns:
        Dictionary with page numbers as keys and text content as values
    """
    pdf_content = {}
    
    # Create a memory buffer from file data
    pdf_data = pdf_file.read()
    pdf_file.seek(0)  # Reset file pointer
    
    # Use PyMuPDF to extract text
    with fitz.open(stream=pdf_data, filetype="pdf") as doc:
        for page_num, page in enumerate(doc):
            text = page.get_text()
            pdf_content[page_num] = text
    
    return pdf_content

def extract_images_from_pdf(pdf_file):
    """
    Extract images from PDF using PyMuPDF
    
    Args:
        pdf_file: File object containing PDF data
    
    Returns:
        Dictionary with page numbers as keys and lists of images as values
    """
    pdf_images = {}
    
    # Create a memory buffer from file data
    pdf_data = pdf_file.read()
    pdf_file.seek(0)  # Reset file pointer
    
    # Use PyMuPDF to extract images
    with fitz.open(stream=pdf_data, filetype="pdf") as doc:
        for page_num, page in enumerate(doc):
            images = []
            image_list = page.get_images(full=True)
            
            for img_index, img in enumerate(image_list):
                xref = img[0]
                base_image = doc.extract_image(xref)
                image_bytes = base_image["image"]
                
                # Convert to PIL Image
                image = Image.open(io.BytesIO(image_bytes))
                images.append(image)
            
            pdf_images[page_num] = images
    
    return pdf_images

def extract_layout_info_from_pdf(pdf_file):
    """
    Extract detailed layout information from PDF using pdfplumber
    
    Args:
        pdf_file: File object containing PDF data
    
    Returns:
        Dictionary with detailed layout information
    """
    layout_info = {}
    
    # Use pdfplumber for more detailed layout analysis
    with pdfplumber.open(pdf_file) as pdf:
        for page_num, page in enumerate(pdf.pages):
            # Extract page dimensions
            width, height = page.width, page.height
            
            # Extract text with bounding boxes
            text_elements = page.extract_words(
                keep_blank_chars=True,
                x_tolerance=3,
                y_tolerance=3,
                extra_attrs=["fontname", "size"]
            )
            
            # Extract tables
            tables = page.extract_tables()
            
            # Extract images (bounding boxes only from pdfplumber)
            images = page.images
            
            # Store all layout information
            layout_info[page_num] = {
                "dimensions": {"width": width, "height": height},
                "text_elements": text_elements,
                "tables": tables,
                "images": images
            }
    
    pdf_file.seek(0)  # Reset file pointer
    return layout_info

def classify_pdf_page(page_text, page_layout=None):
    """
    Classify the content of a PDF page
    
    Args:
        page_text: Text content of the page
        page_layout: Optional layout information
    
    Returns:
        Dictionary with classification results
    """
    # Convert to lowercase for case-insensitive matching
    text_lower = page_text.lower()
    
    # Define classification patterns
    patterns = {
        "ab_test": [
            r"a/b\s+test", r"ab\s+test", r"test\s+type", r"control", r"variant",
            r"hypothesis", r"segment", r"goal", r"kpi", r"conversion rate",
            r"success criteria"
        ],
        "analytics": [
            r"analytics", r"dashboard", r"metrics", r"performance", r"data",
            r"chart", r"graph", r"trend", r"report", r"stats"
        ],
        "ui_mockup": [
            r"mockup", r"wireframe", r"layout", r"design", r"interface",
            r"ui", r"ux", r"screen", r"page", r"prototype"
        ]
    }
    
    # Count pattern matches for each category
    category_scores = {}
    for category, category_patterns in patterns.items():
        category_scores[category] = 0
        for pattern in category_patterns:
            matches = re.findall(pattern, text_lower)
            category_scores[category] += len(matches)
    
    # Determine primary and secondary classifications
    classifications = sorted(category_scores.items(), key=lambda x: x[1], reverse=True)
    
    # Check for layout-based signals if available
    if page_layout:
        # Example: check for side-by-side comparison layout typical of A/B tests
        # This would be more sophisticated in a real implementation
        pass
    
    result = {
        "primary_classification": classifications[0][0] if classifications[0][1] > 0 else "unknown",
        "secondary_classification": classifications[1][0] if len(classifications) > 1 and classifications[1][1] > 0 else None,
        "is_ab_test": classifications[0][0] == "ab_test" and classifications[0][1] >= 3,
        "confidence": min(classifications[0][1] / 5, 1.0) if classifications[0][1] > 0 else 0,
        "all_scores": category_scores
    }
    
    return result

def extract_test_parameters_from_page(page_text, page_layout=None):
    """
    Extract A/B test parameters from a page classified as an A/B test
    
    Args:
        page_text: Text content of the page
        page_layout: Optional layout information
    
    Returns:
        Dictionary with extracted test parameters
    """
    # Initialize parameters with None values
    test_params = {
        "test_type": None,
        "segment": None,
        "hypothesis": None,
        "goal": None,
        "kpi": None,
        "timeline": None,
        "success_criteria": None
    }
    
    # Convert to lowercase for patterns but keep original for extraction
    text_lower = page_text.lower()
    
    # Extract test type
    test_type_patterns = [
        r"test\s+type:?\s*(.*?)(?:\n|$)",
        r"test:?\s*(.*?)(?:\n|$)",
        r"a/b\s+test:?\s*(.*?)(?:\n|$)"
    ]
    
    for pattern in test_type_patterns:
        match = re.search(pattern, text_lower)
        if match:
            test_params["test_type"] = match.group(1).strip()
            break
    
    # Extract segment
    segment_patterns = [
        r"segment:?\s*(.*?)(?:\n|$)",
        r"target\s+(?:group|audience):?\s*(.*?)(?:\n|$)",
        r"users?:?\s*(.*?)(?:\n|$)"
    ]
    
    for pattern in segment_patterns:
        match = re.search(pattern, text_lower)
        if match:
            test_params["segment"] = match.group(1).strip()
            break
    
    # Extract hypothesis
    hypothesis_pattern = r"hypothesis:?\s*(.*?)(?:\n\n|\n[A-Z])"
    match = re.search(hypothesis_pattern, text_lower, re.DOTALL)
    if match:
        test_params["hypothesis"] = match.group(1).strip()
    
    # Extract goal
    goal_patterns = [
        r"goal:?\s*(.*?)(?:\n|$)",
        r"objective:?\s*(.*?)(?:\n|$)"
    ]
    
    for pattern in goal_patterns:
        match = re.search(pattern, text_lower)
        if match:
            test_params["goal"] = match.group(1).strip()
            break
    
    # Extract KPI
    kpi_patterns = [
        r"kpi:?\s*(.*?)(?:\n|$)",
        r"key\s+performance\s+indicator:?\s*(.*?)(?:\n|$)",
        r"metric:?\s*(.*?)(?:\n|$)"
    ]
    
    for pattern in kpi_patterns:
        match = re.search(pattern, text_lower)
        if match:
            test_params["kpi"] = match.group(1).strip()
            break
    
    # Extract timeline
    timeline_patterns = [
        r"timeline:?\s*(.*?)(?:\n|$)",
        r"duration:?\s*(.*?)(?:\n|$)",
        r"test\s+period:?\s*(.*?)(?:\n|$)"
    ]
    
    for pattern in timeline_patterns:
        match = re.search(pattern, text_lower)
        if match:
            test_params["timeline"] = match.group(1).strip()
            break
    
    # Extract success criteria
    criteria_patterns = [
        r"success\s+criteria:?\s*(.*?)(?:\n\n|\n[A-Z]|$)",
        r"success\s+metric:?\s*(.*?)(?:\n\n|\n[A-Z]|$)"
    ]
    
    for pattern in criteria_patterns:
        match = re.search(pattern, text_lower, re.DOTALL)
        if match:
            test_params["success_criteria"] = match.group(1).strip()
            break
    
    return test_params

def extract_control_variant_images(pdf_file, page_num):
    """
    Attempt to extract control and variant images from a PDF page
    
    Args:
        pdf_file: File object containing PDF data
        page_num: Page number to process
    
    Returns:
        Tuple of (control_image, variant_image) or (None, None) if not found
    """
    control_image = None
    variant_image = None
    
    # Create a memory buffer from file data
    pdf_data = pdf_file.read()
    pdf_file.seek(0)  # Reset file pointer
    
    try:
        with fitz.open(stream=pdf_data, filetype="pdf") as doc:
            if page_num < len(doc):
                page = doc[page_num]
                
                # Get text with coordinates
                text_blocks = page.get_text("blocks")
                
                # Find blocks that might contain "control" or "variant"
                control_block = None
                variant_block = None
                
                for block in text_blocks:
                    block_text = block[4].lower()
                    if "control" in block_text:
                        control_block = block
                    elif "variant" in block_text:
                        variant_block = block
                
                # Extract images
                images = []
                image_list = page.get_images(full=True)
                
                for img_index, img in enumerate(image_list):
                    xref = img[0]
                    base_image = doc.extract_image(xref)
                    image_bytes = base_image["image"]
                    
                    # Convert to PIL Image
                    pil_img = Image.open(io.BytesIO(image_bytes))
                    
                    # Get image rect on page
                    for img_name in page.get_image_info():
                        if img_name["xref"] == xref:
                            if img_name.get("bbox"):
                                rect = img_name["bbox"]
                                images.append({
                                    "image": pil_img,
                                    "rect": rect
                                })
                            break
                
                # Try to match images to control/variant blocks
                if control_block and variant_block and images:
                    # Simple heuristic: assign images by proximity to text blocks
                    control_rect = fitz.Rect(control_block[:4])
                    variant_rect = fitz.Rect(variant_block[:4])
                    
                    # Find image closest to control block
                    control_image_info = min(
                        images, 
                        key=lambda x: abs(x["rect"].y0 - control_rect.y1),
                        default=None
                    )
                    
                    # Find image closest to variant block
                    variant_image_info = min(
                        images, 
                        key=lambda x: abs(x["rect"].y0 - variant_rect.y1),
                        default=None
                    )
                    
                    # Get the images
                    if control_image_info:
                        control_image = control_image_info["image"]
                    
                    if variant_image_info:
                        variant_image = variant_image_info["image"]
                
                # If we couldn't find by text association, try side-by-side layout detection
                if not control_image and not variant_image and len(images) >= 2:
                    # Sort images by x-coordinate (left to right)
                    sorted_images = sorted(images, key=lambda x: x["rect"].x0)
                    
                    # If we have two images side by side, assume left=control, right=variant
                    if len(sorted_images) >= 2:
                        control_image = sorted_images[0]["image"]
                        variant_image = sorted_images[1]["image"]
    
    except Exception as e:
        print(f"Error extracting control/variant images: {e}")
    
    return control_image, variant_image

def is_slide_ab_test(text, layout_info=None, image_count=0):
    """
    Determine if a slide is likely to be an A/B test slide
    
    Args:
        text: Text content of the slide
        layout_info: Optional layout information
        image_count: Number of images on the slide
    
    Returns:
        Boolean indicating if the slide is an A/B test
    """
    # Convert to lowercase for case-insensitive matching
    text_lower = text.lower()
    
    # Define A/B test-specific keywords
    ab_test_keywords = [
        "a/b test", "ab test", "variant", "control", "hypothesis", 
        "test type", "segment", "goal", "kpi", "cvr", "conversion rate",
        "revenue per visitor", "rpv", "average order value", "aov"
    ]
    
    # Count keyword occurrences
    keyword_matches = sum(1 for keyword in ab_test_keywords if keyword in text_lower)
    
    # Check for side-by-side layout which is common in A/B test slides
    has_side_by_side = False
    if layout_info and "text_elements" in layout_info:
        # Look for "control" and "variant" labels with similar y-positions
        control_element = None
        variant_element = None
        
        for element in layout_info["text_elements"]:
            element_text = element.get("text", "").lower()
            if "control" in element_text:
                control_element = element
            elif "variant" in element_text:
                variant_element = element
        
        # Check if control and variant labels are side by side (similar y-position)
        if control_element and variant_element:
            y_diff = abs(control_element.get("top", 0) - variant_element.get("top", 0))
            if y_diff < 20:  # Arbitrary threshold for "same height"
                has_side_by_side = True
    
    # A/B test slides often have exactly 2-3 images (control, variant, and maybe supporting data)
    ideal_image_count = (image_count == 2 or image_count == 3)
    
    # Make classification decision
    # Strong signals: many keyword matches or side-by-side layout with control/variant labels
    if keyword_matches >= 3 or has_side_by_side:
        return True
    # Medium signals: some keywords and typical image count
    elif keyword_matches >= 2 and ideal_image_count:
        return True
    # Not enough evidence
    else:
        return False