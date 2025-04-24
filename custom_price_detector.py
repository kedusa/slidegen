import cv2
import numpy as np
from PIL import Image
import re
import pytesseract

class ShippingPriceDetector:
    """
    Specialized class for detecting and modifying shipping prices in checkout forms
    """
    
    def __init__(self):
        """Initialize the shipping price detector"""
        # OCR config for better price detection
        self.config = '--oem 3 --psm 7 -c tessedit_char_whitelist="$0123456789."'
        
        # Standard shipping patterns
        self.standard_shipping_patterns = [
            "standard shipping",
            "standard & processing",
            "standard delivery"
        ]
        
        # Rush shipping patterns
        self.rush_shipping_patterns = [
            "rush shipping",
            "expedited shipping",
            "express shipping",
            "priority shipping"
        ]
    
    def detect_prices(self, image, target_price=None):
        """
        Detect shipping prices in the image with specialized detection
        
        Args:
            image: PIL Image
            target_price: Optional price to look for
            
        Returns:
            List of detected price locations with metadata
        """
        # Convert to OpenCV format for processing
        img_cv = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
        
        # Get image dimensions
        h, w = img_cv.shape[:2]
        
        # Process the image for better OCR
        gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)
        _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY)
        
        # Detect text regions
        detected_regions = self._detect_text_regions(thresh)
        
        # Find shipping method sections
        shipping_method_regions = self._find_shipping_method_sections(detected_regions, thresh)
        
        # Detect prices in the shipping method regions
        price_locations = []
        
        # If we found shipping method regions, look for prices in them
        if shipping_method_regions:
            for region in shipping_method_regions:
                # Extract the region of interest
                x, y, w, h = region["x"], region["y"], region["width"], region["height"]
                roi = thresh[y:y+h, x:x+w]
                
                # Look for price pattern in the region
                price_info = self._find_price_in_region(roi, region, target_price)
                
                if price_info:
                    price_locations.append(price_info)
        else:
            # Fall back to full image price detection
            price_locations = self._detect_prices_in_full_image(thresh, target_price)
        
        # If still no results, use hardcoded positions
        if not price_locations:
            # Get image dimensions
            img_height, img_width = thresh.shape
            
            # Standard positions for shipping prices
            # Standard shipping (top option)
            price_locations.append({
                "x": int(img_width * 0.85) - 70,
                "y": int(img_height * 0.3) - 10,
                "width": 65,
                "height": 25,
                "detected_price": target_price or "$7.95",
                "shipping_type": "standard"
            })
            
            # Rush shipping (bottom option)
            price_locations.append({
                "x": int(img_width * 0.85) - 70,
                "y": int(img_height * 0.65) - 10,
                "width": 65,
                "height": 25,
                "detected_price": "$24.95",
                "shipping_type": "rush"
            })
        
        return price_locations
    
    def _detect_text_regions(self, image):
        """Find text regions in the image"""
        # Apply OCR with bounding box data
        ocr_data = pytesseract.image_to_data(image, output_type=pytesseract.Output.DICT)
        
        # Process OCR results to find text regions
        regions = []
        for i, text in enumerate(ocr_data['text']):
            if text.strip():
                x = ocr_data['left'][i]
                y = ocr_data['top'][i]
                w = ocr_data['width'][i]
                h = ocr_data['height'][i]
                conf = ocr_data['conf'][i]
                
                # Skip low confidence or very small regions
                if conf < 50 or w < 5 or h < 5:
                    continue
                
                regions.append({
                    "x": x,
                    "y": y,
                    "width": w,
                    "height": h,
                    "text": text.lower(),
                    "confidence": conf
                })
        
        return regions
    
    def _find_shipping_method_sections(self, regions, image):
        """Identify regions that contain shipping method information"""
        shipping_regions = []
        
        # Check each region to see if it contains shipping-related text
        for region in regions:
            text = region.get("text", "").lower()
            # Check for shipping-related keywords
            if any(pattern in text for pattern in self.standard_shipping_patterns + self.rush_shipping_patterns):
                # Expand the region to include the price on the right
                # This is based on the typical layout of shipping method sections
                x, y, w, h = region["x"], region["y"], region["width"], region["height"]
                
                # Get image dimensions
                img_height, img_width = image.shape
                
                # Expand width to likely include the price
                expanded_width = min(img_width - x, w + 200)
                
                shipping_regions.append({
                    "x": x,
                    "y": y,
                    "width": expanded_width,
                    "height": h,
                    "text": text,
                    "shipping_type": "standard" if any(pattern in text for pattern in self.standard_shipping_patterns) else "rush"
                })
        
        return shipping_regions
    
    def _find_price_in_region(self, roi, region, target_price=None):
        """Find price in a region of interest"""
        # Extract text from the region using OCR focused on digits and $ sign
        roi_text = pytesseract.image_to_string(roi, config=self.config)
        
        # Search for price pattern
        price_pattern = r'\$\d+\.\d{2}|\$\d+'
        price_match = re.search(price_pattern, roi_text)
        
        if price_match:
            detected_price = price_match.group(0)
            
            # If we have a target price and it doesn't match, skip this region
            if target_price and target_price != detected_price:
                # Check if the numbers match even if formatting differs
                target_value = re.sub(r'[^\d.]', '', target_price)
                detected_value = re.sub(r'[^\d.]', '', detected_price)
                
                if target_value != detected_value:
                    return None
            
            # Find the price position within the region
            # For simplicity, we'll assume it's in the right side of the region
            price_x = region["x"] + int(region["width"] * 0.7)
            price_y = region["y"]
            price_width = 70
            price_height = region["height"]
            
            return {
                "x": price_x,
                "y": price_y,
                "width": price_width,
                "height": price_height,
                "detected_price": detected_price,
                "shipping_type": region.get("shipping_type", "unknown")
            }
        
        return None
    
    def _detect_prices_in_full_image(self, image, target_price=None):
        """Detect prices in the full image"""
        # Apply OCR to the full image
        ocr_data = pytesseract.image_to_data(image, config=self.config, output_type=pytesseract.Output.DICT)
        
        # Look for price patterns
        price_locations = []
        price_pattern = r'\$\d+\.\d{2}|\$\d+'
        
        for i, text in enumerate(ocr_data['text']):
            if not text.strip():
                continue
                
            price_match = re.search(price_pattern, text)
            if not price_match:
                continue
                
            detected_price = price_match.group(0)
            
            # If we have a target price and it doesn't match, skip
            if target_price and target_price != detected_price:
                continue
                
            # Get coordinates
            x = ocr_data['left'][i]
            y = ocr_data['top'][i]
            w = ocr_data['width'][i]
            h = ocr_data['height'][i]
            
            # Determine if this is likely a standard or rush shipping price
            # For simplicity, assume first price is standard, second is rush
            shipping_type = "standard" if len(price_locations) == 0 else "rush"
            
            price_locations.append({
                "x": x,
                "y": y,
                "width": w,
                "height": h,
                "detected_price": detected_price,
                "shipping_type": shipping_type
            })
        
        return price_locations
    
    def create_price_variant(self, control_image, old_price, new_price):
        """
        Create a price variant by replacing the old price with the new price
        
        Args:
            control_image: PIL Image of control
            old_price: Original price
            new_price: New price to replace it with
            
        Returns:
            PIL Image with the price replaced
        """
        # Create a copy of the control image
        variant_image = control_image.copy()
        variant_array = np.array(variant_image)
        
        # Convert to CV2 format
        img_cv = cv2.cvtColor(variant_array, cv2.COLOR_RGB2BGR)
        
        # Detect price locations
        price_locations = self.detect_prices(control_image, old_price)
        
        # Track if we've modified any prices
        price_modified = False
        
        # Filter to only standard shipping prices if needed
        standard_price_locations = [loc for loc in price_locations 
                                   if loc.get("shipping_type") == "standard"]
        
        # If we have standard shipping price locations, use those, otherwise use all
        target_locations = standard_price_locations if standard_price_locations else price_locations
        
        # Replace prices
        for location in target_locations:
            x = location["x"]
            y = location["y"]
            w = location["width"]
            h = location["height"]
            
            # Create a clean rectangle to cover old price
            cv2.rectangle(img_cv, (x-5, y-5), (x+w+5, y+h+5), (255, 255, 255), -1)
            
            # Add new price text with proper alignment
            font = cv2.FONT_HERSHEY_DUPLEX
            font_scale = 0.6
            font_thickness = 1