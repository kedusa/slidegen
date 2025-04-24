# Use Python 3.9 slim version as the base image
FROM python:3.9-slim

# Set working directory
WORKDIR /app

# Install system dependencies
# - Update package lists
# - Install Tesseract OCR engine, English language pack, Leptonica dev files
# - Install libraries often required by OpenCV and Pillow
# - Install Chromium (headless browser needed by html2image) and its dependencies
# - Install common fonts (MS Core Fonts installer, Liberation fonts)
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
    tesseract-ocr \
    tesseract-ocr-eng \
    libleptonica-dev \
    libgl1-mesa-glx \
    libglib2.0-0 \
    # Dependencies for Chromium / html2image
    chromium \
    chromium-driver \
    # Font packages (Accept EULA for mscorefonts)
    cabextract \
    ttf-mscorefonts-installer \
    fonts-liberation \
    && echo "ttf-mscorefonts-installer msttcorefonts/accepted-mscorefonts-eula select true" | debconf-set-selections \
    && apt-get install -y --no-install-recommends ttf-mscorefonts-installer \
    # Clean up apt cache
    && rm -rf /var/lib/apt/lists/*

# Copy the requirements file into the container first
COPY requirements.txt .

# Install Python dependencies specified in requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

# Copy the rest of your application code
COPY . .

# Add a health check endpoint
HEALTHCHECK CMD curl --fail http://localhost:8501/_stcore/health || exit 1

# Command to run the Streamlit application
CMD ["streamlit", "run", "app.py", "--server.port", "8501", "--server.address", "0.0.0.0", "--server.enableCORS", "false", "--server.enableXsrfProtection", "false"]
