# Use Python 3.9 slim version as the base image
FROM python:3.9-slim

# Set the working directory inside the container
WORKDIR /app

# Install system dependencies
RUN apt-get update && \
    # Set frontend to noninteractive to avoid prompts during apt-get install
    export DEBIAN_FRONTEND=noninteractive && \
    # Install all packages in a single layer
    apt-get install -y --no-install-recommends \
        # Tesseract OCR dependencies
        tesseract-ocr \
        tesseract-ocr-eng \
        libleptonica-dev \
        # Graphics libraries (potential dependencies for headless browser/other tools)
        libgl1-mesa-glx \
        libglib2.0-0 \
        # Dependencies for html2image (headless browser)
        chromium \
        chromium-driver \
        # Font packages: Install Liberation fonts as a metrically compatible fallback for Arial/Times/etc.
        # This is the recommended approach for font consistency on Linux servers.
        fonts-liberation \
    # Clean up apt cache afterwards to keep image size smaller
    && rm -rf /var/lib/apt/lists/*

# Copy the requirements file into the container first
COPY requirements.txt .

# Install Python dependencies specified in requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

# Copy the rest of your application code
COPY . .

# Add a health check endpoint targeting the Streamlit port
HEALTHCHECK CMD curl --fail http://localhost:8501/_stcore/health || exit 1

# Command to run the Streamlit application when the container starts
# - Binds to 0.0.0.0 to be accessible from outside the container on Render
# - Uses port 8501
# - Disables CORS/XSRF protection (handled by Render's proxy)
# Ensure 'app.py' is the correct name of your main Streamlit script
CMD ["streamlit", "run", "app.py", "--server.port", "8501", "--server.address", "0.0.0.0", "--server.enableCORS", "false", "--server.enableXsrfProtection", "false"]
