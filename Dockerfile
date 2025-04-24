# Use Python 3.9 slim version as the base image (matches previous intent)
FROM python:3.9-slim

# Set the working directory inside the container
WORKDIR /app

# Install system dependencies
# - Update package lists
# - Install Tesseract OCR engine, English language pack, Leptonica development files (often needed by Tesseract wrappers)
# - Install libraries often required by OpenCV and Pillow for image processing
# - Clean up apt cache afterwards to keep image size smaller
RUN apt-get update && apt-get install -y --no-install-recommends \
    tesseract-ocr \
    tesseract-ocr-eng \
    libleptonica-dev \
    libgl1-mesa-glx \
    libglib2.0-0 \
 && rm -rf /var/lib/apt/lists/*

# Copy the requirements file into the container first
# This leverages Docker's layer caching - if requirements.txt doesn't change,
# this layer won't be rebuilt, speeding up subsequent builds.
COPY requirements.txt .

# Install Python dependencies specified in requirements.txt
# --no-cache-dir reduces image size
RUN pip install --no-cache-dir -r requirements.txt

# Copy the rest of your application code (app.py, etc.) into the container
COPY . .

# Add a health check endpoint (good practice for Render)
# Streamlit has a built-in health check endpoint
# Render will use the $PORT variable automatically
HEALTHCHECK CMD curl --fail http://localhost:${PORT}/_stcore/health || exit 1

# Command to run the Streamlit application when the container starts
# - Uses the PORT environment variable provided by Render
# - Binds to 0.0.0.0 to be accessible from outside the container
# - Disables CORS/XSRF protection (common setting for Render proxy)
# - Ensure 'app.py' is the correct name of your main Streamlit script
CMD ["streamlit", "run", "app.py", "--server.port", "${PORT}", "--server.address", "0.0.0.0", "--server.enableCORS", "false", "--server.enableXsrfProtection", "false"]
