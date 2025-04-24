# Use Python 3.9 slim version as the base image
FROM python:3.9-slim

# Set the working directory inside the container
WORKDIR /app

# Install system dependencies
RUN apt-get update && \
    # Set frontend to noninteractive to avoid prompts during apt-get install
    export DEBIAN_FRONTEND=noninteractive && \
    # Install all packages in a single layer
    # REMOVED: ttf-mscorefonts-installer and cabextract to avoid EULA/install issues
    # KEPT: fonts-liberation as a fallback font set
    apt-get install -y --no-install-recommends \
    tesseract-ocr \
    tesseract-ocr-eng \
    libleptonica-dev \
    libgl1-mesa-glx \
    libglib2.0-0 \
    # Dependencies for Chromium / html2image
    chromium \
    chromium-driver \
    # Font packages (Using Liberation fonts as fallback)
    fonts-liberation \
    # Clean up apt cache afterwards to keep image size smaller
    && rm -rf /var/lib/apt/lists/*

# Copy the requirements file into the container first
COPY requirements.txt .

# Install Python dependencies specified in requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

# Copy the rest of your application code
COPY . .

# Add a health check endpoint
# Using 8501 here as well for consistency with the CMD instruction.
HEALTHCHECK CMD curl --fail http://localhost:8501/_stcore/health || exit 1

# Command to run the Streamlit application when the container starts
# Explicitly using port 8501 as requested
# Binds to 0.0.0.0 to be accessible from outside the container
# Disables CORS/XSRF protection (common setting for Render proxy)
# Ensure 'app.py' is the correct name of your main Streamlit script
CMD ["streamlit", "run", "app.py", "--server.port", "8501", "--server.address", "0.0.0.0", "--server.enableCORS", "false", "--server.enableXsrfProtection", "false"]
