# Automated A/B Test Slide Generator

This Streamlit application automatically generates professional A/B test slides from user inputs. The tool helps Customer Success Managers (CSMs) create consistent, high-quality test slides without manual design work.

## 🔍 Overview

The Automated A/B Test Slide Generator solves several key problems:
- Eliminates inconsistencies in slide structure and design
- Prevents wasted time recreating identical templates
- Reduces delays in getting customer approval and launching tests

## 🚀 Key Features

- Extract and process text from uploaded PNG images using OCR
- Parse PDF files to extract text, images, and layout information
- Classify PDF content to identify A/B test-related slides
- Generate coherent hypotheses using templates
- Create intelligent side-by-side control/variant mockups
- Infer test goals, KPIs, and tags from test descriptions
- Generate downloadable PPTX slides with professional purple theme
- Support for various test types (price, shipping, messaging, layout)

## 📋 Requirements

- Python 3.8+
- Tesseract OCR engine (for image text extraction)
- Required Python packages (see `requirements.txt`)

## 🛠️ Installation

1. Install Tesseract OCR engine:
   - **Windows**: Download and install from [here](https://github.com/UB-Mannheim/tesseract/wiki)
   - **macOS**: `brew install tesseract`
   - **Linux**: `sudo apt-get install tesseract-ocr`

2. Clone this repository:
   ```
   git clone https://github.com/yourusername/ab-test-slide-generator.git
   cd ab-test-slide-generator
   ```

3. Install Python dependencies:
   ```
   pip install -r requirements.txt
   ```

4. Run the application:
   ```
   streamlit run app.py
   ```

## 📊 Usage

1. Upload a supporting data file (PNG or PDF)
2. Upload a control layout image (PNG)
3. Enter the segment information
4. Enter the test type description
5. Click "Generate A/B Test Slide"
6. Preview the generated slide
7. Download the slide in PPTX format

## 📂 Project Structure

```
ab-test-slide-generator/
├── app.py                      # Main Streamlit application
├── enhanced_slide_generator.py # Slide generation logic with AI options
├── image_processor.py          # Advanced image processing utilities
├── requirements.txt            # Python dependencies
└── README.md                   # Project documentation
```

## 🎯 Variant Mockup Generation

The application intelligently creates variant mockups based on test type:

- **Price Tests**: Detects and replaces price information with new values
- **Shipping Tests**: Adds appropriate shipping messaging based on test parameters
- **Messaging Tests**: Modifies text elements to show new copy
- **Layout Tests**: Highlights areas of layout changes with visual indicators

## 🎨 Slide Design

The generated slides follow a professional design with:

- Dark purple background theme
- Clear section organization (Hypothesis, Segment, Timeline, Goal, Elements)
- Side-by-side Control/Variant mockups
- Success criteria and statistical significance information
- Supporting data visualization (when available)

## 🧠 AI-Enhanced Content Generation (Optional)

The application includes an experimental AI-enhanced content generation option that:

- Creates more nuanced hypotheses
- Better interprets test descriptions
- Generates more contextually appropriate success criteria

## 🔧 Advanced Configuration

The application includes several advanced options:
- Choose between PDF or PPTX output formats (PPTX is default)
- Provide a custom hypothesis if the auto-generated one doesn't fit your needs
- Enable AI-enhanced content generation

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.