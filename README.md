# Document Conversion Service

A FastAPI-based document conversion service that supports conversion between multiple document formats.
---

English | [中文](README_zh.md)

## Features

- 📄 **PDF Conversion**: PDF to Word, Excel, PowerPoint, Images, Markdown
- 📝 **Office Conversion**: High-quality document conversion based on LibreOffice
- 🖼️ **Image Conversion**: Images to PDF, Office documents
- 📝 **Markdown Conversion**: Markdown to PDF, Word, Excel, PowerPoint
- 🔄 **Batch Conversion**: Support for batch file conversion with concurrent processing
- 🚀 **High Performance**: Based on LibreOffice process pool, asynchronous processing
- 📊 **Format Preservation**: Perfect preservation of original document formats and styles
- 🛡️ **Secure & Reliable**: File validation, error handling, logging
- 🐳 **Docker Support**: Complete containerized deployment solution

## Supported Formats

### Input Formats
- **PDF**: `.pdf`
- **Word**: `.doc`, `.docx`, `.odt`, `.rtf`
- **Excel**: `.xls`, `.xlsx`, `.ods`
- **PowerPoint**: `.ppt`, `.pptx`, `.odp`
- **Markdown**: `.md`, `.markdown`
- **Others**: `.html`, `.txt`
- **Images**: `.jpg`, `.jpeg`, `.png`, `.gif`, `.bmp`, `.tiff`, `.tif`, `.webp`

### Output Formats
- **PDF**: `.pdf` (all formats)
- **Word**: `.docx`, `.doc`, `.odt`, `.rtf`
- **Excel**: `.xlsx`, `.xls`, `.ods`
- **PowerPoint**: `.pptx`, `.ppt`, `.odp`
- **Markdown**: `.md`
- **Others**: `.html`, `.txt`
- **Images**: `.jpg`, `.png`, `.gif`, `.bmp`, `.tiff`

## Quick Start

### Requirements

- Python 3.11+
- System dependencies:
  - Tesseract OCR
  - Poppler (PDF tools)
  - LibreOffice (Office document processing)
  - libmagic (file type detection)

### Installation

1. **Clone the repository**
```bash
git clone <repository-url>
cd doc-conversion-service
```

2. **Create virtual environment**
```bash
python -m venv .ven
source .ven/bin/activate  # Linux/Mac
# or
.ven\Scripts\activate  # Windows
```

3. **Install Python dependencies**
```bash
pip install -r requirements.txt
```

4. **Install system dependencies**

Ubuntu/Debian:
```bash
sudo apt update
sudo apt install -y tesseract-ocr poppler-utils libreoffice libmagic1
```

CentOS/RHEL:
```bash
sudo yum install -y tesseract poppler-utils libreoffice libmagic
```

macOS:
```bash
brew install tesseract poppler libreoffice libmagic
```

5. **Configure environment variables**
```bash
# Project uses default configuration, no additional setup required
```

### Running the Service

1. **Development mode**
```bash
python -m uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

2. **Production mode**
```bash
uvicorn app.main:app --host 0.0.0.0 --port 8000
```

3. **Using Docker**
```bash
# Build image
docker build -t doc-conversion-service .

# Run container
docker run -d \
  --name doc-conversion \
  -p 8000:8000 \
  -v $(pwd)/outputs:/app/outputs \
  -v $(pwd)/temp:/app/temp \
  doc-conversion-service

# Using Docker Compose
docker-compose up -d
```

## API Documentation

After starting the service, visit the following addresses to view API documentation:

- **Swagger UI**: http://localhost:8000/docs
- **ReDoc**: http://localhost:8000/redoc

### Main API Endpoints

- `POST /api/convert` - Document conversion (returns download link)
- `POST /api/convert/download` - Document conversion (directly returns file)
- `GET /api/download/{file_id}` - Download conversion result
- `GET /health` - Health check

### API Usage Examples

#### 1. Basic conversion (returns download link)
```bash
curl -X POST "http://localhost:8000/api/convert" \
  -H "accept: application/json" \
  -H "Content-Type: multipart/form-data" \
  -F "file=@document.pdf" \
  -F "target_format=docx"
```

#### 2. Direct download of conversion result
```bash
curl -X POST "http://localhost:8000/api/convert/download" \
  -H "accept: application/octet-stream" \
  -H "Content-Type: multipart/form-data" \
  -F "file=@document.pdf" \
  -F "target_format=docx" \
  --output converted_document.docx
```

## Project Structure

```
doc-conversion-service/
├── app/
│   ├── api/           # API routes and models
│   ├── config/        # Configuration management
│   ├── converters/    # Converter modules
│   │   ├── libreoffice_converter.py  # LibreOffice converter
│   │   ├── pdf_converter.py          # PDF converter
│   │   ├── image_converter.py        # Image converter
│   │   ├── markdown_converter.py     # Markdown converter
│   │   └── document_to_image_converter.py  # Document to image
│   ├── services/      # Business logic
│   ├── utils/         # Utility functions
│   └── main.py        # Application entry point
├── bugs/              # Bug fix related files
│   ├── tests/         # Bug fix test files
│   ├── docs/          # Bug fix summary documents
│   └── fixes/         # Bug fix code backups
├── tests/             # Test files
├── logs/              # Log files
├── temp/              # Temporary files
├── outputs/           # Output files
├── requirements.txt   # Python dependencies
├── Dockerfile         # Docker image configuration
├── docker-compose.yml # Docker Compose configuration
├── check_dependencies.py  # Dependency check script
└── README.md         # Project documentation
```

## Development Guide

### Code Standards

- Use **Black** for code formatting
- Use **Flake8** for code linting
- Use **MyPy** for type checking
- Follow **PEP 8** coding standards

### Development Commands

```bash
# Code formatting
black app/ tests/

# Code linting
flake8 app/ tests/

# Type checking
mypy app/

# Run tests
pytest

# Dependency check
python3 check_dependencies.py

# Run tests with coverage report
pytest --cov=app --cov-report=html
```



### Adding New Converters

1. Create a new converter class in `app/converters/` directory
2. Implement conversion methods
3. Register the converter in `app/services/conversion.py`
4. Update supported conversion formats in `app/config/settings.py`
5. Add corresponding tests

## Deployment

### Docker Deployment

```bash
# Build image
docker build -t doc-conversion-service .

# Run container
docker run -d \
  --name doc-conversion \
  -p 8000:8000 \
  -v $(pwd)/outputs:/app/outputs \
  -v $(pwd)/temp:/app/temp \
  doc-conversion-service
```

### Using Docker Compose

```bash
# Start all services
docker-compose up -d

# View logs
docker-compose logs -f

# Stop services
docker-compose down
```

### Production Deployment

```bash
# Build production image
docker build -t xxxxx/library/wcl/doc-conversion-service .

# Push to image registry
docker push xxxxx/library/wcl/doc-conversion-service:latest

# Using Podman (if needed)
podman build -t xxxxx/library/wcl/doc-conversion-service:latest .
podman push --tls-verify=false xxxxx/library/wcl/doc-conversion-service:latest
```

## Monitoring and Logging

- Log files: `logs/app.log`
- Health check: `GET /health`
- Monitoring metrics: Can be integrated with Prometheus
- Dependency check: `python3 check_dependencies.py`

## Troubleshooting

### Common Issues

1. **Docker deployment failure**
   - Check if system dependencies are completely installed
   - Run `python3 check_dependencies.py` to check dependencies
   - View Docker logs: `docker-compose logs`

2. **Conversion failure**
   - Check if file format is supported
   - View application logs: `tail -f logs/app.log`
   - Verify LibreOffice is correctly installed

3. **Dependency issues**
   - Run dependency check: `python3 check_dependencies.py`
   - Reinstall dependencies: `pip install -r requirements.txt`

## Contributing

1. Fork the project
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Create a Pull Request

## License

This project is licensed under AGPL-3.0 License - see the [LICENSE](LICENSE) file for details.


## Contact

- Project Maintainer: [Your Name]
- Email: [canduecho@gmail.com]
- Project Link: [https://github.com/canduecho/doc-conversion-service]



