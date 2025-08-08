"""
应用配置文件
"""
import os
from pathlib import Path
from typing import Dict, List

# 加载 .env 文件
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    # 如果没有 python-dotenv，尝试手动加载
    env_file = Path(__file__).resolve().parent.parent.parent / '.env'
    if env_file.exists():
        with open(env_file, 'r') as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('#') and '=' in line:
                    key, value = line.split('=', 1)
                    os.environ[key] = value

# 项目根目录
BASE_DIR = Path(__file__).resolve().parent.parent.parent

# 文件存储配置
TEMP_DIR = BASE_DIR / 'temp'
OUTPUT_DIR = BASE_DIR / 'outputs'

# 确保目录存在
TEMP_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

# 文件上传配置
MAX_FILE_SIZE = 100 * 1024 * 1024  # 100MB
ALLOWED_EXTENSIONS = {
    # PDF 格式
    'pdf': ['pdf'],
    
    # Office 格式 (LibreOffice 支持)
    'word': ['doc', 'docx', 'odt', 'rtf'],
    'excel': ['xls', 'xlsx', 'ods'],
    'powerpoint': ['ppt', 'pptx', 'odp'],
    
    # 文本格式
    'text': ['txt', 'html', 'md', 'markdown'],
    
    # 图片格式
    'image': ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'tiff', 'tif', 'webp']
}

# 支持的转换格式 (基于 LibreOffice 能力优化 + 跨类型转换)
SUPPORTED_CONVERSIONS = {
    # LibreOffice 支持的转换 (包括图片输出)
    'doc': ['pdf', 'docx', 'odt', 'rtf', 'html', 'jpg', 'png', 'gif'],
    'docx': ['pdf', 'doc', 'odt', 'rtf', 'html', 'jpg', 'png', 'gif', 'xlsx', 'pptx'],  # 支持跨类型转换
    'xls': ['pdf', 'xlsx', 'ods', 'html', 'jpg', 'png', 'gif'],
    'xlsx': ['pdf', 'xls', 'ods', 'html', 'jpg', 'png', 'gif', 'docx', 'pptx'],         # 支持跨类型转换
    'ppt': ['pdf', 'pptx', 'odp', 'html', 'jpg', 'png', 'gif'],
    'pptx': ['pdf', 'ppt', 'odp', 'html', 'jpg', 'png', 'gif', 'docx', 'xlsx'],         # 支持跨类型转换
    'odt': ['pdf', 'docx', 'doc', 'rtf', 'html', 'jpg', 'png', 'gif'],
    'ods': ['pdf', 'xlsx', 'xls', 'html', 'jpg', 'png', 'gif'],
    'odp': ['pdf', 'pptx', 'ppt', 'html', 'jpg', 'png', 'gif'],
    'rtf': ['pdf', 'docx', 'odt', 'html', 'jpg', 'png', 'gif'],
    'html': ['pdf', 'docx', 'odt', 'rtf', 'jpg', 'png', 'gif'],
    'txt': ['pdf', 'docx', 'odt', 'rtf', 'html', 'jpg', 'png', 'gif'],
    
    # PDF 转换 (使用专用转换器)
    'pdf': ['docx', 'xlsx', 'pptx', 'jpg', 'png', 'gif', 'md', 'markdown'],
    
    # Markdown 转换
    'md': ['pdf', 'docx', 'xlsx', 'pptx'],
    'markdown': ['pdf', 'docx', 'xlsx', 'pptx'],
    
    # 图片转换 (包括转 Office 格式)
    'jpg': ['pdf', 'png', 'gif', 'bmp', 'tiff', 'docx', 'pptx', 'xlsx'],
    'jpeg': ['pdf', 'png', 'gif', 'bmp', 'tiff', 'docx', 'pptx', 'xlsx'],
    'png': ['pdf', 'jpg', 'gif', 'bmp', 'tiff', 'docx', 'pptx', 'xlsx'],
    'gif': ['pdf', 'jpg', 'png', 'bmp', 'tiff', 'docx', 'pptx', 'xlsx'],
    'bmp': ['pdf', 'jpg', 'png', 'gif', 'tiff', 'docx', 'pptx', 'xlsx'],
    'tiff': ['pdf', 'jpg', 'png', 'gif', 'bmp', 'docx', 'pptx', 'xlsx'],
    'tif': ['pdf', 'jpg', 'png', 'gif', 'bmp', 'docx', 'pptx', 'xlsx'],
    'webp': ['pdf', 'jpg', 'png', 'gif', 'bmp', 'docx', 'pptx', 'xlsx']
}

# 转换选项配置
CONVERSION_OPTIONS = {
    'quality': ['high', 'medium', 'low'],
    'page_range': None,  # 支持页面范围，如 "1-5" 或 "1,3,5"
    'output_size': ['A4', 'letter', 'original'],
    'maintain_formatting': True,
    'ocr_enabled': False,
    'libreoffice_workers': 3  # LibreOffice 并发工作进程数
}

# 服务器配置
HOST = os.getenv('HOST', '0.0.0.0')
PORT = int(os.getenv('PORT', 8008))
DEBUG = os.getenv('DEBUG', 'True').lower() == 'true'

# 日志配置
LOG_LEVEL = os.getenv('LOG_LEVEL', 'INFO')
LOG_FILE = BASE_DIR / 'logs' / 'app.log'

# 创建日志目录
LOG_FILE.parent.mkdir(exist_ok=True)

# 安全配置
SECRET_KEY = os.getenv('SECRET_KEY', 'your-secret-key-here')
ACCESS_TOKEN_EXPIRE_MINUTES = int(os.getenv('ACCESS_TOKEN_EXPIRE_MINUTES', 30))

# 数据库配置（如果需要）
DATABASE_URL = os.getenv('DATABASE_URL', 'sqlite:///./doc_conversion.db')

# 外部服务配置
TESSERACT_PATH = os.getenv('TESSERACT_PATH', '/usr/bin/tesseract')
POPPLER_PATH = os.getenv('POPPLER_PATH', '/usr/bin/pdftoppm')
LIBREOFFICE_PATH = os.getenv('LIBREOFFICE_PATH', '/usr/bin/libreoffice')

# 缓存配置
CACHE_TTL = int(os.getenv('CACHE_TTL', 3600))  # 1小时

# 任务队列配置
REDIS_URL = os.getenv('REDIS_URL', 'redis://localhost:6379')
CELERY_BROKER_URL = os.getenv('CELERY_BROKER_URL', REDIS_URL)
CELERY_RESULT_BACKEND = os.getenv('CELERY_RESULT_BACKEND', REDIS_URL)

# 文件清理配置
TEMP_FILE_RETENTION_HOURS = int(os.getenv('TEMP_FILE_RETENTION_HOURS', 24))
OUTPUT_FILE_RETENTION_HOURS = int(os.getenv('OUTPUT_FILE_RETENTION_HOURS', 168))  # 7天 

# PDF 转换配置
USE_PDF2DOCX = os.getenv('USE_PDF2DOCX', 'false').lower() == 'true'
PDF2DOCX_FALLBACK = os.getenv('PDF2DOCX_FALLBACK', 'true').lower() == 'true' 