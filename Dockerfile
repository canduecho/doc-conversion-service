# 使用 Python 3.11 作为基础镜像
FROM python:3.11-slim

# 设置工作目录
WORKDIR /app

# 设置环境变量
ENV PYTHONPATH=/app
ENV PYTHONUNBUFFERED=1
ENV DEBIAN_FRONTEND=noninteractive

# 安装系统依赖
RUN apt-get update && apt-get install -y \
    # LibreOffice 相关依赖
    libreoffice \
    libreoffice-writer \
    libreoffice-calc \
    libreoffice-impress \
    # PDF 处理依赖
    poppler-utils \
    # 图片处理依赖
    imagemagick \
    # OCR 依赖（可选）
    tesseract-ocr \
    tesseract-ocr-chi-sim \
    tesseract-ocr-eng \
    # 字体支持
    fonts-liberation \
    fonts-dejavu \
    fonts-wqy-microhei \
    fonts-wqy-zenhei \
    # 文件类型检测依赖
    libmagic1 \
    # 其他必要工具
    curl \
    wget \
    unzip \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# 复制 requirements.txt
COPY requirements.txt .

# 安装 Python 依赖
RUN pip install --no-cache-dir --upgrade pip \
    && pip install --no-cache-dir -r requirements.txt

# 复制应用代码
COPY . .

# 创建必要的目录
RUN mkdir -p /app/uploads /app/downloads /app/temp /app/logs /app/outputs

# 设置权限
RUN chmod +x /app/start.sh

# 暴露端口
EXPOSE 8000

# 健康检查
HEALTHCHECK --interval=30s --timeout=30s --start-period=5s --retries=3 \
    CMD curl -f http://localhost:8000/health || exit 1

# 启动命令
CMD ["python", "-m", "uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8000"] 