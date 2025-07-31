# 文档转换服务

一个基于 FastAPI 的文档转换服务，支持多种格式文档的相互转换。

## 功能特性

- 📄 **PDF 转换**: PDF 转 Word、Excel、PowerPoint、图片、Markdown
- 📝 **Office 转换**: 基于 LibreOffice 的高质量文档转换
- 🖼️ **图片转换**: 图片转 PDF、Office 文档
- 📝 **Markdown 转换**: Markdown 转 PDF、Word、Excel、PowerPoint
- 🔄 **批量转换**: 支持批量文件转换，并发处理
- 🚀 **高性能**: 基于 LibreOffice 进程池，异步处理
- 📊 **格式保持**: 完美保持原始文档格式和样式
- 🛡️ **安全可靠**: 文件验证、错误处理、日志记录
- 🐳 **Docker 支持**: 完整的容器化部署方案

## 支持的格式

### 输入格式
- **PDF**: `.pdf`
- **Word**: `.doc`, `.docx`, `.odt`, `.rtf`
- **Excel**: `.xls`, `.xlsx`, `.ods`
- **PowerPoint**: `.ppt`, `.pptx`, `.odp`
- **Markdown**: `.md`, `.markdown`
- **其他**: `.html`, `.txt`
- **图片**: `.jpg`, `.jpeg`, `.png`, `.gif`, `.bmp`, `.tiff`, `.tif`, `.webp`

### 输出格式
- **PDF**: `.pdf` (所有格式)
- **Word**: `.docx`, `.doc`, `.odt`, `.rtf`
- **Excel**: `.xlsx`, `.xls`, `.ods`
- **PowerPoint**: `.pptx`, `.ppt`, `.odp`
- **Markdown**: `.md`
- **其他**: `.html`, `.txt`
- **图片**: `.jpg`, `.png`, `.gif`, `.bmp`, `.tiff`

## 快速开始

### 环境要求

- Python 3.11+
- 系统依赖:
  - Tesseract OCR
  - Poppler (PDF 工具)
  - LibreOffice (Office 文档处理)
  - libmagic (文件类型检测)

### 安装依赖

1. **克隆项目**
```bash
git clone <repository-url>
cd doc-conversion-service
```

2. **创建虚拟环境**
```bash
python -m venv .ven
source .ven/bin/activate  # Linux/Mac
# 或
.ven\Scripts\activate  # Windows
```

3. **安装 Python 依赖**
```bash
pip install -r requirements.txt
```

4. **安装系统依赖**

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

5. **配置环境变量**
```bash
# 项目使用默认配置，无需额外配置
```

### 运行服务

1. **开发模式**
```bash
python -m uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

2. **生产模式**
```bash
uvicorn app.main:app --host 0.0.0.0 --port 8000
```

3. **使用 Docker**
```bash
# 构建镜像
docker build -t doc-conversion-service .

# 运行容器
docker run -d \
  --name doc-conversion \
  -p 8000:8000 \
  -v $(pwd)/outputs:/app/outputs \
  -v $(pwd)/temp:/app/temp \
  doc-conversion-service

# 使用 Docker Compose
docker-compose up -d
```

## API 文档

启动服务后，访问以下地址查看 API 文档：

- **Swagger UI**: http://localhost:8000/docs
- **ReDoc**: http://localhost:8000/redoc

### 主要 API 端点

- `POST /api/convert` - 文档转换（返回下载链接）
- `POST /api/convert/download` - 文档转换（直接返回文件）
- `GET /api/download/{file_id}` - 下载转换结果
- `GET /health` - 健康检查

### API 使用示例

#### 1. 基本转换（返回下载链接）
```bash
curl -X POST "http://localhost:8000/api/convert" \
  -H "accept: application/json" \
  -H "Content-Type: multipart/form-data" \
  -F "file=@document.pdf" \
  -F "target_format=docx"
```

#### 2. 直接下载转换结果
```bash
curl -X POST "http://localhost:8000/api/convert/download" \
  -H "accept: application/octet-stream" \
  -H "Content-Type: multipart/form-data" \
  -F "file=@document.pdf" \
  -F "target_format=docx" \
  --output converted_document.docx
```

## 项目结构

```
doc-conversion-service/
├── app/
│   ├── api/           # API 路由和模型
│   ├── config/        # 配置管理
│   ├── converters/    # 转换器模块
│   │   ├── libreoffice_converter.py  # LibreOffice 转换器
│   │   ├── pdf_converter.py          # PDF 转换器
│   │   ├── image_converter.py        # 图片转换器
│   │   ├── markdown_converter.py     # Markdown 转换器
│   │   └── document_to_image_converter.py  # 文档转图片
│   ├── services/      # 业务逻辑
│   ├── utils/         # 工具函数
│   └── main.py        # 应用入口
├── bugs/              # Bug 修复相关文件
│   ├── tests/         # Bug 修复测试文件
│   ├── docs/          # Bug 修复总结文档
│   └── fixes/         # Bug 修复代码备份
├── tests/             # 测试文件
├── logs/              # 日志文件
├── temp/              # 临时文件
├── outputs/           # 输出文件
├── requirements.txt   # Python 依赖
├── Dockerfile         # Docker 镜像配置
├── docker-compose.yml # Docker Compose 配置
├── check_dependencies.py  # 依赖检查脚本
└── README.md         # 项目说明
```

## 开发指南

### 代码规范

- 使用 **Black** 进行代码格式化
- 使用 **Flake8** 进行代码检查
- 使用 **MyPy** 进行类型检查
- 遵循 **PEP 8** 编码规范

### 开发命令

```bash
# 代码格式化
black app/ tests/

# 代码检查
flake8 app/ tests/

# 类型检查
mypy app/

# 运行测试
pytest

# 依赖检查
python3 check_dependencies.py

# 运行测试并生成覆盖率报告
pytest --cov=app --cov-report=html
```

### Bug 修复流程

项目使用结构化的 Bug 修复流程：

1. **创建测试文件**: 在 `bugs/tests/` 目录下创建测试文件
2. **修复代码**: 在相应模块中修复问题
3. **创建总结文档**: 在 `bugs/docs/` 目录下创建修复总结
4. **备份修复代码**: 在 `bugs/fixes/` 目录下备份重要修复

### 添加新的转换器

1. 在 `app/converters/` 目录下创建新的转换器类
2. 实现转换方法
3. 在 `app/services/conversion.py` 中注册转换器
4. 在 `app/config/settings.py` 中更新支持的转换格式
5. 添加相应的测试

## 部署

### Docker 部署

```bash
# 构建镜像
docker build -t doc-conversion-service .

# 运行容器
docker run -d \
  --name doc-conversion \
  -p 8000:8000 \
  -v $(pwd)/outputs:/app/outputs \
  -v $(pwd)/temp:/app/temp \
  doc-conversion-service
```

### 使用 Docker Compose

```bash
# 启动所有服务
docker-compose up -d

# 查看日志
docker-compose logs -f

# 停止服务
docker-compose down
```

### 生产环境部署

```bash
# 构建生产镜像
docker build -t XXX/library/wcl/doc-conversion-service .

# 推送到镜像仓库
docker push XXX/library/wcl/doc-conversion-service:latest

# 使用 Podman（如果需要）
podman build -t XXX/library/wcl/doc-conversion-service:latest .
podman push --tls-verify=false XXX/library/wcl/doc-conversion-service:latest
```

## 监控和日志

- 日志文件: `logs/app.log`
- 健康检查: `GET /health`


## 故障排除

### 常见问题

1. **Docker 部署失败**
   - 检查系统依赖是否安装完整
   - 运行 `python3 check_dependencies.py` 检查依赖
   - 查看 Docker 日志: `docker-compose logs`

2. **转换失败**
   - 检查文件格式是否支持
   - 查看应用日志: `tail -f logs/app.log`
   - 验证 LibreOffice 是否正确安装

3. **依赖问题**
   - 运行依赖检查: `python3 check_dependencies.py`
   - 重新安装依赖: `pip install -r requirements.txt`



## 许可证

本项目采用 AGPL-3.0 许可证 - 查看 [LICENSE](LICENSE) 文件了解详情。


## 联系方式

- 项目维护者: [canduecho]
- 邮箱: [canduecho@gmail.com]
- 项目链接: [https://github.com/canduecho/doc-conversion-service]

---
