.PHONY: help install dev test lint format clean docker-build docker-run docker-compose-up docker-compose-down

# 默认目标
help: ## 显示帮助信息
	@echo "可用的命令:"
	@grep -E '^[a-zA-Z_-]+:.*?## .*$$' $(MAKEFILE_LIST) | sort | awk 'BEGIN {FS = ":.*?## "}; {printf "\033[36m%-20s\033[0m %s\n", $$1, $$2}'

# 安装依赖
install: ## 安装项目依赖
	python -m pip install --upgrade pip
	pip install -r requirements.txt

# 创建虚拟环境
venv: ## 创建虚拟环境
	python -m venv venv
	@echo "虚拟环境已创建，请运行: source venv/bin/activate"

# 开发模式运行
dev: ## 开发模式运行服务
	python -m app.main

# 生产模式运行
prod: ## 生产模式运行服务
	uvicorn app.main:app --host 0.0.0.0 --port 8000 --workers 4

# 运行测试
test: ## 运行测试
	pytest -v

# 运行测试并生成覆盖率报告
test-cov: ## 运行测试并生成覆盖率报告
	pytest --cov=app --cov-report=html --cov-report=term

# 代码检查
lint: ## 运行代码检查
	flake8 app/ tests/
	mypy app/

# 代码格式化
format: ## 格式化代码
	black app/ tests/
	isort app/ tests/

# 清理缓存文件
clean: ## 清理缓存和临时文件
	find . -type d -name "__pycache__" -exec rm -rf {} +
	find . -type f -name "*.pyc" -delete
	find . -type f -name "*.pyo" -delete
	find . -type f -name "*.pyd" -delete
	find . -type d -name "*.egg-info" -exec rm -rf {} +
	find . -type d -name ".pytest_cache" -exec rm -rf {} +
	find . -type d -name ".mypy_cache" -exec rm -rf {} +
	rm -rf build/ dist/ *.egg-info/

# Docker 相关命令
docker-build: ## 构建 Docker 镜像
	docker build -t doc-conversion-service .

docker-run: ## 运行 Docker 容器
	docker run -d --name doc-conversion -p 8000:8000 doc-conversion-service

docker-stop: ## 停止 Docker 容器
	docker stop doc-conversion
	docker rm doc-conversion

docker-compose-up: ## 启动所有 Docker 服务
	docker-compose up -d

docker-compose-down: ## 停止所有 Docker 服务
	docker-compose down

docker-compose-logs: ## 查看 Docker 服务日志
	docker-compose logs -f

# 系统依赖安装
install-system-deps: ## 安装系统依赖 (Ubuntu/Debian)
	sudo apt update
	sudo apt install -y tesseract-ocr tesseract-ocr-chi-sim poppler-utils libreoffice

# 数据库迁移
migrate: ## 运行数据库迁移
	alembic upgrade head

migrate-create: ## 创建新的迁移文件
	alembic revision --autogenerate -m "$(message)"

# 健康检查
health: ## 检查服务健康状态
	curl -f http://localhost:8000/health || echo "服务未运行"

# 完整设置
setup: venv install install-system-deps ## 完整设置开发环境
	@echo "开发环境设置完成！"
	@echo "请运行: source venv/bin/activate"

# 开发工作流
dev-workflow: format lint test ## 完整的开发工作流

# 部署前检查
pre-deploy: clean format lint test ## 部署前检查 