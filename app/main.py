"""
文档转换服务主应用
"""
import time
from contextlib import asynccontextmanager
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import JSONResponse
from loguru import logger

from app.config import (
    DEBUG, LOG_LEVEL, LOG_FILE, TEMP_FILE_RETENTION_HOURS, OUTPUT_FILE_RETENTION_HOURS,
    TEMP_DIR, OUTPUT_DIR
)
from app.utils import FileUtils
from app.api.routes import router as api_router


# 配置日志
logger.remove()
logger.add(
    LOG_FILE,
    level=LOG_LEVEL,
    rotation="10 MB",
    retention="7 days",
    format="{time:YYYY-MM-DD HH:mm:ss} | {level} | {name}:{function}:{line} | {message}"
)
logger.add(
    lambda msg: print(msg, end=""),
    level=LOG_LEVEL,
    format="{time:HH:mm:ss} | {level} | {message}"
)


# 应用启动和关闭事件
@asynccontextmanager
async def lifespan(app: FastAPI):
    """应用生命周期管理"""
    # 启动时执行
    logger.info("文档转换服务启动中...")
    start_time = time.time()
    
    # 确保目录存在
    FileUtils.ensure_directory_exists(TEMP_DIR)
    FileUtils.ensure_directory_exists(OUTPUT_DIR)
    
    # 清理旧文件
    temp_cleaned = FileUtils.cleanup_old_files(TEMP_DIR, TEMP_FILE_RETENTION_HOURS)
    output_cleaned = FileUtils.cleanup_old_files(OUTPUT_DIR, OUTPUT_FILE_RETENTION_HOURS)
    
    logger.info(f"启动完成，清理了 {temp_cleaned} 个临时文件和 {output_cleaned} 个输出文件")
    
    yield
    
    # 关闭时执行
    logger.info("文档转换服务关闭中...")
    # 清理临时文件
    temp_cleaned = FileUtils.cleanup_old_files(TEMP_DIR, 0)
    logger.info(f"服务关闭，清理了 {temp_cleaned} 个临时文件")


# 创建 FastAPI 应用
app = FastAPI(
    title="文档转换服务",
    description="支持多种格式文档相互转换的 RESTful API 服务",
    version="1.0.0",
    debug=DEBUG,
    lifespan=lifespan
)

# 配置 CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # 生产环境中应该限制具体域名
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 挂载静态文件目录
app.mount("/outputs", StaticFiles(directory="outputs"), name="outputs")

# 注册路由
app.include_router(api_router, prefix="/api", tags=["文档转换"])


@app.get("/", tags=["根路径"])
async def root():
    """根路径"""
    return {
        "message": "文档转换服务",
        "version": "1.0.0",
        "docs": "/docs",
        "health": "/health"
    }


@app.get("/health", tags=["健康检查"])
async def health_check():
    """健康检查"""
    try:
        # 检查关键目录
        temp_dir_ok = TEMP_DIR.exists()
        output_dir_ok = OUTPUT_DIR.exists()
        
        # 检查依赖服务（这里可以添加更多检查）
        dependencies = {
            "temp_directory": "healthy" if temp_dir_ok else "unhealthy",
            "output_directory": "healthy" if output_dir_ok else "unhealthy",
            "file_system": "healthy"
        }
        
        return {
            "status": "healthy",
            "version": "1.0.0",
            "uptime": time.time(),
            "dependencies": dependencies
        }
    except Exception as e:
        logger.error(f"健康检查失败: {e}")
        raise HTTPException(status_code=503, detail="服务不可用")


@app.exception_handler(404)
async def not_found_handler(request, exc):
    """404 错误处理"""
    return JSONResponse(
        status_code=404,
        content={
            "error": "Not Found",
            "message": "请求的资源不存在",
            "path": str(request.url.path)
        }
    )


@app.exception_handler(500)
async def internal_error_handler(request, exc):
    """500 错误处理"""
    logger.error(f"内部服务器错误: {exc}")
    return JSONResponse(
        status_code=500,
        content={
            "error": "Internal Server Error",
            "message": "服务器内部错误",
            "path": str(request.url.path)
        }
    )


if __name__ == "__main__":
    import uvicorn
    from app.config import HOST, PORT
    
    logger.info(f"启动服务器: {HOST}:{PORT}")
    uvicorn.run(
        "app.main:app",
        host=HOST,
        port=PORT,
        reload=DEBUG,
        log_level=LOG_LEVEL.lower()
    ) 