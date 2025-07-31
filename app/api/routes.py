"""
API 路由定义
"""
import time
from typing import List
from fastapi import APIRouter, UploadFile, File, Form, HTTPException, Depends
from fastapi.responses import FileResponse
from loguru import logger

from app.api.models import (
    ConversionRequest, ConversionResponse, SupportedFormatsResponse,
    FormatInfo, HealthResponse, ErrorResponse, FileInfo,
    BatchConversionRequest, BatchConversionResponse
)
from app.config import SUPPORTED_CONVERSIONS, ALLOWED_EXTENSIONS, TEMP_DIR, OUTPUT_DIR
from app.utils import FileUtils
from app.services.conversion import ConversionService
from app.services.validation import ValidationService

router = APIRouter()


@router.post("/convert", response_model=ConversionResponse)
async def convert_document(
    file: UploadFile = File(..., description="要转换的文件"),
    target_format: str = Form(..., description="目标格式"),
    quality: str = Form("medium", description="转换质量"),
    page_range: str = Form(None, description="页面范围"),
    output_size: str = Form("original", description="输出尺寸"),
    maintain_formatting: bool = Form(True, description="是否保持格式"),
    ocr_enabled: bool = Form(False, description="是否启用 OCR")
):
    """
    转换文档格式
    
    - **file**: 要转换的文件
    - **target_format**: 目标格式 (pdf, docx, xlsx, pptx, jpg, png, etc.)
    - **quality**: 转换质量 (high, medium, low)
    - **page_range**: 页面范围，如 "1-5" 或 "1,3,5"
    - **output_size**: 输出尺寸 (A4, letter, original)
    - **maintain_formatting**: 是否保持格式
    - **ocr_enabled**: 是否启用 OCR
    """
    start_time = time.time()
    
    try:
        # 验证文件
        if not file.filename:
            raise HTTPException(status_code=400, detail="文件名不能为空")
        
        if not FileUtils.is_allowed_extension(file.filename):
            raise HTTPException(status_code=400, detail="不支持的文件格式")
        
        # 验证目标格式
        source_extension = FileUtils.get_file_extension(file.filename)
        if target_format not in SUPPORTED_CONVERSIONS.get(source_extension, []):
            raise HTTPException(
                status_code=400, 
                detail=f"不支持从 {source_extension} 转换到 {target_format}"
            )
        
        # 保存上传的文件
        temp_path, unique_filename = FileUtils.create_temp_file(file.filename)
        if not await FileUtils.save_uploaded_file(file, temp_path):
            raise HTTPException(status_code=500, detail="文件保存失败")
        
        # 验证文件大小
        if not FileUtils.validate_file_size(temp_path):
            FileUtils.cleanup_temp_file(temp_path)
            raise HTTPException(status_code=400, detail="文件大小超出限制")
        
        # 创建转换选项
        conversion_options = {
            'quality': quality,
            'page_range': page_range,
            'output_size': output_size,
            'maintain_formatting': maintain_formatting,
            'ocr_enabled': ocr_enabled
        }
        
        # 执行转换
        conversion_service = ConversionService()
        result = await conversion_service.convert(
            temp_path, 
            target_format, 
            conversion_options
        )
        
        if result['success']:
            # 获取输出文件信息
            output_file_path = result['output_path']
            output_filename = result['output_filename']
            file_size = FileUtils.get_file_size(output_file_path)
            
            # 清理临时文件
            FileUtils.cleanup_temp_file(temp_path)
            
            conversion_time = time.time() - start_time
            
            return ConversionResponse(
                success=True,
                message="文档转换成功",
                file_url=f"/outputs/{output_filename}",
                file=output_filename,
                file_size=file_size,
                conversion_time=conversion_time
            )
        else:
            # 清理临时文件
            FileUtils.cleanup_temp_file(temp_path)
            raise HTTPException(status_code=500, detail=result['error'])
            
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"文档转换失败: {e}")
        raise HTTPException(status_code=500, detail="文档转换失败")


@router.post("/convert/download")
async def convert_and_download_document(
    file: UploadFile = File(..., description="要转换的文件"),
    target_format: str = Form(..., description="目标格式"),
    quality: str = Form("medium", description="转换质量"),
    page_range: str = Form(None, description="页面范围"),
    output_size: str = Form("original", description="输出尺寸"),
    maintain_formatting: bool = Form(True, description="是否保持格式"),
    ocr_enabled: bool = Form(False, description="是否启用 OCR")
):
    """
    转换文档格式并直接下载
    
    - **file**: 要转换的文件
    - **target_format**: 目标格式 (pdf, docx, xlsx, pptx, jpg, png, etc.)
    - **quality**: 转换质量 (high, medium, low)
    - **page_range**: 页面范围，如 "1-5" 或 "1,3,5"
    - **output_size**: 输出尺寸 (A4, letter, original)
    - **maintain_formatting**: 是否保持格式
    - **ocr_enabled**: 是否启用 OCR
    
    注意：此端点直接返回转换后的文件，适合小文件快速转换
    """
    try:
        # 验证文件
        if not file.filename:
            raise HTTPException(status_code=400, detail="文件名不能为空")
        
        if not FileUtils.is_allowed_extension(file.filename):
            raise HTTPException(status_code=400, detail="不支持的文件格式")
        
        # 验证目标格式
        source_extension = FileUtils.get_file_extension(file.filename)
        if target_format not in SUPPORTED_CONVERSIONS.get(source_extension, []):
            raise HTTPException(
                status_code=400, 
                detail=f"不支持从 {source_extension} 转换到 {target_format}"
            )
        
        # 保存上传的文件
        temp_path, unique_filename = FileUtils.create_temp_file(file.filename)
        if not await FileUtils.save_uploaded_file(file, temp_path):
            raise HTTPException(status_code=500, detail="文件保存失败")
        
        # 验证文件大小
        if not FileUtils.validate_file_size(temp_path):
            FileUtils.cleanup_temp_file(temp_path)
            raise HTTPException(status_code=400, detail="文件大小超出限制")
        
        # 创建转换选项
        conversion_options = {
            'quality': quality,
            'page_range': page_range,
            'output_size': output_size,
            'maintain_formatting': maintain_formatting,
            'ocr_enabled': ocr_enabled
        }
        
        # 执行转换
        conversion_service = ConversionService()
        result = await conversion_service.convert(
            temp_path, 
            target_format, 
            conversion_options
        )
        
        if result['success']:
            # 获取输出文件信息
            output_file_path = result['output_path']
            output_filename = result['output_filename']
            
            # 清理临时文件
            FileUtils.cleanup_temp_file(temp_path)
            
            # 直接返回文件
            return FileResponse(
                path=output_file_path,
                filename=output_filename,
                media_type='application/octet-stream'
            )
        else:
            # 清理临时文件
            FileUtils.cleanup_temp_file(temp_path)
            raise HTTPException(status_code=500, detail=result['error'])
            
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"文档转换并下载失败: {e}")
        raise HTTPException(status_code=500, detail="文档转换并下载失败")


@router.get("/formats", response_model=SupportedFormatsResponse)
async def get_supported_formats():
    """
    获取支持的格式信息
    """
    try:
        # 构建输入格式信息
        input_formats = {}
        for format_type, extensions in ALLOWED_EXTENSIONS.items():
            input_formats[format_type] = FormatInfo(
                format_name=format_type.upper(),
                extensions=extensions,
                description=f"支持 {format_type} 格式的文件"
            )
        
        # 构建输出格式信息
        output_formats = {}
        all_output_formats = set()
        for formats in SUPPORTED_CONVERSIONS.values():
            all_output_formats.update(formats)
        
        for format_name in all_output_formats:
            output_formats[format_name] = FormatInfo(
                format_name=format_name.upper(),
                extensions=[format_name],
                description=f"输出 {format_name} 格式的文件"
            )
        
        return SupportedFormatsResponse(
            input_formats=input_formats,
            output_formats=output_formats,
            conversion_matrix=SUPPORTED_CONVERSIONS
        )
    except Exception as e:
        logger.error(f"获取支持格式失败: {e}")
        raise HTTPException(status_code=500, detail="获取支持格式失败")


@router.get("/download/{filename}")
async def download_file(filename: str):
    """
    下载转换后的文件
    
    - **filename**: 文件名
    """
    try:
        file_path = OUTPUT_DIR / filename
        
        if not file_path.exists():
            raise HTTPException(status_code=404, detail="文件不存在")
        
        return FileResponse(
            path=str(file_path),
            filename=filename,
            media_type='application/octet-stream'
        )
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"文件下载失败: {e}")
        raise HTTPException(status_code=500, detail="文件下载失败")


@router.post("/batch-convert", response_model=BatchConversionResponse)
async def batch_convert_documents(
    request: BatchConversionRequest
):
    """
    批量转换文档
    
    - **files**: 文件 ID 列表
    - **target_format**: 目标格式
    - **options**: 转换选项
    """
    try:
        # TODO: 实现批量转换逻辑
        # 这里可以使用 Celery 或其他异步任务队列
        
        return BatchConversionResponse(
            batch_id="batch_123",
            total_files=len(request.files),
            completed_files=0,
            failed_files=0,
            results=[],
            status="pending"
        )
    except Exception as e:
        logger.error(f"批量转换失败: {e}")
        raise HTTPException(status_code=500, detail="批量转换失败")


@router.get("/file-info/{filename}", response_model=FileInfo)
async def get_file_info(filename: str):
    """
    获取文件信息
    
    - **filename**: 文件名
    """
    try:
        # 检查临时文件
        temp_file_path = TEMP_DIR / filename
        if temp_file_path.exists():
            file_info = FileUtils.get_file_info(str(temp_file_path), filename)
            return FileInfo(**file_info)
        
        # 检查输出文件
        output_file_path = OUTPUT_DIR / filename
        if output_file_path.exists():
            file_info = FileUtils.get_file_info(str(output_file_path), filename)
            return FileInfo(**file_info)
        
        raise HTTPException(status_code=404, detail="文件不存在")
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"获取文件信息失败: {e}")
        raise HTTPException(status_code=500, detail="获取文件信息失败")


@router.delete("/files/{filename}")
async def delete_file(filename: str):
    """
    删除文件
    
    - **filename**: 文件名
    """
    try:
        # 检查临时文件
        temp_file_path = TEMP_DIR / filename
        if temp_file_path.exists():
            if FileUtils.cleanup_temp_file(str(temp_file_path)):
                return {"message": "临时文件删除成功"}
        
        # 检查输出文件
        output_file_path = OUTPUT_DIR / filename
        if output_file_path.exists():
            if FileUtils.cleanup_temp_file(str(output_file_path)):
                return {"message": "输出文件删除成功"}
        
        raise HTTPException(status_code=404, detail="文件不存在")
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"删除文件失败: {e}")
        raise HTTPException(status_code=500, detail="删除文件失败")


@router.post("/cleanup")
async def cleanup_files():
    """
    清理过期文件
    """
    try:
        from app.config import TEMP_FILE_RETENTION_HOURS, OUTPUT_FILE_RETENTION_HOURS
        
        temp_cleaned = FileUtils.cleanup_old_files(
            TEMP_DIR, 
            TEMP_FILE_RETENTION_HOURS
        )
        output_cleaned = FileUtils.cleanup_old_files(
            OUTPUT_DIR, 
            OUTPUT_FILE_RETENTION_HOURS
        )
        
        return {
            "message": "文件清理完成",
            "temp_files_cleaned": temp_cleaned,
            "output_files_cleaned": output_cleaned
        }
    except Exception as e:
        logger.error(f"文件清理失败: {e}")
        raise HTTPException(status_code=500, detail="文件清理失败") 