"""
API 数据模型
"""
from typing import Optional, Dict, Any, List
from pydantic import BaseModel, Field
from enum import Enum


class ConversionFormat(str, Enum):
    """支持的转换格式枚举"""
    PDF = 'pdf'
    DOCX = 'docx'
    XLSX = 'xlsx'
    PPTX = 'pptx'
    JPG = 'jpg'
    PNG = 'png'
    GIF = 'gif'
    BMP = 'bmp'
    TIFF = 'tiff'


class QualityLevel(str, Enum):
    """转换质量等级"""
    HIGH = 'high'
    MEDIUM = 'medium'
    LOW = 'low'


class OutputSize(str, Enum):
    """输出尺寸"""
    A4 = 'A4'
    LETTER = 'letter'
    ORIGINAL = 'original'


class ConversionOptions(BaseModel):
    """转换选项"""
    quality: QualityLevel = Field(default=QualityLevel.MEDIUM, description='转换质量')
    page_range: Optional[str] = Field(default=None, description='页面范围，如 "1-5" 或 "1,3,5"')
    output_size: OutputSize = Field(default=OutputSize.ORIGINAL, description='输出尺寸')
    maintain_formatting: bool = Field(default=True, description='是否保持格式')
    ocr_enabled: bool = Field(default=False, description='是否启用 OCR')


class ConversionRequest(BaseModel):
    """转换请求模型"""
    target_format: ConversionFormat = Field(..., description='目标格式')
    options: Optional[ConversionOptions] = Field(default=None, description='转换选项')


class ConversionResponse(BaseModel):
    """转换响应模型"""
    success: bool = Field(..., description='转换是否成功')
    message: str = Field(..., description='响应消息')
    file_url: Optional[str] = Field(default=None, description='转换后文件的下载链接')
    file: Optional[str] = Field(default=None, description='转换后的文件名')
    file_size: Optional[int] = Field(default=None, description='文件大小（字节）')
    conversion_time: Optional[float] = Field(default=None, description='转换耗时（秒）')
    error_details: Optional[str] = Field(default=None, description='错误详情')


class FormatInfo(BaseModel):
    """格式信息"""
    format_name: str = Field(..., description='格式名称')
    extensions: List[str] = Field(..., description='支持的扩展名')
    description: str = Field(..., description='格式描述')


class SupportedFormatsResponse(BaseModel):
    """支持的格式响应"""
    input_formats: Dict[str, FormatInfo] = Field(..., description='支持的输入格式')
    output_formats: Dict[str, FormatInfo] = Field(..., description='支持的输出格式')
    conversion_matrix: Dict[str, List[str]] = Field(..., description='转换矩阵')


class HealthResponse(BaseModel):
    """健康检查响应"""
    status: str = Field(..., description='服务状态')
    version: str = Field(..., description='服务版本')
    uptime: float = Field(..., description='运行时间（秒）')
    dependencies: Dict[str, str] = Field(..., description='依赖服务状态')


class ErrorResponse(BaseModel):
    """错误响应模型"""
    error: str = Field(..., description='错误类型')
    message: str = Field(..., description='错误消息')
    details: Optional[Dict[str, Any]] = Field(default=None, description='错误详情')


class FileInfo(BaseModel):
    """文件信息"""
    filename: str = Field(..., description='文件名')
    original_name: str = Field(..., description='原始文件名')
    file_size: int = Field(..., description='文件大小（字节）')
    content_type: str = Field(..., description='内容类型')
    extension: str = Field(..., description='文件扩展名')
    upload_time: str = Field(..., description='上传时间')


class BatchConversionRequest(BaseModel):
    """批量转换请求"""
    files: List[str] = Field(..., description='文件 ID 列表')
    target_format: ConversionFormat = Field(..., description='目标格式')
    options: Optional[ConversionOptions] = Field(default=None, description='转换选项')


class BatchConversionResponse(BaseModel):
    """批量转换响应"""
    batch_id: str = Field(..., description='批量转换 ID')
    total_files: int = Field(..., description='总文件数')
    completed_files: int = Field(..., description='已完成文件数')
    failed_files: int = Field(..., description='失败文件数')
    results: List[ConversionResponse] = Field(..., description='转换结果列表')
    status: str = Field(..., description='批量转换状态') 