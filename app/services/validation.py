"""
文件验证服务
"""
from typing import Dict, Any, Optional, Tuple
from loguru import logger

from app.utils import FileUtils
from app.config import ALLOWED_EXTENSIONS, SUPPORTED_CONVERSIONS


class ValidationService:
    """文件验证服务类"""
    
    def __init__(self):
        """初始化验证服务"""
        pass
    
    def validate_file(
        self, 
        file_path: str, 
        original_name: str = None
    ) -> Tuple[bool, str]:
        """
        验证文件
        
        Args:
            file_path: 文件路径
            original_name: 原始文件名
            
        Returns:
            (是否有效, 错误消息)
        """
        try:
            # 检查文件是否存在
            if not FileUtils.get_file_size(file_path):
                return False, "文件不存在或为空"
            
            # 检查文件大小
            if not FileUtils.validate_file_size(file_path):
                return False, "文件大小超出限制"
            
            # 检查文件扩展名
            filename = original_name or FileUtils.get_file_extension(file_path)
            if not FileUtils.is_allowed_extension(filename):
                return False, "不支持的文件格式"
            
            # 检查 MIME 类型
            mime_type = FileUtils.detect_mime_type(file_path)
            if not self._is_valid_mime_type(mime_type, filename):
                return False, "文件类型不匹配"
            
            return True, "文件验证通过"
            
        except Exception as e:
            logger.error(f"文件验证失败: {e}")
            return False, f"文件验证失败: {str(e)}"
    
    def validate_conversion(
        self, 
        source_format: str, 
        target_format: str
    ) -> Tuple[bool, str]:
        """
        验证转换是否支持
        
        Args:
            source_format: 源格式
            target_format: 目标格式
            
        Returns:
            (是否支持, 错误消息)
        """
        try:
            # 检查源格式是否支持
            if source_format not in SUPPORTED_CONVERSIONS:
                return False, f"不支持的源格式: {source_format}"
            
            # 检查目标格式是否支持
            supported_targets = SUPPORTED_CONVERSIONS.get(source_format, [])
            if target_format not in supported_targets:
                return False, f"不支持从 {source_format} 转换到 {target_format}"
            
            return True, "转换验证通过"
            
        except Exception as e:
            logger.error(f"转换验证失败: {e}")
            return False, f"转换验证失败: {str(e)}"
    
    def validate_conversion_options(
        self, 
        options: Dict[str, Any]
    ) -> Tuple[bool, str]:
        """
        验证转换选项
        
        Args:
            options: 转换选项
            
        Returns:
            (是否有效, 错误消息)
        """
        try:
            if not options:
                return True, "选项验证通过"
            
            # 验证质量选项
            if 'quality' in options:
                quality = options['quality']
                if quality not in ['high', 'medium', 'low']:
                    return False, f"无效的质量选项: {quality}"
            
            # 验证页面范围
            if 'page_range' in options and options['page_range']:
                page_range = options['page_range']
                if not self._is_valid_page_range(page_range):
                    return False, f"无效的页面范围: {page_range}"
            
            # 验证输出尺寸
            if 'output_size' in options:
                output_size = options['output_size']
                if output_size not in ['A4', 'letter', 'original']:
                    return False, f"无效的输出尺寸: {output_size}"
            
            return True, "选项验证通过"
            
        except Exception as e:
            logger.error(f"选项验证失败: {e}")
            return False, f"选项验证失败: {str(e)}"
    
    def _is_valid_mime_type(self, mime_type: str, filename: str) -> bool:
        """
        检查 MIME 类型是否有效
        
        Args:
            mime_type: MIME 类型
            filename: 文件名
            
        Returns:
            是否有效
        """
        # 定义 MIME 类型映射
        mime_type_mapping = {
            'application/pdf': ['pdf'],
            'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['docx'],
            'application/msword': ['doc'],
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['xlsx'],
            'application/vnd.ms-excel': ['xls'],
            'application/vnd.openxmlformats-officedocument.presentationml.presentation': ['pptx'],
            'application/vnd.ms-powerpoint': ['ppt'],
            'image/jpeg': ['jpg', 'jpeg'],
            'image/png': ['png'],
            'image/gif': ['gif'],
            'image/bmp': ['bmp'],
            'image/tiff': ['tiff', 'tif'],
            'image/webp': ['webp']
        }
        
        # 获取文件扩展名
        extension = FileUtils.get_file_extension(filename)
        
        # 检查 MIME 类型是否匹配
        if mime_type in mime_type_mapping:
            return extension in mime_type_mapping[mime_type]
        
        # 如果 MIME 类型不在映射中，检查是否为通用类型
        generic_types = [
            'application/octet-stream',
            'text/plain',
            'application/zip'
        ]
        
        return mime_type in generic_types
    
    def _is_valid_page_range(self, page_range: str) -> bool:
        """
        验证页面范围格式
        
        Args:
            page_range: 页面范围字符串
            
        Returns:
            是否有效
        """
        try:
            if not page_range:
                return True
            
            # 支持格式: "1-5", "1,3,5", "1-3,5-7"
            parts = page_range.split(',')
            
            for part in parts:
                part = part.strip()
                if '-' in part:
                    # 范围格式: "1-5"
                    start, end = part.split('-')
                    if not start.isdigit() or not end.isdigit():
                        return False
                    if int(start) > int(end):
                        return False
                else:
                    # 单个页面: "1"
                    if not part.isdigit():
                        return False
            
            return True
            
        except Exception:
            return False
    
    def get_file_validation_info(self, file_path: str) -> Dict[str, Any]:
        """
        获取文件验证信息
        
        Args:
            file_path: 文件路径
            
        Returns:
            验证信息
        """
        try:
            file_info = FileUtils.get_file_info(file_path)
            
            # 执行验证
            is_valid, message = self.validate_file(file_path)
            
            return {
                'file_info': file_info,
                'is_valid': is_valid,
                'message': message,
                'validation_details': {
                    'file_exists': bool(FileUtils.get_file_size(file_path)),
                    'file_size_valid': FileUtils.validate_file_size(file_path),
                    'extension_valid': FileUtils.is_allowed_extension(file_info.get('filename', '')),
                    'mime_type_valid': self._is_valid_mime_type(
                        file_info.get('content_type', ''), 
                        file_info.get('filename', '')
                    )
                }
            }
            
        except Exception as e:
            logger.error(f"获取文件验证信息失败: {e}")
            return {
                'file_info': {},
                'is_valid': False,
                'message': f"获取验证信息失败: {str(e)}",
                'validation_details': {}
            }
    
    def get_supported_conversions_for_format(self, source_format: str) -> list:
        """
        获取指定格式支持的所有转换
        
        Args:
            source_format: 源格式
            
        Returns:
            支持的目标格式列表
        """
        return SUPPORTED_CONVERSIONS.get(source_format, [])
    
    def get_all_supported_formats(self) -> Dict[str, list]:
        """
        获取所有支持的格式
        
        Returns:
            支持的格式字典
        """
        return {
            'input_formats': ALLOWED_EXTENSIONS,
            'conversion_matrix': SUPPORTED_CONVERSIONS
        } 