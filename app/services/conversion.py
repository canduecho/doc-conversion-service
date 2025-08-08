"""
文档转换服务
集成多种转换器，提供统一的转换接口
"""
import asyncio
import time
from typing import Dict, Any, Optional, List
from pathlib import Path
from loguru import logger

from app.config import OUTPUT_DIR, SUPPORTED_CONVERSIONS
from app.utils import FileUtils
from app.converters.libreoffice_converter import LibreOfficeConverter
from app.converters.pdf_converter import PDFConverter
from app.converters.image_converter import ImageConverter
from app.converters.document_to_image_converter import DocumentToImageConverter
from app.converters.markdown_converter import MarkdownConverter
from app.converters.cross_type_converter import CrossTypeConverter


class ConversionService:
    """文档转换服务"""
    
    def __init__(self):
        """初始化转换服务"""
        # 初始化转换器
        self.libreoffice_converter = LibreOfficeConverter(max_workers=3)
        self.pdf_converter = PDFConverter()
        self.image_converter = ImageConverter()
        self.document_to_image_converter = DocumentToImageConverter()
        self.markdown_converter = MarkdownConverter()
        self.cross_type_converter = CrossTypeConverter()
        
        # 转换器映射
        self.converter_mapping = {
            'libreoffice': self.libreoffice_converter,
            'pdf': self.pdf_converter,
            'image': self.image_converter,
            'document_to_image': self.document_to_image_converter,
            'markdown': self.markdown_converter,
            'cross_type': self.cross_type_converter
        }
    
    async def convert(
        self,
        input_path: str,
        target_format: str,
        options: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """
        转换文档
        
        Args:
            input_path: 输入文件路径
            target_format: 目标格式
            options: 转换选项
            
        Returns:
            转换结果
        """
        start_time = time.time()
        
        try:
            # 验证输入文件
            if not Path(input_path).exists():
                return {
                    'success': False,
                    'error': f'输入文件不存在: {input_path}'
                }
            
            # 获取文件扩展名
            input_ext = Path(input_path).suffix.lower().lstrip('.')
            
            # 验证转换支持
            if not self._is_conversion_supported(input_ext, target_format):
                return {
                    'success': False,
                    'error': f'不支持的转换: {input_ext} -> {target_format}'
                }
            
            # 生成输出文件路径
            output_filename = self._generate_output_filename(input_path, target_format)
            output_path = OUTPUT_DIR / output_filename
            
            # 选择转换器
            converter = self._select_converter(input_ext, target_format)
            if not converter:
                return {
                    'success': False,
                    'error': f'未找到合适的转换器: {input_ext} -> {target_format}'
                }
            
            # 执行转换
            result = await self._execute_conversion(
                converter, input_path, str(output_path), target_format, options
            )
            
            if result['success']:
                conversion_time = time.time() - start_time
                result['conversion_time'] = conversion_time
                result['output_filename'] = output_filename
                result['output_path'] = str(output_path)
            
            return result
            
        except Exception as e:
            logger.error(f"转换服务异常: {e}")
            return {
                'success': False,
                'error': f'转换服务异常: {str(e)}'
            }
    
    def _is_conversion_supported(self, input_ext: str, target_format: str) -> bool:
        """检查转换是否支持"""
        return (input_ext in SUPPORTED_CONVERSIONS and 
                target_format in SUPPORTED_CONVERSIONS[input_ext])
    
    def _select_converter(self, input_ext: str, target_format: str):
        """选择转换器"""
        # 跨类型转换检查
        if self._is_cross_type_conversion(input_ext, target_format):
            return self.cross_type_converter
        
        # Markdown 转换
        if input_ext in ['md', 'markdown']:
            return self.markdown_converter
        
        # PDF 转 Markdown
        elif input_ext == 'pdf' and target_format in ['md', 'markdown']:
            return self.markdown_converter
        
        # 文档到图片转换 (使用 PDF 中转方案)
        elif (input_ext in ['doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx', 'odt', 'ods', 'odp', 'rtf', 'txt', 'html'] and 
              target_format in ['jpg', 'png', 'gif']):
            return self.document_to_image_converter
        
        # 优先使用 LibreOffice 转换器处理 Office 文档和文本文件
        elif input_ext in ['doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx', 'odt', 'ods', 'odp', 'rtf', 'txt', 'html']:
            return self.libreoffice_converter
        
        # PDF 转换器处理 PDF 相关转换
        elif input_ext == 'pdf':
            return self.pdf_converter
        
        # 图片转换器处理图片相关转换
        elif input_ext in ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'tiff', 'tif', 'webp']:
            return self.image_converter
        
        return None
    
    def _is_cross_type_conversion(self, input_ext: str, target_format: str) -> bool:
        """检查是否为跨类型转换"""
        cross_type_conversions = {
            'docx': ['xlsx', 'pptx'],
            'xlsx': ['docx', 'pptx'],
            'pptx': ['docx', 'xlsx']
        }
        return (input_ext in cross_type_conversions and 
                target_format in cross_type_conversions[input_ext])
    
    async def _execute_conversion(
        self,
        converter,
        input_path: str,
        output_path: str,
        target_format: str,
        options: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """执行转换"""
        try:
            # 根据转换器类型调用相应方法
            if isinstance(converter, LibreOfficeConverter):
                return await converter.convert_document(
                    input_path, output_path, target_format, options
                )
            elif isinstance(converter, PDFConverter):
                # 根据目标格式选择合适的方法
                if target_format in ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'tiff', 'tif']:
                    return await converter.pdf_to_image(
                        input_path, output_path, target_format, options
                    )
                else:
                    return await converter.pdf_to_office(
                        input_path, output_path, target_format, options
                    )
            elif isinstance(converter, ImageConverter):
                # 根据目标格式选择合适的方法
                if target_format in ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'tiff', 'tif', 'webp']:
                    return await converter.image_to_image(
                        input_path, output_path, target_format, options
                    )
                elif target_format == 'pdf':
                    return await converter.image_to_pdf(
                        input_path, output_path, options
                    )
                elif target_format in ['docx', 'pptx', 'xlsx']:
                    return await converter.image_to_office(
                        input_path, output_path, target_format, options
                    )
                else:
                    return {
                        'success': False,
                        'error': f'图片转换器不支持的目标格式: {target_format}'
                    }
            elif hasattr(converter, 'convert_to_image'):  # DocumentToImageConverter
                return await converter.convert_to_image(
                    input_path, output_path, target_format, options
                )
            elif hasattr(converter, 'convert_cross_type'):  # CrossTypeConverter
                return await converter.convert_cross_type(
                    input_path, output_path, target_format, options
                )
            elif hasattr(converter, 'markdown_to_pdf'):  # MarkdownConverter
                # Markdown 转换器
                if target_format == 'pdf':
                    return await converter.markdown_to_pdf(
                        input_path, output_path, options
                    )
                elif target_format == 'docx':
                    return await converter.markdown_to_docx(
                        input_path, output_path, options
                    )
                elif target_format == 'xlsx':
                    return await converter.markdown_to_xlsx(
                        input_path, output_path, options
                    )
                elif target_format == 'pptx':
                    return await converter.markdown_to_pptx(
                        input_path, output_path, options
                    )
                elif target_format in ['md', 'markdown']:
                    return await converter.pdf_to_markdown(
                        input_path, output_path, options
                    )
                else:
                    return {
                        'success': False,
                        'error': f'Markdown 转换器不支持的目标格式: {target_format}'
                    }
            else:
                return {
                    'success': False,
                    'error': f'未知的转换器类型: {type(converter)}'
                }
                
        except Exception as e:
            logger.error(f"转换执行失败: {e}")
            return {
                'success': False,
                'error': f'转换执行失败: {str(e)}'
            }
    
    def _generate_output_filename(self, input_path: str, target_format: str) -> str:
        """生成输出文件名"""
        input_file = Path(input_path)
        timestamp = int(time.time())
        return f"{input_file.stem}_{timestamp}.{target_format}"
    
    async def batch_convert(
        self,
        files: List[Dict[str, Any]],
        options: Optional[Dict[str, Any]] = None
    ) -> List[Dict[str, Any]]:
        """
        批量转换文档
        
        Args:
            files: 文件列表，每个文件包含 input_path, target_format
            options: 转换选项
            
        Returns:
            转换结果列表
        """
        tasks = []
        
        # 创建转换任务
        for file_info in files:
            task = self.convert(
                file_info['input_path'],
                file_info['target_format'],
                options
            )
            tasks.append(task)
        
        # 并发执行转换
        results = await asyncio.gather(*tasks, return_exceptions=True)
        
        # 处理结果
        processed_results = []
        for i, result in enumerate(results):
            if isinstance(result, Exception):
                processed_results.append({
                    'success': False,
                    'error': f'转换异常: {str(result)}',
                    'file': files[i]
                })
            else:
                processed_results.append(result)
        
        return processed_results
    
    def get_supported_conversions(self) -> Dict[str, List[str]]:
        """获取支持的转换格式"""
        return SUPPORTED_CONVERSIONS
    
    def get_converter_status(self) -> Dict[str, bool]:
        """获取转换器状态"""
        return {
            'libreoffice': self.libreoffice_converter._check_libreoffice(),
            'pdf': True,  # PDF 转换器总是可用
            'image': True  # 图片转换器总是可用
        }
    
    def cleanup(self):
        """清理资源"""
        self.libreoffice_converter.cleanup() 