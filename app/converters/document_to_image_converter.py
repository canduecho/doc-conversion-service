"""
文档到图片转换器
使用 PDF 中转方案实现文档到图片的转换
"""

import os
import tempfile
from pathlib import Path
from typing import Dict, Any, Optional
from loguru import logger

from app.converters.libreoffice_converter import LibreOfficeConverter
from app.converters.pdf_converter import PDFConverter


class DocumentToImageConverter:
    """
    文档到图片转换器
    使用 PDF 中转方案：文档 → PDF → 图片
    """
    
    def __init__(self):
        self.libreoffice_converter = LibreOfficeConverter()
        self.pdf_converter = PDFConverter()
    
    async def convert_to_image(
        self,
        input_path: str,
        output_path: str,
        target_format: str,
        options: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """
        将文档转换为图片
        
        Args:
            input_path: 输入文件路径
            output_path: 输出文件路径
            target_format: 目标图片格式 (jpg, png, gif)
            options: 转换选项
            
        Returns:
            转换结果字典
        """
        try:
            logger.info(f"开始文档到图片转换: {input_path} → {target_format}")
            
            # 验证目标格式
            if target_format not in ['jpg', 'png', 'gif']:
                return {
                    'success': False,
                    'error': f'不支持的目标格式: {target_format}'
                }
            
            # 创建临时目录
            with tempfile.TemporaryDirectory() as temp_dir:
                temp_pdf_path = os.path.join(temp_dir, 'temp_document.pdf')
                
                # 步骤1: 文档 → PDF
                logger.info("步骤1: 文档转换为 PDF")
                pdf_result = await self._convert_document_to_pdf(
                    input_path, temp_pdf_path, options
                )
                
                if not pdf_result['success']:
                    return pdf_result
                
                # 步骤2: PDF → 图片
                logger.info("步骤2: PDF 转换为图片")
                image_result = await self._convert_pdf_to_image(
                    temp_pdf_path, output_path, target_format, options
                )
                
                if not image_result['success']:
                    return image_result
                
                logger.info(f"文档到图片转换成功: {output_path}")
                return {
                    'success': True,
                    'output_path': output_path,
                    'output_filename': os.path.basename(output_path)
                }
                
        except Exception as e:
            logger.error(f"文档到图片转换失败: {e}")
            return {
                'success': False,
                'error': f'文档到图片转换失败: {str(e)}'
            }
    
    async def _convert_document_to_pdf(
        self,
        input_path: str,
        output_path: str,
        options: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """
        将文档转换为 PDF
        
        Args:
            input_path: 输入文件路径
            output_path: 输出 PDF 路径
            options: 转换选项
            
        Returns:
            转换结果字典
        """
        try:
            # 使用 LibreOffice 转换器
            result = await self.libreoffice_converter.convert_to_pdf(
                input_path, output_path, options
            )
            
            if result['success']:
                logger.info(f"文档转 PDF 成功: {output_path}")
                return result
            else:
                logger.error(f"文档转 PDF 失败: {result['error']}")
                return result
                
        except Exception as e:
            logger.error(f"文档转 PDF 异常: {e}")
            return {
                'success': False,
                'error': f'文档转 PDF 异常: {str(e)}'
            }
    
    async def _convert_pdf_to_image(
        self,
        input_path: str,
        output_path: str,
        target_format: str,
        options: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """
        将 PDF 转换为图片
        
        Args:
            input_path: 输入 PDF 路径
            output_path: 输出图片路径
            target_format: 目标图片格式
            options: 转换选项
            
        Returns:
            转换结果字典
        """
        try:
            # 使用 PDF 转换器
            result = await self.pdf_converter.pdf_to_image(
                input_path, output_path, target_format, options
            )
            
            if result['success']:
                logger.info(f"PDF 转图片成功: {output_path}")
                return result
            else:
                logger.error(f"PDF 转图片失败: {result['error']}")
                return result
                
        except Exception as e:
            logger.error(f"PDF 转图片异常: {e}")
            return {
                'success': False,
                'error': f'PDF 转图片异常: {str(e)}'
            }
    
    def get_supported_formats(self) -> Dict[str, list]:
        """
        获取支持的转换格式
        
        Returns:
            支持的转换格式字典
        """
        return {
            'input_formats': [
                'doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx',
                'odt', 'ods', 'odp', 'rtf', 'html', 'txt'
            ],
            'output_formats': ['jpg', 'png', 'gif']
        }
    
    def is_supported_conversion(self, source_format: str, target_format: str) -> bool:
        """
        检查是否支持指定的转换
        
        Args:
            source_format: 源格式
            target_format: 目标格式
            
        Returns:
            是否支持
        """
        supported_formats = self.get_supported_formats()
        
        return (
            source_format in supported_formats['input_formats'] and
            target_format in supported_formats['output_formats']
        ) 