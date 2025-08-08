"""
跨类型转换器
使用 PDF 作为中间格式实现跨文档类型转换
"""
import asyncio
import tempfile
import shutil
from pathlib import Path
from typing import Dict, Any, Optional, List
from loguru import logger

from app.config import TEMP_DIR
from app.converters.libreoffice_converter import LibreOfficeConverter
from app.converters.pdf_converter import PDFConverter


class CrossTypeConverter:
    """跨类型转换器 - 使用 PDF 作为中间格式"""
    
    def __init__(self):
        """初始化转换器"""
        self.libreoffice_converter = LibreOfficeConverter()
        self.pdf_converter = PDFConverter()
        
        # 定义跨类型转换映射
        self.cross_type_conversions = {
            # Word 文档跨类型转换
            'docx': {
                'xlsx': self._convert_docx_to_xlsx,
                'pptx': self._convert_docx_to_pptx
            },
            # Excel 表格跨类型转换
            'xlsx': {
                'docx': self._convert_xlsx_to_docx,
                'pptx': self._convert_xlsx_to_pptx
            },
            # PowerPoint 演示文稿跨类型转换
            'pptx': {
                'docx': self._convert_pptx_to_docx,
                'xlsx': self._convert_pptx_to_xlsx
            }
        }
    
    async def convert_cross_type(
        self,
        input_path: str,
        output_path: str,
        target_format: str,
        options: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """
        执行跨类型转换
        
        Args:
            input_path: 输入文件路径
            output_path: 输出文件路径
            target_format: 目标格式
            options: 转换选项
            
        Returns:
            转换结果
        """
        try:
            # 获取输入文件扩展名
            input_ext = Path(input_path).suffix.lower().lstrip('.')
            
            # 检查是否支持跨类型转换
            if (input_ext in self.cross_type_conversions and 
                target_format in self.cross_type_conversions[input_ext]):
                
                # 执行跨类型转换
                converter_func = self.cross_type_conversions[input_ext][target_format]
                return await converter_func(input_path, output_path, options)
            else:
                return {
                    'success': False,
                    'error': f'不支持的跨类型转换: {input_ext} → {target_format}'
                }
                
        except Exception as e:
            logger.error(f"跨类型转换失败: {e}")
            return {
                'success': False,
                'error': f'跨类型转换失败: {str(e)}'
            }
    
    async def _convert_docx_to_xlsx(
        self,
        input_path: str,
        output_path: str,
        options: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """docx → xlsx 转换"""
        try:
            # 创建临时目录
            temp_dir = tempfile.mkdtemp(dir=TEMP_DIR)
            
            try:
                # 步骤1: docx → pdf
                pdf_temp_path = Path(temp_dir) / f"{Path(input_path).stem}_temp.pdf"
                pdf_result = await self.libreoffice_converter.convert_document(
                    input_path, str(pdf_temp_path), 'pdf', options
                )
                
                if not pdf_result['success']:
                    return pdf_result
                
                # 步骤2: pdf → xlsx
                xlsx_result = await self.pdf_converter.pdf_to_office(
                    str(pdf_temp_path), output_path, 'xlsx', options
                )
                
                if xlsx_result['success']:
                    return {
                        'success': True,
                        'output_path': output_path,
                        'output_filename': Path(output_path).name,
                        'conversion_type': 'cross_type',
                        'intermediate_format': 'pdf'
                    }
                else:
                    return xlsx_result
                    
            finally:
                # 清理临时文件
                if Path(temp_dir).exists():
                    shutil.rmtree(temp_dir)
                    
        except Exception as e:
            logger.error(f"docx → xlsx 转换失败: {e}")
            return {
                'success': False,
                'error': f'docx → xlsx 转换失败: {str(e)}'
            }
    
    async def _convert_docx_to_pptx(
        self,
        input_path: str,
        output_path: str,
        options: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """docx → pptx 转换"""
        try:
            # 创建临时目录
            temp_dir = tempfile.mkdtemp(dir=TEMP_DIR)
            
            try:
                # 步骤1: docx → pdf
                pdf_temp_path = Path(temp_dir) / f"{Path(input_path).stem}_temp.pdf"
                pdf_result = await self.libreoffice_converter.convert_document(
                    input_path, str(pdf_temp_path), 'pdf', options
                )
                
                if not pdf_result['success']:
                    return pdf_result
                
                # 步骤2: pdf → pptx
                pptx_result = await self.pdf_converter.pdf_to_office(
                    str(pdf_temp_path), output_path, 'pptx', options
                )
                
                if pptx_result['success']:
                    return {
                        'success': True,
                        'output_path': output_path,
                        'output_filename': Path(output_path).name,
                        'conversion_type': 'cross_type',
                        'intermediate_format': 'pdf'
                    }
                else:
                    return pptx_result
                    
            finally:
                # 清理临时文件
                if Path(temp_dir).exists():
                    shutil.rmtree(temp_dir)
                    
        except Exception as e:
            logger.error(f"docx → pptx 转换失败: {e}")
            return {
                'success': False,
                'error': f'docx → pptx 转换失败: {str(e)}'
            }
    
    async def _convert_xlsx_to_docx(
        self,
        input_path: str,
        output_path: str,
        options: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """xlsx → docx 转换"""
        try:
            # 创建临时目录
            temp_dir = tempfile.mkdtemp(dir=TEMP_DIR)
            
            try:
                # 步骤1: xlsx → pdf
                pdf_temp_path = Path(temp_dir) / f"{Path(input_path).stem}_temp.pdf"
                pdf_result = await self.libreoffice_converter.convert_document(
                    input_path, str(pdf_temp_path), 'pdf', options
                )
                
                if not pdf_result['success']:
                    return pdf_result
                
                # 步骤2: pdf → docx
                docx_result = await self.pdf_converter.pdf_to_office(
                    str(pdf_temp_path), output_path, 'docx', options
                )
                
                if docx_result['success']:
                    return {
                        'success': True,
                        'output_path': output_path,
                        'output_filename': Path(output_path).name,
                        'conversion_type': 'cross_type',
                        'intermediate_format': 'pdf'
                    }
                else:
                    return docx_result
                    
            finally:
                # 清理临时文件
                if Path(temp_dir).exists():
                    shutil.rmtree(temp_dir)
                    
        except Exception as e:
            logger.error(f"xlsx → docx 转换失败: {e}")
            return {
                'success': False,
                'error': f'xlsx → docx 转换失败: {str(e)}'
            }
    
    async def _convert_xlsx_to_pptx(
        self,
        input_path: str,
        output_path: str,
        options: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """xlsx → pptx 转换"""
        try:
            # 创建临时目录
            temp_dir = tempfile.mkdtemp(dir=TEMP_DIR)
            
            try:
                # 步骤1: xlsx → pdf
                pdf_temp_path = Path(temp_dir) / f"{Path(input_path).stem}_temp.pdf"
                pdf_result = await self.libreoffice_converter.convert_document(
                    input_path, str(pdf_temp_path), 'pdf', options
                )
                
                if not pdf_result['success']:
                    return pdf_result
                
                # 步骤2: pdf → pptx
                pptx_result = await self.pdf_converter.pdf_to_office(
                    str(pdf_temp_path), output_path, 'pptx', options
                )
                
                if pptx_result['success']:
                    return {
                        'success': True,
                        'output_path': output_path,
                        'output_filename': Path(output_path).name,
                        'conversion_type': 'cross_type',
                        'intermediate_format': 'pdf'
                    }
                else:
                    return pptx_result
                    
            finally:
                # 清理临时文件
                if Path(temp_dir).exists():
                    shutil.rmtree(temp_dir)
                    
        except Exception as e:
            logger.error(f"xlsx → pptx 转换失败: {e}")
            return {
                'success': False,
                'error': f'xlsx → pptx 转换失败: {str(e)}'
            }
    
    async def _convert_pptx_to_docx(
        self,
        input_path: str,
        output_path: str,
        options: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """pptx → docx 转换"""
        try:
            # 创建临时目录
            temp_dir = tempfile.mkdtemp(dir=TEMP_DIR)
            
            try:
                # 步骤1: pptx → pdf
                pdf_temp_path = Path(temp_dir) / f"{Path(input_path).stem}_temp.pdf"
                pdf_result = await self.libreoffice_converter.convert_document(
                    input_path, str(pdf_temp_path), 'pdf', options
                )
                
                if not pdf_result['success']:
                    return pdf_result
                
                # 步骤2: pdf → docx
                docx_result = await self.pdf_converter.pdf_to_office(
                    str(pdf_temp_path), output_path, 'docx', options
                )
                
                if docx_result['success']:
                    return {
                        'success': True,
                        'output_path': output_path,
                        'output_filename': Path(output_path).name,
                        'conversion_type': 'cross_type',
                        'intermediate_format': 'pdf'
                    }
                else:
                    return docx_result
                    
            finally:
                # 清理临时文件
                if Path(temp_dir).exists():
                    shutil.rmtree(temp_dir)
                    
        except Exception as e:
            logger.error(f"pptx → docx 转换失败: {e}")
            return {
                'success': False,
                'error': f'pptx → docx 转换失败: {str(e)}'
            }
    
    async def _convert_pptx_to_xlsx(
        self,
        input_path: str,
        output_path: str,
        options: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """pptx → xlsx 转换"""
        try:
            # 创建临时目录
            temp_dir = tempfile.mkdtemp(dir=TEMP_DIR)
            
            try:
                # 步骤1: pptx → pdf
                pdf_temp_path = Path(temp_dir) / f"{Path(input_path).stem}_temp.pdf"
                pdf_result = await self.libreoffice_converter.convert_document(
                    input_path, str(pdf_temp_path), 'pdf', options
                )
                
                if not pdf_result['success']:
                    return pdf_result
                
                # 步骤2: pdf → xlsx
                xlsx_result = await self.pdf_converter.pdf_to_office(
                    str(pdf_temp_path), output_path, 'xlsx', options
                )
                
                if xlsx_result['success']:
                    return {
                        'success': True,
                        'output_path': output_path,
                        'output_filename': Path(output_path).name,
                        'conversion_type': 'cross_type',
                        'intermediate_format': 'pdf'
                    }
                else:
                    return xlsx_result
                    
            finally:
                # 清理临时文件
                if Path(temp_dir).exists():
                    shutil.rmtree(temp_dir)
                    
        except Exception as e:
            logger.error(f"pptx → xlsx 转换失败: {e}")
            return {
                'success': False,
                'error': f'pptx → xlsx 转换失败: {str(e)}'
            }
    
    def get_supported_cross_type_conversions(self) -> Dict[str, List[str]]:
        """获取支持的跨类型转换"""
        return {
            'docx': ['xlsx', 'pptx'],
            'xlsx': ['docx', 'pptx'],
            'pptx': ['docx', 'xlsx']
        } 