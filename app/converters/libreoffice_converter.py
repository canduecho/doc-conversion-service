"""
基于 LibreOffice 的文档转换器
支持高性能的文档格式转换
"""
import asyncio
import subprocess
import tempfile
import shutil
from pathlib import Path
from typing import Dict, Any, Optional, List
from concurrent.futures import ProcessPoolExecutor, as_completed
from loguru import logger
import os

from app.config import TEMP_DIR


class LibreOfficeConverter:
    """LibreOffice 文档转换器"""
    
    def __init__(self, max_workers: int = 3):
        """
        初始化转换器
        
        Args:
            max_workers: 最大并发转换进程数
        """
        self.max_workers = max_workers
        # 暂时不使用 ProcessPoolExecutor，避免序列化问题
        # self.executor = ProcessPoolExecutor(max_workers=max_workers)
        
        # 支持的格式映射
        self.format_mapping = {
            # 输入格式
            'doc': 'MS Word 97',
            'docx': 'MS Word 2007 XML',
            'xls': 'MS Excel 97',
            'xlsx': 'MS Excel 2007 XML',
            'ppt': 'MS PowerPoint 97',
            'pptx': 'MS PowerPoint 2007 XML',
            'odt': 'writer8',
            'ods': 'calc8',
            'odp': 'impress8',
            'rtf': 'Rich Text Format',
            'txt': 'Text',
            'html': 'HTML',
            'pdf': 'writer_pdf_Export',
            
            # 输出格式
            'pdf': 'writer_pdf_Export',
            'docx': 'MS Word 2007 XML',
            'xlsx': 'MS Excel 2007 XML',
            'pptx': 'MS PowerPoint 2007 XML',
            'odt': 'writer8',
            'ods': 'calc8',
            'odp': 'impress8',
            'rtf': 'Rich Text Format',
            'html': 'HTML',
            'txt': 'Text'
        }
        
        # 检查 LibreOffice 是否可用
        self._check_libreoffice()
    
    def _check_libreoffice(self) -> bool:
        """检查 LibreOffice 是否可用"""
        try:
            result = subprocess.run(
                ['libreoffice', '--version'],
                capture_output=True,
                text=True,
                timeout=10
            )
            if result.returncode == 0:
                logger.info(f"LibreOffice 可用: {result.stdout.strip()}")
                return True
            else:
                logger.error(f"LibreOffice 检查失败: {result.stderr}")
                return False
        except (subprocess.TimeoutExpired, FileNotFoundError) as e:
            logger.error(f"LibreOffice 未安装或不可用: {e}")
            return False
    
    async def convert_document(
        self,
        input_path: str,
        output_path: str,
        target_format: str,
        options: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """
        转换文档格式
        
        Args:
            input_path: 输入文件路径
            output_path: 输出文件路径
            target_format: 目标格式
            options: 转换选项
            
        Returns:
            转换结果
        """
        try:
            # 验证输入文件
            if not os.path.exists(input_path):
                return {
                    'success': False,
                    'error': f'输入文件不存在: {input_path}'
                }
            
            # 获取文件扩展名
            input_ext = Path(input_path).suffix.lower().lstrip('.')
            output_ext = target_format.lower()
            
            # 验证格式支持
            if not self._is_format_supported(input_ext, output_ext):
                return {
                    'success': False,
                    'error': f'不支持的转换: {input_ext} → {output_ext}'
                }
            
            # 创建临时目录
            temp_dir = tempfile.mkdtemp(dir=TEMP_DIR)
            
            try:
                # 构建转换命令
                cmd = self._build_conversion_command(
                    input_path, temp_dir, output_ext, options
                )
                
                # 执行转换
                result = await self._execute_conversion(cmd, temp_dir)
                
                if result['success']:
                    # 查找输出文件
                    output_file = self._find_output_file(temp_dir, output_ext)
                    
                    if output_file and output_file.exists():
                        # 移动到目标位置
                        shutil.move(str(output_file), output_path)
                        
                        return {
                            'success': True,
                            'output_path': output_path,
                            'output_filename': Path(output_path).name
                        }
                    else:
                        return {
                            'success': False,
                            'error': f'未找到输出文件: {output_ext}'
                        }
                else:
                    return result
                    
            finally:
                # 清理临时目录
                if os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
                    
        except Exception as e:
            logger.error(f"文档转换失败: {e}")
            return {
                'success': False,
                'error': f'文档转换失败: {str(e)}'
            }
    
    async def convert_to_pdf(
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
            转换结果
        """
        return await self.convert_document(input_path, output_path, 'pdf', options)
    
    def _is_format_supported(self, input_ext: str, output_ext: str) -> bool:
        """检查格式是否支持"""
        # 定义支持的转换
        supported_conversions = {
            # Office 文档转换
            'doc': ['pdf', 'docx', 'odt', 'rtf', 'html', 'txt'],
            'docx': ['pdf', 'doc', 'odt', 'rtf', 'html', 'txt'],
            'xls': ['pdf', 'xlsx', 'ods', 'html', 'txt'],
            'xlsx': ['pdf', 'xls', 'ods', 'html', 'txt'],
            'ppt': ['pdf', 'pptx', 'odp', 'html'],
            'pptx': ['pdf', 'ppt', 'odp', 'html'],
            
            # OpenDocument 格式
            'odt': ['pdf', 'docx', 'doc', 'rtf', 'html', 'txt'],
            'ods': ['pdf', 'xlsx', 'xls', 'html', 'txt'],
            'odp': ['pdf', 'pptx', 'ppt', 'html'],
            
            # 其他格式
            'rtf': ['pdf', 'docx', 'odt', 'html', 'txt'],
            'html': ['pdf', 'docx', 'odt', 'rtf'],
            'txt': ['pdf', 'docx', 'odt', 'rtf', 'html']
        }
        
        return (input_ext in supported_conversions and 
                output_ext in supported_conversions[input_ext])
    
    def _build_conversion_command(
        self,
        input_path: str,
        output_dir: str,
        output_format: str,
        options: Optional[Dict[str, Any]] = None
    ) -> List[str]:
        """构建转换命令"""
        cmd = [
            'libreoffice',
            '--headless',  # 无界面模式
            '--convert-to', output_format,
            '--outdir', output_dir
        ]
        
        # 添加页面范围（如果支持）
        if options and options.get('page_range'):
            page_range = options['page_range']
            if page_range:
                # 注意：LibreOffice 的 --pages 参数可能不是所有版本都支持
                # 暂时注释掉，避免错误
                # cmd.extend(['--pages', page_range])
                pass
        
        # 添加输入文件
        cmd.append(input_path)
        
        return cmd
    
    async def _execute_conversion(
        self,
        cmd: List[str],
        temp_dir: str
    ) -> Dict[str, Any]:
        """执行转换命令"""
        try:
            # 直接执行转换，不使用进程池
            result = await asyncio.get_event_loop().run_in_executor(
                None,
                self._run_conversion_process,
                cmd
            )
            
            return result
            
        except Exception as e:
            logger.error(f"转换执行失败: {e}")
            return {
                'success': False,
                'error': f'转换执行失败: {str(e)}'
            }
    
    def _run_conversion_process(self, cmd: List[str]) -> Dict[str, Any]:
        """在进程池中运行转换"""
        try:
            logger.info(f"执行转换命令: {' '.join(cmd)}")
            
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=300  # 5分钟超时
            )
            
            logger.info(f"转换命令返回码: {result.returncode}")
            logger.info(f"转换命令标准输出: {result.stdout}")
            logger.info(f"转换命令错误输出: {result.stderr}")
            
            if result.returncode == 0:
                return {
                    'success': True,
                    'stdout': result.stdout,
                    'stderr': result.stderr
                }
            else:
                return {
                    'success': False,
                    'error': f'转换失败: {result.stderr}',
                    'stdout': result.stdout,
                    'stderr': result.stderr
                }
                
        except subprocess.TimeoutExpired:
            logger.error("转换超时")
            return {
                'success': False,
                'error': '转换超时'
            }
        except Exception as e:
            logger.error(f"转换异常: {e}")
            return {
                'success': False,
                'error': f'转换异常: {str(e)}'
            }
    
    def _find_output_file(self, temp_dir: str, output_ext: str) -> Optional[Path]:
        """查找输出文件"""
        temp_path = Path(temp_dir)
        
        logger.info(f"在临时目录中查找输出文件: {temp_dir}")
        logger.info(f"查找扩展名: .{output_ext}")
        
        # 列出临时目录中的所有文件
        all_files = list(temp_path.iterdir())
        logger.info(f"临时目录中的文件: {[f.name for f in all_files]}")
        
        # 查找匹配的文件
        for file_path in all_files:
            if file_path.is_file():
                logger.info(f"检查文件: {file_path.name}, 扩展名: {file_path.suffix.lower()}")
                if file_path.suffix.lower() == f'.{output_ext}':
                    logger.info(f"找到匹配的输出文件: {file_path}")
                    return file_path
        
        logger.warning(f"未找到扩展名为 .{output_ext} 的输出文件")
        return None
    
    async def batch_convert(
        self,
        files: List[Dict[str, str]],
        options: Optional[Dict[str, Any]] = None
    ) -> List[Dict[str, Any]]:
        """
        批量转换文档
        
        Args:
            files: 文件列表，每个文件包含 input_path, output_path, target_format
            options: 转换选项
            
        Returns:
            转换结果列表
        """
        tasks = []
        
        # 创建转换任务
        for file_info in files:
            task = self.convert_document(
                file_info['input_path'],
                file_info['output_path'],
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
    
    def get_supported_formats(self) -> Dict[str, List[str]]:
        """获取支持的格式"""
        return {
            'input_formats': ['doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx', 
                            'odt', 'ods', 'odp', 'rtf', 'html', 'txt'],
            'output_formats': ['pdf', 'docx', 'xlsx', 'pptx', 'odt', 'ods', 
                             'odp', 'rtf', 'html', 'txt']
        }
    
    def cleanup(self):
        """清理资源"""
        if hasattr(self, 'executor') and self.executor:
            self.executor.shutdown(wait=True) 