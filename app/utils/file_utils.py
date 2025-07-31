"""
文件工具模块
"""
import os
import shutil
import uuid
import magic
from pathlib import Path
from typing import Optional, Tuple, List
from datetime import datetime
import aiofiles
from loguru import logger

from app.config import TEMP_DIR, OUTPUT_DIR, MAX_FILE_SIZE, ALLOWED_EXTENSIONS


class FileUtils:
    """文件工具类"""
    
    @staticmethod
    def generate_unique_filename(original_name: str) -> str:
        """
        生成唯一的文件名
        
        Args:
            original_name: 原始文件名
            
        Returns:
            唯一文件名
        """
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        unique_id = str(uuid.uuid4())[:8]
        extension = Path(original_name).suffix
        original_name = original_name.replace(extension, '')
        return f"{original_name}_{timestamp}_{unique_id}{extension}"
    
    @staticmethod
    def get_file_extension(filename: str) -> str:
        """
        获取文件扩展名
        
        Args:
            filename: 文件名
            
        Returns:
            文件扩展名（小写）
        """
        return Path(filename).suffix.lower().lstrip('.')
    
    @staticmethod
    def is_allowed_extension(filename: str) -> bool:
        """
        检查文件扩展名是否允许
        
        Args:
            filename: 文件名
            
        Returns:
            是否允许
        """
        extension = FileUtils.get_file_extension(filename)
        all_extensions = []
        for ext_list in ALLOWED_EXTENSIONS.values():
            all_extensions.extend(ext_list)
        return extension in all_extensions
    
    @staticmethod
    def get_file_type(filename: str) -> Optional[str]:
        """
        根据文件名获取文件类型
        
        Args:
            filename: 文件名
            
        Returns:
            文件类型
        """
        extension = FileUtils.get_file_extension(filename)
        
        for file_type, extensions in ALLOWED_EXTENSIONS.items():
            if extension in extensions:
                return file_type
        return None
    
    @staticmethod
    def detect_mime_type(file_path: str) -> str:
        """
        检测文件的 MIME 类型
        
        Args:
            file_path: 文件路径
            
        Returns:
            MIME 类型
        """
        try:
            mime = magic.from_file(file_path, mime=True)
            return mime
        except Exception as e:
            logger.warning(f"无法检测文件 MIME 类型: {e}")
            return 'application/octet-stream'
    
    @staticmethod
    def get_file_size(file_path: str) -> int:
        """
        获取文件大小
        
        Args:
            file_path: 文件路径
            
        Returns:
            文件大小（字节）
        """
        try:
            return os.path.getsize(file_path)
        except OSError:
            return 0
    
    @staticmethod
    def validate_file_size(file_path: str) -> bool:
        """
        验证文件大小是否在允许范围内
        
        Args:
            file_path: 文件路径
            
        Returns:
            是否有效
        """
        file_size = FileUtils.get_file_size(file_path)
        return file_size <= MAX_FILE_SIZE
    
    @staticmethod
    def create_temp_file(original_name: str) -> Tuple[str, str]:
        """
        创建临时文件
        
        Args:
            original_name: 原始文件名
            
        Returns:
            (临时文件路径, 唯一文件名)
        """
        unique_filename = FileUtils.generate_unique_filename(original_name)
        temp_path = TEMP_DIR / unique_filename
        return str(temp_path), unique_filename
    
    @staticmethod
    def create_output_file(filename: str, target_format: str) -> Tuple[str, str]:
        """
        创建输出文件路径
        
        Args:
            filename: 文件名
            target_format: 目标格式
            
        Returns:
            (输出文件路径, 输出文件名)
        """
        # 移除原始扩展名，添加目标格式扩展名
        name_without_ext = Path(filename).stem
        output_filename = f"{name_without_ext}.{target_format}"
        output_path = OUTPUT_DIR / output_filename
        return str(output_path), output_filename
    
    @staticmethod
    async def save_uploaded_file(uploaded_file, temp_path: str) -> bool:
        """
        保存上传的文件到临时目录
        
        Args:
            uploaded_file: 上传的文件对象
            temp_path: 临时文件路径
            
        Returns:
            是否保存成功
        """
        try:
            async with aiofiles.open(temp_path, 'wb') as f:
                content = await uploaded_file.read()
                await f.write(content)
            return True
        except Exception as e:
            logger.error(f"保存上传文件失败: {e}")
            return False
    
    @staticmethod
    def cleanup_temp_file(file_path: str) -> bool:
        """
        清理临时文件
        
        Args:
            file_path: 文件路径
            
        Returns:
            是否清理成功
        """
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
                logger.info(f"临时文件已清理: {file_path}")
                return True
        except Exception as e:
            logger.error(f"清理临时文件失败: {e}")
        return False
    
    @staticmethod
    def cleanup_old_files(directory: Path, max_age_hours: int) -> int:
        """
        清理指定目录中的旧文件
        
        Args:
            directory: 目录路径
            max_age_hours: 最大保留时间（小时）
            
        Returns:
            清理的文件数量
        """
        if not directory.exists():
            return 0
        
        current_time = datetime.now()
        cleaned_count = 0
        
        try:
            for file_path in directory.iterdir():
                if file_path.is_file():
                    file_age = current_time - datetime.fromtimestamp(file_path.stat().st_mtime)
                    if file_age.total_seconds() > max_age_hours * 3600:
                        file_path.unlink()
                        cleaned_count += 1
                        logger.info(f"清理旧文件: {file_path}")
        except Exception as e:
            logger.error(f"清理旧文件时出错: {e}")
        
        return cleaned_count
    
    @staticmethod
    def ensure_directory_exists(directory: Path) -> bool:
        """
        确保目录存在
        
        Args:
            directory: 目录路径
            
        Returns:
            是否成功
        """
        try:
            directory.mkdir(parents=True, exist_ok=True)
            return True
        except Exception as e:
            logger.error(f"创建目录失败: {e}")
            return False
    
    @staticmethod
    def get_file_info(file_path: str, original_name: str = None) -> dict:
        """
        获取文件信息
        
        Args:
            file_path: 文件路径
            original_name: 原始文件名
            
        Returns:
            文件信息字典
        """
        try:
            stat = os.stat(file_path)
            filename = Path(file_path).name
            extension = FileUtils.get_file_extension(filename)
            mime_type = FileUtils.detect_mime_type(file_path)
            
            return {
                'filename': filename,
                'original_name': original_name or filename,
                'file_size': stat.st_size,
                'content_type': mime_type,
                'extension': extension,
                'upload_time': datetime.fromtimestamp(stat.st_mtime).isoformat(),
                'file_type': FileUtils.get_file_type(filename)
            }
        except Exception as e:
            logger.error(f"获取文件信息失败: {e}")
            return {}
    
    @staticmethod
    def copy_file(src_path: str, dst_path: str) -> bool:
        """
        复制文件
        
        Args:
            src_path: 源文件路径
            dst_path: 目标文件路径
            
        Returns:
            是否复制成功
        """
        try:
            shutil.copy2(src_path, dst_path)
            return True
        except Exception as e:
            logger.error(f"复制文件失败: {e}")
            return False
    
    @staticmethod
    def move_file(src_path: str, dst_path: str) -> bool:
        """
        移动文件
        
        Args:
            src_path: 源文件路径
            dst_path: 目标文件路径
            
        Returns:
            是否移动成功
        """
        try:
            shutil.move(src_path, dst_path)
            return True
        except Exception as e:
            logger.error(f"移动文件失败: {e}")
            return False 