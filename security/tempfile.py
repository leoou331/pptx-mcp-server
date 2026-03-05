"""
临时文件管理模块
自动清理临时文件
"""
import os
import tempfile
import shutil
import threading
import atexit
import logging
from typing import Set

log = logging.getLogger("pptx-server")


class TempFileManager:
    """临时文件管理器（线程安全单例）"""
    
    _instance = None
    _lock = threading.Lock()
    
    def __new__(cls):
        if cls._instance is None:
            with cls._lock:
                if cls._instance is None:
                    cls._instance = super().__new__(cls)
                    cls._instance._initialized = False
        return cls._instance
    
    def __init__(self):
        if self._initialized:
            return
        
        self._files: Set[str] = set()
        self._file_lock = threading.Lock()
        self._temp_dir = tempfile.mkdtemp(prefix='pptx_mcp_')
        self._initialized = True
        
        # 注册退出清理
        atexit.register(self.cleanup)
        
        log.info(f"Temp directory created: {self._temp_dir}")
    
    @property
    def temp_dir(self) -> str:
        """临时目录路径"""
        return self._temp_dir
    
    def create(self, suffix: str = '.pptx', prefix: str = 'pptx_') -> str:
        """
        创建临时文件
        
        Args:
            suffix: 文件后缀
            prefix: 文件前缀
            
        Returns:
            临时文件路径
        """
        # 检查磁盘空间
        stat = shutil.disk_usage(self._temp_dir)
        free_mb = stat.free / 1024 / 1024
        if free_mb < 500:
            raise RuntimeError(f"磁盘空间不足: {free_mb:.0f}MB")
        
        fd, path = tempfile.mkstemp(
            suffix=suffix,
            prefix=prefix,
            dir=self._temp_dir
        )
        os.close(fd)
        
        with self._file_lock:
            self._files.add(path)
        
        return path
    
    def register(self, path: str):
        """
        注册外部临时文件
        
        Args:
            path: 文件路径
        """
        with self._file_lock:
            self._files.add(path)
    
    def release(self, path: str):
        """
        释放并删除临时文件
        
        Args:
            path: 文件路径
        """
        with self._file_lock:
            if path in self._files:
                try:
                    os.remove(path)
                except FileNotFoundError:
                    pass
                except Exception as e:
                    log.warning(f"Failed to delete temp file: {path} - {e}")
                self._files.discard(path)
    
    def cleanup(self):
        """清理所有临时文件"""
        with self._file_lock:
            failed = []
            for path in list(self._files):
                try:
                    os.remove(path)
                except FileNotFoundError:
                    pass
                except Exception as e:
                    log.warning(f"Failed to delete: {path} - {e}")
                    failed.append(path)
            
            self._files.clear()
            
            if failed:
                log.warning(f"{len(failed)} temp files failed to delete")
        
        # 清理临时目录
        try:
            shutil.rmtree(self._temp_dir, ignore_errors=True)
            log.info(f"Temp directory cleaned: {self._temp_dir}")
        except Exception as e:
            log.warning(f"Failed to clean temp dir: {e}")
    
    def get_stats(self) -> dict:
        """获取统计信息"""
        with self._file_lock:
            file_count = len(self._files)
            
            total_size = 0
            for path in self._files:
                try:
                    total_size += os.path.getsize(path)
                except:
                    pass
            
            stat = shutil.disk_usage(self._temp_dir)
            
            return {
                "temp_dir": self._temp_dir,
                "tracked_files": file_count,
                "total_size_mb": round(total_size / 1024 / 1024, 2),
                "disk_free_mb": round(stat.free / 1024 / 1024, 2)
            }


# 全局实例
temp_manager = TempFileManager()
