import os
import logging
import time
import shutil
from datetime import datetime, timedelta

logger = logging.getLogger(__name__)


def cleanup_old_files(directory, max_age_hours=24):
    """
    清理指定目录中超过一定时间的文件和子目录

    Args:
        directory: 目录路径
        max_age_hours: 最大保存时间（小时）
    """
    if not os.path.exists(directory):
        logger.warning(f"目录不存在，无法清理: {directory}")
        return

    now = datetime.now()
    threshold = now - timedelta(hours=max_age_hours)

    files_count = 0
    dirs_count = 0
    try:
        # 先处理顶层文件
        for item in os.listdir(directory):
            item_path = os.path.join(directory, item)
            
            try:
                item_mod_time = datetime.fromtimestamp(os.path.getmtime(item_path))
                
                if item_mod_time < threshold:
                    if os.path.isfile(item_path):
                        os.remove(item_path)
                        files_count += 1
                    elif os.path.isdir(item_path):
                        shutil.rmtree(item_path)
                        dirs_count += 1
            except Exception as e:
                logger.error(f"清理 {item_path} 时出错: {str(e)}")

        if files_count > 0 or dirs_count > 0:
            logger.info(f"已清理{files_count}个文件和{dirs_count}个子目录，超过{max_age_hours}小时，目录: {directory}")
    except Exception as e:
        logger.error(f"清理目录时出错: {str(e)}")


def get_file_size_str(size_in_bytes):
    """获取友好的文件大小表示"""
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size_in_bytes < 1024:
            return f"{size_in_bytes:.2f} {unit}"
        size_in_bytes /= 1024
    return f"{size_in_bytes:.2f} TB"