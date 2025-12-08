"""
自动更新管理模块
负责检查更新、下载更新和安装更新
"""

import os
import sys
import json
import requests
import shutil
import zipfile
import subprocess
from datetime import datetime
import logging

# 配置日志
import os

# 获取程序所在目录
app_dir = os.path.dirname(os.path.abspath(__file__))
log_file = os.path.join(app_dir, 'update.log')

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file, encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class UpdateManager:
    def __init__(self, app_name: str, current_version: str, update_server_url: str):
        """
        初始化更新管理器
        
        :param app_name: 应用程序名称
        :param current_version: 当前版本号
        :param update_server_url: 更新服务器URL
        """
        self.app_name = app_name
        self.current_version = current_version
        self.update_server_url = update_server_url
        self.update_info = None
        
    def check_for_updates(self) -> dict:
        """
        检查是否有可用更新
        
        :return: 更新信息字典，如果没有更新返回空字典
        """
        try:
            logger.info(f"检查更新... 当前版本: {self.current_version}")
            
            # 从服务器获取最新版本信息
            response = requests.get(f"{self.update_server_url}/version.json", timeout=10)
            response.raise_for_status()
            
            self.update_info = response.json()
            latest_version = self.update_info.get("version", "")
            
            # 比较版本号
            if self._is_newer_version(latest_version, self.current_version):
                logger.info(f"发现新版本: {latest_version}")
                return self.update_info
            else:
                logger.info("当前已是最新版本")
                return {}
                
        except requests.RequestException as e:
            logger.error(f"检查更新失败: {e}")
            return {}
        except json.JSONDecodeError as e:
            logger.error(f"解析更新信息失败: {e}")
            return {}
    
    def download_update(self, download_path: str = None) -> str:
        """
        下载更新包
        
        :param download_path: 下载路径，默认使用临时目录
        :return: 下载的更新包路径，如果下载失败返回空字符串
        """
        if not self.update_info:
            logger.error("没有更新信息，无法下载")
            return ""
        
        try:
            update_url = self.update_info.get("download_url", "")
            if not update_url:
                logger.error("更新包下载URL为空")
                return ""
            
            # 设置下载路径
            if not download_path:
                download_path = os.path.join(os.getcwd(), "updates")
            
            if not os.path.exists(download_path):
                os.makedirs(download_path)
            
            # 下载文件名
            file_name = f"{self.app_name}_update_{self.update_info['version']}.zip"
            download_file_path = os.path.join(download_path, file_name)
            
            logger.info(f"开始下载更新包: {file_name}")
            logger.info(f"下载地址: {update_url}")
            
            # 下载更新包
            with requests.get(update_url, stream=True, timeout=30) as r:
                r.raise_for_status()
                total_size = int(r.headers.get('content-length', 0))
                downloaded_size = 0
                
                with open(download_file_path, 'wb') as f:
                    for chunk in r.iter_content(chunk_size=8192):
                        if chunk:
                            f.write(chunk)
                            downloaded_size += len(chunk)
                            # 记录下载进度
                            if total_size > 0:
                                progress = (downloaded_size / total_size) * 100
                                if progress % 10 == 0:  # 每10%记录一次
                                    logger.info(f"下载进度: {progress:.1f}%")
            
            logger.info(f"更新包下载完成: {download_file_path}")
            return download_file_path
            
        except requests.RequestException as e:
            logger.error(f"下载更新失败: {e}")
            return ""
        except Exception as e:
            logger.error(f"下载更新时发生错误: {e}")
            return ""
    
    def install_update(self, update_file_path: str) -> bool:
        """
        安装更新
        
        :param update_file_path: 更新包路径
        :return: 安装是否成功
        """
        try:
            if not os.path.exists(update_file_path):
                logger.error(f"更新包不存在: {update_file_path}")
                return False
            
            logger.info(f"开始安装更新: {update_file_path}")
            
            # 创建临时解压目录
            extract_dir = os.path.join(os.path.dirname(update_file_path), "temp_extract")
            if os.path.exists(extract_dir):
                shutil.rmtree(extract_dir)
            os.makedirs(extract_dir)
            
            # 解压更新包
            logger.info(f"解压更新包到: {extract_dir}")
            with zipfile.ZipFile(update_file_path, 'r') as zip_ref:
                zip_ref.extractall(extract_dir)
            
            # 获取当前程序目录
            current_dir = os.path.dirname(os.path.abspath(sys.executable))
            logger.info(f"当前程序目录: {current_dir}")
            
            # 创建备份目录
            backup_dir = os.path.join(current_dir, "backup")
            if not os.path.exists(backup_dir):
                os.makedirs(backup_dir)
            
            # 备份当前程序文件
            logger.info("备份当前程序文件...")
            for root, dirs, files in os.walk(current_dir):
                for file in files:
                    if file.endswith('.exe') or file.endswith('.py') or file.endswith('.json') or file.endswith('.yaml') or file.endswith('.html'):
                        src_file = os.path.join(root, file)
                        # 计算相对路径用于备份
                        rel_path = os.path.relpath(src_file, current_dir)
                        backup_file = os.path.join(backup_dir, rel_path)
                        
                        # 确保备份目录存在
                        backup_file_dir = os.path.dirname(backup_file)
                        if not os.path.exists(backup_file_dir):
                            os.makedirs(backup_file_dir)
                        
                        # 备份文件
                        shutil.copy2(src_file, backup_file)
            
            # 复制更新文件到当前目录
            logger.info("复制更新文件到当前目录...")
            for root, dirs, files in os.walk(extract_dir):
                for file in files:
                    src_file = os.path.join(root, file)
                    # 计算目标路径
                    rel_path = os.path.relpath(src_file, extract_dir)
                    dest_file = os.path.join(current_dir, rel_path)
                    
                    # 确保目标目录存在
                    dest_dir = os.path.dirname(dest_file)
                    if not os.path.exists(dest_dir):
                        os.makedirs(dest_dir)
                    
                    # 替换文件
                    shutil.copy2(src_file, dest_file)
            
            logger.info("更新安装完成")
            
            # 清理临时文件
            shutil.rmtree(extract_dir)
            
            return True
            
        except zipfile.BadZipFile as e:
            logger.error(f"更新包格式错误: {e}")
            return False
        except PermissionError as e:
            logger.error(f"权限错误，无法安装更新: {e}")
            return False
        except Exception as e:
            logger.error(f"安装更新时发生错误: {e}")
            return False
    
    def restart_application(self) -> bool:
        """
        重启应用程序
        
        :return: 重启是否成功
        """
        try:
            logger.info("重启应用程序...")
            
            # 获取当前可执行文件路径
            executable_path = sys.executable
            logger.info(f"可执行文件路径: {executable_path}")
            
            # 启动新实例
            subprocess.Popen([executable_path], close_fds=True)
            
            # 退出当前实例
            sys.exit(0)
            
        except Exception as e:
            logger.error(f"重启应用程序失败: {e}")
            return False
    
    def _is_newer_version(self, latest_version: str, current_version: str) -> bool:
        """
        比较版本号是否为新版本
        
        :param latest_version: 最新版本号
        :param current_version: 当前版本号
        :return: 如果latest_version是新版本返回True，否则返回False
        """
        try:
            # 将版本号拆分为整数列表
            latest_parts = list(map(int, latest_version.strip('v').split('.')))
            current_parts = list(map(int, current_version.strip('v').split('.')))
            
            # 确保版本号部分数量相同
            max_length = max(len(latest_parts), len(current_parts))
            latest_parts += [0] * (max_length - len(latest_parts))
            current_parts += [0] * (max_length - len(current_parts))
            
            # 逐部分比较
            for i in range(max_length):
                if latest_parts[i] > current_parts[i]:
                    return True
                elif latest_parts[i] < current_parts[i]:
                    return False
            
            # 版本号相同
            return False
            
        except ValueError:
            # 版本号格式错误，返回False
            logger.error(f"版本号格式错误: latest={latest_version}, current={current_version}")
            return False

# 测试代码
if __name__ == "__main__":
    # 测试更新检查
    update_manager = UpdateManager(
        app_name="AutoReport",
        current_version="1.0.0",
        update_server_url="https://example.com/updates"
    )
    
    # 模拟版本检查
    print("检查更新...")
    update_info = update_manager.check_for_updates()
    print(f"更新信息: {update_info}")
