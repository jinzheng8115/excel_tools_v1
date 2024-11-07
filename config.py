import os

# 基础路径配置
BASE_DIR = os.path.dirname(os.path.abspath(__file__))  # 获取项目根目录的绝对路径
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')      # 上传文件的临时存储目录
RESULT_FOLDER = os.path.join(BASE_DIR, 'results')      # 处理结果的存储目录
DOWNLOAD_FOLDER = os.path.join(BASE_DIR, 'downloads')  # 下载文件存储目录

# 确保必要的目录存在
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)
os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

# 文件配置
ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'txt', 'csv'}  # 添加 csv 格式
MAX_CONTENT_LENGTH = 16 * 1024 * 1024  # 最大文件大小限制：16MB

# 数据处理配置
MAX_ROWS = 100000        # 最大处理行数
CHUNK_SIZE = 10000       # 分块处理的大小
RESULT_EXPIRY = 3600    # 结果文件过期时间（秒）

# 错误消息配置
ERROR_MESSAGES = {
    'file_not_found': '未找到文件',
    'invalid_format': '不支持的文件格式',
    'file_too_large': '文件大小超过限制',
    'no_sheets': '文件中没有工作表',
    'process_failed': '处理失败，请检查数据格式'
}

# 开发环境配置
DEBUG = True
HOST = '0.0.0.0'
PORT = 5000