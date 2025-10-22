# 导出新的 snake_case 函数名（推荐使用）
from .WordWriter import word_writer
from .WordWriter import merge_table_row

# 导出旧的驼峰命名（向后兼容，将在未来版本弃用）
from .WordWriter import WordWriter
from .WordWriter import MergeTableRow

# 版本信息
__version__ = '3.3.0'

# 定义公共 API
__all__ = [
    # 新命名（推荐）
    'word_writer',
    'merge_table_row',
    # 旧命名（兼容）
    'WordWriter',
    'MergeTableRow',
]