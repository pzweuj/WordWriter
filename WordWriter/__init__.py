# ============================================================================
# v4.0.0 新的面向对象 API（推荐使用）
# ============================================================================
from .core import WordWriter as WordWriterClass
from .core import TagSearcher, ContentReplacer

# ============================================================================
# 函数式 API（向后兼容）
# ============================================================================

# 新的 snake_case 函数名
from .WordWriter import word_writer
from .WordWriter import merge_table_row

# 旧的驼峰命名（兼容）
from .WordWriter import WordWriter as word_writer_func
from .WordWriter import MergeTableRow

# ============================================================================
# 版本信息
# ============================================================================
__version__ = '4.0.4'

# ============================================================================
# 公共 API
# ============================================================================
__all__ = [
    # v4.0 面向对象 API（推荐）
    'WordWriter',  # 包装器，同时支持 v3.x 和 v4.0
    'WordWriterClass',
    'TagSearcher',
    'ContentReplacer',
    
    # 函数式 API（向后兼容）
    'word_writer',
    'merge_table_row',
    'word_writer_func',  # 旧的 WordWriter 函数
    'MergeTableRow',
]

# ============================================================================
# 便捷导入：WordWriter 默认指向新的类（但支持向后兼容）
# ============================================================================
class _WordWriterWrapper:
    """WordWriter 包装器，同时支持 v3.x 函数式调用和 v4.0 类式调用
    
    这个包装器确保真正的向后兼容性：
    - v3.x 调用: WordWriter(input_docx, output_docx, replace_dict, logs)
    - v4.0 调用: WordWriter(template_path).replace(...).save(...)
    """
    
    def __new__(cls, *args, **kwargs):
        """根据参数数量决定是函数式调用还是类式调用"""
        # 判断是否为 v3.x 函数式调用
        # v3.x: WordWriter(input_docx, output_docx, replace_dict, logs=True)
        if len(args) >= 3 or 'output_docx' in kwargs or 'replace_dict' in kwargs:
            # 这是 v3.x 的函数式调用
            # 提取参数
            if len(args) >= 3:
                input_docx = args[0]
                output_docx = args[1]
                replace_dict = args[2]
                logs = args[3] if len(args) >= 4 else kwargs.get('logs', True)
            else:
                input_docx = args[0] if len(args) >= 1 else kwargs.get('input_docx')
                output_docx = kwargs.get('output_docx')
                replace_dict = kwargs.get('replace_dict')
                logs = kwargs.get('logs', True)
            
            # 调用旧的函数式 API
            word_writer_func(input_docx, output_docx, replace_dict, logs)
            return None
        else:
            # 这是 v4.0 的类式调用
            # 创建并返回 WordWriterClass 实例
            return WordWriterClass(*args, **kwargs)
    
    def __init__(self, *args, **kwargs):
        """初始化方法（仅用于类式调用）"""
        # 对于函数式调用，__new__ 返回 None，不会调用 __init__
        # 对于类式调用，__new__ 返回 WordWriterClass 实例，也不需要这个 __init__
        pass

# 使用包装器作为默认的 WordWriter
WordWriter = _WordWriterWrapper