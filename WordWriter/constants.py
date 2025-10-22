# coding=utf-8
"""WordWriter 常量定义模块

此模块包含 WordWriter 中使用的所有常量，包括标签前缀、特殊值、
单位转换常量和默认样式等。

Author: pzweuj
Since: v3.3.0
"""


class TagPrefix:
    """标签前缀常量
    
    定义了 WordWriter 支持的各种标签类型的前缀标识。
    """
    # 标签标记
    TAG_START = "#["
    TAG_END = "]#"
    
    # 标签类型前缀
    TABLE = "#[TABLE"
    IMAGE = "#[IMAGE"
    TABLE_IMAGE = "#[TBIMG"
    TEXTBOX = "#[TX"


class SpecialValue:
    """特殊值常量
    
    定义了具有特殊含义的值，用于触发特定操作。
    """
    DELETE_PARAGRAPH = "#DELETETHISPARAGRAPH#"
    DELETE_TABLE = "#DELETETHISTABLE#"


class Conversion:
    """单位转换常量
    
    定义了各种单位转换系数。
    """
    # EMU (English Metric Unit) 转换
    EMU_PER_CM = 360000  # EMU 每厘米
    EMU_PER_INCH = 914400  # EMU 每英寸
    CM_TO_EMU = 360000  # 厘米转 EMU（向后兼容）


class DefaultBorder:
    """默认边框样式常量
    
    定义了表格和单元格的默认边框样式。
    """
    SIZE = '0'
    COLOR = 'auto'
    SPACE = '0'
    VAL = 'single'
    
    @classmethod
    def get_style_dict(cls):
        """获取默认边框样式字典
        
        Returns:
            dict: 包含边框样式的字典
        """
        return {
            'size': cls.SIZE,
            'color': cls.COLOR,
            'space': cls.SPACE,
            'val': cls.VAL
        }


class XMLNamespace:
    """XML 命名空间常量
    
    定义了 Office Open XML 中使用的命名空间前缀。
    """
    WORD = "w"
    WORD_MAIN = "main"
    
    # 常用的标签后缀
    TAG_RUN = "main}r"
    TAG_PARAGRAPH_PROPERTIES = "main}pPr"
    TAG_ALTERNATE_CONTENT = "AlternateContent"
    TAG_TEXTBOX = "textbox"


class LogMessage:
    """日志消息常量
    
    定义了日志输出中使用的消息模板。
    """
    MISSING_TAG = "【Missing Tag】 "
    FILLING_TAG = "【Filling Tag】 "
    ERROR_TAG = "【Error】 "
