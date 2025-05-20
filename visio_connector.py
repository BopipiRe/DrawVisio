from dataclasses import dataclass, field
from typing import Optional, Tuple
import re

@dataclass
class VisioConnector:
    """Visio连接线类（支持坐标单位转换和样式控制）"""
    page: object
    from_point: str  # 起点坐标（如"100px,200px"或"2.5in,3cm"）
    to_point: str    # 终点坐标（格式同起点）
    line_style: str = "solid"  # 支持solid/dashed/dotted/dash_dot
    label: Optional[str] = None
    text_visible: bool = True
    connector: object = field(init=False, repr=False)
    arrow_type: str = "arrow"  # none/arrow/diamond/circle
    line_weight: str = "1pt"   # 线宽（支持px/pt单位）

    # 单位转换常量
    UNITS = {
        "px": lambda x: x / 96,    # 像素转英寸（96dpi）
        "in": lambda x: x,         # 英寸直接使用
        "cm": lambda x: x / 2.54,  # 厘米转英寸
        "mm": lambda x: x / 25.4,
        "pt": lambda x: x / 72      # 磅转英寸
    }

    def __post_init__(self):
        """解析坐标单位并创建连接线"""
        self.from_x, self.from_y = self._parse_point(self.from_point)
        self.to_x, self.to_y = self._parse_point(self.to_point)
        self._create_connector()
        self._apply_style()

    def _parse_point(self, point_str: str) -> Tuple[float, float]:
        """解析带单位的坐标字符串（如'100px,50mm'）"""
        parts = [p.strip() for p in point_str.split(",")]
        if len(parts) != 2:
            raise ValueError(f"坐标格式错误，应为'x单位,y单位'，如'100px,2.5in'")

        x = self._parse_length(parts[0])
        y = self._parse_length(parts[1])
        return x, y

    def _parse_length(self, value: str) -> float:
        """解析带单位的长度值（如'100px'）"""
        match = re.match(r"^([\d.]+)\s*(px|in|cm|mm|pt)?$", value, re.IGNORECASE)
        if not match:
            raise ValueError(f"无效的长度格式: {value}")

        num, unit = float(match.group(1)), (match.group(2) or "px").lower()
        return self.UNITS[unit](num)

    def _create_connector(self):
        """创建连接线（自动转换为Visio内部英寸单位）"""
        self.connector = self.page.DrawLine(
            self.from_x, self.from_y,
            self.to_x, self.to_y
        )

    def _apply_style(self):
        """应用样式（线条、箭头、文本）"""
        # 线条样式
        pattern = {
            "solid": 1,
            "dashed": 2,
            "dotted": 3,
            "dash_dot": 4
        }.get(self.line_style.lower(), 1)
        self.connector.CellsU("LinePattern").FormulaU = str(pattern)

        # 箭头样式
        arrow_code = {
            "none": 0,
            "arrow": 1,
            "diamond": 2,
            "circle": 3
        }.get(self.arrow_type.lower(), 1)
        self.connector.CellsU("EndArrow").FormulaU = str(arrow_code)

        # 线宽（支持px/pt单位）
        weight_pt = self._parse_length(self.line_weight) * 72  # 转换为磅
        self.connector.CellsU("LineWeight").FormulaU = f"{weight_pt} pt"

        # 文本标签
        if self.label:
            self.connector.Text = self.label
            transparency = 100 if not self.text_visible else 0
            self.connector.CellsU("Char.Transparency").FormulaU = str(transparency)

    def toggle_text_visibility(self):
        """切换文本显隐状态"""
        self.text_visible = not self.text_visible
        transparency = 100 if not self.text_visible else 0
        self.connector.CellsU("Char.Transparency").FormulaU = str(transparency)
