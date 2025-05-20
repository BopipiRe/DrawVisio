from dataclasses import dataclass, field
from typing import Optional


@dataclass
class VisioConnector:
    """Visio连接线类（支持样式控制和显示/隐藏）"""
    page: object
    from_shape: object
    to_shape: object
    line_style: str = "solid"  # 支持solid/dashed/dotted
    label: Optional[str] = None  # 连接线文本标签
    text_visible: bool = True  # 控制显示/隐藏
    connector: object = field(init=False, repr=False)  # 隐藏内部对象
    arrow_type: str = "arrow"  # 箭头类型（none/arrow/diamond等）

    # 样式常量（类属性）
    LINE_STYLES = {
        "solid": 1,
        "dashed": 2,
        "dotted": 3,
        "dash_dot": 4
    }
    ARROW_TYPES = {
        "none": 0,
        "arrow": 1,
        "diamond": 2,
        "circle": 3
    }

    def __post_init__(self):
        """自动创建连接线并应用样式"""
        self._create_connector()
        self._apply_style()
        self._set_text_visibility()

    def _create_connector(self):
        """创建连接线（基于形状中心点）"""
        self.connector = self.page.DrawLine(
            self.from_shape.center_x,
            self.from_shape.center_y,
            self.to_shape.center_x,
            self.to_shape.center_y
        )

    def _apply_style(self):
        """应用线条样式和箭头"""
        # 设置线条样式
        pattern = self.LINE_STYLES.get(self.line_style.lower(), 1)
        self.connector.CellsU("LinePattern").FormulaU = str(pattern)

        # 设置箭头
        arrow_code = self.ARROW_TYPES.get(self.arrow_type.lower(), 0)
        # self.connector.CellsU("BeginArrow").FormulaU = str(arrow_code)
        self.connector.CellsU("EndArrow").FormulaU = str(arrow_code)

        # 设置文本标签
        if self.label:
            self.connector.Text = self.label

    def _set_text_visibility(self):
        """控制文字显隐（通过Text字段和字符透明度实现）"""
        if not self.label:
            return

        transparency = 100 if not self.text_visible else 0
        self.connector.CellsU("Char.Transparency").FormulaU = str(transparency)

    def toggle_visibility(self):
        """切换显示状态"""
        self.is_visible = not self.text_visible
        self._set_text_visibility()

    @property
    def points(self) -> tuple:
        """获取连接线端点坐标（返回像素单位）"""
        return (
            (self.from_shape.center_x, self.from_shape.center_y),
            (self.to_shape.center_x, self.to_shape.center_y)
        )
