from dataclasses import dataclass, field
from typing import Optional, Tuple

@dataclass
class VisioShape:
    """Visio图形基类（矩形），使用@dataclass自动生成初始化方法"""
    page: object  # Visio页面对象
    shape_id: str
    x: float
    y: float
    width: float
    height: float
    text: Optional[str] = None
    bg_rgb: Optional[Tuple[int, int, int]] = None  # 填充颜色(RGB元组)
    line_color: Optional[Tuple[int, int, int]] = None  # 线条颜色(RGB)
    line_pattern:  int = 1  #  线条模式（0=无，1=实线，2=虚线，3=长虚线，4=点线，5=点划线，6=双点线）
    line_weight: float = 0.75  # 线宽（单位：pt）
    fill_pattern: int = 1  # 填充模式（1=实心）
    shape: object = field(init=False, repr=False)  # 不包含在__init__和__repr__中

    def __post_init__(self):
        """在自动生成的__init__后执行形状创建和样式设置"""
        self._create_shape()

    def _create_shape(self):
        """创建矩形并设置样式（与原方法一致）"""
        self.shape = self.page.DrawRectangle(
            self.x, self.y,
            self.x + self.width,
            self.y + self.height
        )
        if self.text:
            self.shape.Text = self.text
        if self.bg_rgb:
            self.shape.CellsU("FillForegnd").FormulaU = f"RGB({self.bg_rgb[0]},{self.bg_rgb[1]},{self.bg_rgb[2]})"
        if self.line_color:
            self.shape.CellsU("LineColor").FormulaU = f"RGB({self.line_color[0]},{self.line_color[1]},{self.line_color[2]})"

        self.shape.CellsU("LinePattern").FormulaU = str(self.line_pattern)
        self.shape.CellsU("LineWeight").FormulaU = f"{self.line_weight} pt"
        self.shape.CellsU("FillPattern").FormulaU = str(self.fill_pattern)

    @property
    def center_x(self) -> float:
        return self.shape.Cells("PinX").Result("")

    @property
    def center_y(self) -> float:
        return self.shape.Cells("PinY").Result("")


if __name__ == "__main__":
    from visio_diagram import VisioDiagram

    diagram = VisioDiagram()
    # 创建实例（无需手动写__init__）
    shape = VisioShape(
        page=diagram.page,
        shape_id="rect1",
        x=2.0, y=3.0,
        width=4.0, height=2.0,
        text="示例",
        line_weight=2.0,
        bg_rgb=(255, 0, 0),
        line_color=(0, 0, 0),
        fill_pattern=0
    )
    diagram.add_shape(shape)
    diagram.save_and_close("E:\Code\Python\DrawVisio\diagram.vsdx")
