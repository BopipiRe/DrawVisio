from dataclasses import dataclass, field
from typing import Optional

from utils import UnitParser


@dataclass
class TextStyle:
    color: Optional[str] = None
    fontFamily: Optional[str] = None
    overflowMode: Optional[str] = None
    fontSize: Optional[float] = None
    lineHeight: Optional[float] = None
    fontWeight: Optional[str] = None
    transform: Optional[str] = None


@dataclass
class ShapeProperties:
    id: Optional[str] = None
    name: Optional[str] = None
    nodeType: Optional[str] = None
    bindFlow: Optional[str] = None
    paramType: Optional[str] = None
    param: Optional[str] = None
    type: Optional[str] = None
    x: Optional[str] = None
    y: Optional[str] = None
    borderWidth: Optional[float] = None
    stroke: Optional[str] = None
    isDashBorder: Optional[bool] = None
    fill: Optional[str] = None
    width: Optional[float] = None
    height: Optional[float] = None
    zIndex: Optional[float] = None
    rotate: Optional[str] = None
    textStyle: Optional[TextStyle] = None
    borderColor: Optional[str] = None


@dataclass
class VisioShape:
    page: object
    pageHeight: float
    id: str
    x: float
    y: float
    width: float
    height: float
    zIndex: float
    rotate: Optional[str] = None
    type: Optional[str] = None
    properties: Optional[ShapeProperties] = None
    shape: object = field(init=False, repr=False)

    def __post_init__(self):
        """初始化后自动解析单位并创建形状"""
        if isinstance(self.properties, dict):
            self.properties = ShapeProperties(**self.properties)
        if isinstance(self.properties.textStyle, dict):
            self.properties.textStyle = TextStyle(**self.properties.textStyle)
        self._parse_units()
        self._create_shape()

    def _parse_units(self):
        """解析所有带单位的参数"""
        # 坐标和尺寸转换（支持px → in）
        self.y = self.pageHeight - self.y - self.height
        self.x = UnitParser.px_to_in(self.x)
        self.y = UnitParser.px_to_in(self.y)
        self.width = UnitParser.px_to_in(self.width)
        self.height = UnitParser.px_to_in(self.height)

        # 线宽转换（px → pt）
        self.properties.borderWidth = UnitParser.px_to_pt(self.properties.borderWidth)

    def _create_shape(self):
        """创建矩形（使用解析后的英寸单位）"""
        self.shape = self.page.DrawRectangle(
            self.x, self.y,
            self.x + self.width,
            self.y + self.height
        )
        self._set_style()

    def _set_style(self):
        """设置样式（使用解析后的磅单位）"""
        if self.properties.name:
            self.shape.Text = self.properties.name

        if self.properties.fill.find('-') != -1:
            # 渐变色填充
            gradient_colors = self.properties.fill.split('-')

            try:
                # 启用渐变并设置类型（线性渐变）
                self.shape.Cells("FillGradientEnabled").Formula = "1"
                self.shape.Cells("FillGradientDir").Formula = "0"  # 0=水平，1=径向

                # 清除旧的渐变停止点
                self.shape.Cells("FillGradientStops").Formula = "0"

                # 添加新的渐变停止点
                for i, color in enumerate(gradient_colors):
                    stop_pos = i / (len(gradient_colors) - 1)  # 计算位置（0~1）
                    rgb = UnitParser.hex_to_rgb(color)
                    self.shape.Cells("FillGradientStops").Formula = (
                        f"SETATREFEX(FillGradientStops, {stop_pos}, RGB({rgb}))"
                    )
            except:
                self.shape.Cells("FillGradientEnabled").Formula = "0"
                self.shape.Cells("FillForegnd").Formula = f"RGB({UnitParser.hex_to_rgb(gradient_colors[0])})"
                self.shape.Cells("FillBkgnd").Formula = f"RGB({UnitParser.hex_to_rgb(gradient_colors[1])})"
        else:
            # 纯色填充
            self.shape.Cells("FillForegnd").Formula = f"RGB({UnitParser.hex_to_rgb(self.properties.fill)})"

        if self.properties.borderWidth:
            self.shape.Cells("LineWeight").Formula = f"{self.properties.borderWidth} pt"
        if self.properties.stroke:
            self.shape.Cells("LineColor").Formula = f"RGB({UnitParser.hex_to_rgb(self.properties.stroke)})"
        if self.properties.isDashBorder:
            self.shape.Cells("LinePattern").Formula = "2"  # 虚线
        else:
            self.shape.Cells("LinePattern").Formula = "1"  # 实线
