from dataclasses import dataclass, field
from typing import List, Optional, Dict, Any
import win32com.client
from utils import UnitParser


@dataclass
class VisioConnector:
    page: object
    pageHeight: float
    config: Dict[str, Any]
    connector: object = field(init=False, repr=False)

    def __post_init__(self):
        self.draw_connector()

    def draw_connector(self):
        """绘制连接线（带所有可选样式）"""

        # 将点列表转换为一维坐标数组 [x1, y1, x2, y2, ...]
        visio_points = [coord for point in self.config['pointsList']
                                for coord in (UnitParser.px_to_in(point['x']),
                                      UnitParser.px_to_in(self.pageHeight - point['y']))]

        # 使用 DrawPolyline 绘制折线
        connector = self.page.DrawPolyline(visio_points, 0)

        # 设置连接线样式
        self._apply_edge_style(connector, self.config['properties']['edgeStyle'])

        # 添加文字（如果有）
        if self.config['text'].get("value"):
            self._add_connector_text(connector, self.config['text'], self.config['properties']['textStyle'])

        return connector

    def _apply_edge_style(self, connector, style):
        """应用线条样式"""
        # 线条颜色和宽度
        connector.Cells("LineColor").Formula = f"RGB({UnitParser.hex_to_rgb(style['stroke'])})"
        connector.Cells("LineWeight").Formula = f"{style['strokeWidth']} pt"

        # 虚线样式
        if style['strokeDasharray']:
            connector.Cells("LinePattern").Formula = "2"  # 虚线
        else:
            connector.Cells("LinePattern").Formula = "1"  # 实线

        # 箭头设置
        if "targetArrow" in style:
            connector.Cells("EndArrow").Formula = "1"  # 启用箭头
            connector.Cells("EndArrowSize").Formula = "2"  # 中等大小
            # connector.Cells("EndArrowColor").Formula = f"RGB({UnitParser.hex_to_rgb(style['targetArrow']['fill'])})"

    def _add_connector_text(self, connector, text_data, text_style):
        """添加连接线文字"""
        connector.Text = text_data["value"]

        # 设置文字样式
        # char = connector.Characters
        # char.Begin = 0
        # char.End = len(connector.Text)
        # char.Cells("Color").Formula = f"RGB({UnitParser.hex_to_rgb(text_style['color'])})"
        # char.Cells("Size").Formula = f"{text_style['fontSize']} pt"
        # char.Cells("Font").Formula = f'"{text_style.fontFamily}"'

        # 定位文字
        # connector.Cells("TxtPinX").Formula = f"={UnitParser.px_to_in(text_data['x'])}"
        # connector.Cells("TxtPinY").Formula = f"={UnitParser.px_to_in(self.pageHeight - text_data['y'])}"
        # connector.Cells("TxtLocPinX").Formula = "Width*0"
        # connector.Cells("TxtLocPinY").Formula = "Height*0"
