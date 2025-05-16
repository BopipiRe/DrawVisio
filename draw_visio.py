import win32com.client as win32
import json


class VisioBasicDiagram:
    """使用基础矩形和自定义连线绘制Visio图表"""

    def __init__(self):
        self.visio = win32.Dispatch("Visio.Application")
        self.visio.Visible = 0  # 隐藏窗口
        self.doc = self.visio.Documents.Add("")
        self.page = self.visio.ActivePage
        self.shapes = {}  # 存储形状对象的字典

    def add_rectangle(self, shape_id, x, y, width, height, text=None):
        """添加基础矩形"""
        shape = self.page.DrawRectangle(x, y, x + width, y + height)
        if text:
            shape.Text = text
        self.shapes[shape_id] = shape
        return shape

    def add_connector(self, from_id, to_id, line_style="solid"):
        """添加连接线（支持虚实线）"""
        if from_id not in self.shapes or to_id not in self.shapes:
            raise ValueError("无效的形状ID")

        # 创建动态连接线
        connector = self.page.DrawLine(
            self.shapes[from_id].Cells("PinX").Result(""),
            self.shapes[from_id].Cells("PinY").Result(""),
            self.shapes[to_id].Cells("PinX").Result(""),
            self.shapes[to_id].Cells("PinY").Result("")
        )

        # 设置线型
        if line_style == "dashed":
            connector.Cells("LinePattern").Formula = "2"  # 虚线
        elif line_style == "dotted":
            connector.Cells("LinePattern").Formula = "3"  # 点线
        else:  # 默认实线
            connector.Cells("LinePattern").Formula = "1"

        return connector

    def save_and_close(self, output_path):
        """保存并退出"""
        self.doc.SaveAs(output_path)
        self.visio.Quit()


def create_diagram_from_json(json_file_path, output_path):
    """从JSON配置创建图表"""
    diagram = VisioBasicDiagram()

    with open(json_file_path, 'r', encoding='utf-8') as f:
        config = json.load(f)

    # 添加矩形
    for shape_config in config.get("shapes", []):
        diagram.add_rectangle(
            shape_id=shape_config["id"],
            x=shape_config.get("x", 0),
            y=shape_config.get("y", 0),
            width=shape_config.get("width", 1),
            height=shape_config.get("height", 1),
            text=shape_config.get("text")
        )

    # 添加连接线
    for connector_config in config.get("connectors", []):
        diagram.add_connector(
            from_id=connector_config["from"],
            to_id=connector_config["to"],
            line_style=connector_config.get("style", "solid")
        )

    diagram.save_and_close(output_path)


# 使用示例
create_diagram_from_json(
    json_file_path=r"config.json",
    output_path=r"E:\diagram.vsdx"
)