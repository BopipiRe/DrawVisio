import win32com.client as win32


class VisioDiagram:
    """Visio绘图文档管理类"""

    def __init__(self):
        self.visio = win32.Dispatch("Visio.Application")
        self.visio.Visible = 0
        self.doc = self.visio.Documents.Add("")
        self.page = self.visio.ActivePage
        self.shapes = {}  # 存储形状对象的字典
        self.connectors = []  # 存储连接线对象的列表

    def auto_fit_canvas(self, margin=0.5):
        """通过遍历形状计算边界并调整画布"""
        if len(self.page.Shapes) == 0:
            return

        # 初始化边界值
        min_x = float('inf')
        min_y = float('inf')
        max_x = -float('inf')
        max_y = -float('inf')

        # 遍历所有形状计算边界
        for shape in self.page.Shapes:
            x = shape.Cells("PinX").Result("")
            y = shape.Cells("PinY").Result("")
            width = shape.Cells("Width").Result("")
            height = shape.Cells("Height").Result("")

            min_x = min(min_x, x - width / 2)
            min_y = min(min_y, y - height / 2)
            max_x = max(max_x, x + width / 2)
            max_y = max(max_y, y + height / 2)

        # 设置画布尺寸（增加边距）
        required_width = (max_x - min_x) + 2 * margin
        required_height = (max_y - min_y) + 2 * margin

        # 使用绝对数值而非公式
        self.page.PageSheet.Cells("PageWidth").Formula = f"{required_width} in"
        self.page.PageSheet.Cells("PageHeight").Formula = f"{required_height} in"

    def add_shape(self, shape):
        """统一添加形状的方法，支持两种调用方式：
        1. 传入已创建的VisioShape对象
        2. 传入参数动态创建形状（需提供x/y/width/height）
        """
        self.shapes[shape.shape_id] = shape
        return shape

    def add_connector(self, connector):
        """添加连接线"""
        self.connectors.append(connector)
        return connector

    def save_and_close(self, output_path):
        """保存并退出Visio"""
        self.doc.SaveAs(output_path)
        self.visio.Quit()
