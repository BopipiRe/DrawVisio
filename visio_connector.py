class VisioConnector:
    """Visio连接线类（边）"""
    LINE_STYLES = {
        "solid": 1,
        "dashed": 2,
        "dotted": 3
    }

    def __init__(self, page, from_shape, to_shape, line_style="solid"):
        self.page = page
        self.from_shape = from_shape
        self.to_shape = to_shape
        self.line_style = line_style
        self.connector = None
        self._create_connector()

    def _create_connector(self):
        """创建连接线并设置样式"""
        self.connector = self.page.DrawLine(
            self.from_shape.center_x,
            self.from_shape.center_y,
            self.to_shape.center_x,
            self.to_shape.center_y
        )
        pattern = self.LINE_STYLES.get(self.line_style, 1)
        self.connector.Cells("LinePattern").Formula = str(pattern)
