class VisioShape:
    """Visio图形基类（矩形）"""

    def __init__(self, page, shape_id, x, y, width, height, text=None):
        self.page = page
        self.shape_id = shape_id
        self.x = x
        self.y = y
        self.width = width
        self.height = height
        self.text = text
        self.shape = None
        self._create_shape()

    def _create_shape(self):
        """创建矩形形状"""
        self.shape = self.page.DrawRectangle(
            self.x, self.y,
            self.x + self.width,
            self.y + self.height
        )
        if self.text:
            self.shape.Text = self.text

    @property
    def center_x(self):
        """中心点X坐标"""
        return self.shape.Cells("PinX").Result("")

    @property
    def center_y(self):
        """中心点Y坐标"""
        return self.shape.Cells("PinY").Result("")
