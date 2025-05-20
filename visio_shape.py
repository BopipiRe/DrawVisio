from dataclasses import dataclass, field
from typing import Optional, Tuple, Union

import win32com.client as win32


@dataclass
class VisioShape:
    """Visio图形基类（支持像素单位、图片嵌入）"""
    page: object
    shape_id: str
    x: float  # X坐标（像素）
    y: float  # Y坐标（像素）
    width: float  # 宽度（像素）
    height: float  # 高度（像素）
    text: Optional[str] = None
    bg_rgb: Optional[Tuple[int, int, int]] = None
    line_color: Optional[Tuple[int, int, int]] = None
    line_pattern: int = 1  # 线条模式（1=实线）
    line_weight: float = 1.0  # 线宽（像素）
    fill_pattern: int = 1  # 填充模式
    image: Optional[Union[str, bytes]] = None  # 支持URL或Base64编码的图片
    shape: object = field(init=False, repr=False)

    def __post_init__(self):
        """初始化后自动创建形状并设置样式"""
        self._convert_px_to_visio_units()  # 像素转Visio单位
        self._create_shape()
        self._set_image() if self.image else None

    def _convert_px_to_visio_units(self):
        """将像素转换为Visio默认单位（英寸和磅）"""
        self.x = self._px_to_in(self.x)
        self.y = self._px_to_in(self.y)
        self.width = self._px_to_in(self.width)
        self.height = self._px_to_in(self.height)
        self.line_weight = self._px_to_pt(self.line_weight)

    @staticmethod
    def _px_to_in(px: float) -> float:
        """像素转英寸（假设96DPI）"""
        return px / 96  # 1英寸=96像素

    @staticmethod
    def _px_to_pt(px: float) -> float:
        """像素转磅（1磅=1/72英寸）"""
        return px * 72 / 96  # 1pt = (1/72)in = (1/72)*96px ≈ 1.33px

    def _create_shape(self):
        """创建矩形并设置基础样式"""
        self.shape = self.page.DrawRectangle(
            self.x, self.y,
            self.x + self.width,
            self.y + self.height
        )
        if self.text:
            self.shape.Text = self.text
        self._set_style()

    def _set_style(self):
        """设置线条和填充样式"""
        if self.bg_rgb:
            self.shape.CellsU("FillForegnd").FormulaU = f"RGB({self.bg_rgb[0]},{self.bg_rgb[1]},{self.bg_rgb[2]})"
        if self.line_color:
            self.shape.CellsU(
                "LineColor").FormulaU = f"RGB({self.line_color[0]},{self.line_color[1]},{self.line_color[2]})"
        self.shape.CellsU("LinePattern").FormulaU = str(self.line_pattern)
        self.shape.CellsU("LineWeight").FormulaU = f"{self.line_weight} pt"
        self.shape.CellsU("FillPattern").FormulaU = str(self.fill_pattern)

    def _set_image(self):
        """嵌入图片（支持URL、Base64或本地文件路径）"""
        try:
            if isinstance(self.image, str):
                if self.image.startswith(('http://', 'https://')):
                    self._insert_image_from_url(self.image)
                else:
                    # 尝试作为本地文件路径处理
                    with open(self.image, "rb") as image_file:
                        encoded_string = base64.b64encode(image_file.read()).decode()
                    self._insert_image_from_base64(encoded_string)
            elif isinstance(self.image, (str, bytes)):
                self._insert_image_from_base64(self.image)
        except Exception as e:
            print(f"插入{self.shape_id}的图片时发生错误: {e}")
    def _insert_image_from_url(self, url: str):
        """通过URL插入图片"""
        img_shape = self.page.Import(url)  # 插入图片
        img_shape.CellsU("Width").FormulaU = f"{self.width} in"
        img_shape.CellsU("Height").FormulaU = f"{self.height} in"
        img_shape.CellsU("PinX").FormulaU = f"{self.x + self.width / 2} in"
        img_shape.CellsU("PinY").FormulaU = f"{self.y + self.height / 2} in"

    def _insert_image_from_base64(self, data: Union[str, bytes]):
        """通过Base64插入图片（需临时文件）"""
        import tempfile
        if isinstance(data, str):
            data = base64.b64decode(data.encode())
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
            tmp.write(data)
            tmp.close()
            self._insert_image_from_url(tmp.name)

    @property
    def center_x(self) -> float:
        """中心X坐标（像素）"""
        return self._in_to_px(self.shape.Cells("PinX").Result(""))

    @property
    def center_y(self) -> float:
        """中心Y坐标（像素）"""
        return self._in_to_px(self.shape.Cells("PinY").Result(""))

    @staticmethod
    def _in_to_px(inch: float) -> float:
        """英寸转像素"""
        return inch * 96  # 1英寸=96像素


if __name__ == "__main__":
    from visio_diagram import VisioDiagram
    import base64

    # 初始化Visio绘图
    diagram = VisioDiagram()

    # 测试案例1：基础矩形（验证坐标、尺寸、文本）
    shape1 = VisioShape(
        page=diagram.page,
        shape_id="rect_basic",
        x=100, y=150,  # 像素单位（自动转换为英寸）
        width=200, height=100,  # 像素单位
        text="基础矩形",
        line_color=(0, 0, 0),  # 黑色边框
        bg_rgb=(255, 255, 0)  # 黄色填充
    )
    diagram.add_shape(shape1)
    print(f"形状1中心坐标: ({shape1.center_x}px, {shape1.center_y}px)")  # 验证单位转换

    # 测试案例2：带虚线边框的矩形（验证线条样式）
    shape2 = VisioShape(
        page=diagram.page,
        shape_id="rect_dashed",
        x=400, y=150,
        width=150, height=150,
        text="虚线边框",
        line_pattern=2,  # 虚线
        line_weight=2.5  # 2.5像素线宽
    )
    diagram.add_shape(shape2)

    # 测试案例3：嵌入Base64图片（验证图片功能）
    try:
        with open("frogw.png", "rb") as f:
            image_base64 = base64.b64encode(f.read()).decode()
        shape3 = VisioShape(
            page=diagram.page,
            shape_id="image_shape",
            x=100, y=350,
            width=180, height=120,
            image=image_base64  # Base64编码图片
        )
        diagram.add_shape(shape3)
    except FileNotFoundError:
        print("未找到frogw.png文件，跳过图片嵌入测试")
    except Exception as e:
        print(f"图片嵌入测试失败：{e}")

    # 测试案例4：边界值测试（最小尺寸和零值）
    shape4 = VisioShape(
        page=diagram.page,
        shape_id="rect_min_size",
        x=10, y=10,  # 极小坐标
        width=1, height=1,  # 1px尺寸（约0.01英寸）
        text="极小矩形",
        line_color=(255, 0, 0)  # 红色边框
    )
    diagram.add_shape(shape4)

    # 测试案例5：URL图片嵌入（需联网）
    shape5 = VisioShape(
        page=diagram.page,
        shape_id="web_image",
        x=400, y=350,
        width=200, height=150,
        image="https://pic2.zhimg.com/v2-d7ace568e5ecb33ce176d37c2a11a833_1440w.jpg"  # 直接使用URL
    )
    diagram.add_shape(shape5)

    # 保存并验证结果
    diagram.save_and_close(r"E:\Code\Python\DrawVisio\diagram.vsdx")
    print("测试完成，文件已保存至 E:\Code\Python\DrawVisio\diagram.vsdx")
