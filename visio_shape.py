import base64
import os.path
import re
from dataclasses import dataclass, field
from typing import Optional, Tuple, Union


@dataclass
class VisioShape:
    """Visio图形基类（支持带单位的参数输入）"""
    page: object
    shape_id: str
    x: str  # X坐标（如"100px"、"2.5in"）
    y: str  # Y坐标（如"150px"、"3cm"）
    width: str  # 宽度（如"200px"、"1.5in"）
    height: str  # 高度（如"100px"、"50mm"）
    text: Optional[str] = None
    bg_rgb: Optional[Tuple[int, int, int]] = None
    line_color: Optional[Tuple[int, int, int]] = None
    line_pattern: int = 1  # 线条模式（1=实线）
    line_weight: str = "1pt"  # 线宽（如"2pt"、"1.5px"）
    fill_pattern: int = 1  # 填充模式
    image: Optional[Union[str, bytes]] = None
    shape: object = field(init=False, repr=False)

    def __post_init__(self):
        """初始化后自动解析单位并创建形状"""
        self._parse_units()  # 解析带单位的参数
        self._create_shape()
        self._set_image() if self.image else None

    def _parse_units(self):
        """解析所有带单位的参数"""
        # 坐标和尺寸转换（支持px/in/cm/mm → in）
        self.x_in = self._parse_length(self.x)
        self.y_in = self._parse_length(self.y)
        self.width_in = self._parse_length(self.width)
        self.height_in = self._parse_length(self.height)

        # 线宽转换（支持px/pt → pt）
        self.line_weight_pt = self._parse_line_weight(self.line_weight)

    def _parse_length(self, value: str) -> float:
        """
        解析长度单位（支持px/in/cm/mm/ft）
        格式: 数值 + 单位（如 "100px", "2.5in", "50mm"）
        默认单位: px
        """
        match = re.match(r"^([\d.]+)\s*(px|in|cm|mm|ft)?$", str(value).strip(), re.IGNORECASE)
        if not match:
            raise ValueError(f"无效的长度格式: {value}")

        num, unit = float(match.group(1)), (match.group(2) or "px").lower()

        # 转换为英寸（Visio内部单位）
        if unit == "px":
            return self._px_to_in(num)
        elif unit == "in":
            return num
        elif unit == "cm":
            return num / 2.54
        elif unit == "mm":
            return num / 25.4
        elif unit == "ft":
            return num * 12
        else:
            raise ValueError(f"不支持的单位: {unit}")

    def _parse_line_weight(self, value: str) -> float:
        """
        解析线宽单位（支持px/pt）
        格式: 数值 + 单位（如 "2pt", "1.5px"）
        默认单位: pt
        """
        match = re.match(r"^([\d.]+)\s*(px|pt)?$", str(value).strip(), re.IGNORECASE)
        if not match:
            raise ValueError(f"无效的线宽格式: {value}")

        num, unit = float(match.group(1)), (match.group(2) or "pt").lower()

        # 转换为磅（Visio线宽单位）
        if unit == "px":
            return self._px_to_pt(num)
        elif unit == "pt":
            return num
        else:
            raise ValueError(f"不支持的单位: {unit}")

    @staticmethod
    def _px_to_in(px: float) -> float:
        """像素转英寸（假设96DPI）"""
        return px / 96  # 1英寸=96像素

    @staticmethod
    def _px_to_pt(px: float) -> float:
        """像素转磅（1磅=1/72英寸）[[61][65]]"""
        return px * 72 / 96  # 1pt = (1/72)in = (1/72)*96px ≈ 1.33px

    def _create_shape(self):
        """创建矩形（使用解析后的英寸单位）"""
        self.shape = self.page.DrawRectangle(
            self.x_in, self.y_in,
            self.x_in + self.width_in,
            self.y_in + self.height_in
        )
        if self.text:
            self.shape.Text = self.text
        self._set_style()

    def _set_style(self):
        """设置样式（使用解析后的磅单位）"""
        if self.bg_rgb:
            self.shape.CellsU("FillForegnd").FormulaU = f"RGB({self.bg_rgb[0]},{self.bg_rgb[1]},{self.bg_rgb[2]})"
        if self.line_color:
            self.shape.CellsU(
                "LineColor").FormulaU = f"RGB({self.line_color[0]},{self.line_color[1]},{self.line_color[2]})"
        self.shape.CellsU("LinePattern").FormulaU = str(self.line_pattern)
        self.shape.CellsU("LineWeight").FormulaU = f"{self.line_weight_pt} pt"
        self.shape.CellsU("FillPattern").FormulaU = str(self.fill_pattern)

    # 图片嵌入方法保持不变（同之前实现）
    def _set_image(self):
        """嵌入图片（支持URL/Base64/本地路径）"""
        try:
            if self.image.startswith(('http://', 'https://')):
                self._insert_image_from_url(self.image)
            elif os.path.isfile(self.image):
                with open(self.image, "rb") as f:
                    self._insert_image_from_base64(base64.b64encode(f.read()).decode())
            else:
                self._insert_image_from_base64(self.image)
        except Exception as e:
            print(f"图片嵌入失败: {e}")

    def _insert_image_from_url(self, url: str):
        """通过URL插入图片"""
        img_shape = self.page.Import(url)
        img_shape.CellsU("Width").FormulaU = f"{self.width_in} in"
        img_shape.CellsU("Height").FormulaU = f"{self.height_in} in"
        img_shape.CellsU("PinX").FormulaU = f"{self.x_in + self.width_in / 2} in"
        img_shape.CellsU("PinY").FormulaU = f"{self.y_in + self.height_in / 2} in"

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
        return self.shape.Cells("PinX").Result("")

    @property
    def center_y(self) -> float:
        return self.shape.Cells("PinY").Result("")

    @staticmethod
    def _in_to_px(inch: float) -> float:
        """英寸转像素"""
        return inch * 96


if __name__ == "__main__":
    from visio_diagram import VisioDiagram

    # 初始化Visio绘图
    diagram = VisioDiagram()

    # 测试案例1：基础矩形（验证坐标、尺寸、文本）
    shape1 = VisioShape(
        page=diagram.page,
        shape_id="rect_basic",
        x="100px", y="150px",  # 像素单位（自动转换为英寸）
        width="200px", height="100px",  # 像素单位
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
        x="400px", y="150px",
        width="150px", height="150px",
        text="虚线边框",
        line_pattern=2,  # 虚线
        line_weight="2.5px"  # 2.5像素线宽
    )
    diagram.add_shape(shape2)

    # 测试案例3：嵌入Base64图片（验证图片功能）
    with open("frog.png", "rb") as f:
        image_base64 = base64.b64encode(f.read()).decode()
    shape3 = VisioShape(
        page=diagram.page,
        shape_id="image_shape",
        x="100px", y="350px",
        width="180px", height="120px",
        image="frog.png"  # Base64编码图片
    )
    diagram.add_shape(shape3)

    # 测试案例4：边界值测试（最小尺寸和零值）
    shape4 = VisioShape(
        page=diagram.page,
        shape_id="rect_min_size",
        x="10px", y="10px",  # 极小坐标
        width="1px", height="1px",  # 1px尺寸（约0.01英寸）
        text="极小矩形",
        line_color=(255, 0, 0)  # 红色边框
    )
    diagram.add_shape(shape4)

    # 测试案例5：URL图片嵌入（需联网）
    shape5 = VisioShape(
        page=diagram.page,
        shape_id="web_image",
        x="400px", y="350px",
        width="200px", height="150px",
        image="https://pic2.zhimg.com/v2-d7ace568e5ecb33ce176d37c2a11a833_1440w.jpg"  # 直接使用URL
    )
    diagram.add_shape(shape5)

    # 保存并验证结果
    diagram.save_and_close(r"E:\Code\Python\DrawVisio\diagram.vsdx")
    print("测试完成，文件已保存至 E:\Code\Python\DrawVisio\diagram.vsdx")
