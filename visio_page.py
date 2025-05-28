from dataclasses import dataclass, field

import win32com.client as win32

from utils import UnitParser


@dataclass
class VisioPage:
    doc: object
    name: str
    height: float  # 单位默认px
    width: float
    backgroundColor: str = None
    backgroundImage: str = None
    page: object = field(init=False, repr=False)

    def __post_init__(self):
        self.page = self.doc.Pages.Add()
        self.page.Name = self.name

        self._parse_units()

        self.page.PageSheet.Cells("PageWidth").Formula = self.width  # 设置宽度（英寸）
        self.page.PageSheet.Cells("PageHeight").Formula = self.height  # 设置高度

        # if self.backgroundImage:
        #     self._set_background_image()
        # elif self.backgroundColor:
        #     self._set_background_color()

    def _parse_units(self):
        """解析所有带单位的参数"""
        # 坐标和尺寸转换（px → in）
        self.height = UnitParser.px_to_in(self.height)
        self.width = UnitParser.px_to_in(self.width)

    def _set_background_image(self):
        """设置背景"""
        self.page.Background = True
        # 插入图片并设为背景
        image_path = r"E:\OneDrive - hnu.edu.cn\Code\Python\DrawVisio\static\frog.png"  # 替换为实际路径
        pic_shape = self.page.Shapes.AddPicture(
            FileName=image_path,
            LinkToFile=False,  # True=链接到文件，False=嵌入
            SaveWithDocument=True,
            Left=0, Top=0, Width=self.page.PageSheet.Cells("PageWidth").ResultIU,  # 铺满页面
            Height=self.page.PageSheet.Cells("PageHeight").ResultIU
        )

        # 将图片置于底层并锁定（可选）
        pic_shape.SendToBack()
        pic_shape.Cells("LockMoveX").Formula = "1"  # 锁定位置
        pic_shape.Cells("LockMoveY").Formula = "1"
        pic_shape.Cells("LockAspect").Formula = "1"  # 锁定纵横比

    def _set_background_color(self):
        """通过添加矩形形状实现背景色"""
        rect = self.page.DrawRectangle(0, 0, self.width, self.height)

        if self.backgroundColor.startswith("#"):
            r, g, b = UnitParser.hex_to_rgb(self.backgroundColor)
            color_formula = f"RGB({r},{g},{b})"
        else:
            color_formula = f"RGB({self.backgroundColor})"

        rect.Cells("FillForegnd").Formula = color_formula
        rect.Cells("LinePattern").Formula = "0"  # 无边框
        rect.SendToBack()
        rect.Cells("LockMoveX").Formula = "1"
        rect.Cells("LockMoveY").Formula = "1"


# 测试样例（在main中使用）
if __name__ == "__main__":
    # 初始化Visio应用
    visio = win32.Dispatch("Visio.Application")
    doc = visio.Documents.Add("")  # 新建空白文档
    # 测试1：创建白色背景页（十六进制颜色）
    white_page = VisioPage(
        doc=doc,
        name="WhitePage",
        height=800,
        width=600,
        backgroundColor="#FFFFFF"  # 纯白色
    )

    # 测试2：创建蓝色背景页（RGB字符串）
    blue_page = VisioPage(
        doc=doc,
        name="BluePage",
        height=800,
        width=600,
        backgroundColor="173, 216, 230"  # 浅蓝色
    )

    # 测试3：创建带图片背景的页
    img_page = VisioPage(
        doc=doc,
        name="ImagePage",
        height=300,
        width=600,
        backgroundImage=r"C:\path\to\your\image.jpg"  # 替换为实际图片路径
    )

    doc.Pages.Item(1).Delete(0)

    # 保存文档
    output_path = r"E:\OneDrive - hnu.edu.cn\Code\Python\DrawVisio\static\diagram.vsdx"
    doc.SaveAs(output_path)
    print(f"文档已保存至: {output_path}")

    # 关闭Visio（可选）
    visio.Quit()
