import json

import win32com.client as win32

from visio_connector import VisioConnector
from visio_page import VisioPage
from visio_shape import VisioShape


def batch_set_zorder(shapes: list):
    """
    高效批量设置形状层级（通过shape.z_index属性控制）
    参数：
        shapes: Visio形状对象列表（需含z_index属性）
    """
    # 1. 按z_index升序排序（0=最底层，值越大越靠前）
    sorted_shapes = sorted(shapes, key=lambda x: getattr(x, 'zIndex', 0))

    # 2. 批量设置层级
    visio = win32.Dispatch("Visio.Application")
    visio.ScreenUpdating = False  # 关闭刷新提升性能

    try:
        for shape in sorted_shapes:
            shape.shape.BringToFront()  # 每个形状仅需一次置顶
    finally:
        visio.ScreenUpdating = True


def main(json_path):
    # 启动Visio并隐藏窗口
    visio = win32.Dispatch("Visio.Application")
    visio.Visible = 0  # 隐藏窗口

    # 创建新文档和页面
    doc = visio.Documents.Add("")

    with open(json_path, 'r', encoding='utf-8') as f:
        config = json.load(f)

    page_config = config.get("flowData", {})

    # 创建页面
    page = VisioPage(doc, page_config.get("name"), height=page_config.get("height"), width=page_config.get("width"))

    # 创建shape
    nodes = config.get("graphData").get("nodes")
    edges = config.get("graphData").get("edges")
    shapes = []
    if nodes:
        for node in nodes:
            # if node.get('type') != 'act' and node.get('type') != 'text':
            #     continue
            shape = VisioShape(page=page.page, pageHeight=page_config.get("height"), **node)
            shapes.append(shape)

    batch_set_zorder(shapes)
    if edges:
        for edge in edges:
            connector = VisioConnector(page=page.page, pageHeight=page_config.get("height"), config=edge)

    doc.Pages.Item(1).Delete(0)

    # 保存文档
    output_path = "E:\\OneDrive - hnu.edu.cn\\Code\\Python\\DrawVisio\\static\\" + page_config.get("name") + ".vsdx"
    doc.SaveAs(output_path)
    visio.Quit()
    print(f"文档已保存至: {output_path}")


if __name__ == "__main__":
    main(json_path=r"E:\OneDrive - hnu.edu.cn\Code\Python\DrawVisio\static\graph.json")
