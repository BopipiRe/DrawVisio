import json

from visio_connector import VisioConnector
from visio_diagram import VisioDiagram
from visio_shape import VisioShape


def create_diagram_from_json(json_file_path, output_path):
    """从JSON配置创建Visio图表"""
    diagram = VisioDiagram()

    with open(json_file_path, 'r', encoding='utf-8') as f:
        config = json.load(f)

    # 添加所有形状
    for shape_config in config.get("shapes", []):
        shape = VisioShape(
            page=diagram.page,
            shape_id=shape_config.pop("id"),  # 移除已单独处理的键
            **shape_config  # 剩余键值自动解包为命名参数
        )
        diagram.add_shape(shape)
        diagram.auto_fit_canvas()

    # 添加所有连接线
    for connector_config in config.get("connectors", []):
        connector = VisioConnector(
            page=diagram.page,
            **connector_config
        )
        diagram.add_connector(connector)

    diagram.save_and_close(output_path)


# 使用示例
if __name__ == "__main__":
    create_diagram_from_json(
        json_file_path="config.json",
        output_path=r"E:\Code\Python\DrawVisio\diagram.vsdx"
    )
