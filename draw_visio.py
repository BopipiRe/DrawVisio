import json

from visio_diagram import VisioDiagram


def create_diagram_from_json(json_file_path, output_path):
    """从JSON配置创建Visio图表"""
    diagram = VisioDiagram()

    with open(json_file_path, 'r', encoding='utf-8') as f:
        config = json.load(f)

    # 添加所有形状
    for shape_config in config.get("shapes", []):
        diagram.add_shape(
            shape_id=shape_config["id"],
            x=shape_config.get("x", 0),
            y=shape_config.get("y", 0),
            width=shape_config.get("width", 1),
            height=shape_config.get("height", 1),
            text=shape_config.get("text")
        )
        diagram.auto_fit_canvas()

    # 添加所有连接线
    for connector_config in config.get("connectors", []):
        diagram.add_connector(
            from_id=connector_config["from"],
            to_id=connector_config["to"],
            line_style=connector_config.get("style", "solid")
        )

    diagram.save_and_close(output_path)


# 使用示例
if __name__ == "__main__":
    create_diagram_from_json(
        json_file_path="config.json",
        output_path=r"E:\diagram.vsdx"
    )
