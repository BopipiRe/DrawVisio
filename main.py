import json

import win32com.client as win32

import visio_page


def main(json_path):
    # 启动Visio并隐藏窗口
    visio = win32.Dispatch("Visio.Application")
    visio.Visible = 0  # 隐藏窗口

    # 创建新文档和页面
    doc = visio.Documents.Add("")

    with open(json_path, 'r', encoding='utf-8') as f:
        config = json.load(f)

    page_config = config.get("flowData", {})

    visio_page.VisioPage(doc, page_config.get("name"), height=page_config.get("height"), width=page_config.get("width"))

    doc.Pages.Item(1).Delete(0)

    # 保存文档
    output_path = "E:\\OneDrive - hnu.edu.cn\\Code\\Python\\DrawVisio\\static\\" + page_config.get("name") + ".vsdx"
    doc.SaveAs(output_path)
    visio.Quit()
    print(f"文档已保存至: {output_path}")


if __name__ == "__main__":
    main(json_path=r"E:\OneDrive - hnu.edu.cn\Code\Python\DrawVisio\static\graph.json")
