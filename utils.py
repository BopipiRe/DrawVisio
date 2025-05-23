import re


class UnitParser:
    @staticmethod
    def parse_length(value: str) -> float:
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
            return UnitParser.px_to_in(num)
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

    @staticmethod
    def parse_line_weight(value: str) -> float:
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
            return UnitParser.px_to_pt(num)
        elif unit == "pt":
            return num
        else:
            raise ValueError(f"不支持的单位: {unit}")

    @staticmethod
    def in_to_px(inch: float) -> float:
        """英寸转像素"""
        return inch * 96

    @staticmethod
    def px_to_in(px: float) -> float:
        """像素转英寸（假设96DPI）"""
        return px / 96  # 1英寸=96像素

    @staticmethod
    def px_to_pt(px: float) -> float:
        """像素转磅（1磅=1/72英寸）[[61][65]]"""
        return px * 72 / 96  # 1pt = (1/72)in = (1/72)*96px ≈ 1.33px

    @staticmethod
    def hex_to_rgb(hex_color: str) -> tuple:
        """将 #FFFFFF 格式转换为 (R, G, B) 元组"""
        hex_color = hex_color.lstrip("#")
        if len(hex_color) == 3:  # 处理缩写格式如 #FFF
            hex_color = "".join([c * 2 for c in hex_color])
        return tuple(int(hex_color[i:i + 2], 16) for i in (0, 2, 4))