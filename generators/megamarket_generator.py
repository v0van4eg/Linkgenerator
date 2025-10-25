from .base_generator import BaseGenerator
from openpyxl.utils import get_column_letter


class MegamarketGenerator(BaseGenerator):

    def __init__(self):
        super().__init__('megamarket.xlsx')

    def get_worksheet_title(self):
        return "Megamarket Images"

    def get_headers(self):
        headers = ["Код товара СММ(обязательно)", "Ссылка на основное фото"]
        for i in range(1, 10):
            headers.append(f"Ссылка на доп. фото №{i}")
        return headers

    def generate_row_data(self, article, urls, client_name):
        row_data = [article]  # Код товара СММ

        # Добавляем ссылки на изображения
        if urls:
            row_data.append(urls[0])  # Основное фото
            # Дополнительные фото (максимум 9)
            for i in range(1, 10):
                if i < len(urls):
                    row_data.append(urls[i])
                else:
                    row_data.append("")
        else:
            # Если нет изображений, добавляем пустые ячейки
            row_data.append("")  # Основное фото
            for i in range(1, 10):
                row_data.append("")

        return row_data

    def adjust_column_widths(self, ws):
        column_widths = {
            'A': 20,  # Код товара
            'B': 40,  # Основное фото
        }
        for i in range(3, 12):  # Столбцы C-K для доп. фото
            column_letter = get_column_letter(i)
            column_widths[column_letter] = 40

        for col, width in column_widths.items():
            ws.column_dimensions[col].width = width
