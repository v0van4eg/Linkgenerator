from .base_generator import BaseGenerator
from openpyxl.utils import get_column_letter


class MagnitcosmeticGenerator(BaseGenerator):

    def __init__(self):
        super().__init__('magnitcosmetic.xlsx')

    def get_worksheet_title(self):
        return "MagnitCosmetic Images"

    def get_headers(self):
        headers = ["Код товара", "Артикул", "Основное фото"]
        for i in range(1, 16):
            headers.append(f"Фото {i}")
        return headers

    def generate_row_data(self, article, urls, client_name):
        row_data = [article, article]  # Код товара и Артикул

        # Добавляем ссылки на изображения
        for i in range(16):  # Основное + 15 дополнительных
            if i < len(urls):
                row_data.append(urls[i])
            else:
                row_data.append("")

        return row_data

    def adjust_column_widths(self, ws):
        for col in range(1, 19):
            column_letter = get_column_letter(col)
            ws.column_dimensions[column_letter].width = 30
