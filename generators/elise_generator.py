from .base_generator import BaseGenerator
from openpyxl.utils import get_column_letter


class EliseGenerator(BaseGenerator):

    def __init__(self):
        super().__init__('elise.xlsx')

    def get_worksheet_title(self):
        return "ELISE Images"

    def get_headers(self):
        headers = [
            "Бренд", "Номенклатура", "Артикул", "Базовый Штрихкод",
            "Измерение", "ЗначениеИзмерения",
            "Заголовок (Без бренда)", "КраткоеОписание", "СпособПрименения",
            "ЛинейкаБренда", "Каждый цвет"
        ]
        # Добавляем столбцы для изображений
        max_images = 20
        for i in range(1, max_images + 1):
            headers.append(f"ИмяФайлаКартинки{i}")
        return headers

    def generate_row_data(self, article, urls, client_name):
        row_data = [
            "ARTDECO",  # Бренд
            f"ARTDECO Товар артикул {article}",  # Номенклатура
            article,  # Артикул
            "",  # Базовый Штрихкод
            "",  # Измерение
            "",  # ЗначениеИзмерения
            f"Товар артикул {article}",  # Заголовок (Без бренда)
            f"Изображения товара артикул {article} для клиента {client_name}",  # КраткоеОписание
            "Нанести",  # СпособПрименения
            "АРТДЕКО",  # Линейка Бренда
            "Каждый цвет"  # Каждый цвет
        ]

        # Добавляем ссылки на изображения
        max_images = 20
        for i, url in enumerate(urls):
            if i >= max_images:
                break
            row_data.append(url)

        # Заполняем оставшиеся ячейки для изображений пустыми значениями
        remaining_images = max_images - min(len(urls), max_images)
        for _ in range(remaining_images):
            row_data.append("")

        return row_data

    def adjust_column_widths(self, ws):
        column_widths = {
            'A': 15, 'B': 30, 'C': 15, 'D': 20, 'E': 15,
            'F': 25, 'G': 25, 'H': 40, 'I': 20, 'J': 20,
            'K': 25
        }
        # Устанавливаем ширину для столбцов с изображениями
        for i in range(12, 12 + 20):  # Столбцы L-AE для изображений
            column_letter = get_column_letter(i)
            column_widths[column_letter] = 30

        for col, width in column_widths.items():
            ws.column_dimensions[col].width = width
