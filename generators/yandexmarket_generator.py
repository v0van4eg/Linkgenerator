from .base_generator import BaseGenerator


class YandexmarketGenerator(BaseGenerator):

    def __init__(self):
        super().__init__('yandexmarket.xlsx')

    def get_start_row(self):
        return 4  # ЯндексМаркет начинает данные с 4 строки

    def get_worksheet_title(self):
        return "YandexMarket Catalog"

    def get_headers(self):
        # Для ЯндексМаркета заголовки уже есть в шаблоне
        return []

    def create_new_workbook(self):
        """Создает базовый шаблон для ЯндексМаркета"""
        wb = Workbook()
        ws = wb.active
        ws.title = self.get_worksheet_title()

        # Базовые заголовки (упрощенная версия)
        headers = [
            "Ваш SKU *", "Название товара *", "Ссылка на изображение *",
            "Описание товара *", "Категория на Маркете *", "Бренд *",
            "Штрихкод *", "Страна производства", "Ссылка на страницу товара на вашем сайте",
            "Артикул производителя *"
        ]
        ws.append(headers)

        # Стили для заголовков
        header_font = Font(bold=True)
        for cell in ws[1]:
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')

        return wb, ws, 2

    def generate_row_data(self, article, urls, client_name):
        return [
            article,  # A: Ваш SKU *
            f"Товар артикул {article}",  # B: Название товара *
            urls[0] if urls else "",  # C: Ссылка на изображение * (первая)
            f"Описание товара артикул {article}",  # D: Описание товара *
            "Категория по умолчанию",  # E: Категория на Маркете *
            "Бренд по умолчанию",  # F: Бренд *
            "Штрихкод заглушка",  # G: Штрихкод *
            "",  # H: Страна производства
            "",  # I: Ссылка на страницу товара на вашем сайте
            article,  # J: Артикул производителя *
        ]
