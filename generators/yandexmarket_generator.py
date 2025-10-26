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

        # Стили для заголовков
        header_font = Font(bold=True)
        for cell in ws[1]:
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')

        return wb, ws, 2

    def generate_row_data(self, article, urls, client_name):
        # строка из 30 ссылок, разделённых запятыми
        if urls:
            urls = ','.join(urls)
        else:
            urls = ""
        return [
            "", # A
            "", # B
            article,  # C: Ваш SKU *
            f"Товар артикул {article}",  # D: Название товара *
            urls,  # спсиорк ссылок на изображения
            f"Описание товара артикул {article}",  # E: Описание товара *
        ]
