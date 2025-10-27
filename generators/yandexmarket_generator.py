# generators/yandexmarket_generator.py
from .base_generator import BaseGenerator
class YandexmarketGenerator(BaseGenerator):
    def __init__(self):
        # Имя шаблона может совпадать с именем в списке TEMPLATES
        super().__init__('yandexmarket.xlsx')

    def get_start_row(self):
        return 4  # ЯндексМаркет начинает данные с 4 строки

    def get_worksheet_title(self): # Переопределяем для конкретного шаблона
        return "YandexMarket Catalog"

    def get_headers(self):
        # Для ЯндексМаркета заголовки уже есть в шаблоне
        return []

    def create_new_workbook(self):
        """Создает базовый шаблон для ЯндексМаркета"""
        wb = Workbook()
        ws = wb.active
        ws.title = self.get_worksheet_title() # Используем переопределенный метод
        # Стили для заголовков (если нужно добавить в новые строки)
        header_font = Font(bold=True)
        # Предполагаем, что заголовки в шаблоне есть, иначе добавим их
        # ws.cell(row=1, column=1, value="Название").font = header_font
        # ws.cell(row=1, column=2, value="Описание").font = header_font
        # и т.д.
        return wb, ws, 2 # или get_start_row()

    def generate_row_data(self, article, urls, template_name): # template_name теперь доступен, но не используется в этом генераторе
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
