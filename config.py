import os


class Config:
    SECRET_KEY = 'your-secret-key-here'
    UPLOAD_FOLDER = 'uploads'
    MAX_CONTENT_LENGTH = 5 * 1024 * 1024 * 1024  # 1GB max file size
    ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif', 'webp'}
    BASE_URL = 'http://vladimir-hp.gradient.ru'  # Замените на ваш домен

    # Список клиентов
    CLIENTS = [
        'ЭЛИЗЕ',
        'Мегамаркет',
        'ЯндексМаркет',
        'МагнитКосметик'
    ]

    # Пути к шаблонам XLSX
    TEMPLATE_PATHS = {
        'ЭЛИЗЕ': 'templates/elise.xlsx',
        'Мегамаркет': 'templates/megamarket.xlsx',
        'ЯндексМаркет': 'templates/yandexmarket.xlsx',
        'МагнитКосметик': 'templates/magnitcosmetic.xlsx'
    }


def allowed_file(filename):
    return '.' in filename and \
        filename.rsplit('.', 1)[1].lower() in Config.ALLOWED_EXTENSIONS
