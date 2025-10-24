import os

class Config:
    SECRET_KEY = 'your-secret-key-here'
    UPLOAD_FOLDER = 'uploads'
    RESULTS_FOLDER = 'results'  # <-- Добавляем папку для результатов
    MAX_CONTENT_LENGTH = 5 * 1024 * 1024 * 1024  # 1GB max file size
    ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif', 'webp'}
    BASE_URL = 'http://tecnobook'  # Замените на ваш домен
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

    # Убедимся, что папки существуют
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    os.makedirs(RESULTS_FOLDER, exist_ok=True) # <-- Добавляем создание папки результатов

def allowed_file(filename):
    return '.' in filename and \
        filename.rsplit('.', 1)[1].lower() in Config.ALLOWED_EXTENSIONS
