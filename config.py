# config.py
import os

class Config:
    SECRET_KEY = 'your-secret-key-here'
    UPLOAD_FOLDER = 'uploads'
    RESULTS_FOLDER = 'results'  # <-- Добавляем папку для результатов
    MAX_CONTENT_LENGTH = 5 * 1024 * 1024 * 1024  # 1GB max file size
    ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif', 'webp'}
    BASE_URL = os.getenv('BASE_URL', 'http://tecnobook')

    # Список шаблонов (вместо клиентов)
    TEMPLATES = [
        'Мегамаркет',
        'ЯндексМаркет',
        # Добавьте другие шаблоны по мере необходимости
    ]

    # Пути к шаблонам XLSX (обновлены, если имя шаблона отличается от имени файла)
    TEMPLATE_PATHS = {
        'Мегамаркет': 'templates/megamarket.xlsx',
        'ЯндексМаркет': 'templates/yandexmarket.xlsx',
        # Убедитесь, что имена ключей соответствуют именам в TEMPLATES
    }

    # Убедимся, что папки существуют
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    os.makedirs(RESULTS_FOLDER, exist_ok=True) # <-- Добавляем создание папки результатов

def allowed_file(filename):
    return '.' in filename and \
        filename.rsplit('.', 1)[1].lower() in Config.ALLOWED_EXTENSIONS
