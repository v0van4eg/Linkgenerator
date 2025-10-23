from flask import Flask, render_template, request, jsonify, send_file
import os
import uuid
from werkzeug.utils import secure_filename
from config import Config, allowed_file
import re
import unicodedata
from urllib.parse import quote
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
import io

app = Flask(__name__)
app.config.from_object(Config)

# Создаем директорию для загрузок если не существует
os.makedirs(Config.UPLOAD_FOLDER, exist_ok=True)


def safe_folder_name(name: str) -> str:
    """
    Преобразует строку в безопасное имя папки:
    - Оставляет буквы (включая кириллицу), цифры, дефис, подчёркивание
    - Заменяет пробелы и другие разделители на дефис
    - Удаляет потенциально опасные символы
    - Убирает начальные/конечные дефисы и точки
    """
    if not name:
        return "unnamed"
    # Нормализуем Unicode (например, буквы с диакритикой)
    name = unicodedata.normalize('NFKD', name)
    # Разрешённые символы: буквы (все алфавиты), цифры, дефис, подчёркивание
    # Заменяем всё остальное на дефис
    name = re.sub(r'[^\w\s-]', '', name, flags=re.UNICODE)
    name = re.sub(r'[-\s]+', '-', name, flags=re.UNICODE).strip('-_')
    # Ограничим длину, чтобы избежать проблем с ОС
    name = name[:255] if name else "unnamed"
    return name


def generate_xlsx_document(image_urls, client_name, product_name):
    """
    Генерирует XLSX документ с заданной структурой
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Product Images"

    # Заголовки столбцов
    headers = [
        "Бренд", "Номенклатура", "Артикул", "Базовый Штрихкод",
        "Измерение", "ЗначениеИзмерения (Если косметика то колористика)",
        "Заголовок (Без бренда)", "КраткоеОписание", "СпособПрименения",
        "ЛинейкаБренда", "Каждый цвет (ПО мише еще и по объёмам)"
    ]

    # Добавляем столбцы для изображений
    max_images = 10  # Максимальное количество столбцов для изображений
    for i in range(1, max_images + 1):
        headers.append(f"ИмяФайлаКартинки{i}")

    # Записываем заголовки
    ws.append(headers)

    # Стили для заголовков
    header_font = Font(bold=True)
    for cell in ws[1]:
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')

    # Данные для строки
    # В реальном приложении эти данные можно получать из формы или базы данных
    row_data = [
        "",  # Бренд
        product_name,  # Номенклатура
        "",  # Артикул
        "",  # Базовый Штрихкод
        "",  # Измерение
        "",  # ЗначениеИзмерения
        product_name,  # Заголовок (Без бренда)
        f"Изображения товара {product_name} для клиента {client_name}",  # КраткоеОписание
        "",  # СпособПрименения
        "",  # ЛинейкаБренда
        ""  # Каждый цвет
    ]

    # Добавляем ссылки на изображения
    for i, url in enumerate(image_urls):
        if i >= max_images:
            break
        row_data.append(url)

    # Заполняем оставшиеся ячейки для изображений пустыми значениями
    remaining_images = max_images - min(len(image_urls), max_images)
    for _ in range(remaining_images):
        row_data.append("")

    # Записываем данные
    ws.append(row_data)

    # Настраиваем ширину столбцов
    column_widths = {
        'A': 15, 'B': 30, 'C': 15, 'D': 20, 'E': 15,
        'F': 25, 'G': 25, 'H': 40, 'I': 20, 'J': 20,
        'K': 25
    }

    # Устанавливаем ширину для столбцов с изображениями
    for i in range(12, 12 + max_images):
        column_letter = chr(64 + i)  # Преобразуем номер в букву столбца
        column_widths[column_letter] = 30

    for col, width in column_widths.items():
        ws.column_dimensions[col].width = width

    # Сохраняем в буфер
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    return buffer


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Получаем данные из формы
        client_name = request.form.get('client_name', '').strip()
        product_name = request.form.get('product_name', '').strip()

        if not client_name or not product_name:
            return render_template('index.html', error='Заполните все поля')

        # Создаем пути для клиента и товара
        client_folder = safe_folder_name(client_name)
        product_folder = safe_folder_name(product_name)

        full_path = os.path.join(Config.UPLOAD_FOLDER, client_folder, product_folder)
        os.makedirs(full_path, exist_ok=True)

        uploaded_files = request.files.getlist('images')
        image_urls = []

        for file in uploaded_files:
            if file and allowed_file(file.filename):
                # Генерируем уникальное имя файла
                random_hex = uuid.uuid4().hex[:6]
                file_extension = os.path.splitext(file.filename)[1]
                file_name = os.path.splitext(file.filename)[0]
                unique_filename = f"{file_name}-{random_hex}{file_extension}"

                # Сохраняем файл
                file_path = os.path.join(full_path, unique_filename)
                file.save(file_path)

                # Генерируем URL
                image_url = "{}/images/{}/{}/{}".format(
                    Config.BASE_URL,
                    quote(client_folder, safe=''),
                    quote(product_folder, safe=''),
                    quote(unique_filename, safe='')
                )
                image_urls.append(image_url)

        return render_template('index.html',
                               image_urls=image_urls,
                               client_name=client_name,
                               product_name=product_name)

    return render_template('index.html')


@app.route('/download-links')
def download_links():
    urls = request.args.getlist('urls')
    if not urls:
        return "No URLs provided", 400

    # Создаем временный файл
    temp_file = f"temp_links_{uuid.uuid4().hex}.txt"
    with open(temp_file, 'w', encoding='utf-8') as f:
        f.write('\n'.join(urls))

    return send_file(temp_file,
                     as_attachment=True,
                     download_name='image_links.txt',
                     mimetype='text/plain')


@app.route('/download-xlsx')
def download_xlsx():
    urls = request.args.getlist('urls')
    client_name = request.args.get('client_name', '')
    product_name = request.args.get('product_name', '')

    if not urls:
        return "No URLs provided", 400

    # Генерируем XLSX документ
    xlsx_buffer = generate_xlsx_document(urls, client_name, product_name)

    # Создаем имя файла
    filename = f"{safe_folder_name(client_name)}_{safe_folder_name(product_name)}_images.xlsx"

    return send_file(
        xlsx_buffer,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
