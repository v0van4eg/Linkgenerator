from flask import Flask, render_template, request, jsonify, send_file
import os
import uuid
import zipfile
from werkzeug.utils import secure_filename
from config import Config, allowed_file
import re
import unicodedata
from urllib.parse import quote
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter
import io
import tempfile
import shutil

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


def process_zip_archive(zip_file, client_name):
    """
    Обрабатывает ZIP-архив и извлекает изображения в структуру каталогов клиента
    Возвращает список URL созданных изображений
    """
    image_urls = []

    # Создаем временную директорию для распаковки
    with tempfile.TemporaryDirectory() as temp_dir:
        # Сохраняем ZIP файл
        zip_path = os.path.join(temp_dir, zip_file.filename)
        zip_file.save(zip_path)

        # Распаковываем архив
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        # Обходим все файлы в распакованной директории
        for root, dirs, files in os.walk(temp_dir):
            for file in files:
                # Пропускаем системные файлы
                if file.lower() in ['thumbs.db', '.ds_store']:
                    continue

                # Проверяем расширение файла
                if not allowed_file(file):
                    continue

                # Определяем артикул из имени папки
                relative_path = os.path.relpath(root, temp_dir)
                if relative_path == '.':
                    # Файлы в корне архива - пропускаем
                    continue

                # Первая папка в пути - это артикул
                article = os.path.basename(root) if relative_path.count(os.sep) == 0 else relative_path.split(os.sep)[0]

                # Создаем пути для клиента и артикула
                client_folder = safe_folder_name(client_name)
                article_folder = safe_folder_name(article)

                full_path = os.path.join(Config.UPLOAD_FOLDER, client_folder, article_folder)
                os.makedirs(full_path, exist_ok=True)

                # Копируем файл в целевую директорию
                source_file = os.path.join(root, file)
                target_file = os.path.join(full_path, file)

                shutil.copy2(source_file, target_file)

                # Генерируем URL
                image_url = "{}/images/{}/{}/{}".format(
                    Config.BASE_URL,
                    quote(client_folder, safe=''),
                    quote(article_folder, safe=''),
                    quote(file, safe='')
                )
                image_urls.append({
                    'url': image_url,
                    'article': article,
                    'filename': file
                })

    return image_urls


def generate_xlsx_document(image_data, client_name):
    """
    Генерирует XLSX документ с заданной структурой
    image_data - список словарей с ключами: url, article, filename
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Product Images"

    # Заголовки столбцов
    headers = [
        "Бренд", "Номенклатура", "Артикул", "Базовый Штрихкод",
        "Измерение", "ЗначениеИзмерения",
        "Заголовок (Без бренда)", "КраткоеОписание", "СпособПрименения",
        "ЛинейкаБренда", "Каждый цвет"
    ]

    # Добавляем столбцы для изображений
    max_images = 30  # Максимальное количество столбцов для изображений
    for i in range(1, max_images + 1):
        headers.append(f"ИмяФайлаКартинки{i}")

    # Записываем заголовки
    ws.append(headers)

    # Стили для заголовков
    header_font = Font(bold=True)
    for cell in ws[1]:
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')

    # Группируем изображения по артикулам
    articles = {}
    for item in image_data:
        article = item['article']
        if article not in articles:
            articles[article] = []
        articles[article].append(item['url'])

    print(f"Найдено артикулов: {len(articles)}")
    for article, urls in articles.items():
        print(f"Артикул {article}: {len(urls)} изображений")

    # Создаем строки для каждого артикула
    for article, urls in articles.items():
        # Данные для строки
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
        for i, url in enumerate(urls):
            if i >= max_images:
                break
            row_data.append(url)

        # Заполняем оставшиеся ячейки для изображений пустыми значениями
        remaining_images = max_images - min(len(urls), max_images)
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
        column_letter = get_column_letter(i)
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
        # Проверяем, какая форма отправлена - отдельные файлы или архив
        if 'archive' in request.files and request.files['archive'].filename != '':
            # Обработка ZIP архива
            return handle_archive_upload(request)
        else:
            # Обработка отдельных изображений
            return handle_single_upload(request)

    # GET запрос - отображаем форму с пустыми полями
    return render_template('index.html',
                           clients=Config.CLIENTS,
                           selected_client='',
                           product_name='',
                           image_urls=[])


def handle_single_upload(request):
    """Обработка загрузки отдельных изображений"""
    client_name = request.form.get('client_name', '').strip()
    product_name = request.form.get('product_name', '').strip()

    if not client_name or not product_name:
        return render_template('index.html',
                               error='Заполните все поля',
                               clients=Config.CLIENTS,
                               selected_client=client_name,
                               product_name=product_name)

    # Проверяем, что выбранный клиент есть в списке
    if client_name not in Config.CLIENTS:
        return render_template('index.html',
                               error='Выберите клиента из списка',
                               clients=Config.CLIENTS,
                               selected_client=client_name,
                               product_name=product_name)

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
            image_urls.append({
                'url': image_url,
                'article': product_name,
                'filename': unique_filename
            })

    return render_template('index.html',
                           image_urls=image_urls,
                           clients=Config.CLIENTS,
                           selected_client=client_name,
                           product_name=product_name)


def handle_archive_upload(request):
    """Обработка загрузки ZIP архива"""
    client_name = request.form.get('client_name', '').strip()
    archive_file = request.files['archive']

    if not client_name or not archive_file or archive_file.filename == '':
        return render_template('index.html',
                               error='Выберите клиента и архив',
                               clients=Config.CLIENTS,
                               selected_client=client_name,
                               product_name='')

    # Проверяем, что выбранный клиент есть в списке
    if client_name not in Config.CLIENTS:
        return render_template('index.html',
                               error='Выберите клиента из списка',
                               clients=Config.CLIENTS,
                               selected_client=client_name,
                               product_name='')

    # Проверяем, что файл является ZIP архивом
    if not archive_file.filename.lower().endswith('.zip'):
        return render_template('index.html',
                               error='Файл должен быть ZIP архивом',
                               clients=Config.CLIENTS,
                               selected_client=client_name,
                               product_name='')

    try:
        # Обрабатываем ZIP архив
        image_data = process_zip_archive(archive_file, client_name)

        if not image_data:
            return render_template('index.html',
                                   error='В архиве не найдено подходящих изображений',
                                   clients=Config.CLIENTS,
                                   selected_client=client_name,
                                   product_name='')

        return render_template('index.html',
                               image_urls=image_data,
                               clients=Config.CLIENTS,
                               selected_client=client_name,
                               product_name='')

    except Exception as e:
        return render_template('index.html',
                               error=f'Ошибка при обработке архива: {str(e)}',
                               clients=Config.CLIENTS,
                               selected_client=client_name,
                               product_name='')


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


@app.route('/download-xlsx', methods=['POST'])
def download_xlsx():
    try:
        data = request.get_json()
        if not data:
            return jsonify({'error': 'No data provided'}), 400

        image_data = data.get('image_data', [])
        client_name = data.get('client_name', '')

        if not image_data:
            return jsonify({'error': 'No image data provided'}), 400

        print(f"Получено данных для XLSX: {len(image_data)} элементов")
        print(f"Клиент: {client_name}")

        # Генерируем XLSX документ
        xlsx_buffer = generate_xlsx_document(image_data, client_name)

        # Создаем имя файла
        filename = f"{safe_folder_name(client_name)}_images.xlsx"

        # Сохраняем файл временно
        temp_file = f"temp_xlsx_{uuid.uuid4().hex}.xlsx"
        with open(temp_file, 'wb') as f:
            f.write(xlsx_buffer.getvalue())

        response = send_file(
            temp_file,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

        # Удаляем временный файл после отправки
        @response.call_on_close
        def remove_file():
            try:
                os.remove(temp_file)
            except:
                pass
        return response

    except Exception as e:
        app.logger.error(f"Error generating XLSX: {str(e)}")
        return jsonify({'error': f'Ошибка при генерации XLSX-файла: {str(e)}'}), 500


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
