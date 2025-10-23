from flask import Flask, render_template, request, jsonify, send_file
import os
import uuid
from werkzeug.utils import secure_filename
from config import Config, allowed_file
import re
import unicodedata
from urllib.parse import quote

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
                # image_url = f"{Config.BASE_URL}/images/{client_folder}/{product_folder}/{unique_filename}"
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





if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
