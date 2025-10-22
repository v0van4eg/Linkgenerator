from flask import Flask, render_template, request, jsonify, send_file
import os
import uuid
from werkzeug.utils import secure_filename
from config import Config, allowed_file

app = Flask(__name__)
app.config.from_object(Config)

# Создаем директорию для загрузок если не существует
os.makedirs(Config.UPLOAD_FOLDER, exist_ok=True)


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Получаем данные из формы
        client_name = request.form.get('client_name', '').strip()
        product_name = request.form.get('product_name', '').strip()

        if not client_name or not product_name:
            return render_template('index.html', error='Заполните все поля')

        # Создаем пути для клиента и товара
        client_folder = secure_filename(client_name)
        product_folder = secure_filename(product_name)
        full_path = os.path.join(Config.UPLOAD_FOLDER, client_folder, product_folder)
        os.makedirs(full_path, exist_ok=True)

        uploaded_files = request.files.getlist('images')
        image_urls = []

        for file in uploaded_files:
            if file and allowed_file(file.filename):
                # Генерируем уникальное имя файла
                file_extension = os.path.splitext(file.filename)[1]
                unique_filename = f"{uuid.uuid4().hex}{file_extension}"

                # Сохраняем файл
                file_path = os.path.join(full_path, unique_filename)
                file.save(file_path)

                # Генерируем URL
                image_url = f"{Config.BASE_URL}/images/{client_folder}/{product_folder}/{unique_filename}"
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
