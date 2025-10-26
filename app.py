from flask import Flask, render_template, request, jsonify, send_file, redirect, url_for, session
import os
import uuid
import zipfile
from config import Config, allowed_file
import re
import unicodedata
from urllib.parse import quote
import tempfile
import shutil
import json
from datetime import datetime

# Импортируем фабрику генераторов
from generators import GeneratorFactory

app = Flask(__name__)
app.config.from_object(Config)

RESULTS_FOLDER = 'results'

# Инициализируем папку для результатов
if not os.path.exists(RESULTS_FOLDER):
    os.makedirs(RESULTS_FOLDER)
    print(f"Папка для результатов создана: {RESULTS_FOLDER}")

def safe_folder_name(name: str) -> str:
    """Преобразует строку в безопасное имя папки"""
    if not name:
        return "unnamed"
    name = unicodedata.normalize('NFKD', name)
    name = re.sub(r'[^\w\s-]', '', name, flags=re.UNICODE)
    name = re.sub(r'[-\s]+', '-', name, flags=re.UNICODE).strip('-_')
    return name[:255] if name else "unnamed"


def process_zip_archive(zip_file, client_name):
    """Обрабатывает ZIP-архив и извлекает изображения"""
    image_urls = []
    with tempfile.TemporaryDirectory() as temp_dir:
        zip_path = os.path.join(temp_dir, zip_file.filename)
        zip_file.save(zip_path)

        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        for root, dirs, files in os.walk(temp_dir):
            for file in files:
                if file.lower() in ['thumbs.db', '.ds_store']:
                    continue
                if not allowed_file(file):
                    continue

                relative_path = os.path.relpath(root, temp_dir)
                if relative_path == '.':
                    continue

                article = os.path.basename(root) if relative_path.count(os.sep) == 0 else relative_path.split(os.sep)[0]

                client_folder = safe_folder_name(client_name)
                article_folder = safe_folder_name(article)
                full_path = os.path.join(Config.UPLOAD_FOLDER, client_folder, article_folder)
                os.makedirs(full_path, exist_ok=True)

                file_extension = os.path.splitext(file)[1]
                file_name_base = os.path.splitext(file)[0]
                unique_filename = f"{file_name_base}_{uuid.uuid4().hex[:6]}{file_extension}"
                target_file_unique = os.path.join(full_path, unique_filename)

                source_file = os.path.join(root, file)
                shutil.copy2(source_file, target_file_unique)

                image_url = "{}/images/{}/{}/{}".format(
                    Config.BASE_URL,
                    quote(client_folder, safe=''),
                    quote(article_folder, safe=''),
                    quote(unique_filename, safe='')
                )
                image_urls.append({
                    'url': image_url,
                    'article': article,
                    'filename': unique_filename
                })
    return image_urls


def generate_xlsx_document(image_data, client_name):
    """Генерирует XLSX документ используя фабрику генераторов"""
    generator = GeneratorFactory.create_generator(client_name)
    return generator.generate(image_data, client_name)


# Остальные функции (save_results_to_file, load_results_from_file, маршруты)
# остаются без изменений, как в вашем исходном коде

def save_results_to_file(image_data, client_name, product_name=None):
    """Сохраняет результаты обработки в JSON-файл"""
    result_id = uuid.uuid4().hex
    results_data = {
        'image_data': image_data,
        'client_name': client_name,
        'product_name': product_name or '',
        'timestamp': datetime.now().isoformat()
    }
    filename = f"results_{result_id}.json"
    filepath = os.path.join(Config.RESULTS_FOLDER, filename)
    with open(filepath, 'w', encoding='utf-8') as f:
        json.dump(results_data, f, ensure_ascii=False, indent=4)
    return result_id


def load_results_from_file(result_id):
    """Загружает результаты из JSON-файла"""
    filename = f"results_{result_id}.json"
    filepath = os.path.join(Config.RESULTS_FOLDER, filename)
    if os.path.exists(filepath):
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                data = json.load(f)
            if 'image_data' in data and 'client_name' in data:
                return data
        except (json.JSONDecodeError, IOError) as e:
            print(f"Ошибка чтения файла {filepath}: {e}")
    return None


# Маршруты Flask остаются без изменений
@app.route('/admin', methods=['GET'])
def index():
    return render_template('index.html',
                           clients=Config.CLIENTS,
                           selected_client='',
                           product_name='',
                           image_urls=[],
                           error='')


@app.route('/admin', methods=['POST'])
def handle_upload():
    if 'archive' in request.files and request.files['archive'].filename != '':
        result_id, error = handle_archive_upload_logic(request)
    else:
        result_id, error = handle_single_upload_logic(request)

    if error:
        session['error'] = error
        return redirect(url_for('index'))

    if result_id:
        return redirect(url_for('view_results', result_id=result_id))

    return redirect(url_for('index'))



@app.route('/admin/results/<result_id>', methods=['GET'])
def view_results(result_id):
    results_data = load_results_from_file(result_id)
    if results_data:
        image_urls = results_data.get('image_data', [])
        client_name = results_data.get('client_name', '')
        product_name = results_data.get('product_name', '')
        return render_template('index.html',
                               image_urls=image_urls,
                               clients=Config.CLIENTS,
                               selected_client=client_name,
                               product_name=product_name,
                               error='')
    else:
        error = 'Результаты не найдены или срок их действия истек.'
        return render_template('index.html',
                               image_urls=[],
                               clients=Config.CLIENTS,
                               selected_client='',
                               product_name='',
                               error=error)


# Новый маршрут для корня - отображает hello.html
@app.route('/')
def hello():
    return render_template('hello.html') # Убедитесь, что файл hello.html находится в папке templates


def handle_single_upload_logic(request):
    """Логика обработки отдельных изображений"""
    client_name = request.form.get('client_name', '').strip()
    product_name = request.form.get('product_name', '').strip()

    if not client_name or not product_name:
        return None, 'Заполните все поля'

    if client_name not in Config.CLIENTS:
        return None, 'Выберите клиента из списка'

    client_folder = safe_folder_name(client_name)
    product_folder = safe_folder_name(product_name)
    full_path = os.path.join(Config.UPLOAD_FOLDER, client_folder, product_folder)
    os.makedirs(full_path, exist_ok=True)

    uploaded_files = request.files.getlist('images')
    image_urls = []

    for file in uploaded_files:
        if file and allowed_file(file.filename):
            random_hex = uuid.uuid4().hex[:6]
            file_extension = os.path.splitext(file.filename)[1]
            file_name = os.path.splitext(file.filename)[0]
            unique_filename = f"{file_name}-{random_hex}{file_extension}"
            file_path = os.path.join(full_path, unique_filename)
            file.save(file_path)
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

    if not image_urls:
        return None, 'Не загружено ни одного подходящего изображения'

    result_id = save_results_to_file(image_urls, client_name, product_name)
    return result_id, None


def handle_archive_upload_logic(request):
    """Логика обработки ZIP архива"""
    client_name = request.form.get('client_name', '').strip()
    archive_file = request.files['archive']

    if not client_name or not archive_file or archive_file.filename == '':
        return None, 'Выберите клиента и архив'

    if client_name not in Config.CLIENTS:
        return None, 'Выберите клиента из списка'

    if not archive_file.filename.lower().endswith('.zip'):
        return None, 'Файл должен быть ZIP архивом'

    try:
        image_data = process_zip_archive(archive_file, client_name)
        if not image_data:
            return None, 'В архиве не найдено подходящих изображений'

        result_id = save_results_to_file(image_data, client_name)
        return result_id, None

    except Exception as e:
        return None, f'Ошибка при обработке архива: {str(e)}'


@app.route('/admin/download-links')
def download_links():
    urls = request.args.getlist('urls')
    if not urls:
        return "No URLs provided", 400
    temp_file = f"temp_links_{uuid.uuid4().hex}.txt"
    with open(temp_file, 'w', encoding='utf-8') as f:
        f.write('\n'.join(urls))
    return send_file(temp_file,
                     as_attachment=True,
                     download_name='image_links.txt',
                     mimetype='text/plain')


@app.route('/admin/download-xlsx', methods=['POST'])
def download_xlsx():
    try:
        data = request.get_json()
        if not data:
            return jsonify({'error': 'No data provided'}), 400

        image_data = data.get('image_data', [])
        client_name = data.get('client_name', '')
        if not image_data:
            return jsonify({'error': 'No image data provided'}), 400

        print(f"Генерация XLSX для клиента: {client_name}, элементов: {len(image_data)}")

        xlsx_buffer = generate_xlsx_document(image_data, client_name)
        filename = f"{safe_folder_name(client_name)}_images.xlsx"

        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx', mode='w+b') as temp_file:
            temp_file.write(xlsx_buffer.getvalue())
            temp_file_path = temp_file.name

        try:
            response = send_file(
                temp_file_path,
                as_attachment=True,
                download_name=filename,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
            return response
        finally:
            try:
                if os.path.exists(temp_file_path):
                    os.remove(temp_file_path)
            except Exception as e:
                print(f"Ошибка при удалении временного файла: {e}")

    except Exception as e:
        app.logger.error(f"Error generating XLSX: {str(e)}")
        return jsonify({'error': f'Ошибка при генерации XLSX-файла: {str(e)}'}), 500


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
