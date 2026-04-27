"""checkr_gui.py — Web-GUI для автоматичної валідації товарного фіду e-commerce.

Веб-інтерфейс користувача для програми Checkr.
Дозволяє завантажити файл через браузер та зберігає результат поруч із вихідним файлом.

Використання:
    python checkr_gui.py
    
Потім відкрийте браузер за адресою: http://localhost:5000
"""

import sys
import os
import io
import tempfile
import re
import time
import webbrowser
import threading
from pathlib import Path
from contextlib import redirect_stdout

try:
    from flask import Flask, render_template_string, request, jsonify, send_file
    from werkzeug.utils import secure_filename
    from werkzeug.exceptions import RequestEntityTooLarge
except ImportError:
    print("Помилка: не вдалося імпортувати Flask або werkzeug.", file=sys.stderr)
    print("Встановіть їх командою: pip install Flask", file=sys.stderr)
    sys.exit(1)

# Імпортуємо основну функцію валідації з checkr.py
try:
    from checkr import validate_feed
except ImportError:
    print("Помилка: не вдалося імпортувати модуль checkr.py", file=sys.stderr)
    sys.exit(1)


app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max file size
# Використовуємо системну тимчасову директорію для кросплатформеності
app.config['UPLOAD_FOLDER'] = os.path.join(tempfile.gettempdir(), 'checkr_uploads')

# Створюємо директорію для завантажень, якщо вона не існує
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)


@app.errorhandler(RequestEntityTooLarge)
def handle_file_too_large(e):
    """Обробка помилки перевищення розміру файлу."""
    return jsonify({
        'success': False,
        'error': 'Файл завеликий. Максимальний розмір: 50MB'
    }), 413


# HTML шаблон для інтерфейсу
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="uk">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Checkr — Валідація товарного фіду</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            display: flex;
            justify-content: center;
            align-items: center;
            padding: 20px;
        }
        .container {
            background: white;
            border-radius: 15px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            padding: 40px;
            max-width: 700px;
            width: 100%;
        }
        h1 {
            color: #333;
            text-align: center;
            margin-bottom: 10px;
            font-size: 28px;
        }
        .subtitle {
            color: #666;
            text-align: center;
            margin-bottom: 30px;
            font-size: 14px;
            line-height: 1.6;
        }
        .upload-area {
            border: 3px dashed #667eea;
            border-radius: 10px;
            padding: 40px;
            text-align: center;
            margin-bottom: 20px;
            cursor: pointer;
            transition: all 0.3s;
        }
        .upload-area:hover {
            border-color: #764ba2;
            background: #f8f9ff;
        }
        .upload-area.dragover {
            border-color: #764ba2;
            background: #f0f0ff;
        }
        #fileInput {
            display: none;
        }
        .file-label {
            color: #667eea;
            font-size: 16px;
            font-weight: 600;
            cursor: pointer;
        }
        .file-name {
            margin-top: 15px;
            color: #333;
            font-weight: 500;
        }
        .btn {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border: none;
            padding: 15px 40px;
            font-size: 16px;
            font-weight: 600;
            border-radius: 25px;
            cursor: pointer;
            width: 100%;
            transition: transform 0.2s;
            margin-top: 20px;
        }
        .btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 10px 20px rgba(0,0,0,0.2);
        }
        .btn:disabled {
            background: #ccc;
            cursor: not-allowed;
            transform: none;
        }
        .progress {
            display: none;
            margin-top: 20px;
        }
        .progress-bar {
            width: 100%;
            height: 30px;
            background: #f0f0f0;
            border-radius: 15px;
            overflow: hidden;
        }
        .progress-fill {
            height: 100%;
            background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
            width: 0%;
            transition: width 0.3s;
            display: flex;
            align-items: center;
            justify-content: center;
            color: white;
            font-weight: 600;
        }
        .output {
            display: none;
            margin-top: 20px;
            padding: 20px;
            background: #f8f9fa;
            border-radius: 10px;
            max-height: 300px;
            overflow-y: auto;
        }
        .output pre {
            margin: 0;
            font-family: 'Courier New', monospace;
            font-size: 12px;
            white-space: pre-wrap;
            word-wrap: break-word;
        }
        .success {
            color: #28a745;
            font-weight: 600;
        }
        .error {
            color: #dc3545;
            font-weight: 600;
        }
        .download-link {
            display: none;
            margin-top: 20px;
            text-align: center;
        }
        .download-link a {
            color: #667eea;
            text-decoration: none;
            font-weight: 600;
            font-size: 16px;
        }
        .download-link a:hover {
            text-decoration: underline;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>🔍 Checkr</h1>
        <p class="subtitle">
            Автоматична валідація товарного фіду e-commerce<br>
            Перевірка логічних розбіжностей між назвою, описами та характеристиками товарів
        </p>
        
        <div class="upload-area" id="uploadArea">
            <label for="fileInput" class="file-label">
                📁 Клацніть або перетягніть файл сюди<br>
                <small style="color: #999; font-weight: normal;">Підтримуються формати: CSV, XLSX, XLS</small>
            </label>
            <input type="file" id="fileInput" accept=".csv,.xlsx,.xls">
            <div class="file-name" id="fileName"></div>
        </div>
        
        <button class="btn" id="checkBtn" disabled>Перевірити файл</button>
        
        <div class="progress" id="progress">
            <div class="progress-bar">
                <div class="progress-fill" id="progressFill">0%</div>
            </div>
        </div>
        
        <div class="output" id="output">
            <pre id="outputText"></pre>
        </div>
        
        <div class="download-link" id="downloadLink"></div>
    </div>

    <script>
        const fileInput = document.getElementById('fileInput');
        const uploadArea = document.getElementById('uploadArea');
        const fileName = document.getElementById('fileName');
        const checkBtn = document.getElementById('checkBtn');
        const progress = document.getElementById('progress');
        const progressFill = document.getElementById('progressFill');
        const output = document.getElementById('output');
        const outputText = document.getElementById('outputText');
        const downloadLink = document.getElementById('downloadLink');
        
        let selectedFile = null;
        
        // Обробка вибору файлу
        fileInput.addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                selectedFile = e.target.files[0];
                fileName.textContent = `Обрано: ${selectedFile.name}`;
                checkBtn.disabled = false;
            }
        });
        
        // Drag and drop
        uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            uploadArea.classList.add('dragover');
        });
        
        uploadArea.addEventListener('dragleave', () => {
            uploadArea.classList.remove('dragover');
        });
        
        uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadArea.classList.remove('dragover');
            
            if (e.dataTransfer.files.length > 0) {
                selectedFile = e.dataTransfer.files[0];
                fileName.textContent = `Обрано: ${selectedFile.name}`;
                checkBtn.disabled = false;
                fileInput.files = e.dataTransfer.files;
            }
        });
        
        // Обробка перевірки файлу
        checkBtn.addEventListener('click', async () => {
            if (!selectedFile) return;
            
            checkBtn.disabled = true;
            progress.style.display = 'block';
            output.style.display = 'none';
            downloadLink.style.display = 'none';
            outputText.textContent = '';
            
            // Анімація прогресу
            let progressValue = 0;
            const progressInterval = setInterval(() => {
                progressValue = Math.min(progressValue + 2, 90);
                progressFill.style.width = progressValue + '%';
                progressFill.textContent = progressValue + '%';
            }, 100);
            
            const formData = new FormData();
            formData.append('file', selectedFile);
            
            try {
                const response = await fetch('/validate', {
                    method: 'POST',
                    body: formData
                });
                
                clearInterval(progressInterval);
                progressFill.style.width = '100%';
                progressFill.textContent = '100%';
                
                const result = await response.json();
                
                output.style.display = 'block';
                
                if (result.success) {
                    outputText.innerHTML = `<span class="success">✅ Перевірка завершена успішно!</span>\n\n${result.output}\n\n<span class="success">📁 Результат збережено: ${result.output_file}</span>`;
                    downloadLink.innerHTML = `<a href="/download/${result.download_id}" download>⬇️ Завантажити результат</a>`;
                    downloadLink.style.display = 'block';
                } else {
                    outputText.innerHTML = `<span class="error">❌ Помилка: ${result.error}</span>`;
                }
            } catch (error) {
                clearInterval(progressInterval);
                output.style.display = 'block';
                outputText.innerHTML = `<span class="error">❌ Помилка з'єднання: ${error.message}</span>`;
            }
            
            setTimeout(() => {
                progress.style.display = 'none';
                progressFill.style.width = '0%';
                checkBtn.disabled = false;
            }, 1000);
        });
    </script>
</body>
</html>
"""


# Словник для зберігання шляхів до результатів
results_storage = {}


@app.route('/')
def index():
    """Головна сторінка."""
    return render_template_string(HTML_TEMPLATE)


@app.route('/validate', methods=['POST'])
def validate():
    """Обробка завантаженого файлу та валідація."""
    if 'file' not in request.files:
        return jsonify({'success': False, 'error': 'Файл не знайдено'})
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'success': False, 'error': 'Файл не обрано'})
    
    # Перевіряємо розширення файлу
    allowed_extensions = {'.csv', '.xlsx', '.xls'}
    file_ext = Path(file.filename).suffix.lower()
    if file_ext not in allowed_extensions:
        return jsonify({
            'success': False, 
            'error': f'Непідтримуваний формат файлу. Підтримуються: {", ".join(allowed_extensions)}'
        })
    
    try:
        # Зберігаємо оригінальне розширення перед очищенням імені файлу
        original_ext = Path(file.filename).suffix.lower()
        
        # Очищаємо ім'я файлу для безпеки
        safe_filename_base = secure_filename(file.filename)
        if not safe_filename_base:
            return jsonify({'success': False, 'error': 'Некоректне ім\'я файлу'})
        
        # Додаємо timestamp для унікальності імені файлу та запобігання перезапису
        timestamp = str(int(time.time() * 1000))
        # Використовуємо оригінальне розширення, якщо secure_filename його видалив
        safe_ext = Path(safe_filename_base).suffix or original_ext
        safe_stem = Path(safe_filename_base).stem
        unique_filename = f"{safe_stem}_{timestamp}{safe_ext}"
        
        # Зберігаємо завантажений файл
        input_path = Path(app.config['UPLOAD_FOLDER']) / unique_filename
        file.save(str(input_path))
        
        # Створюємо ім'я вихідного файлу
        output_file = input_path.parent / f"{input_path.stem}_result.xlsx"
        
        # Перенаправляємо stdout у буфер
        output_buffer = io.StringIO()
        with redirect_stdout(output_buffer):
            validate_feed(str(input_path), str(output_file))
        
        # Отримуємо вивід
        output = output_buffer.getvalue()
        
        # Створюємо безпечний ID для завантаження (лише букви, цифри, підкреслення)
        download_id = re.sub(r'[^\w\-]', '_', input_path.stem) + "_result"
        results_storage[download_id] = str(output_file)
        
        return jsonify({
            'success': True,
            'output': output,
            'output_file': output_file.name,
            'download_id': download_id
        })
        
    except Exception as exc:
        # Не викриваємо деталі помилки користувачу для безпеки
        import logging
        logging.error(f"Validation error: {exc}", exc_info=True)
        return jsonify({'success': False, 'error': 'Виникла помилка під час обробки файлу'})


@app.route('/download/<download_id>')
def download(download_id):
    """Завантаження результату."""
    # Валідуємо download_id - дозволяємо тільки букви, цифри, підкреслення та дефіси
    if not re.match(r'^[\w\-]+$', download_id):
        return "Некоректний ID файлу", 400
    
    if download_id not in results_storage:
        return "Файл не знайдено", 404
    
    file_path = results_storage[download_id]
    
    # Переконуємось, що шлях знаходиться в межах upload folder
    try:
        resolved_path = Path(file_path).resolve()
        upload_folder = Path(app.config['UPLOAD_FOLDER']).resolve()
        if not resolved_path.is_relative_to(upload_folder):
            return "Доступ заборонено", 403
    except (ValueError, OSError):
        return "Некоректний шлях", 400
    
    # Видаляємо запис зі словника після першого завантаження
    # Примітка: файли залишаються на диску для можливого повторного використання
    # але видаляються зі словника для обмеження споживання пам'яті
    del results_storage[download_id]
    
    return send_file(file_path, as_attachment=True)


def main():
    """Точка входу web-GUI програми."""
    print("=" * 70)
    print("🔍 Checkr — Валідація товарного фіду e-commerce")
    print("=" * 70)
    print("\n✅ Сервер запущено!")
    print("🌐 Відкриваємо браузер за адресою: http://localhost:5000")
    print("\n💡 Для зупинки сервера натисніть Ctrl+C")
    print("⚠️  УВАГА: Максимальний розмір файлу - 50MB\n")
    print("=" * 70 + "\n")
    
    # Відкриваємо браузер автоматично після короткої затримки
    def open_browser():
        time.sleep(1.5)  # Даємо серверу час запуститися
        webbrowser.open('http://localhost:5000')
    
    browser_thread = threading.Thread(target=open_browser)
    browser_thread.daemon = True
    browser_thread.start()
    
    # Прив'язуємося до localhost (127.0.0.1) для безпеки
    app.run(host='127.0.0.1', port=5000, debug=False)


if __name__ == "__main__":
    main()
