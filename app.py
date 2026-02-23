import json
import os
import re
from datetime import datetime
from flask import Flask, render_template, request, redirect, url_for, flash, send_file
from openpyxl import Workbook
from io import BytesIO
import yadisk
from dotenv import load_dotenv

from export_utils import create_excel_report

# Загружаем переменные окружения
load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv('SECRET_KEY', 'default-secret-key-change-me')

# Фильтр для форматирования даты в шаблонах
@app.template_filter('format_date')
def format_date_filter(iso_date):
    if iso_date and re.match(r'^\d{4}-\d{2}-\d{2}$', iso_date):
        year, month, day = iso_date.split('-')
        return f"{day}.{month}.{year}"
    return iso_date

# Константы из окружения
YANDEX_TOKEN = os.getenv('YANDEX_TOKEN')
REMOTE_PATH = os.getenv('REMOTE_PATH', '/remedies_journal.json')
LOCAL_FALLBACK = os.getenv('LOCAL_FALLBACK', 'data.json')

# --- Работа с Яндекс.Диском ---
def ensure_remote_dir(y, remote_path):
    """Создаёт директорию на Яндекс.Диске, если она не существует."""
    if '/' in remote_path:
        remote_dir = '/'.join(remote_path.split('/')[:-1])
        if remote_dir and not remote_dir.startswith('/'):
            remote_dir = '/' + remote_dir
        if remote_dir and remote_dir != '/':
            try:
                if not y.exists(remote_dir):
                    y.mkdir(remote_dir)
            except Exception:
                pass

def read_data():
    """Читает данные с Яндекс.Диска, при ошибке - из локального файла."""
    if YANDEX_TOKEN:
        try:
            y = yadisk.YaDisk(token=YANDEX_TOKEN)
            if y.exists(REMOTE_PATH):
                buf = BytesIO()
                y.download(REMOTE_PATH, buf)
                buf.seek(0)
                json_str = buf.getvalue().decode('utf-8')
                return json.loads(json_str)
            else:
                return []
        except Exception:
            if os.path.exists(LOCAL_FALLBACK):
                with open(LOCAL_FALLBACK, 'r', encoding='utf-8') as f:
                    try:
                        return json.load(f)
                    except json.JSONDecodeError:
                        return []
            return []
    else:
        if os.path.exists(LOCAL_FALLBACK):
            with open(LOCAL_FALLBACK, 'r', encoding='utf-8') as f:
                try:
                    return json.load(f)
                except json.JSONDecodeError:
                    return []
        return []

def write_data(data):
    """Записывает данные на Яндекс.Диск, при ошибке сохраняет локально."""
    if YANDEX_TOKEN:
        try:
            y = yadisk.YaDisk(token=YANDEX_TOKEN)
            ensure_remote_dir(y, REMOTE_PATH)
            json_str = json.dumps(data, ensure_ascii=False, indent=2)
            buf = BytesIO(json_str.encode('utf-8'))
            y.upload(buf, REMOTE_PATH, overwrite=True)
            return True
        except Exception:
            with open(LOCAL_FALLBACK, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            return False
    else:
        with open(LOCAL_FALLBACK, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return False

def generate_id():
    return datetime.now().strftime('%Y%m%d%H%M%S%f')

# --- Функции для работы с датой и временем ---
def parse_date(date_str):
    """Преобразует дату из форматов ДД.ММ.ГГГГ, ГГГГ-ММ-ДД или ДДММГГГГ в ISO."""
    date_str = date_str.strip()
    if not date_str:
        return ''
    # ДД.ММ.ГГГГ
    match = re.match(r'^(\d{2})\.(\d{2})\.(\d{4})$', date_str)
    if match:
        day, month, year = match.groups()
        return f"{year}-{month}-{day}"
    # ГГГГ-ММ-ДД
    match = re.match(r'^(\d{4})-(\d{2})-(\d{2})$', date_str)
    if match:
        return date_str
    # ДДММГГГГ (8 цифр)
    match = re.match(r'^(\d{2})(\d{2})(\d{4})$', date_str)
    if match:
        day, month, year = match.groups()
        return f"{year}-{month}-{day}"
    return ''

def format_date_for_display(iso_date):
    """Преобразует ISO дату в ДД.ММ.ГГГГ."""
    if iso_date and re.match(r'^\d{4}-\d{2}-\d{2}$', iso_date):
        year, month, day = iso_date.split('-')
        return f"{day}.{month}.{year}"
    return iso_date

def parse_time(time_str):
    """Преобразует время из форматов ЧЧ:ММ или ЧЧММ в ЧЧ:ММ."""
    time_str = time_str.strip()
    if not time_str:
        return ''
    # ЧЧ:ММ
    if re.match(r'^([0-1]?[0-9]|2[0-3]):[0-5][0-9]$', time_str):
        parts = time_str.split(':')
        return f"{int(parts[0]):02d}:{parts[1]}"
    # ЧЧММ (4 цифры)
    match = re.match(r'^([0-1][0-9]|2[0-3])([0-5][0-9])$', time_str)
    if match:
        return f"{match.group(1)}:{match.group(2)}"
    return ''

def validate_time(time_str):
    """Проверяет корректность времени (ЧЧ:ММ)."""
    return bool(re.match(r'^([0-1][0-9]|2[0-3]):[0-5][0-9]$', time_str))

# --- Маршруты ---
@app.route('/')
def index():
    records = read_data()
    records.sort(key=lambda x: x.get('date-time', ''), reverse=True)
    return render_template('index.html', records=records)

@app.route('/add', methods=['GET', 'POST'])
def add():
    if request.method == 'POST':
        date_raw = request.form['date']
        time_raw = request.form['time']
        remedy = request.form['remedy']
        potency = request.form['potency']

        # Дата обязательна
        date_iso = parse_date(date_raw)
        if not date_iso:
            flash('Неверный формат даты. Используйте ДД.ММ.ГГГГ или ДДММГГГГ (8 цифр)', 'danger')
            return render_template('add_edit.html', record=None,
                                   date_display=date_raw, time=time_raw,
                                   remedy=remedy, potency=potency)

        # Время опционально
        time_fixed = ''
        if time_raw:
            time_fixed = parse_time(time_raw)
            if not time_fixed:
                flash('Неверный формат времени. Используйте ЧЧ:ММ или ЧЧММ (4 цифры)', 'danger')
                return render_template('add_edit.html', record=None,
                                       date_display=date_raw, time=time_raw,
                                       remedy=remedy, potency=potency)

        date_time = f"{date_iso} {time_fixed}" if time_fixed else date_iso

        new_record = {
            'id': generate_id(),
            'date-time': date_time,
            'remedy': remedy,
            'potency': potency,
            'events': []
        }
        records = read_data()
        records.append(new_record)
        write_data(records)
        flash('Запись добавлена', 'success')
        return redirect(url_for('index'))

    return render_template('add_edit.html', record=None)

@app.route('/edit/<record_id>', methods=['GET', 'POST'])
def edit(record_id):
    records = read_data()
    record = next((r for r in records if r['id'] == record_id), None)
    if not record:
        flash('Запись не найдена', 'danger')
        return redirect(url_for('index'))

    if request.method == 'POST':
        date_raw = request.form['date']
        time_raw = request.form['time']
        remedy = request.form['remedy']
        potency = request.form['potency']

        date_iso = parse_date(date_raw)
        if not date_iso:
            flash('Неверный формат даты. Используйте ДД.ММ.ГГГГ или ДДММГГГГ (8 цифр)', 'danger')
            return render_template('add_edit.html', record=record,
                                   date_display=date_raw, time=time_raw,
                                   remedy=remedy, potency=potency)

        time_fixed = ''
        if time_raw:
            time_fixed = parse_time(time_raw)
            if not time_fixed:
                flash('Неверный формат времени. Используйте ЧЧ:ММ или ЧЧММ (4 цифры)', 'danger')
                return render_template('add_edit.html', record=record,
                                       date_display=date_raw, time=time_raw,
                                       remedy=remedy, potency=potency)

        record['date-time'] = f"{date_iso} {time_fixed}" if time_fixed else date_iso
        record['remedy'] = remedy
        record['potency'] = potency
        write_data(records)
        flash('Запись обновлена', 'success')
        return redirect(url_for('index'))

    # Разделяем date-time на дату и время для отображения
    dt = record.get('date-time', '')
    parts = dt.split(' ')
    date_part = parts[0] if len(parts) > 0 else ''
    time_part = parts[1] if len(parts) > 1 else ''
    date_display = format_date_for_display(date_part)
    return render_template('add_edit.html', record=record,
                           date_display=date_display, time=time_part,
                           remedy=record.get('remedy', ''),
                           potency=record.get('potency', ''))

@app.route('/delete/<record_id>', methods=['POST'])
def delete(record_id):
    records = read_data()
    records = [r for r in records if r['id'] != record_id]
    write_data(records)
    flash('Запись удалена', 'success')
    return redirect(url_for('index'))

# --- Управление событиями (с датой) ---
@app.route('/add_event/<record_id>', methods=['POST'])
def add_event(record_id):
    records = read_data()
    record = next((r for r in records if r['id'] == record_id), None)
    if record:
        date_raw = request.form.get('event_date', '')
        time_raw = request.form.get('event_time', '')
        description = request.form.get('description', '')

        if not description:
            flash('Описание события обязательно', 'danger')
            return redirect(url_for('index'))

        # Парсим дату, если она есть
        date_iso = ''
        if date_raw:
            date_iso = parse_date(date_raw)
            if not date_iso:
                flash('Неверный формат даты события. Используйте ДД.ММ.ГГГГ или ДДММГГГГ (8 цифр)', 'danger')
                return redirect(url_for('index'))

        # Парсим время, если оно есть
        time_fixed = ''
        if time_raw:
            time_fixed = parse_time(time_raw)
            if not time_fixed:
                flash('Неверный формат времени события. Используйте ЧЧ:ММ или ЧЧММ (4 цифры)', 'danger')
                return redirect(url_for('index'))

        # Создаём событие (поля date и time опциональны)
        new_event = {
            'date': date_iso,
            'time': time_fixed,
            'description': description
        }
        if 'events' not in record:
            record['events'] = []
        record['events'].append(new_event)
        write_data(records)
        flash('Событие добавлено', 'success')
    return redirect(url_for('index'))

@app.route('/delete_event/<record_id>/<int:event_index>', methods=['POST'])
def delete_event(record_id, event_index):
    records = read_data()
    record = next((r for r in records if r['id'] == record_id), None)
    if record and 'events' in record and 0 <= event_index < len(record['events']):
        del record['events'][event_index]
        write_data(records)
        flash('Событие удалено', 'success')
    return redirect(url_for('index'))

@app.route('/edit_event/<record_id>/<int:event_index>', methods=['GET', 'POST'])
def edit_event(record_id, event_index):
    records = read_data()
    record = next((r for r in records if r['id'] == record_id), None)
    if not record or 'events' not in record or event_index >= len(record['events']):
        flash('Событие не найдено', 'danger')
        return redirect(url_for('index'))

    event = record['events'][event_index]

    if request.method == 'POST':
        date_raw = request.form.get('event_date', '')
        time_raw = request.form.get('event_time', '')
        description = request.form.get('description', '')

        if not description:
            flash('Описание события обязательно', 'danger')
            return render_template('edit_event.html',
                                   record_id=record_id,
                                   event_index=event_index,
                                   event={'date': date_raw, 'time': time_raw, 'description': description})

        date_iso = ''
        if date_raw:
            date_iso = parse_date(date_raw)
            if not date_iso:
                flash('Неверный формат даты события', 'danger')
                return render_template('edit_event.html',
                                       record_id=record_id,
                                       event_index=event_index,
                                       event={'date': date_raw, 'time': time_raw, 'description': description})

        time_fixed = ''
        if time_raw:
            time_fixed = parse_time(time_raw)
            if not time_fixed:
                flash('Неверный формат времени события', 'danger')
                return render_template('edit_event.html',
                                       record_id=record_id,
                                       event_index=event_index,
                                       event={'date': date_raw, 'time': time_raw, 'description': description})

        # Обновляем событие
        event['date'] = date_iso
        event['time'] = time_fixed
        event['description'] = description
        write_data(records)
        flash('Событие обновлено', 'success')
        return redirect(url_for('index'))

    # GET: подготавливаем данные для формы
    # Преобразуем дату события для отображения (если есть)
    display_date = format_date_for_display(event.get('date', '')) if event.get('date') else ''
    return render_template('edit_event.html',
                           record_id=record_id,
                           event_index=event_index,
                           event_date=display_date,
                           event_time=event.get('time', ''),
                           event_description=event.get('description', ''))

# --- Экспорт в Excel ---
@app.route('/export')
def export():
    records = read_data()
    excel_data = create_excel_report(records)
    return send_file(excel_data, download_name='remedies_journal.xlsx', as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True, port=5009)
