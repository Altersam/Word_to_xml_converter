# -*- coding: utf-8 -*-
"""
Moodle XML Converter — GUI (PyQt5)
Графический интерфейс для конвертера Word -> Moodle XML.
"""

import os
import sys
import re
import unicodedata
import traceback
from typing import List, Optional
from copy import deepcopy

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QGroupBox, QPushButton, QLabel, QLineEdit, QFileDialog,
    QComboBox, QCheckBox, QTextEdit, QProgressBar, QSplitter,
    QMessageBox, QTreeWidget, QTreeWidgetItem, QHeaderView,
    QAbstractItemView
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QColor, QFont, QBrush

# Подключаем конвертер
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from universal_moodle_converter_v3_stable import (
    VALID_MARKERS, MARKER_TO_QTYPE, MoodleConverter,
    QuestionTypeDetector,
)

try:
    import docxlatex
except ImportError:
    docxlatex = None

from lxml import etree


# ============================================================
# МАРКЕРЫ И ИХ ОПИСАНИЯ
# ============================================================

MARKER_DESCRIPTIONS = {
    '': '(авто)',
    'multichoice_one': 'Один правильный',
    'multichoice_many': 'Несколько правильных (-100%)',
    'shortanswer_phrase': 'Текстовый ввод',
    'numerical_partial': 'Цифры partial scoring',
    'numerical_numcombo': 'Цифры любой порядок',
    'matching': 'Соотношение L/R',
    'match_123': 'Последовательность',
    'match': 'Соотношение (=matching)',
    'ddmatch': 'Drag-and-drop',
    'gapselect': 'Выпадающие списки',
    'cloze': 'Встроенные ответы',
    'numerical': 'Числовой (. = ,)',
}

MARKER_COLORS = {
    'multichoice_one':      QColor(200, 230, 255),
    'multichoice_many':     QColor(180, 220, 255),
    'shortanswer_phrase':   QColor(200, 255, 200),
    'numerical_partial':    QColor(220, 255, 200),
    'numerical_numcombo':   QColor(240, 255, 200),
    'matching':             QColor(255, 230, 200),
    'match_123':            QColor(255, 220, 180),
    'match':                QColor(255, 230, 200),
    'ddmatch':              QColor(255, 210, 210),
    'gapselect':            QColor(230, 200, 255),
    'cloze':                QColor(255, 255, 200),
    'numerical':            QColor(200, 255, 255),
}

COLOR_CORRECT = QColor(0, 130, 0)
COLOR_WRONG   = QColor(180, 0, 0)
COLOR_TEXT    = QColor(50, 50, 50)
COLOR_META    = QColor(100, 100, 140)
COLOR_ERROR   = QColor(255, 220, 220)


# ============================================================
# ПАРСЕР DOCX ДЛЯ ПРЕДПРОСМОТРА
# ============================================================

class ParsedQuestion:
    """Один вопрос, извлечённый из docx."""
    __slots__ = ('name', 'grade', 'marker', 'subcategory',
                 'content', 'auto_type', 'errors', 'line_num', 'selected')

    def __init__(self):
        self.name: str = ''
        self.grade: float = 1.0
        self.marker: str = ''
        self.subcategory: str = ''
        self.content: List[str] = []
        self.auto_type: str = ''
        self.errors: List[str] = []
        self.line_num: int = 0
        self.selected: bool = True  # По умолчанию все вопросы выбраны


def parse_docx_preview(docx_path: str) -> tuple:
    """Парсит docx и возвращает (questions, global_errors)."""
    questions: List[ParsedQuestion] = []
    errors: List[str] = []

    if docxlatex is None:
        errors.append('docxlatex не установлен (pip install docxlatex)')
        return questions, errors

    try:
        doc = docxlatex.Document(docx_path)
        text = doc.get_text()
        lines = text.split('\n')
    except Exception as e:
        errors.append(f'Ошибка чтения файла: {e}')
        return questions, errors

    current_marker = ''
    current_subcategory = ''
    current_q: Optional[ParsedQuestion] = None

    for line_idx, line in enumerate(lines):
        line = unicodedata.normalize('NFC', line.strip())
        if not line:
            continue

        if line.startswith('V1:'):
            continue

        m = re.match(r'^\{(\w+)\}V2:(.*)', line)
        if m:
            current_marker = m.group(1)
            current_subcategory = m.group(2).strip()
            if current_marker not in VALID_MARKERS:
                errors.append(f'Строка {line_idx+1}: неизвестный маркер {{{current_marker}}}')
            continue

        if line.startswith('V2:'):
            current_subcategory = line[3:].strip()
            continue

        # Определяем начало вопроса (7 форматов)
        is_question = line.startswith('I:')
        if not is_question:
            if line.startswith('I I:'):
                line = 'I:' + line[4:]; is_question = True
            elif line.startswith('I ') and 'Задание' in line:
                line = 'I:' + line[2:]; is_question = True
            elif line.startswith(':Задание'):
                line = 'I:' + line[1:]; is_question = True
            elif line.startswith('Задание ') and ', b=' in line:
                if re.search(r'ТЗ\d+-\d+|Т\d+-\d+|T\d+-\d+', line):
                    line = 'I:' + line; is_question = True
            elif re.match(r'^[A-Za-z0-9_\-=]+Задание', line) and ', b=' in line:
                if re.search(r'ТЗ\d+-\d+|Т\d+-\d+|T\d+-\d+', line):
                    mm = re.match(r'^[A-Za-z0-9_\-=]+(Задание.*)', line)
                    if mm:
                        line = 'I:' + mm.group(1); is_question = True
            elif re.match(r'^[А-Яа-яёЁ]+\s+[А-Яа-яёЁ]\.[А-Яа-яёЁ]\.,', line) and ', b=' in line:
                tz = re.search(r'ТЗ(\d+)-(\d+)', line)
                if tz:
                    line = f'I:Задание {tz.group(2)}. {line}'; is_question = True

        if is_question:
            if current_q is not None:
                current_q.auto_type = QuestionTypeDetector.detect(
                    current_q.content, '', current_q.marker)
                questions.append(current_q)
            current_q = ParsedQuestion()
            current_q.name = line
            current_q.marker = current_marker
            current_q.subcategory = current_subcategory
            current_q.line_num = line_idx + 1
            gm = re.search(r'b=(\d+)', line)
            current_q.grade = float(gm.group(1)) if gm else 1.0
            continue

        if current_q is not None:
            is_ans = (line.startswith('+:') or line.startswith('-:') or
                      (len(line) > 1 and line[0] in '+-' and line[1] in ' \t'))
            if line.startswith(('S:', '->', '=>')) or is_ans or \
               not line.startswith(('V1:', 'V2:', 'I:')):
                current_q.content.append(line)

    if current_q is not None:
        current_q.auto_type = QuestionTypeDetector.detect(
            current_q.content, '', current_q.marker)
        questions.append(current_q)

    # Предобработка ошибок
    for q in questions:
        has_plus = any(c.startswith('+:') or c.startswith('+') for c in q.content)
        has_lr = any(re.match(r'^[LR]\d+:', c) for c in q.content)
        has_num = any(re.match(r'^\d+:', c) for c in q.content)
        skip = ('cloze', 'matching', 'match_123', 'match',
                'shortanswer_phrase', 'numerical_numcombo')
        if not has_plus and q.auto_type not in skip and not has_lr and not has_num:
            q.errors.append('Нет правильного ответа (+:)')
        if not q.content:
            q.errors.append('Нет текста вопроса')
        if not q.name:
            q.errors.append('Пустое имя вопроса')

    return questions, errors


# ============================================================
# ВАЛИДАТОР XML (ПОСТОБРАБОТКА)
# ============================================================

class XMLValidator:
    MOODLE_QTYPES = {
        'multichoice', 'shortanswer', 'matching', 'cloze',
        'ddmatch', 'gapselect', 'numerical', 'category',
    }

    @staticmethod
    def validate(xml_path: str) -> List[str]:
        issues: List[str] = []
        try:
            tree = etree.parse(xml_path)
            root = tree.getroot()
        except Exception as e:
            issues.append(f'XML не парсится: {e}')
            return issues

        if root.tag != 'quiz':
            issues.append(f'Корневой элемент: {root.tag} (ожидается quiz)')

        n = 0
        for q in root.findall('.//question'):
            qtype = q.get('type', '')
            if qtype == 'category':
                continue
            n += 1
            ne = q.find('name/text')
            qn = ne.text if ne is not None and ne.text else f'#{n}'

            if qtype not in XMLValidator.MOODLE_QTYPES:
                issues.append(f'{qn}: неизвестный тип "{qtype}"')

            qt = q.find('questiontext/text')
            if qt is None or not qt.text:
                issues.append(f'{qn}: пустой текст')
            else:
                txt = qt.text
                for pat in ['_IMAGE_', '@@PLUGINFILE@@']:
                    if pat in txt and not any(f.text for f in q.findall('.//file')):
                        issues.append(f'{qn}: маркер {pat} без файла')

            ans = q.findall('answer')
            if qtype in ('multichoice', 'shortanswer', 'numerical') and not ans:
                issues.append(f'{qn}: нет ответов')

            for fe in q.findall('.//file'):
                if fe.get('encoding') == 'base64' and len(fe.text or '') < 20:
                    issues.append(f'{qn}: пустая base64 картинка')

            if qtype == 'matching':
                sqs = q.findall('subquestion')
                if not sqs:
                    issues.append(f'{qn}: matching без subquestion')
                for sq in sqs:
                    if sq.find('answer') is None:
                        issues.append(f'{qn}: subquestion без answer')

            if qtype == 'ddmatch' and not q.findall('subquestion'):
                issues.append(f'{qn}: ddmatch без subquestion')

            if qtype == 'gapselect' and not q.findall('selectoption'):
                issues.append(f'{qn}: gapselect без selectoption')

        if n == 0:
            issues.append('Файл не содержит вопросов')
        return issues


# ============================================================
# РАЗДЕЛЕНИЕ XML НА ЧАСТИ
# ============================================================

def split_xml_by_size(xml_path: str, max_bytes: int = 1_000_000) -> List[str]:
    """Разделяет XML файл на части до max_bytes.
    
    Каждая часть начинается с дубликата категорий.
    """
    import re
    
    sz = os.path.getsize(xml_path)
    if sz <= max_bytes:
        return [xml_path]

    with open(xml_path, 'r', encoding='utf-8') as f:
        content = f.read()

    quiz_start = content.find('<quiz>')
    quiz_end = content.rfind('</quiz>')
    
    xml_decl = content[:quiz_start]
    inner = content[quiz_start + 6:quiz_end]
    
    # Извлекаем первые две категории для дублирования
    category_matches = list(re.finditer(r'(<question type="category">.*?</question>)', inner, re.DOTALL))
    categories = [m.group(1) for m in category_matches[:2]]
    
    # Все вопросы
    all_items = list(re.finditer(r'(<question[^>]*>.*?</question>)', inner, re.DOTALL))

    parts = []
    current_part = ""
    current_size = 0
    part_num = 1

    for match in all_items:
        item = match.group(1)
        
        # Пропускаем категории (они уже сохранены)
        if '<question type="category">' in item:
            continue
        
        item_size = len(item.encode('utf-8'))
        
        if current_size + item_size > max_bytes and current_size > 0:
            parts.append(current_part)
            part_num += 1
            current_part = item
            current_size = item_size
        else:
            current_part += item
            current_size += item_size

    if current_part:
        parts.append(current_part)

    base, ext = os.path.splitext(xml_path)
    result_parts = []
    
    for i, part_content in enumerate(parts, 1):
        full_xml = f"{xml_decl}<quiz>{''.join(categories)}\n{part_content}</quiz>"
        pp = f'{base}_part{i}{ext}'
        with open(pp, 'w', encoding='utf-8') as out:
            out.write(full_xml)
        result_parts.append(pp)

    return result_parts


# ============================================================
# РАБОЧИЙ ПОТОК
# ============================================================

class ConvertWorker(QThread):
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(bool, str)
    validation = pyqtSignal(list)

    def __init__(self, docx_path, output_path, questions, split_xml, selected_indices=None, parent=None):
        super().__init__(parent)
        self.docx_path = docx_path
        self.output_path = output_path
        self.questions = questions
        self.split_xml = split_xml
        self.selected_indices = selected_indices

    def run(self):
        try:
            self.progress.emit(10, 'Запуск конвертации...')
            conv = MoodleConverter(self.docx_path, self.output_path, self.selected_indices)
            conv._marker_overrides = {}
            for q in self.questions:
                if q.marker and getattr(q, 'selected', True):
                    conv._marker_overrides[q.name] = q.marker

            self.progress.emit(30, 'Конвертация...')
            ok = conv.convert()
            if not ok:
                self.finished.emit(False, '\n'.join(conv.errors))
                return

            self.progress.emit(70, 'Валидация XML...')
            issues = XMLValidator.validate(self.output_path)
            self.validation.emit(issues)

            if self.split_xml:
                self.progress.emit(85, 'Разделение на части...')
                parts = split_xml_by_size(self.output_path)
                if len(parts) > 1:
                    self.progress.emit(95, f'Разделено на {len(parts)} частей')
                else:
                    self.progress.emit(95, 'Файл < 1МБ, разделение не нужно')

            self.progress.emit(100, 'Готово!')
            self.finished.emit(True, self.output_path)
        except Exception as e:
            self.finished.emit(False, f'Ошибка: {e}\n{traceback.format_exc()}')


# ============================================================
# ГЛАВНОЕ ОКНО
# ============================================================

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Moodle XML Converter')
        self.setMinimumSize(1100, 780)
        self.questions: List[ParsedQuestion] = []
        self.worker = None
        self._init_ui()

    # ---------- UI ----------
    def _init_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        lay = QVBoxLayout(central)
        lay.setContentsMargins(8, 8, 8, 8)

        # --- Файлы ---
        grp = QGroupBox('Файлы')
        gl = QVBoxLayout(grp)

        r1 = QHBoxLayout()
        r1.addWidget(QLabel('Файл DOCX:'))
        self.input_edit = QLineEdit()
        self.input_edit.setPlaceholderText('Выберите файл .docx...')
        r1.addWidget(self.input_edit, 1)
        b1 = QPushButton('Обзор...')
        b1.clicked.connect(self._browse_input)
        r1.addWidget(b1)
        gl.addLayout(r1)

        r2 = QHBoxLayout()
        r2.addWidget(QLabel('Папка вывода:'))
        self.output_edit = QLineEdit()
        self.output_edit.setPlaceholderText('Папка для сохранения XML...')
        r2.addWidget(self.output_edit, 1)
        b2 = QPushButton('Обзор...')
        b2.clicked.connect(self._browse_output)
        r2.addWidget(b2)
        gl.addLayout(r2)

        r3 = QHBoxLayout()
        self.chk_split = QCheckBox('Разделить XML на части до 1 МБ')
        r3.addWidget(self.chk_split)
        r3.addStretch()
        bp = QPushButton('  Предпросмотр  ')
        bp.setStyleSheet('font-weight:bold; padding:6px 18px;')
        bp.clicked.connect(self._do_preview)
        r3.addWidget(bp)
        bc = QPushButton('  Конвертировать  ')
        bc.setStyleSheet('font-weight:bold; padding:6px 18px; background:#4CAF50; color:white;')
        bc.clicked.connect(self._do_convert)
        r3.addWidget(bc)
        gl.addLayout(r3)
        lay.addWidget(grp)

        # --- Центр: дерево (50%) + предпросмотр Moodle (50%) ---
        h_splitter = QSplitter(Qt.Horizontal)
        h_splitter.setHandleWidth(1)

        # Левая панель: дерево вопросов + лог
        v_splitter = QSplitter(Qt.Vertical)
        v_splitter.setHandleWidth(1)

        # Компактная панель с чекбоксом и счётчиком (в одну строку)
        header_widget = QWidget()
        header_layout = QHBoxLayout(header_widget)
        header_layout.setContentsMargins(5, 2, 5, 2)
        header_layout.setSpacing(10)
        
        self.chk_select_all = QCheckBox('Выделить все')
        self.chk_select_all.setChecked(True)
        self.chk_select_all.stateChanged.connect(self._toggle_select_all)
        header_layout.addWidget(self.chk_select_all)
        
        self.lbl_selected = QLabel('Выбрано: 0 / 0')
        header_layout.addWidget(self.lbl_selected)
        
        header_layout.addStretch()
        
        header_widget.setStyleSheet('background: #f0f0f0; border-bottom: 1px solid #ccc;')
        header_widget.setFixedHeight(32)
        v_splitter.addWidget(header_widget)

        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(['✓', '#', 'Имя / Содержимое', 'Маркер', 'Тип', 'Балл', 'Ошибки'])
        self.tree.setColumnCount(7)
        hdr = self.tree.header()
        hdr.setStretchLastSection(False)
        hdr.setSectionResizeMode(0, QHeaderView.Fixed)
        hdr.setSectionResizeMode(1, QHeaderView.Fixed)
        hdr.setSectionResizeMode(2, QHeaderView.Stretch)
        hdr.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(6, QHeaderView.ResizeToContents)
        self.tree.setColumnWidth(0, 30)
        self.tree.setColumnWidth(1, 45)
        self.tree.setColumnWidth(3, 180)
        self.tree.setColumnWidth(4, 140)
        self.tree.setColumnWidth(5, 50)
        self.tree.setAlternatingRowColors(True)
        self.tree.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.tree.setRootIsDecorated(True)
        self.tree.setAnimated(True)
        self.tree.setFont(QFont('Segoe UI', 9))
        self.tree.itemChanged.connect(self._on_item_check_changed)
        self.tree.itemClicked.connect(self._on_tree_item_clicked)
        
        v_splitter.addWidget(self.tree)

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(200)
        self.log_text.setFont(QFont('Consolas', 9))
        v_splitter.addWidget(self.log_text)
        v_splitter.setSizes([550, 200])
        
        # Правая панель: предпросмотр Moodle
        preview_group = QGroupBox('Предпросмотр Moodle')
        preview_layout = QVBoxLayout(preview_group)
        
        self.preview_text = QTextEdit()
        self.preview_text.setReadOnly(True)
        self.preview_text.setFont(QFont('Segoe UI', 10))
        self.preview_text.setStyleSheet('background: #fafafa; border: 1px solid #ccc;')
        preview_layout.addWidget(self.preview_text)
        
        h_splitter.addWidget(v_splitter)
        h_splitter.addWidget(preview_group)
        h_splitter.setSizes([500, 500])
        
        lay.addWidget(h_splitter, 1)

        # --- Низ ---
        bot = QHBoxLayout()
        self.progress = QProgressBar()
        self.progress.setTextVisible(True)
        bot.addWidget(self.progress, 1)
        self.status_label = QLabel('Готов')
        bot.addWidget(self.status_label)
        lay.addLayout(bot)

    # ---------- Browse ----------
    def _browse_input(self):
        p, _ = QFileDialog.getOpenFileName(self, 'Файл DOCX', '',
                                           'Word (*.docx);;Все (*)')
        if p:
            self.input_edit.setText(p)
            if not self.output_edit.text():
                self.output_edit.setText(os.path.dirname(p))

    def _browse_output(self):
        p = QFileDialog.getExistingDirectory(self, 'Папка для сохранения')
        if p:
            self.output_edit.setText(p)

    # ---------- Preview ----------
    def _do_preview(self):
        dp = self.input_edit.text().strip()
        if not dp or not os.path.isfile(dp):
            QMessageBox.warning(self, 'Ошибка', 'Выберите файл DOCX')
            return

        self.log_text.clear()
        self.log('Парсинг файла...')
        self.status_label.setText('Парсинг...')
        QApplication.processEvents()

        self.questions, errs = parse_docx_preview(dp)
        for e in errs:
            self.log(f'[ОШИБКА] {e}', 'red')
        self.log(f'Найдено вопросов: {len(self.questions)}')
        ec = sum(1 for q in self.questions if q.errors)
        if ec:
            self.log(f'Вопросов с ошибками: {ec}', 'orange')

        self._fill_tree()
        self._update_selected_count()

    def _fill_tree(self):
        self.tree.clear()
        marker_list = [''] + sorted(VALID_MARKERS)
        
        current_category = None
        category_indices = []  # Индексы вопросов текущей категории
        
        for idx, q in enumerate(self.questions):
            # Проверяем смену категории (маркера блока)
            if q.marker != current_category:
                # Если была предыдущая категория с вопросами - добавляем её чекбокс перед новой
                if category_indices and current_category is not None:
                    self._add_category_separator(current_category, category_indices)
                current_category = q.marker
                category_indices = [idx]
            else:
                category_indices.append(idx)
            
            # --- Родительский элемент (заголовок вопроса) ---
            item = QTreeWidgetItem()
            item.setCheckState(0, Qt.Checked)  # Галочка по умолчанию
            item.setData(0, Qt.UserRole, idx)  # Сохраняем индекс вопроса
            
            item.setText(1, str(idx + 1))
            item.setTextAlignment(1, Qt.AlignCenter)

            name_clean = q.name.replace('I:', '').strip()
            if len(name_clean) > 90:
                name_clean = name_clean[:87] + '...'
            item.setText(2, name_clean)
            item.setToolTip(2, q.name)

            # Маркер — текст (комбобокс встроим позже)
            item.setText(3, '')
            item.setText(4, q.auto_type)
            item.setText(5, str(q.grade))
            item.setTextAlignment(5, Qt.AlignCenter)

            err_text = '; '.join(q.errors) if q.errors else ''
            item.setText(6, err_text)

            # Цвета
            bg = MARKER_COLORS.get(q.marker, QColor(255, 255, 255))
            if q.errors:
                bg = COLOR_ERROR
            for c in range(7):
                item.setBackground(c, bg)
            if q.errors:
                item.setForeground(6, QBrush(QColor(180, 0, 0)))

            # Делаем расширяемым
            item.setChildIndicatorPolicy(QTreeWidgetItem.ShowIndicator)

            # --- Дочерние элементы: тело вопроса ---
            for line in q.content:
                child = QTreeWidgetItem()
                child.setFlags(child.flags() & ~Qt.ItemIsSelectable)
                
                # Галочка не нужна для дочерних элементов
                child.setCheckState(0, Qt.Unchecked)

                display = line
                if len(display) > 150:
                    display = display[:147] + '...'
                child.setText(2, display)
                child.setToolTip(2, line)

                # Окраска строк содержимого
                if line.startswith('+:') or (line.startswith('+') and len(line) > 1 and line[1] in ' \t'):
                    child.setForeground(2, QBrush(COLOR_CORRECT))
                    child.setText(1, '+')
                    child.setForeground(1, QBrush(COLOR_CORRECT))
                elif line.startswith('-:') or (line.startswith('-') and len(line) > 1 and line[1] in ' \t'):
                    child.setForeground(2, QBrush(COLOR_WRONG))
                    child.setText(1, '-')
                    child.setForeground(1, QBrush(COLOR_WRONG))
                elif line.startswith('S:'):
                    child.setForeground(2, QBrush(COLOR_TEXT))
                    child.setText(1, 'S')
                    child.setForeground(1, QBrush(COLOR_META))
                elif re.match(r'^[LR]\d+:', line):
                    child.setForeground(2, QBrush(COLOR_META))
                    child.setText(1, line[:2])
                elif re.match(r'^\d+:', line):
                    child.setForeground(2, QBrush(COLOR_META))
                    child.setText(1, '#')
                else:
                    child.setForeground(2, QBrush(COLOR_TEXT))

                item.addChild(child)

            self.tree.addTopLevelItem(item)

            # Комбобокс маркера
            combo = QComboBox()
            combo.setFont(QFont('Segoe UI', 8))
            for m in marker_list:
                desc = MARKER_DESCRIPTIONS.get(m, m)
                label = f'{m}  ({desc})' if m else desc
                combo.addItem(label, m)
            ci = marker_list.index(q.marker) if q.marker in marker_list else 0
            combo.setCurrentIndex(ci)
            combo.currentIndexChanged.connect(
                lambda i, row=idx, cb=combo: self._marker_changed(row, cb))
            self.tree.setItemWidget(item, 2, combo)

        self.tree.expandAll()
        # Свернём обратно, чтобы пользователь раскрывал вручную
        self.tree.collapseAll()
        
        # Добавляем последнюю категорию
        if category_indices and current_category is not None:
            self._add_category_separator(current_category, category_indices)
    
    def _add_category_separator(self, marker, indices):
        """Добавляет разделитель категории с чекбоксом для выбора всех вопросов категории."""
        desc = MARKER_DESCRIPTIONS.get(marker, marker) if marker else 'Без маркера'
        label = f'✓ Выбрать все ({desc})'
        
        sep = QTreeWidgetItem()
        sep.setText(1, '▸')
        sep.setText(2, label)
        sep.setCheckState(0, Qt.Checked)
        sep.setData(0, Qt.UserRole, {'category': True, 'indices': indices})
        
        # Серый фон для разделителя
        for c in range(7):
            sep.setBackground(c, QColor(220, 220, 220))
        sep.setForeground(2, QBrush(QColor(80, 80, 80)))
        
        self.tree.addTopLevelItem(sep)

    def _marker_changed(self, row, combo):
        mk = combo.currentData()
        if row >= len(self.questions):
            return
        q = self.questions[row]
        q.marker = mk
        q.auto_type = QuestionTypeDetector.detect(q.content, '', mk)

        item = self.tree.topLevelItem(row)
        if item:
            item.setText(3, q.auto_type)
            bg = MARKER_COLORS.get(mk, QColor(255, 255, 255))
            if q.errors:
                bg = COLOR_ERROR
            for c in range(6):
                item.setBackground(c, bg)

    # ---------- Convert ----------
    def _do_convert(self):
        dp = self.input_edit.text().strip()
        od = self.output_edit.text().strip()
        if not dp or not os.path.isfile(dp):
            QMessageBox.warning(self, 'Ошибка', 'Выберите файл DOCX')
            return
        if not od:
            QMessageBox.warning(self, 'Ошибка', 'Укажите папку для сохранения')
            return

        os.makedirs(od, exist_ok=True)
        bn = os.path.splitext(os.path.basename(dp))[0]
        out = os.path.join(od, bn + '.xml')

        if not self.questions:
            self._do_preview()

        # Проверяем выбранные вопросы
        selected_indices = [i for i, q in enumerate(self.questions) if getattr(q, 'selected', True)]
        if not selected_indices:
            QMessageBox.warning(self, 'Ошибка', 'Не выбрано ни одного вопроса')
            return

        eq = [q for q in self.questions if q.errors and getattr(q, 'selected', True)]
        if eq:
            r = QMessageBox.question(
                self, 'Предупреждение',
                f'{len(eq)} выбранных вопросов с ошибками.\nПродолжить?',
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if r != QMessageBox.Yes:
                return

        self.log('\n' + '=' * 60)
        self.log(f'Запуск конвертации ({len(selected_indices)} вопросов)...')
        self.progress.setValue(0)
        self.status_label.setText('Конвертация...')

        self.worker = ConvertWorker(dp, out, self.questions, self.chk_split.isChecked(), selected_indices)
        self.worker.progress.connect(self._on_progress)
        self.worker.finished.connect(self._on_finished)
        self.worker.validation.connect(self._on_validation)
        self.worker.start()

    def _on_progress(self, pct, msg):
        self.progress.setValue(pct)
        self.status_label.setText(msg)
        self.log(msg)

    def _on_validation(self, issues):
        if issues:
            self.log(f'\n--- Постобработка: {len(issues)} проблем ---', 'orange')
            for i in issues[:50]:
                self.log(f'  [!] {i}', 'orange')
            if len(issues) > 50:
                self.log(f'  ... и ещё {len(issues) - 50}', 'orange')
        else:
            self.log('Постобработка: XML корректен', 'green')

    def _on_finished(self, ok, result):
        if ok:
            self.log(f'\nСохранено: {result}', 'green')
            mb = os.path.getsize(result) / (1024 * 1024)
            self.log(f'Размер: {mb:.2f} МБ')
            self.status_label.setText('Готово!')
            QMessageBox.information(self, 'Готово',
                                    f'Конвертация завершена!\n\nФайл: {result}\nРазмер: {mb:.2f} МБ')
        else:
            self.log(f'\nОШИБКА: {result}', 'red')
            self.status_label.setText('Ошибка!')
            QMessageBox.critical(self, 'Ошибка', result)
        self.worker = None

    # ---------- Log ----------
    def log(self, text, color='black'):
        self.log_text.append(f'<span style="color:{color}">{text}</span>')
        sb = self.log_text.verticalScrollBar()
        sb.setValue(sb.maximum())

    # ---------- Checkbox ----------
    def _on_item_check_changed(self, item, column):
        if column == 0:
            idx = item.data(0, Qt.UserRole)
            checked = item.checkState(0) == Qt.Checked
            
            # Проверяем: это категория или отдельный вопрос
            if isinstance(idx, dict) and idx.get('category'):
                # Это чекбокс категории - выбрать/снять все вопросы в ней
                indices = idx.get('indices', [])
                for q_idx in indices:
                    if q_idx < len(self.questions):
                        self.questions[q_idx].selected = checked
                # Обновляем чекбоксы вопросов в дереве
                self._refresh_question_checkboxes(indices)
            elif idx is not None:
                # Обычный вопрос
                self.questions[idx].selected = checked
            
            self._update_selected_count()
    
    def _refresh_question_checkboxes(self, indices):
        """Обновляет чекбоксы конкретных вопросов в дереве."""
        for i in range(self.tree.topLevelItemCount()):
            item = self.tree.topLevelItem(i)
            if item:
                item_idx = item.data(0, Qt.UserRole)
                if item_idx is not None and not isinstance(item_idx, dict):
                    if item_idx in indices:
                        q = self.questions[item_idx]
                        item.setCheckState(0, Qt.Checked if q.selected else Qt.Unchecked)

    def _toggle_select_all(self, state):
        checked = state == Qt.Checked
        for i, q in enumerate(self.questions):
            q.selected = checked
        # Обновляем чекбоксы в дереве
        for i in range(self.tree.topLevelItemCount()):
            item = self.tree.topLevelItem(i)
            if item:
                item.setCheckState(0, Qt.Checked if checked else Qt.Unchecked)
        self._update_selected_count()

    def _update_selected_count(self):
        count = sum(1 for q in self.questions if getattr(q, 'selected', True))
        self.lbl_selected.setText(f'Выбрано: {count} / {len(self.questions)}')
        self.status_label.setText(f'Выбрано вопросов: {count} / {len(self.questions)}')

    def _on_tree_item_clicked(self, item, column):
        idx = item.data(0, Qt.UserRole)
        if idx is not None and not isinstance(idx, dict):
            if idx < len(self.questions):
                self._update_preview(idx)
    
    def _update_preview(self, q_idx):
        q = self.questions[q_idx]
        
        if q.errors:
            self.preview_text.setHtml(f'<div style="color:red; padding:10px;"><b>Ошибки в вопросе:</b><br>{ "<br>".join(q.errors) }</div>')
            return
        
        try:
            from universal_moodle_converter_v3_stable import XMLGenerator
            
            gen = XMLGenerator()
            
            marker = q.marker
            if not marker:
                from universal_moodle_converter_v3_stable import QuestionTypeDetector
                marker = QuestionTypeDetector.detect(q.content, '', marker)
            
            grade = q.grade if q.grade else 1.0
            
            if marker == 'multichoice_one':
                gen.create_multichoice(q.name, q.content, grade, single=True, penalty_wrong=1.0)
            elif marker == 'multichoice_many':
                gen.create_multichoice(q.name, q.content, grade, single=False, penalty_wrong=1.0)
            elif marker == 'matching':
                gen.create_matching(q.name, q.content, grade)
            elif marker == 'ddmatch':
                gen.create_ddmatch(q.name, q.content, grade)
            elif marker == 'gapselect':
                gen.create_gapselect(q.name, q.content, grade)
            elif marker == 'cloze':
                gen.create_cloze(q.name, q.content, grade)
            elif marker in ('numerical_partial', 'numerical_numcombo'):
                gen.create_numerical(q.name, q.content, grade)
            elif marker == 'shortanswer_phrase':
                gen.create_shortanswer(q.name, q.content, grade)
            else:
                gen.create_multichoice(q.name, q.content, grade, single=True, penalty_wrong=1.0)
            
            xml_str = etree.tostring(gen.root, pretty_print=True, encoding='unicode')
            
            html = self._xml_to_moodle_preview(q.name, q.content, marker, xml_str)
            self.preview_text.setHtml(html)
            
        except Exception as e:
            self.preview_text.setHtml(f'<div style="color:red; padding:10px;">Ошибка: {str(e)}</div>')
    
    def _xml_to_moodle_preview(self, name, content, marker, xml_str):
        tree = etree.fromstring(xml_str)
        
        q_elem = tree.find('.//question')
        if q_elem is None:
            return '<div style="color:red;">XML не содержит вопросов</div>'
        
        qtype = q_elem.get('type', '')
        
        if qtype == 'multichoice':
            single = q_elem.find('.//single')
            single_val = single.text if single is not None else 'true'
            is_single = single_val == 'true'
            
            qt = q_elem.find('.//questiontext')
            qtext = ''.join(qt.itertext()) if qt is not None else ''
            
            answers = q_elem.findall('.//answer')
            
            html = f'<div style="font-family:Segoe UI; max-width:100%;">'
            html += f'<h3 style="color:#333;">{name}</h3>'
            html += f'<p style="color:#555;">{qtext}</p>'
            html += '<div style="margin:10px 0 20px 20px;">'
            
            for ans in answers:
                fraction = ans.get('fraction', '0')
                ans_text = ''.join(ans.itertext())
                
                if fraction == '100':
                    html += f'<div style="color:green; margin:4px 0;">● {ans_text}</div>'
                elif fraction == '50':
                    html += f'<div style="color:#c90;">● {ans_text} (50%)</div>'
                else:
                    html += f'<div style="color:#999; margin:4px 0;">○ {ans_text}</div>'
            
            html += '</div></div>'
            
        elif qtype == 'matching':
            qt = q_elem.find('.//questiontext')
            qtext = ''.join(qt.itertext()) if qt is not None else ''
            
            subqs = q_elem.findall('.//subquestion')
            
            html = f'<div style="font-family:Segoe UI; max-width:100%;">'
            html += f'<h3 style="color:#333;">{name}</h3>'
            html += f'<p style="color:#555;">{qtext}</p>'
            
            for sq in subqs:
                text = ''.join(sq.find('text').itertext()) if sq.find('text') is not None else ''
                ans_text = ''.join(sq.find('.//answer/text').itertext()) if sq.find('.//answer') is not None else ''
                html += f'<div style="margin:5px 0; padding:5px; background:#f5f5f5; border-radius:4px;">'
                html += f'<b>{text}</b> → <span style="color:green;">{ans_text}</span>'
                html += '</div>'
            
            html += '</div>'
            
        elif qtype == 'gapselect':
            qt = q_elem.find('.//questiontext')
            qtext = ''.join(qt.itertext()) if qt is not None else ''
            
            selectopts = q_elem.findall('.//selectoption')
            
            html = f'<div style="font-family:Segoe UI; max-width:100%;">'
            html += f'<h3 style="color:#333;">{name}</h3>'
            html += f'<p style="color:#555;">{qtext}</p>'
            
            groups = {}
            for opt in selectopts:
                text = ''.join(opt.find('text').itertext()).strip()
                group = opt.get('group', '1')
                if group not in groups:
                    groups[group] = []
                groups[group].append(text)
            
            for g, opts in sorted(groups.items()):
                html += f'<div style="margin:10px 0;"><b>Варианты (группа {g}):</b><br>'
                for o in opts:
                    html += f'<span style="display:inline-block; margin:2px 5px; padding:2px 8px; background:#e0e0e0; border-radius:3px;">{o}</span>'
                html += '</div>'
            
            html += '</div>'
            
        elif qtype == 'cloze':
            qt = q_elem.find('.//questiontext')
            qtext = ''.join(qt.itertext()) if qt is not None else ''
            
            html = f'<div style="font-family:Segoe UI; max-width:100%;">'
            html += f'<h3 style="color:#333;">{name}</h3>'
            html += f'<p style="color:#555; font-style:italic;">{qtext}</p>'
            html += '</div>'
            
        elif qtype in ('numerical', 'shortanswer'):
            qt = q_elem.find('.//questiontext')
            qtext = ''.join(qt.itertext()) if qt is not None else ''
            
            answers = q_elem.findall('.//answer')
            
            html = f'<div style="font-family:Segoe UI; max-width:100%;">'
            html += f'<h3 style="color:#333;">{name}</h3>'
            html += f'<p style="color:#555;">{qtext}</p>'
            html += '<div style="margin:10px 0;">'
            
            for ans in answers:
                fraction = ans.get('fraction', '0')
                ans_text = ''.join(ans.itertext())
                
                if fraction == '100':
                    html += f'<div style="color:green;">Правильный ответ: {ans_text}</div>'
                else:
                    html += f'<div style="color:#999;">{ans_text}</div>'
            
            html += '<input type="text" style="width:200px; padding:5px; border:1px solid #ccc; border-radius:4px;" placeholder="Ваш ответ..." disabled>'
            html += '</div></div>'
        
        else:
            html = f'<div style="font-family:Segoe UI;">Тип: {qtype} — предпросмотр недоступен</div>'
        
        return html


# ============================================================
# ТОЧКА ВХОДА
# ============================================================

def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    w = MainWindow()
    w.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
