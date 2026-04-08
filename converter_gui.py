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
    'shortanswer_partial': 'Цифры partial scoring',
    'shortanswer_numcombo': 'Цифры любой порядок',
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
    'shortanswer_partial':  QColor(220, 255, 200),
    'shortanswer_numcombo': QColor(240, 255, 200),
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
                 'content', 'auto_type', 'errors', 'line_num')

    def __init__(self):
        self.name: str = ''
        self.grade: float = 1.0
        self.marker: str = ''
        self.subcategory: str = ''
        self.content: List[str] = []
        self.auto_type: str = ''
        self.errors: List[str] = []
        self.line_num: int = 0


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
                'shortanswer_phrase', 'shortanswer_numcombo')
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
    sz = os.path.getsize(xml_path)
    if sz <= max_bytes:
        return [xml_path]

    tree = etree.parse(xml_path)
    root = tree.getroot()

    base, ext = os.path.splitext(xml_path)
    parts = []

    cur = etree.Element('quiz')
    cur_sz = len(etree.tostring(cur, encoding='utf-8', xml_declaration=True))
    last_category = None
    pn = 1

    for ch in root:
        # Пропускаем элементы без тега (например, комментарии)
        if ch.tag is None:
            continue
            
        qb = len(etree.tostring(ch, encoding='utf-8'))
        
        # Проверяем, не превысим ли лимит
        if cur_sz + qb > max_bytes and len(cur) > 0:
            # Нужно проверить: если текущий элемент - вопрос, 
            # а не категория, то можно разделять
            is_category = ch.get('type') == 'category'
            
            # Если это вопрос (не категория), то сохраняем текущий файл
            # и создаём новый
            if not is_category:
                # Сохраняем текущую часть
                pp = f'{base}_part{pn}{ext}'
                etree.ElementTree(cur).write(pp, encoding='utf-8', xml_declaration=True)
                parts.append(pp)
                pn += 1
                
                # Начинаем новую часть с последней категорией
                cur = etree.Element('quiz')
                if last_category is not None:
                    cur.append(deepcopy(last_category))
                cur_sz = len(etree.tostring(cur, encoding='utf-8', xml_declaration=True))
        
        # Добавляем элемент
        cur.append(ch)
        
        # Запоминаем категорию для следующего файла
        if ch.get('type') == 'category':
            last_category = ch
            cur_sz = len(etree.tostring(cur, encoding='utf-8', xml_declaration=True))
        else:
            cur_sz += qb

    # Сохраняем последнюю часть
    if len(cur) > 0:
        pp = f'{base}_part{pn}{ext}'
        etree.ElementTree(cur).write(pp, encoding='utf-8', xml_declaration=True)
        parts.append(pp)

    return parts


# ============================================================
# РАБОЧИЙ ПОТОК
# ============================================================

class ConvertWorker(QThread):
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(bool, str)
    validation = pyqtSignal(list)

    def __init__(self, docx_path, output_path, questions, split_xml, parent=None):
        super().__init__(parent)
        self.docx_path = docx_path
        self.output_path = output_path
        self.questions = questions
        self.split_xml = split_xml

    def run(self):
        try:
            self.progress.emit(10, 'Запуск конвертации...')
            conv = MoodleConverter(self.docx_path, self.output_path)
            conv._marker_overrides = {}
            for q in self.questions:
                if q.marker:
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

        # --- Центр: дерево + лог ---
        splitter = QSplitter(Qt.Vertical)

        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(['#', 'Имя / Содержимое', 'Маркер', 'Тип', 'Балл', 'Ошибки'])
        self.tree.setColumnCount(6)
        hdr = self.tree.header()
        hdr.setStretchLastSection(False)
        hdr.setSectionResizeMode(0, QHeaderView.Fixed)
        hdr.setSectionResizeMode(1, QHeaderView.Stretch)
        hdr.setSectionResizeMode(2, QHeaderView.Fixed)
        hdr.setSectionResizeMode(3, QHeaderView.Fixed)
        hdr.setSectionResizeMode(4, QHeaderView.Fixed)
        hdr.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        self.tree.setColumnWidth(0, 45)
        self.tree.setColumnWidth(2, 200)
        self.tree.setColumnWidth(3, 150)
        self.tree.setColumnWidth(4, 50)
        self.tree.setAlternatingRowColors(True)
        self.tree.setSelectionMode(QAbstractItemView.SingleSelection)
        self.tree.setRootIsDecorated(True)
        self.tree.setAnimated(True)
        self.tree.setFont(QFont('Segoe UI', 9))
        splitter.addWidget(self.tree)

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(200)
        self.log_text.setFont(QFont('Consolas', 9))
        splitter.addWidget(self.log_text)
        splitter.setSizes([550, 200])
        lay.addWidget(splitter, 1)

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
        self.status_label.setText(f'Предпросмотр: {len(self.questions)} вопросов')

    def _fill_tree(self):
        self.tree.clear()
        marker_list = [''] + sorted(VALID_MARKERS)

        for idx, q in enumerate(self.questions):
            # --- Родительский элемент (заголовок вопроса) ---
            item = QTreeWidgetItem()
            item.setText(0, str(idx + 1))
            item.setTextAlignment(0, Qt.AlignCenter)

            name_clean = q.name.replace('I:', '').strip()
            if len(name_clean) > 90:
                name_clean = name_clean[:87] + '...'
            item.setText(1, name_clean)
            item.setToolTip(1, q.name)

            # Маркер — текст (комбобокс встроим позже)
            item.setText(2, '')
            item.setText(3, q.auto_type)
            item.setText(4, str(q.grade))
            item.setTextAlignment(4, Qt.AlignCenter)

            err_text = '; '.join(q.errors) if q.errors else ''
            item.setText(5, err_text)

            # Цвета
            bg = MARKER_COLORS.get(q.marker, QColor(255, 255, 255))
            if q.errors:
                bg = COLOR_ERROR
            for c in range(6):
                item.setBackground(c, bg)
            if q.errors:
                item.setForeground(5, QBrush(QColor(180, 0, 0)))

            # Делаем расширяемым
            item.setChildIndicatorPolicy(QTreeWidgetItem.ShowIndicator)

            # --- Дочерние элементы: тело вопроса ---
            for line in q.content:
                child = QTreeWidgetItem()
                child.setFlags(child.flags() & ~Qt.ItemIsSelectable)

                display = line
                if len(display) > 150:
                    display = display[:147] + '...'
                child.setText(1, display)
                child.setToolTip(1, line)

                # Окраска строк содержимого
                if line.startswith('+:') or (line.startswith('+') and len(line) > 1 and line[1] in ' \t'):
                    child.setForeground(1, QBrush(COLOR_CORRECT))
                    child.setText(0, '+')
                    child.setForeground(0, QBrush(COLOR_CORRECT))
                elif line.startswith('-:') or (line.startswith('-') and len(line) > 1 and line[1] in ' \t'):
                    child.setForeground(1, QBrush(COLOR_WRONG))
                    child.setText(0, '-')
                    child.setForeground(0, QBrush(COLOR_WRONG))
                elif line.startswith('S:'):
                    child.setForeground(1, QBrush(COLOR_TEXT))
                    child.setText(0, 'S')
                    child.setForeground(0, QBrush(COLOR_META))
                elif re.match(r'^[LR]\d+:', line):
                    child.setForeground(1, QBrush(COLOR_META))
                    child.setText(0, line[:2])
                elif re.match(r'^\d+:', line):
                    child.setForeground(1, QBrush(COLOR_META))
                    child.setText(0, '#')
                else:
                    child.setForeground(1, QBrush(COLOR_TEXT))

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

        eq = [q for q in self.questions if q.errors]
        if eq:
            r = QMessageBox.question(
                self, 'Предупреждение',
                f'{len(eq)} вопросов с ошибками.\nПродолжить?',
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if r != QMessageBox.Yes:
                return

        self.log('\n' + '=' * 60)
        self.log('Запуск конвертации...')
        self.progress.setValue(0)
        self.status_label.setText('Конвертация...')

        self.worker = ConvertWorker(dp, out, self.questions, self.chk_split.isChecked())
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
