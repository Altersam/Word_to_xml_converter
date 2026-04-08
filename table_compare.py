# -*- coding: utf-8 -*-
import sys
import os
sys.stdout.reconfigure(encoding='utf-8')

from lxml import etree
from collections import Counter

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def analyze_xml(path, name):
    """Анализирует XML файл и возвращает статистику по типам вопросов."""
    full_path = os.path.join(BASE_DIR, path)
    try:
        tree = etree.parse(full_path)
        root = tree.getroot()
    except Exception as e:
        print(f"Ошибка чтения {path}: {e}", file=sys.stderr)
        return {}
    
    results = {}
    for q in root.findall('.//question'):
        qtype = q.get('type', 'unknown')
        if qtype == 'category':
            continue
        
        if qtype not in results:
            results[qtype] = {
                'count': 0,
                'answers': 0,
                'fractions': set(),
                'images': 0
            }
        
        results[qtype]['count'] += 1
        
        # Answers
        for a in q.findall('answer'):
            results[qtype]['answers'] += 1
            frac = a.get('fraction', '0')
            if frac not in ('0', '-'):
                results[qtype]['fractions'].add(frac)
        
        # Images
        for f in q.findall('.//file'):
            results[qtype]['images'] += 1
    
    return results

def format_fractions(fractions):
    """Форматирует фракции для отображения."""
    if not fractions:
        return '-'
    return ','.join(sorted(fractions, key=lambda x: -float(x) if x.replace('.','').isdigit() else 0))

# Пары для сравнения
comparisons = [
    ('АЯ 10кл', 'test_output\\вопросы-АЯ  10кл  ВИ ЛФУ.xml', '..\\output\\вопросы-АЯ  10кл  ВИ ЛФУ  2026-Английский язык  10 класс  Вступительные испытания ЛИЦЕЙ  2026-20260320-1515.xml'),
    ('РЯ 10кл', 'test_output\\вопросы-РЯ  10кл  ВИ ЛФУ.xml', '..\\output\\вопросы-РЯ  10кл  ВИ ЛФУ  2026-Русский язык  10 класс  Вступительные испытания ЛИЦЕЙ  2026-20260320-1747.xml'),
    ('МАТ 10кл', 'test_output\\вопросы-МАТ  10кл  ВИ ЛФУ.xml', '..\\output\\вопросы-МАТ  10кл  ВИ ЛФУ  2026-Математика  10 класс  Вступительные испытания ЛИЦЕЙ  2026-20260320-1515.xml'),
    ('ФИЗ 10кл', 'test_output\\вопросы-ФИЗ  10кл  ВИ ЛФУ.xml', '..\\output\\вопросы-ФИЗ  10кл  ВИ ЛФУ  2026-Физика  10 класс  Вступительные испытания ЛИЦЕЙ  2026-20260320-1748.xml'),
    ('НЯ 10кл', 'test_output\\вопросы-НЯ  10кл  ВИ ЛФУ.xml', '..\\output\\вопросы-НЯ  10кл  ВИ ЛФУ  2026-Нем.язык  10 класс  Вступительные испытания ЛИЦЕЙ  2026-20260320-1516.xml'),
    ('ИЯ 10кл', 'test_output\\вопросы-ИЯ  10кл  ВИ ЛФУ.xml', '..\\output\\вопросы-ИЯ  10кл  ВИ ЛФУ  2026-Исп.язык  10 класс  Вступительные испытания ЛИЦЕЙ  2026-20260320-1515.xml'),
    ('ФЯ 10кл', 'test_output\\вопросы-ФЯ  10кл  ВИ ЛФУ  2026.xml', '..\\output\\вопросы-ФЯ  10кл  ВИ ЛФУ  2026-Франц.язык  10 класс  Вступительные испытания ЛИЦЕЙ  2026-20260320-1514.xml'),
    ('МАТ 8кл', 'test_output\\вопросы-МАТ  8кл  ВИ ЛФУ.xml', '..\\output\\вопросы-МАТ  8кл  ВИ ЛФУ  2026.xml'),
    ('АЯ 8кл', 'test_output\\вопросы-АЯ  8кл  ВИ ЛФУ.xml', '..\\output\\вопросы-АЯ  8кл  ВИ ЛФУ.xml'),
]

# Заголовок таблицы
print(f"{'Предмет':<12} | {'Эталон':^40} | {'Конвертер':^40}")
print(f"{'':12} | {'тип/кол-во/fraction/ответы/картинки':^40} | {'тип/кол-во/fraction/ответы/картинки':^40}")
print("-" * 95)

for name, gen_path, ref_path in comparisons:
    gen_stats = analyze_xml(gen_path, name)
    ref_stats = analyze_xml(ref_path, name)
    
    # Получаем все типы из обоих файлов
    all_types = set(gen_stats.keys()) | set(ref_stats.keys())
    if 'category' in all_types:
        all_types.discard('category')
    
    if not all_types:
        print(f"{name:<12} | {'ПУСТОЙ ФАЙЛ':^40} | {'ПУСТОЙ ФАЙЛ':^40}")
        continue
    
    # Если только один тип, выводим одну строку
    if len(all_types) == 1:
        qtype = list(all_types)[0]
        ref = ref_stats.get(qtype, {'count': 0, 'answers': 0, 'fractions': set(), 'images': 0})
        gen = gen_stats.get(qtype, {'count': 0, 'answers': 0, 'fractions': set(), 'images': 0})
        
        ref_str = f"{qtype}/{ref['count']}/{format_fractions(ref['fractions'])}/{ref['answers']}/{ref['images']}"
        gen_str = f"{qtype}/{gen['count']}/{format_fractions(gen['fractions'])}/{gen['answers']}/{gen['images']}"
        
        match = "✓" if ref['count'] == gen['count'] and ref['answers'] == gen['answers'] else "✗"
        print(f"{name:<12} | {ref_str:^40} | {gen_str:^40} {match}")
    else:
        # Несколько типов - выводим первую строку с названием
        print(f"{name:<12} |", end="")
        
        first = True
        for qtype in sorted(all_types):
            if not first:
                print(f"{'':12} |", end="")
            
            ref = ref_stats.get(qtype, {'count': 0, 'answers': 0, 'fractions': set(), 'images': 0})
            gen = gen_stats.get(qtype, {'count': 0, 'answers': 0, 'fractions': set(), 'images': 0})
            
            ref_str = f"{qtype}/{ref['count']}/{format_fractions(ref['fractions'])}/{ref['answers']}/{ref['images']}"
            gen_str = f"{qtype}/{gen['count']}/{format_fractions(gen['fractions'])}/{gen['answers']}/{gen['images']}"
            
            match = "✓" if ref['count'] == gen['count'] and ref['answers'] == gen['answers'] else "✗"
            print(f" {ref_str:^38} | {gen_str:^38} {match}")
            first = False
