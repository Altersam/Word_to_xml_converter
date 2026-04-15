import sys
with open(r'C:\Users\alter\YandexDisk\Лицей\Цифровизация\9009\word_to_xml\.Final\converter_gui.py', 'r') as f:
    lines = f.readlines()

# Check line 920 - this is the nested function def
print('Line 920:', repr(lines[919]))
print('Line 921:', repr(lines[920]))
print('Line 922:', repr(lines[921]))

# Check indentation of nested function def
indent = len(lines[919]) - len(lines[919].lstrip())
print('Line 919 indent:', indent)

# Check the parent function
func_def = None
for i in range(910):
    if 'def _xml_to_moodle_preview' in lines[i]:
        func_def = i
        break
print('Function def at', func_def)
func_indent = len(lines[func_def]) - len(lines[func_def].lstrip())
print('Function def indent:', func_indent)