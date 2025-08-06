import re
import os
import zipfile
from pathlib import Path
import shutil
import tempfile

def remove_sheet_protection(content):
    """Удаляет все вхождения <sheetProtection.../> в тексте."""
    return re.sub(r'<sheetProtection.*?/>', '', content, flags=re.DOTALL)

def process_xml_file(file_path):
    """Перезаписывает XML-файл без <sheetProtection.../>."""
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    modified_content = remove_sheet_protection(content)
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(modified_content)

def process_xlsx_ods(file_path):
    """Перезаписывает XLSX/ODS без <sheetProtection.../>."""
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_dir_path = Path(temp_dir)
        
        # Распаковка
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir_path)
        
        # Поиск и изменение sheet*.xml
        sheets_dir = temp_dir_path / 'xl' / 'worksheets'
        if sheets_dir.exists():
            for sheet_file in sheets_dir.glob('*.xml'):
                process_xml_file(sheet_file)
        
        # Перезапись исходного файла
        with zipfile.ZipFile(file_path, 'w', zipfile.ZIP_DEFLATED) as zip_out:
            for root, _, files in os.walk(temp_dir):
                for file in files:
                    file_path_local = Path(root) / file
                    arcname = os.path.relpath(file_path_local, temp_dir)
                    zip_out.write(file_path_local, arcname)

def process_folder(folder_path):
    """Обрабатывает все XML/XLSX/ODS в указанной папке."""
    folder_path = Path(folder_path)
    if not folder_path.exists():
        print(f"❌ Ошибка: папка {folder_path} не существует!")
        return

    for file in folder_path.glob('*'):
        if file.suffix.lower() == '.xml':
            process_xml_file(file)
            print(f"XML обработан: {file.name}")
        elif file.suffix.lower() in ('.xlsx', '.ods'):
            process_xlsx_ods(file)
            print(f"XLSX/ODS обработан: {file.name}")

if __name__ == "__main__":
    print("🛠 Удаление <sheetProtection.../> из XML/XLSX/ODS файлов")
    folder_path = input("Введите путь к папке с файлами: ").strip()
    process_folder(folder_path)
    print("✅ Готово! Все файлы перезаписаны.")