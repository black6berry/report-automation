import os
from pathlib import Path
import datetime
import openpyxl
from docxtpl import DocxTemplate

from work_folder import WorkFolder
from work_template import DocumentGenerator


def main():
    template_paths = {
        "УП12": {

            "1": {
                "path_template": "templates\УП12\отчет_ФИО_шаблон_УП12.docx",
                "path_dir": "Отчеты",
            },
            "2": {
                "path_template": "templates\УП12\дневник_ФИО_шаблон_УП12.docx",
                "path_dir": "Дневники",
            },
            "3": {
                "path_template": "templates\УП12\атт.лист_характеристика_ФИО_шаблон_УП12.docx",
                "path_dir": "Характеристика",
            },
            "4": {
                "path_template": "templates\УП12\инд.задание_ФИО_шаблон_УП12.docx",
                "path_dir": "Индивидуальное задание",
            },
        },
        "УП13": {
            "1": {
                "path_template": "templates\УП13\отчет_ФИО_шаблон_УП13.docx",
                "path_dir": "Отчеты",
            },
            "2": {
                "path_template": "templates\УП13\дневник_ФИО_шаблон_УП13.docx",
                "path_dir": "Дневники",
            },
            "3": {
                "path_template": "templates\УП13\атт.лист_характеристика_ФИО_шаблон_УП13.docx",
                "path_dir": "Характеристика",
            },
            "4": {
                "path_template": "templates\УП13\инд.задание_ФИО_шаблон_УП13.docx",
                "path_dir": "Индивидуальное задание",
            },
        }

    }

    while True:
        print("""
            Документы для какой практики создаем ?
            УП12 | УП13
        """)

        while True:
            try:
                name_practice = str(input("Введите название практики: ")).upper()
                if name_practice == "УП12" or name_practice == "УП13":
                    break
                else:
                    raise ValueError
            except ValueError:
                print("Такой практики нет, будь внимательнее :D")

        print("""
           Какой файл вы хотите сгенерировать ? 
           1 - Отчеты
           2 - Дневники
           3 - Характеристика
           4 - Индивидуальное задание
        """)

        while True:
            try:
                type_file_int = int(input("Выберите цифру: "))
                if type_file_int == 1 or type_file_int == 2 or type_file_int == 3 or type_file_int == 4:
                    break
                else:
                    raise ValueError
            except ValueError:
                print("Вы ввели не число, введите число")

        try:
            path_dir = template_paths[f"{name_practice}"][f"{type_file_int}"]["path_dir"]
        except:
            print(f"Не удается найти шаблон - {path_dir}")

        home_dir = Path.cwd()
        # print(home_dir)

        full_path_to_dir = Path(home_dir) / f"{path_dir}"
        # Создаем папку для хранения в ней файлов
        WorkFolder.create_folder(full_path_to_dir)

        path_template = template_paths[f"{name_practice}"][f"{type_file_int}"]["path_template"]
        # print(path_template)

        doc_template = DocxTemplate(f"{path_template}")
        # print(doc_template)

        # Загружаем шаблон excel
        wb = openpyxl.load_workbook(filename='excel_template.xlsx')
        # Определяем страницу с рабочей областью
        sheet = wb['Практика']

        new_doc = DocumentGenerator().create_doc(type_file=type_file_int,
                                                 sheet=sheet,
                                                 full_path_to_dir=full_path_to_dir,
                                                 doc_template=doc_template,
                                                 )
        # DocumentGenerator.create_doc(path_dir)
        # DocumentGenerator.create_doc_diary(path_dir)

        repeat = str(input("Нужно сгенерировать еще файлы д\н? "))
        if repeat == "Н" or repeat == "н" or repeat == "F" or repeat == "f":
            break



if __name__ == '__main__':
    main()
    # create_doc_diary()
    pass
