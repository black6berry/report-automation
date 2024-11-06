import os
from pathlib import Path
import datetime
import openpyxl
from docxtpl import DocxTemplate



def main():
    template_paths = {
        "1": {
            "path_template": "templates\отчет_ФИО_шаблон.docx",
            "path_dir": "Отчеты",
        },
        "2": {
            "path_template": "templates\дневник_ФИО_шаблон.docx",
            "path_dir": "Дневники",
        },
        "3": {
            "path_template": "templates\атт лист_характеристика_ФИО_шаблон.docx",
            "path_dir": "Характеристика",
        },
        "4": {
            "path_template": "templates\инд.задание_ФИО_шаблон.docx",
            "path_dir": "Индивидуальное задание",
        },
    }

    print("""
       Какой файл вы хотите сгенерировать ? 
       1 - Отчеты
       2 - Дневники
       3 - Характеристика
       4 - Индивидуальное задание
       """)

    file_int = int(input("Выберите цифру: "))
    print(file_int)

    path_dir = template_paths[f"{file_int}"]["path_dir"]
    print(path_dir)

    home_dir = Path.cwd()
    print(home_dir)

    full_path_to_dir = Path(home_dir) / f"{path_dir}"
    # Проверяем создана ли дериктория
    if not os.path.exists(full_path_to_dir):
        # Создаем директориию
        os.makedirs(full_path_to_dir)

    path_template = template_paths[f"{file_int}"]["path_template"]
    print(path_template)

    doc_template = DocxTemplate(f"{path_template}")
    print(doc_template)


    wb = openpyxl.load_workbook(filename='excel_template.xlsx')
    sheet = wb['Практика']




    print(f"""
        Файлы {path_dir} успешно созданы
        Список должниов - {not_finished}
        """)





# Загрузка шаблона дневника
# doc_template = DocxTemplate("/templates/дневник_ФИО_шаблон.docx")
# Загрузка шаблона характеристика
# doc_charcteristics = DocxTemplate("templates/атт лист_характеристика_ФИО_шаблон.docx")
# Загрузка шаблона инд.задание
# doc_individual_task = DocxTemplate("templates/инд.задание_ФИО_шаблон.docx")
# Загрузка шаблона отчета
# doc_report = DocxTemplate("templates/отчет_ФИО_шаблон.docx")



# Сохранение документа отчета
# doc_report.save(f"{FIO}_отчет.docx")

# Сохранение документа характеристика
# doc_charcteristics.save(f"{FIO}_атт.характеристика.docx")

# Сохранение докуммента индивидуальное задание
# doc_individual_task.save(f"{FIO}_инд.задание.docx")





if __name__ == '__main__':
    main()
    # create_doc_diary()
    pass

