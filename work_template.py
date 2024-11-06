import datetime

class TemplateFile():
    """
    Класс для работы с шаблонами файлов
    вызываемый метод генерирет определенный тип шаблона
    Принимает параметры:
    sheet - лист excel с данными
    full_path_to_dir - полный путь до директории
    path_folder - путь до папки куда сохранять готовые документы
    path_template - путь до шаблона файла куда поставлять значения из excel файла
    user_not_finished - список должников
    """
    user_not_finished = list()
    def __init__(self, sheet, full_path_to_dir, path_folder, path_template, doc_templates):
        self.sheet = sheet
        self.full_path_to_dir = full_path_to_dir
        self.path_folder = path_folder
        self.path_template = path_template
        self.doc_templates = doc_templates


    @staticmethod
    def create_type_files():
        """Выбор типа файла для генерации"""


    @staticmethod
    def create_doc_diary(sheet, full_path_to_dir, path_folder, path_template, doc_template):
        """Генерация дневника"""
        for num in range(2, len(list(sheet.rows)) + 1):
            FIO = sheet['A' + str(num)].value
            signature = FIO.split(' ')[0] if FIO else "Нет данных"
            # 3 (удовлетворительно)
            # 4 (хорошо)
            # 5 (отлично)
            score = sheet['B' + str(num)].value

            if score == 3:
                score = f"{score} (удовл.)"
                description_for_student = sheet['G2'].value
            elif score == 4:
                score = f"{score} (хор.)"
                description_for_student = sheet['H2'].value
            elif score == 5:
                score = f"{score} (отл.)"
                description_for_student = sheet['I2'].value
            else:
                user_not_finished.append(FIO)
                description_for_student = "Нет данных"

            start_date = sheet['D2'].value.date()
            # print(start_date)
            if isinstance(start_date, datetime.date):
                start_month = start_date.month
                # print(start_month)
                start_day = start_date.day
                # print(start_day)
            else:
                start_date = start_month = start_day = "Нет данных"

            end_date = sheet['E2'].value.date()
            # print(end_date)
            if isinstance(end_date, datetime.date):
                end_month = end_date.month
                # print(end_month)
                end_day = end_date.day
                # print(end_day)
            else:
                end_date = end_month = end_day = "Нет данных"

            group = sheet['E6'].value
            course = sheet['D6'].value
            # Данные для заполнения шаблона
            context = {
                'FIO': FIO,
                'signature': signature,
                'score': score,
                'start_date': start_date,
                'start_month': start_month,
                'start_day': start_day,
                'end_date': end_date,
                'end_month': end_month,
                'end_day': end_day,
                'group': group,
                'course': course,
                'description_for_student': description_for_student
            }
            # Заполнение шаблона данными
            doc_template.render(context)

            # Сохранение документа дневник
            doc_template.save(full_path_to_dir / f"{FIO}_дневник.docx")

    @staticmethod
    def create_doc_charcteristics():
        """Генерация характеристики"""
        pass

    @staticmethod
    def create_doc_individual_task():
        """Генерация индивидуальных задания"""
        pass

    @staticmethod
    def create_doc_report():
        """Генерация отчеты"""
        pass