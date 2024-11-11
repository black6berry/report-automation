import datetime
from pathlib import Path


class DocumentGenerator:
    """
    Класс для работы с шаблонами файлов
    вызываемый метод генерирет определенный тип шаблона
    """
    user_not_finished = list()


    @classmethod
    def create_doc_diary(
            cls,
            sheet,
            full_path_to_dir: str,
            doc_template: str,
            user_not_finished=user_not_finished):
        """
        Генерация дневника

        :param
        sheet: Лист из Шаблона Excel файла

        :param
        path_dir: Название директории к папке в которой будут лежать созданные файлы

        :param
        full_path_to_dir: Полный путь до папки с файлами

        :param
        doc_template: Путь до шаблон файла на основе чего будет идти генерация

        :param
        user_not_finished: Список пользователей кто не получил оценку, берется из файла excel_templte
        """
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
            full_path_to_file = full_path_to_dir / f"{FIO}_дневник.docx"
            doc_template.save(full_path_to_file)

            print(f"""
                   Файл {full_path_to_file} успешно создан
                   Список должниов - {user_not_finished}
                   """)


    @classmethod
    def create_doc_characteristics(cls):
        """Генерация характеристики"""
        pass

    @classmethod
    def create_doc_individual_task(cls):
        """Генерация индивидуальных задания"""
        pass


    @classmethod
    def create_doc_report(cls):
        """Генерация отчеты"""
        pass


    @classmethod
    def create_doc(
            cls,
            type_file: int,
            sheet,
            full_path_to_dir,
            doc_template,
        ):
        """
        .. note::
            Выбор типа файла для генерации.
            Ф-я принимает тип файла:
        1 - Отчеты

        2 - Дневники

        3 - Характеристика

        4 - Индивидуальное задание
        """
        if type_file == 1:
            DocumentGenerator.create_doc_report()
        elif type_file == 2:
            DocumentGenerator.create_doc_diary(sheet, full_path_to_dir, doc_template)
        elif type_file == 3:
            DocumentGenerator.create_doc_characteristics()
        elif type_file == 4:
            DocumentGenerator.create_doc_individual_task()
        else:
            print(f"""
            Неизвестный тип документа: {type_file}, 
            Укажите правильный тип документа который хотите сгенерировать.
            """)