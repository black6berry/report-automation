import datetime
from pathlib import Path


class DocumentGenerator:
    """
    Класс для работы с шаблонами файлов
    вызываемый метод генерирет определенный тип шаблона
    """
    user_not_finished = list()
    current_year = datetime.datetime.now().year


    @classmethod
    def create_doc_diary(
            cls, sheet,
            full_path_to_dir: str, doc_template: str,
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
            if FIO is not None:
                signature = FIO.split(' ')[0] if FIO else "Нет данных"
                # 3 (удовлетворительно)
                # 4 (хорошо)
                # 5 (отлично)
                score = sheet['C' + str(num)].value

                if score == 3:
                    score = f"{score} (Удовл.)"
                elif score == 4:
                    score = f"{score} (Хор.)"
                elif score == 5:
                    score = f"{score} (Отл.)"
                else:
                    user_not_finished.append(FIO)

                start_date = sheet['F5'].value.date()
                # print(start_date)
                if isinstance(start_date, datetime.date):
                    start_month = start_date.month
                    # print(start_month)
                    start_day = start_date.day
                    if start_day < 9 or start_month < 9:
                        start_day = str(start_day).rjust(2, '0')
                        start_month = str(start_month).rjust(2, '0')
                else:
                    start_date = start_month = start_day = "Нет данных"

                end_date = sheet['F6'].value.date()
                # print(end_date)
                if isinstance(end_date, datetime.date):
                    end_month = end_date.month
                    # print(end_month)
                    end_day = end_date.day
                    if end_day < 9 or end_month < 9:
                        end_day = str(end_day).rjust(2, '0')
                        end_month = str(end_month).rjust(2, '0')
                else:
                    end_date = end_month = end_day = "Нет данных"

                count_day = (end_date - start_date).days
                group = sheet['F4'].value
                course = sheet['F3'].value
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
                    'count_day': count_day
                }
                # Заполнение шаблона данными
                doc_template.render(context)

                # Сохранение документа дневник
                full_path_to_file = full_path_to_dir / f"дневник_{FIO}.docx"
                doc_template.save(full_path_to_file)

                print(f"Файл {full_path_to_file} успешно создан")
        print(f"Список должниов - {user_not_finished}")
        user_not_finished.clear()


    @classmethod
    def create_doc_characteristics(
            cls, sheet,
            full_path_to_dir: str, doc_template: str,
            user_not_finished=user_not_finished):
        """
        Генерация характеристики

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
            if FIO is not None:
                signature = FIO.split(' ')[0] if FIO else "Нет данных"
                # 3 (удовлетворительно)
                # 4 (хорошо)
                # 5 (отлично)
                score = sheet['C' + str(num)].value
                if score == 3:
                    score = f"{score} (Удовл.)"
                    description_for_student = sheet['K2'].value
                elif score == 4:
                    score = f"{score} (Хор.)"
                    description_for_student = sheet['L2'].value
                elif score == 5:
                    score = f"{score} (Отл.)"
                    description_for_student = sheet['M2'].value
                else:
                    user_not_finished.append(FIO)
                    description_for_student = "Нет данных"

                start_date = sheet['F5'].value.date()
                # print(start_date)
                if isinstance(start_date, datetime.date):
                    start_month = start_date.month
                    # print(start_month)
                    start_day = start_date.day
                    if start_day < 9 or start_month < 9:
                        start_day = str(start_day).rjust(2, '0')
                        start_month = str(start_month).rjust(2, '0')
                else:
                    start_date = start_month = start_day = "Нет данных"

                end_date = sheet['F6'].value.date()
                # print(end_date)
                if isinstance(end_date, datetime.date):
                    end_month = end_date.month
                    # print(end_month)
                    end_day = end_date.day
                    if end_day < 9 or end_month < 9:
                        end_day = str(end_day).rjust(2, '0')
                        end_month = str(end_month).rjust(2, '0')
                else:
                    end_date = end_month = end_day = "Нет данных"

                count_day = (end_date - start_date).days
                group = sheet['F4'].value
                course = sheet['F3'].value
                # Данные для заполнения шаблона
                context = {
                    'current_year': cls.current_year,
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
                    'count_day': count_day,
                    'description_for_student': description_for_student
                }
                # Заполнение шаблона данными
                doc_template.render(context)

                # Сохранение документа дневник
                full_path_to_file = full_path_to_dir / f"характеристика_{FIO}.docx"
                doc_template.save(full_path_to_file)

                print(f"Файл {full_path_to_file} успешно создан")
        print(f"Список должниов - {user_not_finished}")
        user_not_finished.clear()


    @classmethod
    def create_doc_individual_task(
            cls, sheet,
            full_path_to_dir: str, doc_template: str,
            user_not_finished=user_not_finished):
        """
        Генерация индивидуальных задания

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
            if FIO is not None:
                signature = FIO.split(' ')[0] if FIO else "Нет данных"
                individual_task = sheet['B' + str(num)].value
                # task_data = []
                # for n in range(2, len(list(sheet.rows)) + 1):
                #     duration_day_work = sheet['H' + str(n)].value # Assuming "Количество дней" is in the first column
                #     task_description = sheet['I' + str(n)].value  # Assuming "Содержание работ" is in the second column
                #     task_data.append({"duration_day_work": duration_day_work, "task_description": task_description})
                # print(task_data)
                # 3 (удовлетворительно)
                # 4 (хорошо)
                # 5 (отлично)
                score = sheet['C' + str(num)].value
                if score == 3:
                    score = f"{score} (Удовл.)"
                    description_for_student = sheet['K2'].value
                elif score == 4:
                    score = f"{score} (Хор.)"
                    description_for_student = sheet['L2'].value
                elif score == 5:
                    score = f"{score} (Отл.)"
                    description_for_student = sheet['M2'].value
                else:
                    user_not_finished.append(FIO)
                    description_for_student = "Нет данных"

                start_date = sheet['F5'].value.date()
                # print(start_date)
                if isinstance(start_date, datetime.date):
                    start_month = start_date.month
                    # print(start_month)
                    start_day = start_date.day
                    if start_day < 9 or start_month < 9:
                        start_day = str(start_day).rjust(2, '0')
                        start_month = str(start_month).rjust(2, '0')
                else:
                    start_date = start_month = start_day = "Нет данных"

                end_date = sheet['F6'].value.date()
                # print(end_date)
                if isinstance(end_date, datetime.date):
                    end_month = end_date.month
                    # print(end_month)
                    end_day = end_date.day
                    if end_day < 9 or end_month < 9:
                        end_day = str(end_day).rjust(2, '0')
                        end_month = str(end_month).rjust(2, '0')
                else:
                    end_date = end_month = end_day = "Нет данных"

                count_day = (end_date - start_date).days
                group = sheet['F4'].value
                course = sheet['F3'].value
                # Данные для заполнения шаблона
                context = {
                    'current_year': cls.current_year,
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
                    'count_day': count_day,
                    'description_for_student': description_for_student,
                    'individual_task': individual_task
                }
                # Заполнение шаблона данными
                doc_template.render(context)

                # Сохранение документа дневник
                full_path_to_file = full_path_to_dir / f"инд.задание_{FIO}.docx"
                doc_template.save(full_path_to_file)

                print(f"Файл {full_path_to_file} успешно создан")
        print(f"Список должниов - {user_not_finished}")
        user_not_finished.clear()


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
            doc_template):
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
            pass
            # DocumentGenerator.create_doc_report(sheet, full_path_to_dir, doc_template)
        elif type_file == 2:
            DocumentGenerator.create_doc_diary(sheet, full_path_to_dir, doc_template)
        elif type_file == 3:
            DocumentGenerator.create_doc_characteristics(sheet, full_path_to_dir, doc_template)
        elif type_file == 4:
            DocumentGenerator.create_doc_individual_task(sheet, full_path_to_dir, doc_template)
        else:
            print(f"""
            Неизвестный тип документа: {type_file}, 
            Укажите правильный тип документа который хотите сгенерировать.
            """)