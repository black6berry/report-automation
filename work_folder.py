import os
from pathlib import Path

class WorkFolder():
    """
    Класс для работы с папкой
    """
    def __init__(self, full_path_to_dir):
        self.full_path_to_dir = full_path_to_dir


    @staticmethod
    def create_folder(full_path_to_dir) -> str:
        """
        Создает дерикторию если она не создана
        """
        if not os.path.exists(full_path_to_dir):
            os.makedirs(full_path_to_dir)
            return f"Директория {full_path_to_dir} создана успешно!"