"""Пользовательские исключения приложения.

Все ошибки, которые может увидеть пользователь, наследуются от AppError
и несут понятное сообщение на русском языке.
"""


class AppError(Exception):
    """Базовое исключение приложения с человеческим сообщением."""

    default_message = "Произошла ошибка."

    def __init__(self, message: str | None = None):
        self.message = message or self.default_message
        super().__init__(self.message)


class PathNotFoundError(AppError):
    """Файл по указанному пути не найден."""

    default_message = "Файл не найден."


class NotExcelFileError(AppError):
    """Файл не является книгой Excel (неверная сигнатура)."""

    default_message = "Файл не является таблицей Excel (.xls или .xlsx)."


class EncryptedFileError(AppError):
    """Файл защищён паролем."""

    default_message = "Файл защищён паролем. Снимите защиту и попробуйте снова."


class CorruptedFileError(AppError):
    """Файл повреждён или не читается."""

    default_message = "Не удалось прочитать файл: он повреждён или имеет неподдерживаемый формат."


class CoordinateParseError(AppError):
    """Координата ячейки не распознана."""

    default_message = "Не удалось распознать координату ячейки (пример: A6)."


class EmptyKeySetError(AppError):
    """По указанной координате не найдено ни одного ключа."""

    default_message = "По указанной координате не найдено ни одного значения ключа."


class NoMatchError(AppError):
    """Не найдено совпадений ключей с базовым набором."""

    default_message = "Не найдено столбцов/строк с совпадающими ключами."


class OutputWriteError(AppError):
    """Не удалось записать итоговый файл."""

    default_message = "Не удалось записать итоговый файл."
