from enum import Enum


class Language(str, Enum):
    RU = "RUS"
    EN = "ENG"
    PL = "PL"
    BOBR = "beaver"


class ValidatorErrComment:

    messages = {
        Language.RU: {
            "missing_column": "Отсутствует колонка: {column}",
            "wrong_header": "Колонка {index}: ожидалось '{expected}', найдено '{found}'",
            "columns_count": "Ожидалось {expected} колонок, найдено {found}",
            "miss_sheet": "Лист с названием {sheet_name} не найден",
        },
        Language.EN: {
            "missing_column": "Missing column: {column}",
            "wrong_header": "Column {index}: expected '{expected}', found '{found}'",
            "columns_count": "Expected {expected} columns, found {found}",
            "miss_sheet": "The sheet named {sheet_name} was not found.",
        },
        Language.PL: {
            "missing_column": "Brak kolumny: {column}",
            "wrong_header": "Kolumna {index}: oczekiwano '{expected}', znaleziono '{found}'",
            "columns_count": "Oczekiwano {expected} kolumn, znaleziono {found}",
            "miss_sheet": "Arkusz o nazwie {sheet_name} nie został znaleziony.",
        },
        Language.BOBR: {
            "missing_column": "🦫 Bóbr nie znalazł kolumny: {column} 🦫",
            "wrong_header": "🦫 Bóbr mówi: kolumna {index} powinna być '{expected}', ale jest '{found}' 🦫",
            "columns_count": "🦫 Bóbr liczył {expected} kolumn, ale znalazł {found} 🦫",
            "miss_sheet": "🦫 Bóbr nie znalazł arkusza o nazwie {sheet_name} 🦫",
        }
    }

    def __init__(self, language: Language):
        self.language = language

    def get(self, key: str, **kwargs):
        template = self.messages[self.language][key]
        return template.format(**kwargs)