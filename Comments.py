from datetime import datetime, date
from dateutil.relativedelta import relativedelta


class Comment():
    count = 0
    start_time = datetime.now()
    actual_datetime_1 = start_time
    actual_datetime_2 = start_time

    def __init__(self, comment: str = '') -> None:
        self.comment = comment

    @classmethod
    def counter(cls) -> None:
        cls.count += 1
        cls.actual_datetime_1, cls.actual_datetime_2 = cls.actual_datetime_1, datetime.now()

    def print_time(self) -> None:
        Comment.counter()

        if Comment.count > 2:
            print(' ' * (len(str(Comment.count - 1)) + 2), '-Выполнено за: ',
                  Comment.actual_datetime_2 - Comment.actual_datetime_1, sep='')
        elif Comment.count == 2:
            print(' ' * (len(str(Comment.count - 1)) + 2),
                  '-Выполнено за: ', datetime.now() - Comment.start_time, sep='')

        print(Comment.count, '. ', self.comment, sep='')

    def print_first_time(self) -> None:
        print(f'\nВремя начала сборки прогноза: {Comment.start_time}')

        Comment.counter()

        print(Comment.count, '. ', self.comment, sep='')

    def print_final_time(self) -> None:
        Comment.counter()

        print(' ' * (len(str(Comment.count - 1)) + 2), '-Выполнено за: ',
              Comment.actual_datetime_2 - Comment.actual_datetime_1, sep='')
        print(f'Время окончания работы скрипта: {Comment.actual_datetime_2}')
        print(
            f'Итоговое время работы скрипта: {Comment.actual_datetime_2 - Comment.start_time}\n')
