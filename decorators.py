from time import time
from collections.abc import Callable


def count_time_of_programm(func: Callable) -> Callable:
    def wrapper(*args, **kwargs):
        print('Программа начала работу.')
        start_time = time()
        res = func(*args, **kwargs)
        end_time = time()
        ex_time = end_time-start_time
        print('Программа закончила работу.')
        print(f'Время выполнения программы {ex_time:.2f} сек.')
        return res
    return wrapper
