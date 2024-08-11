"""Создадим файлы в случае их отсутствия"""

import shutil
from os import path
import folium   # работа с картой


def create_map():
    """Создание карты"""
    if path.exists('places.xlsx') is False:   # проверяем наличие файла
        my_map = folium.Map(location=[55, 55], zoom_start=5)
        my_map.save("map.html")
    else:
        pass


def create_map_path():
    """Создание пути к карте"""
    if path.exists('map_path.txt') is False:   # проверяем наличие файла
        open('map_path.txt', 'w')
    else:
        pass


def create_table():
    """Создание таблицы"""
    if path.exists('places.xlsx') is False:   # проверяем наличие файла
        shutil.copyfile('places_new', 'places.xlsx')
    else:
        pass


def create_table_path():
    """Создание пути к таблице"""
    if path.exists('table_path.txt') is False:   # проверяем наличие файла
        open('table_path.txt', 'w')
    else:
        pass
