"""Main File. Создание окна программы с картой."""

import os
import subprocess           # работа с папками
import map_script as ms     # скрипт с функциями и кодом по созданию карты
import time
from datetime import datetime
import webview              # браузер
import webview.menu as wm
from mss import mss         # скриншоты
import tkinter.messagebox as mb


def active_window_content():
    """Начальное окно с текущей картой"""
    active_window = webview.active_window()
    if active_window:
        active_window.load_url('map.html')


def map_show():
    """Обновить окно с картой"""
    exec(open('map_script.py').read())
    active_window = webview.active_window()
    if active_window:     # перепрыгнем со страницы на страницу
        active_window.load_html('<h1>Загрузка...</h1>')
        time.sleep(0.5)
        active_window.load_url('map.html')


def map_screenshot():
    """Сделать скриншот"""
    with mss() as sct:
        if os.path.exists('screenshots') is True:   # проверим, есть ли папка
            sct.shot(output=f'screenshots/map_screen {datetime.now().strftime("%d%m%y_%H%M%S")}.png')
        else:
            os.mkdir('screenshots')
            sct.shot(output=f'screenshots/map_screen {datetime.now().strftime("%d%m%y_%H%M%S")}.png')
        # msg = "       Скриншот сделан       "
        # mb.showinfo("Скриншот", msg)


def map_screens_folder():
    """Открыть папку со скриншотами"""
    if os.path.exists('screenshots') is True:   # проверим, есть ли папка
        subprocess.call(f'explorer {os.getcwd()}\\screenshots')
    else:
        os.mkdir('screenshots')
        subprocess.call(f'explorer {os.getcwd()}\\screenshots')


def home():
    """Домой, возвращает на карту"""
    active_window = webview.active_window()
    if active_window:     # перепрыгнем со страницы на страницу
        active_window.load_html('<h1>Загрузка...</h1>')
        time.sleep(0.2)
        active_window.load_url('map.html')


def prog_help():
    """Справка"""
    os.startfile('help.pdf')


def prog_about():
    """О программе"""
    msg = ("    << Твоя персональная география >> \n\n"
           "Использовались:   Python v 3.11, OpenStreetMap\n\n"
           "Разработал:   ©️ Рощупкин Д.  aka Demston")
    mb.showinfo("О программе", msg)


exec(open('map_script.py').read())  # выполним скрипт: проверим файлы, создадим карту

if __name__ == '__main__':
    window_main = webview.create_window('Твоя персональная география', 'map.html', )    # создание окна программы

    menu_items = [
        wm.MenuAction('⌂', home),
        wm.Menu(
            'Файл',
            [
                wm.MenuAction('Открыть последнюю таблицу', ms.open_current_file),
                wm.MenuAction('Открыть таблицу с адресами', ms.open_file),
                wm.MenuAction('Выбрать активную таблицу', ms.select_file),
                wm.MenuSeparator(),
                wm.MenuAction('Сохранить карту', ms.save_map),
                wm.MenuAction('Сохранить карту как', ms.save_map_as)
            ],
        ),
        wm.Menu(
            'Карта',
            [
                wm.MenuAction('Сформировать карту', map_show),
                wm.MenuSeparator(),
                wm.MenuAction('Сделать скриншот', map_screenshot),
                wm.MenuAction('Открыть папку скриншотов', map_screens_folder)
             ]
        ),
        wm.Menu(
            'Справка',
            [
                wm.MenuAction('Справка', prog_help),
                wm.MenuSeparator(),
                wm.MenuAction('О программе', prog_about)
            ]
        )
    ]

    webview.start(menu=menu_items)
