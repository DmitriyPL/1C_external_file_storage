# Хранилище внешних обработок 1С

## Описание

Легкое хранилище на базе Python, помогает организовать совместнуб работу программистов над
внешними обработками. Поможет предотвратить затирание кода.

## Настройка

1. Установить Python 
2. Загрузить проект в общую директорию
3. Настроить реестр windows  
4. Настроить ```UsersConfig.json```
5. Настроить ```main.pyw```

### Установить Python

1. У всех пользователей должен быть установлен Python одной версии по одинаковому пути, например ```"C:\Program Files\Python39"```.
2. Скачать питон можно здесь ```https://www.python.org/downloads/```

### Загрузить проект в общую директорию

1. Создайти проект у себя на локальной машине.
2. Настоятельно рекомендую использовать PyCharm, сэкономите кучу нервов! Скачать бесплатную версию можно тут - ```https://www.jetbrains.com/ru-ru/pycharm/download/#section=windows```
3. Не забываем установить используемые в проекте внешние модули ```requirements.txt``` -> pip install requirements.txt (в терминале PyCharm)
4. Копируем файлы репозитория в общую директорию + venv (виртуальное окружение python которое вам создаст PyCharm)

### Настроить реестр windows

1. Добавим в контекстное меню для файлов с расширением .epf две кнопки ```захватить/поместить```
2. Необходимо воспользоваться файлом ```Добавить пункты меню.reg```:
```
Windows Registry Editor Version 5.00

[HKEY_CLASSES_ROOT\V83.ExternalProcessing\Shell\Захватить]
@="Захватить"
"Position"="Bottom"
"Icon"="%SystemRoot%\\System32\\SHELL32.dll,47"

[HKEY_CLASSES_ROOT\V83.ExternalProcessing\Shell\Захватить\command]
@="\"\\\\app1\\1C\\work\\Вложения\\Python_projects\\epf_Controll\\venv\\Scripts\\pythonw.exe\" \"\\\\app1\\1C\\work\\Вложения\\Python_projects\\epf_Controll\\main.pyw\" \"%1\" \"capture\""

[HKEY_CLASSES_ROOT\V83.ExternalProcessing\Shell\Поместить]
@="Поместить"
"Position"="Bottom"
"Icon"="%SystemRoot%\\System32\\SHELL32.dll,132"

[HKEY_CLASSES_ROOT\V83.ExternalProcessing\Shell\Поместить\command]
@="\"\\\\app1\\1C\\work\\Вложения\\Python_projects\\epf_Controll\\venv\\Scripts\\pythonw.exe\" \"\\\\app1\\1C\\work\\Вложения\\Python_projects\\epf_Controll\\main.pyw\" \"%1\" \"put\""
```
Обращаю внимание!!! Мы используем 1С 8.3 в Вашем случае ```HKEY_CLASSES_ROOT\V83.ExternalProcessing``` может иметь иное значение!!!

``` 
[HKEY_CLASSES_ROOT\V83.ExternalProcessing\Shell\Захватить\command]
@="\"\\\\app1\\1C\\work\\Вложения\\Python_projects\\epf_Controll\\venv\\Scripts\\pythonw.exe\" \"\\\\app1\\1C\\work\\Вложения\\Python_projects\\epf_Controll\\main.pyw\" \"%1\" \"capture\""

[HKEY_CLASSES_ROOT\V83.ExternalProcessing\Shell\Захватить\command]
@="[Путь до запускатора python из папки venv] [Путь до скрипта - передается параметром в запускотор] [Путь к файлу который вы захватываете/отпускаете - передается первым параметром в скрипт] [Тип операции захватить/отпустить - передается вторым параметром в скрипт]"

```
Для работы с расширениями можно использовать ```filetypesman```

4. После настройки данного файла вашим колегам останется только стартануть его! 

### Настроить ```UsersConfig.json```

```
{
    "IT-99907262": {
            "ExtForms": "\\\\1C\\1Cv80\\krow42\\Extforms\\",
            "Вложения": "\\\\1C\\1Cv80\\krow42\\Вложения\\"
    }
}

{
    "[Имя компьютера]": {
            "[Ключ 1]": [Путь 1],
            "[Ключ 2]": [Путь 2]
    }
}
```
Ключ - это то, что будет пытаться найти прога в пути захватываемого файла, что бы выбрать один из путей по которому она будет сохранять захватываемый файл

### Настроить ```main.pyw```

```
# -*- coding: utf-8 -*-

import pyperclip
import datetime
import win32con
import win32api
import logging
import shutil
import ctypes
import json
import sys
import os


CAPTURED_FILES_PATH = r"\\app1\1C\work\Вложения\Python_projects\epf_Controll\ЗахваченныеФайлы.json"
USERS_CONFIG        = r"\\app1\1C\work\Вложения\Python_projects\epf_Controll\UsersConfig.json"
OPERATION_LOG_FILE  = r"\\app1\1C\work\Вложения\Python_projects\epf_Controll\log.txt"
MAIN_DIR            = r"\\app1\1C\work"

...
```
MAIN_DIR - корневой каталог рабочей базы, по нему прога проверяет там ли вы пытаетесь захватить файл.
