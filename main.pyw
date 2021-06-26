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

ID_NO = 2


def create_logger(name, log_level=logging.DEBUG, stdout=False, file=None):
    '''
    Создает логера, есть возможность создать логера с выводом в stdout или в файл или туда и туда.
    '''

    logger = logging.getLogger(name)
    logger.setLevel(log_level)
    formatter = logging.Formatter(fmt='[%(asctime)s] - %(name)s - %(levelname).1s - %(message)s',
                                  datefmt='%Y.%m.%d %H:%M:%S')

    if file is not None:
        fh = logging.FileHandler(file, encoding='utf-8-sig')
        fh.setLevel(log_level)
        fh.setFormatter(formatter)
        logger.addHandler(fh)

    if stdout:
        ch = logging.StreamHandler()
        ch.setLevel(log_level)
        ch.setFormatter(formatter)
        logger.addHandler(ch)

    return logger


def get_captured_files():
    with open(CAPTURED_FILES_PATH, 'r', encoding='utf-8-sig') as f:
        return json.load(f)


def get_user_paths():
    with open(USERS_CONFIG, 'r', encoding='utf-8-sig') as f:
        users_data = json.load(f)
        paths = users_data.get(os.environ['COMPUTERNAME'])
    return paths


def copy_file(current_path, operation_type, copy_path=None):
    if operation_type == "capture":
        lst_current_path = current_path.lower().split('\\')

        user_paths = get_user_paths()
        if user_paths is None:
            return False

        target_path = None
        for mask, path in user_paths.items():
            if mask.lower() in lst_current_path:
                poz = current_path.find(mask)
                file_sub_path = current_path[poz + len(mask + '\\'):]
                target_path = user_paths[mask] + file_sub_path

        if target_path is None:
            return False

        path1 = current_path
        path2 = target_path
        pyperclip.copy(path2)

    else:
        path1 = copy_path
        path2 = current_path
        # to force deletion of a file set it to normal
        win32api.SetFileAttributes(path2, win32con.FILE_ATTRIBUTE_NORMAL)

    msg    = 'Файл будет скопирован в: {}'.format(path2)
    answer = ctypes.windll.user32.MessageBoxW(0, msg, "Хотите скопировать?", 1)
    if answer == ID_NO:
        return False

    shutil.copy(path1, path2)
    return path2


def write_file_in_captured_files(files, captured_file, capture_param):
    file_name = os.path.basename(captured_file)

    if capture_param is not None:
        user  = capture_param['user']
        date  = capture_param['date']
        msg   = "Обработка: {0} уже захвачена!\r\r{1} в {2}".format(file_name, user, date)
        title = 'Внимание! Обработка уже захвачена!'
        warning(msg, title)
    else:
        msg    = 'Файл будет захвачен!'
        answer = ctypes.windll.user32.MessageBoxW(0, msg, "Захватить?", 1)
        if answer == ID_NO:
            return False

        comp = os.environ['COMPUTERNAME']
        user = os.environ['USERNAME']

        # Дата(«15.12.2015 20:42:22») // 15.12.2015 20:42:22
        date = str(datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S"))

        copy_path = copy_file(captured_file, "capture")
        win32api.SetFileAttributes(captured_file, win32con.FILE_ATTRIBUTE_READONLY) # make the file read only

        files[captured_file] = {
            'date': date,
            'comp': comp,
            'user': user,
            'copy_path': copy_path
        }

        with open(CAPTURED_FILES_PATH, 'w', encoding='utf-8-sig') as f:
            json.dump(files, f, indent=4, ensure_ascii=False)

        msg   = "Обработка {0} удачно захвачена!".format(file_name)
        title = 'Все прошло успешно!'
        warning(msg, title)
        logger.info('{} захватил {}'.format(user, captured_file))


def remove_captured_file(files, captured_file):
    captured_file_lower = captured_file.lower()
    for path, param in files.items():
        if path.lower() == captured_file_lower or param['copy_path'].lower() == captured_file_lower:
            del files[path]
            return path


def del_file_from_captured_files(files, captured_file, capture_param):
    if capture_param is None:
        msg   = "Обработка не была захвачена!"
        title = 'Внимание обработка не была захвачена!'
        warning(msg, title)
    else:

        file_name = os.path.basename(captured_file)
        user      = capture_param['user']
        date      = capture_param['date']
        copy_path = capture_param['copy_path']

        if os.environ['COMPUTERNAME'] != capture_param['comp']:
            msg = "Обработка была захвачена не Вами, a {0}\r\r" \
                  "в {1}\r\r" \
                  "поэтому Вы не можете ее вернуть!".format(user, date)
            title = 'Ошибочка вышла!'
            warning(msg, title)
        else:
            msg = 'Файл будет помещен!'
            answer = ctypes.windll.user32.MessageBoxW(0, msg, "Поместить?", 1)
            if answer == ID_NO:
                return False

            path = remove_captured_file(files, captured_file)
            copy_file(path, "put", copy_path)

            with open(CAPTURED_FILES_PATH, 'w', encoding='utf-8-sig') as f:
                json.dump(files, f, indent=4, ensure_ascii=False)

            msg   = "Обработка {0} удачно помещена!".format(file_name)
            title = 'Все прошло успешно!'
            warning(msg, title)
            logger.info('{} отпустил {}'.format(user, captured_file))


def warning(msg, title):
    ctypes.windll.user32.MessageBoxW(0, msg, title, 0)


def file_dir_is_ok(file):
    file_lower = file.lower()
    if not file_lower.startswith(MAIN_DIR.lower()):
        msg = "Вы пытаетесь захватить обработку вне рабочего каталога!\r\rДоверенные пути:\r\r" \
              "1.'\\\\app1\\1C\\work'\r\r" \
              "2.'\\\\app1\\1C\\work\\Вложения'\r\r" \
              "3.'\\\\app1\\1C\\work\\ExtForms'"
        title = r'Неверная директория!'
        warning(msg, title)
        return False
    return True


def get_captured_param(files, captured_file, operation_type):
    captured_file_lower = captured_file.lower()
    if operation_type == "capture":
        for path in files.keys():
            if path.lower() == captured_file_lower:
                return files.get(path)
    else:
        for path, param in files.items():
            if path.lower() == captured_file_lower or param['copy_path'].lower() == captured_file_lower:
                return files.get(path)
    return None


def main():
    captured_file  = sys.argv[1]
    operation_type = sys.argv[2]
    files          = get_captured_files()
    capture_param  = get_captured_param(files, captured_file, operation_type)

    if operation_type == "capture":
        if file_dir_is_ok(captured_file):
            write_file_in_captured_files(files, captured_file, capture_param)
    elif operation_type == "put":
        del_file_from_captured_files(files, captured_file, capture_param)
    else:
        warning(captured_file, "else")


if __name__ == '__main__':
    try:
        logger = create_logger("log", file=OPERATION_LOG_FILE)
        main()
    except Exception as e:
        logger.error(e)