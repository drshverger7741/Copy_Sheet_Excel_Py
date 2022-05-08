import configparser
import logging
import os

import openpyxl
import plyer
import xlwings as xw

config = configparser.ConfigParser()
config.read('settings.ini')
logging.basicConfig(filename=config.get('debug',
                                        'log_file'),
                    filemode="w",
                    level=logging.INFO,
                    format="%(asctime)s - %(levelname)s - %(funcName)s: %(lineno)d - %(message)s")
logging.info('Получение данных настройки из settings.ini')
# удаление старого финального файла при наличии
final_file = config.get('final_name_file',
                        'final_file')
if os.path.isfile(final_file):
    os.remove(final_file)

# создание финального файла
wb = openpyxl.Workbook()
wb.save(final_file)
logging.info("Создан новый файл: " + final_file)

# копирование листов
logging.info("Началось копирование листов в " + final_file)
app = xw.App(visible=config.getboolean('final_name_file',
                                       'visible'))
wb2 = xw.Book(final_file,
              update_links=config.getboolean('final_name_file',
                                             'update_links'),
              notify=config.getboolean('final_name_file',
                                       'notify'))
for key, value in config['files_sheet'].items():
    path1 = key
    wb1 = xw.Book(path1,
                  update_links=config.getboolean('every_file',
                                                 'update_links'),
                  notify=config.getboolean('every_file',
                                           'notify'))
    ws1 = wb1.sheets(value)
    ws1.api.Copy(Before=None, After=wb2.sheets(wb2.sheets.count).api)
    wb2.save()
    logging.info("Лист " + value + " из книги " + key + " успешно скопирован в " + final_file)
    wb1.close()
wb2.app.quit()

# удаление ненужных листов
wb = openpyxl.load_workbook(final_file)
sheets = wb.sheetnames
for sheet in sheets:
    if sheet not in config['files_sheet'].values():
        pfd = wb[sheet]
        wb.remove(pfd)
sheets = wb.sheetnames
wb.save(final_file)
logging.info("Листы отсортированы")

# уведомление на W10
try:
    plyer.notification.notify(message="Копирование листов успешно выполнено!",
                              app_name='Copy_Sheet_Excel_by_MFProjects.exe',
                              title="Копирование листов успешно выполнено!")
except:
    logging.error("Невозможно отобразить уведомление виндоус 10 о завершеии работы")
finally:
    logging.info("Копирование листов успешно выполнено!")
