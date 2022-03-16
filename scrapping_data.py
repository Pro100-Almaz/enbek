from pyPythonRPA.Robot import pythonRPA
import datetime

import os
import glob

# Папка для картинок (поиска кнопок)
root_path = str(__file__).split("Sources")[0].replace("/", "\\")
files = root_path + "Files\\"


def convert_date(date):
    year = ""
    month = ""
    day = ""
    for ind in range(4):
        year += date[ind]
    for ind in range(5, 7):
        month += date[ind]
    for ind in range(8, 10):
        day += date[ind]
    return day + "." + month + "." + year


def start(file_filial, code_number, login, password, begin, end):
    # Для ввода данных в SAP,as
    # период на 1 день
    today = convert_date(str(datetime.date.today()))
    yesterday = convert_date(str(datetime.date.today() - datetime.timedelta(days=1)))

    if begin != "none":
        yesterday = begin

    if end != "none":
        today = end

    # Открываем приложение и проверяем на появление
    sap_app = pythonRPA.application(r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe")
    sap_app.start()
    pythonRPA.bySelector([{"title": "SAP Logon 750", "class_name": "#32770", "backend": "uia"}]).wait_appear(30)

    # Вход в КТЖ Тестирование и потом в "Log on"
    # image1 = pythonRPA.byImage(files + "images\\main.png")
    # image1.set_confidence(0.7)
    # image1.click()
    pythonRPA.keyboard.press("enter")
    pythonRPA.bySelector(
        [{"title": "SAP", "class_name": "SAP_FRONTEND_SESSION", "backend": "uia"}, {"ctrl_index": 3}]).wait_appear(50)
    # image3 = pythonRPA.byImage(files + "images\\user_name.png")
    # image3.set_confidence(0.7)
    # image3.click()
    pythonRPA.keyboard.write(login)
    pythonRPA.keyboard.press("tab")
    # image3 = pythonRPA.byImage(files + "images\\password.png")
    # image3.set_confidence(0.7)
    # image3.click()
    pythonRPA.keyboard.write(password)
    pythonRPA.keyboard.press("enter")
    pythonRPA.sleep(1)
    # pythonRPA.bySelector([{"title":"SAP Easy Access  -  User Menu for Robot HR","class_name":"SAP_FRONTEND_SESSION","backend":"uia"}]).wait_appear(7)
    # if not pythonRPA.bySelector([{"title":"SAP Easy Access  -  User Menu for Robot HR","class_name":"SAP_FRONTEND_SESSION","backend":"uia"}]).is_exists():
    #     return "none"
    pythonRPA.keyboard.write("zhr_pa_d503")
    pythonRPA.sleep(1)
    pythonRPA.keyboard.press("enter")
    # pythonRPA.keyboard.press("tab")  # Initial date
    # pythonRPA.keyboard.write("01.11.2021")
    pythonRPA.keyboard.press("tab")
    pythonRPA.keyboard.write(yesterday)  # TODO: In working code change it to (yesterday)
    pythonRPA.keyboard.press("tab")
    pythonRPA.keyboard.write(today)  # Personal Number
    pythonRPA.keyboard.press("tab")
    pythonRPA.keyboard.press("tab")
    pythonRPA.keyboard.press("tab")
    pythonRPA.keyboard.write(code_number)  # Company Code
    pythonRPA.keyboard.press("enter")
    pythonRPA.keyboard.press("tab")
    pythonRPA.keyboard.press("tab")
    pythonRPA.keyboard.press("enter")

    pythonRPA.bySelector([{"title_re": "ZHR_PA_D503-*", "class_name": "XLMAIN", "backend": "uia"}]).wait_appear(100)
    pythonRPA.bySelector([{"title_re": "ZHR_PA_D503-*", "class_name": "XLMAIN", "backend": "uia"}]).click()
    pythonRPA.bySelector([{"title_re": "ZHR_PA_D503-*", "class_name": "XLMAIN", "backend": "uia"}])
    pythonRPA.keyboard.press("f12")
    pythonRPA.keyboard.write("ZHR_PA_D503_" + file_filial)
    pythonRPA.keyboard.press("enter")
    pythonRPA.keyboard.press("ctrl+w")
    pythonRPA.bySelector(
        [{"title": "ZHR_PA_D503", "class_name": "SAP_FRONTEND_SESSION", "backend": "uia"}]).wait_appear(10)
    pythonRPA.bySelector([{"title": "ZHR_PA_D503", "class_name": "SAP_FRONTEND_SESSION", "backend": "uia"}])
    pythonRPA.sleep(5)
    pythonRPA.keyboard.press("ESC")
    pythonRPA.sleep(1)
    pythonRPA.keyboard.write("zhr_pa_d504")
    pythonRPA.keyboard.press("enter")
    pythonRPA.keyboard.press("tab")  # Final date
    pythonRPA.keyboard.write(yesterday) # TODO: In working code change it to (yesterday)
    pythonRPA.keyboard.press("tab")
    pythonRPA.keyboard.write(today)  # Initial Date
    pythonRPA.keyboard.press("tab")
    # pythonRPA.keyboard.write(today) # Personal Number
    pythonRPA.keyboard.press("tab")
    pythonRPA.keyboard.press("tab")
    pythonRPA.keyboard.write(code_number)  # Company Code
    pythonRPA.keyboard.press("enter")
    pythonRPA.keyboard.press("tab")
    pythonRPA.keyboard.press("tab")
    pythonRPA.keyboard.press("enter")

    pythonRPA.sleep(10)
    pythonRPA.bySelector([{"title_re": "ZHR_PA_D504-*", "class_name": "XLMAIN", "backend": "uia"}]).wait_appear(100)
    pythonRPA.bySelector([{"title_re": "ZHR_PA_D504-*", "class_name": "XLMAIN", "backend": "uia"}]).click()
    pythonRPA.bySelector([{"title_re": "ZHR_PA_D504-*", "class_name": "XLMAIN", "backend": "uia"}])
    pythonRPA.keyboard.press("f12")
    pythonRPA.keyboard.write("ZHR_PA_D504_" + file_filial)
    pythonRPA.keyboard.press("enter")
    pythonRPA.keyboard.press("ctrl+w")
    return "done"


def end():
    pythonRPA.bySelector([{"title": "ZHR_PA_D504", "class_name": "SAP_FRONTEND_SESSION", "backend": "uia"}]).wait_appear(1)
    pythonRPA.bySelector([{"title":"ZHR_PA_D504","class_name":"SAP_FRONTEND_SESSION","backend":"uia"}]).close()
    pythonRPA.sleep(1)
    pythonRPA.keyboard.press("tab")
    pythonRPA.keyboard.press("enter")
    pythonRPA.bySelector([{"title": "SAP Logon 750", "class_name": "#32770", "backend": "uia"}]).wait_appear(1)
    pythonRPA.bySelector([{"title": "SAP Logon 750", "class_name": "#32770", "backend": "uia"}]).close()
    pythonRPA.sleep(1)

