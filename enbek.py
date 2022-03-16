import shutil
import glob
import os
import time
import keyboard
from glob import glob
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.ui import WebDriverWait, Select
from webdriver_manager.chrome import ChromeDriverManager

add_ecp = False


class Enbek:

    def __init__(self, robot_path, login, psw, path_to_ecp='', psw_ecp='', debug=False):
        self.debug = debug
        self.driver = None
        self.driver_data = [login, psw]
        self.robot_path = robot_path
        self.downloads_path = os.path.join(os.environ['USERPROFILE'], "Downloads\\").replace("/", "\\")
        self.temp_path = self.robot_path + "Temp\\Enbek"
        self.path = {"downloads": {"title": "Downloads", "path": self.downloads_path, "dir": self.downloads_path[:-1]},
                     "temp": {"title": "Temp", "path": self.temp_path, "dir": self.temp_path[:-1]},
                     "enbek_files": {"title": "Enbek_files", "path": self.temp_path + "Enbek_files\\",
                                     "dir": self.temp_path + "Enbek_files"}}
        self.url = {
            "login": "https://www.enbek.kz/docs/ru/user",
            "list": "https://www.enbek.kz/ru/cabinet/dogovor/list/good",
            "add": "https://www.enbek.kz/ru/cabinet/dogovor/add",
            "vacancies": "https://www.enbek.kz/ru/cabinet/vacint/myvac",
            "all": "https://www.enbek.kz/ru/cabinet/dogovor/list/all"
        }
        # Data containers
        self.anchor = None
        self.path_to_ecp = path_to_ecp
        self.passw = psw_ecp
        print("\t")

    def _fill_new_vac(self, data):
        driver = self.driver

        # TODO Профессиональная область
        select_prof_oblast = driver.find_element(By.XPATH, '//select[@name="profObl"]')
        select_prof_oblast.click()
        prof_obl = data['Профессиональная область']
        print("prof obl", prof_obl)
        try:
            prof_oblast_vybor = driver.find_element(By.XPATH, f'//option[text()="{prof_obl}"]')
        except Exception as e:
            print("Проф область пустая")
            return 0
        prof_oblast_vybor.click()

        # TODO Наименование должности
        time.sleep(2)
        root_dogovor = driver.find_element(By.XPATH, '//div[select[@name="eduProf"]]/span/span/span')
        root_dogovor.click()
        dol = data['Наименование должности']
        li = driver.find_element(By.XPATH, f'//li[text()="{dol}"]')
        li.click()

        # TODO Уточнение должности
        utochnenie_dolj = driver.find_element(By.XPATH, '//input[@name="utprof"]')
        utochnenie_dolj.send_keys(data['Уточнение должности'])

        # TODO Адрес
        obl = data['Адрес']['obl']
        if obl == "Empty":
            raise ValueError("Empty Data")
        center = data['Адрес']['center']
        nas_punkt = data['Адрес']['nas_punkt']
        adres = data['Адрес']['adres']

        button_obl = driver.find_element(By.XPATH, "//Button[text()='Выбрать']")
        button_obl.click()

        self._sel_wait_el(By.XPATH, "//div[@class='modal-content' and //h4[text()='Справочник регионов']]")
        root_adres = driver.find_element(By.XPATH,
                                         "//div[@class='modal-content' and //h4[text()='Справочник регионов']]")

        li_obl = root_adres.find_element(By.XPATH, "//li[span[text()='" + obl + "']]")
        li_obl.click()

        time.sleep(1)
        self._sel_wait_el(By.XPATH, "//li[span[text()='" + center + "']]")
        li_center = root_adres.find_element(By.XPATH, "//li[span[text()='" + center + "']]")
        li_center.click()

        button_adres = root_adres.find_element(By.XPATH, "//button[text()='Выбор']")
        button_adres.click()
        if not self._sel_wait_el(By.XPATH, "//div[@class='modal-content' and //h4[text()='Справочник регионов']]",
                                 appear=False):
            raise ValueError("Время ожидания истекло: Адрес не выбран")

        input_adres = driver.find_element(By.XPATH, "//input[@name='placerab']")
        input_adres.send_keys(adres)

        time.sleep(1)
        span_nas_punkt = driver.find_element(By.XPATH,
                                             '//div[div[label[contains(text(), "Населённый пункт")]]]/div[2]/span/span/span')
        span_nas_punkt.click()

        # root_nas = driver.find_element(By.XPATH, "//span[@class='select2-container select2-container--default select2-container--open']")
        # input_nas_punkt = root_nas.find_element(By.XPATH, "//input[@class='select2-search__field']")
        # input_nas_punkt.send_keys(nas_punkt)

        if self._sel_wait_el(By.XPATH,
                             f"//li[@class='select2-results__option select2-results__option--highlighted' and contains(text(),'{nas_punkt}')]",
                             sec=5):
            li_nas = driver.find_element(By.XPATH,
                                         f"//li[@class='select2-results__option select2-results__option--highlighted' and contains(text(),'{nas_punkt}')]")
        elif self._sel_wait_el(By.XPATH, f"//li[@class='select2-results__option' and contains(text(),'{nas_punkt}')]",
                               sec=5):
            li_nas = driver.find_element(By.XPATH,
                                         f"//li[@class='select2-results__option' and contains(text(),'{nas_punkt}')]")
        else:
            raise ValueError("Время ожидания истекло: Должность не найдена")
        li_nas.click()

        # TODO Режим работы
        rezhim_raboty_select = driver.find_element(By.XPATH, '//div[select[@name="regim"]]/span/span/span')
        rezhim_raboty_select.click()
        rezhim_rab = data['Режим работы']
        vybor = driver.find_element(By.XPATH, f'//li[text()="{rezhim_rab}"]')
        vybor.click()

        # TODO Характер работы
        haracter_raboty_select = driver.find_element(By.XPATH, '//div[select[@name="zan"]]/span/span/span')
        haracter_raboty_select.click()
        haracter_rab = data['Характер работы']
        vybor_haracter_raboty = driver.find_element(By.XPATH, f'//li[text()="{haracter_rab}"]')
        vybor_haracter_raboty.click()

        # TODO Условия труда
        uslovie_truda = driver.find_element(By.XPATH, '//div[select[@name="nac"]]/span/span/span')
        uslovie_truda.click()
        usl_truda = data['Условия труда']
        uslovie_truda_vybor = driver.find_element(By.XPATH, f'//li[text()="{usl_truda}"]')
        uslovie_truda_vybor.click()

        # TODO Стажировка
        stazhirovka_select = driver.find_element(By.XPATH, '//div[select[@name="stager"]]/span/span/span')
        stazhirovka_select.click()
        stazhirovka = data['Стажировка']
        stazhirovka_vybor = driver.find_element(By.XPATH, f'//li[text()="{stazhirovka}"]')
        stazhirovka_vybor.click()

        # TODO Оплата труда (тенге)
        oplata_ot = driver.find_element(By.XPATH, '//input[@name="oplata_ot"]')
        oplata_ot.send_keys(data['Оплата труда (тенге)'])

        # TODO Вакантные места
        vac_mesta = driver.find_element(By.XPATH, '//input[@name="kolvak"]')
        vac_mesta.clear()
        vac_mesta.send_keys(data['Вакантные места'])

        # TODO Стаж работы по специальности
        staj_raboty = driver.find_element(By.XPATH, '//div[select[@name="stag"]]/span/span/span')
        staj_raboty.click()
        staj = data['Стаж работы по специальности']
        staj_raboty_vybor = driver.find_element(By.XPATH, f'//li[text()="{staj}"]')
        staj_raboty_vybor.click()

        # TODO Профессиональные навыки
        navyk = data['Профессиональные навыки']
        for i in range(len(navyk)):
            prof_navyki = driver.find_element(By.XPATH,
                                              f'//div[div[label[contains(text(), "Профессиональные навыки")]]]/div[2]/span/span/span/ul/li[{i + 1}]/input')
            prof_navyki.send_keys(navyk[i])
            time.sleep(1)
            WebDriverWait(driver, 10).until(
                ec.presence_of_element_located((By.XPATH, f'//li[text()="{navyk[i]}"]')))
            prof_navyki_vybor = driver.find_element(By.XPATH, f'//li[text()="{navyk[i]}"]')
            prof_navyki_vybor.click()

        # TODO Личные качества
        lich_kachesctva = data['Личные качества']
        for i in range(len(lich_kachesctva)):
            kachestva = driver.find_element(By.XPATH,
                                            f'//div[div[label[contains(text(), "Личные качества")]]]/div[2]/span/span/span/ul/li[{i + 1}]/input')
            kachestva.send_keys(lich_kachesctva[i])
            time.sleep(1)
            WebDriverWait(driver, 10).until(
                ec.presence_of_element_located((By.XPATH, f'//li[text()="{lich_kachesctva[i]}"]')))
            kachestva_vybor = driver.find_element(By.XPATH, f'//li[text()="{lich_kachesctva[i]}"]')
            kachestva_vybor.click()

        # TODO Уровень образования
        uroven_raboty = driver.find_element(By.XPATH, '//div[select[@name="urobr"]]/span/span/span')
        uroven_raboty.click()
        uroven_obrzovaniya = data['Уровень образования']
        uroven_raboty_vybor = driver.find_element(By.XPATH, f'//li[text()="{uroven_obrzovaniya}"]')
        uroven_raboty_vybor.click()

        # TODO Должностные обязаности
        driver.switch_to.frame(driver.find_element('xpath', '//iframe[@class="tox-edit-area__iframe"]'))
        obiyaz = driver.find_element(By.XPATH, '//body[@contenteditable="true"]')
        obiyaz.send_keys(data['Должностные обязанности'])
        driver.switch_to.default_content()

        # TODO Срок публикации вакансии на портале
        srok_publ = driver.find_element(By.XPATH, '//div[select[@name="prod_hran"]]/span/span/span')
        srok_publ.click()
        srok = data['Срок публикации вакансии на портале']
        srok_publ_vybor = driver.find_element(By.XPATH, f'//li[text()="{srok}"]')
        srok_publ_vybor.click()

        submit_button = driver.find_element(By.XPATH, '//button[contains(text(), "Сохранить вакансию")]')
        # submit_button.click()

    # Files
    def _mkdir(self):
        """Пробегается по self.path, создает их, если не найдены"""
        path = self.path
        for folder in path:
            current_folder = path[folder]["dir"]
            if not os.path.isdir(current_folder):
                os.makedirs(current_folder)

    def _clean(self, var="*", del_sub_dirs=False):
        """Чистит все файлы в self.path"""
        path = self.path
        self._mkdir()
        ignore = ["enbek_files"]
        for folder in path:
            if (var == "*" and folder not in ignore) or var == folder:
                current_folder = path[folder]["path"]
                children = glob.glob(current_folder + "*")
                for child in children:
                    if os.path.isdir(child):
                        print(child, "is a dir")
                        if del_sub_dirs and folder != "temp":
                            shutil.rmtree(child)
                    elif os.path.isfile(child) or os.path.islink(child):
                        print(child, "is a file")
                        os.remove(child)
                if var != "*":
                    break

    # General

    def _sel_wait_el(self, by, selector, sec=60, appear=True):
        """Ожидание элемента появление или изчезновение подается через bool 'appear'"""
        driver = self.driver
        time.sleep(0.3)
        try:
            if appear:
                WebDriverWait(driver, sec).until(ec.presence_of_element_located((by, selector)))
            else:
                WebDriverWait(driver, sec).until_not(ec.presence_of_element_located((by, selector)))
            return True
        except:
            return False
        finally:
            time.sleep(0.2)

    def _sel_init(self):
        """Запуск драйвера и логин в enbek. Объект драйвера создается именно здесь"""
        # Driver init
        try:
            self.driver = webdriver.Chrome(
                executable_path=r'C:\Users\audit_2\.wdm\drivers\chromedriver\win32\92.0.4515.107\chromedriver.exe')
        except Exception as e:
            self.driver = webdriver.Chrome(ChromeDriverManager().install())
        self.driver.maximize_window()
        self.driver.switch_to.window(self.driver.current_window_handle)
        driver = self.driver
        driver.get(self.url["list"])
        login = "//input[@placeholder='Логин или E-mail']"
        login_2 = "//input[@placeholder='Логин или Email']"
        try:
            WebDriverWait(driver, 5).until(ec.presence_of_element_located((By.XPATH, login)))
            driver.find_element(By.XPATH, login).send_keys(self.driver_data[0])
        except:
            WebDriverWait(driver, 5).until(ec.presence_of_element_located((By.XPATH, login_2)))
            driver.find_element(By.XPATH, login_2).send_keys(self.driver_data[0])
        driver.find_element(By.XPATH, "//input[@placeholder='Пароль']").send_keys(self.driver_data[1])
        driver.find_element(By.XPATH, "//input[@placeholder='Пароль']").send_keys(Keys.RETURN)
        driver.get(self.url["list"])

    # create_dogovor
    def _check_iin(self, root, counter=1):
        while counter:
            self._sel_wait_el(By.XPATH,
                              ".//input[@name='IIN']/parent::div/parent::div/div[@class='mdb-ehr-progress' and @style='display: block;']",
                              5)
            self._sel_wait_el(By.XPATH,
                              ".//input[@name='IIN']/parent::div/parent::div/div[@class='mdb-ehr-progress' and @style='display: none;']",
                              5)
            time.sleep(1)
            if counter > 0:
                if self._sel_wait_el(By.XPATH, ".//div[@class = 'mdb-ehr-error-box']/ul/li", 5):
                    raise ValueError("death")
                input_fam = root.find_element(By.XPATH, ".//input[@name='FAM']")
                if len(input_fam.get_attribute("value")) > 1:
                    return True
            else:
                counter -= 1
        return False

    def _fill_iin(self, root, iin):
        self.anchor.click()

        input_iin = root.find_element(By.XPATH, ".//input[@name='IIN']")
        input_iin.send_keys(iin)
        button_iin = root.find_element(By.XPATH, ".//input[@name='IIN']/parent::div/span/button")
        button_iin.click()
        flag = self._check_iin(root, 2)
        if not flag:
            raise ValueError("Время ожидания истекло: не найден по ИИН")
        print("self._fill_iin > done")

    def _fill_string(self, root, num, dol):
        self.anchor.click()

        input_num = root.find_element(By.XPATH, ".//input[@name='numDogovor']")
        input_num.send_keys(num)

        input_dol = root.find_element(By.XPATH, ".//input[@name='shtatDolj']")
        input_dol.send_keys(dol)

        print("self._fill_string > done")

    def _fill_select(self, root, srok, vid):
        self.anchor.click()

        select_srok = Select(root.find_element(By.XPATH, ".//select[@name='dContractCate']"))
        select_srok.select_by_visible_text(srok)

        select_vid = Select(root.find_element(By.XPATH, ".//select[@name='partTime']"))
        select_vid.select_by_visible_text(vid)

        print("self._fill_select > done")

    def login(self):
        try:
            self.driver = webdriver.Chrome(
                r'C:\Users\seikolabs\.wdm\drivers\chromedriver\win32\90.0.4430.24\chromedriver.exe')
        except Exception as e:
            print(e)
            self.driver = webdriver.Chrome(ChromeDriverManager().install())
        self.driver.maximize_window()
        self.driver.switch_to.window(self.driver.current_window_handle)
        driver = self.driver
        driver.get(self.url["vacancies"])
        login = "//input[@placeholder='Логин или E-mail']"
        WebDriverWait(driver, 10).until(ec.presence_of_element_located((By.XPATH, login)))
        driver.find_element(By.XPATH, login).send_keys(self.driver_data[0])
        driver.find_element(By.XPATH, "//input[@placeholder='Пароль']").send_keys(self.driver_data[1])
        driver.find_element(By.XPATH, "//input[@placeholder='Пароль']").send_keys(Keys.RETURN)
        driver.get(self.url["vacancies"])

    def add_vac(self, data):
        if not self.driver:
            self.login()

        driver = self.driver
        try:
            WebDriverWait(driver, 20).until(
                ec.presence_of_element_located((By.XPATH, '//a[contains(text(), "Добавить")]')))
        except:
            driver.get(self.url["vacancies"])
            WebDriverWait(driver, 20).until(
                ec.presence_of_element_located((By.XPATH, '//a[contains(text(), "Добавить")]')))

        add_button = driver.find_element(By.XPATH, '//a[contains(text(), "Добавить")]')
        add_button.click()

        WebDriverWait(driver, 10).until(
            ec.presence_of_element_located((By.XPATH, '//h3[text()="Добавление вакансии"]')))

        try:
            if self._fill_new_vac(data):
                time.sleep(2)
                return 1
            else:
                return 0
        except Exception as e:
            print(e)
            return 0

        # submit_button = driver.find_element(By.XPATH, '//button[contains(text(), "Сохранить вакансию")]')
        # submit_button.click()

    def _fill_date(self, root, dogovor, nachalo, konec=""):
        self.anchor.click()

        input_dogovor = root.find_element(By.XPATH, ".//input[@name='dateZakDogovor']")
        input_dogovor.click()
        input_dogovor.send_keys(dogovor)
        self.anchor.click()

        input_nachalo = root.find_element(By.XPATH, ".//input[@name='dateBegDogovor']")
        input_nachalo.click()
        input_nachalo.send_keys(nachalo)
        self.anchor.click()

        if len(konec):
            input_konec = root.find_element(By.XPATH, ".//input[@name='dateEndDogovor']")
            input_konec.click()
            input_konec.send_keys(konec)
            self.anchor.click()

        print("self._fill_date > done")

    def _fill_dol(self, root, dol):
        driver = self.driver
        self.anchor.click()

        span_dogovor = root.find_element(By.XPATH,
                                         ".//label[text()='Должность ']/parent::div//span[@class='selection']")
        span_dogovor.click()

        root_dogovor = driver.find_element(By.XPATH,
                                           "//span[@class='select2-container select2-container--default select2-container--open']")

        input_dogovor = root_dogovor.find_element(By.XPATH, ".//input[@class='select2-search__field']")

        input_dogovor.send_keys(dol)

        if self._sel_wait_el(By.XPATH,
                             ".//li[@class='select2-results__option select2-results__option--highlighted' and text()='" + dol + "']",
                             sec=10):
            li_dogovor = driver.find_element(By.XPATH,
                                             ".//li[@class='select2-results__option select2-results__option--highlighted' and text()='" + dol + "']")
        elif self._sel_wait_el(By.XPATH, ".//li[@class='select2-results__option' and text()='" + dol + "']", sec=10):
            li_dogovor = driver.find_element(By.XPATH,
                                             ".//li[@class='select2-results__option' and text()='" + dol + "']")
        else:
            raise ValueError("Время ожидания истекло: Должность не найдена")
        li_dogovor.click()
        if not self._sel_wait_el(By.CSS_SELECTOR, "span.select2-container--open input.select2-search__field",
                                 appear=False):
            raise ValueError("Время ожидания истекло: Должность не выбрана")

        print("self._fill_dol > done")

    def _fill_adres(self, root, obl, center, adres, nas_punkt):
        driver = self.driver
        self.anchor.click()

        button_obl = root.find_element(By.XPATH, ".//Button[text()='Выбрать']")
        button_obl.click()

        self._sel_wait_el(By.XPATH, ".//div[@class='modal-content' and //h4[text()='Справочник регионов']]")
        root_adres = driver.find_element(By.XPATH,
                                         ".//div[@class='modal-content' and //h4[text()='Справочник регионов']]")

        li_obl = root_adres.find_element(By.XPATH, ".//li[span[text()='" + obl + "']]")
        li_obl.click()

        time.sleep(1)
        self._sel_wait_el(By.XPATH, ".//li[span[text()='" + center + "']]")
        li_center = root_adres.find_element(By.XPATH, ".//li[span[text()='" + center + "']]")
        li_center.click()

        button_adres = root_adres.find_element(By.XPATH, ".//button[text()='Выбор']")
        button_adres.click()
        if not self._sel_wait_el(By.XPATH, ".//div[@class='modal-content' and //h4[text()='Справочник регионов']]",
                                 appear=False):
            raise ValueError("Адрес не выбран")

        input_adres = root.find_element(By.XPATH, ".//input[@name='workPlace']")
        input_adres.send_keys(adres)

        if obl not in ["Нур-Султан", "Алматы", "Шымкент"]:
            span_nas_punkt = root.find_element(By.XPATH,
                                               ".//label[text()='Населённый пункт ']/parent::div//span[@class='selection']")
            span_nas_punkt.click()

            root_nas = driver.find_element(By.XPATH,
                                           "//span[@class='select2-container select2-container--default select2-container--open']")
            input_nas_punkt = root_nas.find_element(By.XPATH, ".//input[@class='select2-search__field']")
            input_nas_punkt.send_keys(nas_punkt)

            if self._sel_wait_el(By.XPATH,
                                 ".//li[@class='select2-results__option select2-results__option--highlighted' and contains(text(),'{nas_punkt}')]",
                                 sec=10):
                li_nas = driver.find_element(By.XPATH,
                                             ".//li[@class='select2-results__option select2-results__option--highlighted' and contains(text(),'{nas_punkt}')]")
            elif self._sel_wait_el(By.XPATH,
                                   ".//li[@class='select2-results__option' and contains(text(),'{nas_punkt}')]",
                                   sec=10):
                li_nas = driver.find_element(By.XPATH,
                                             ".//li[@class='select2-results__option' and contains(text(),'{nas_punkt}')]")
            else:
                raise ValueError("Время ожидания истекло: Должность не найдена")
            li_nas.click()
            if not self._sel_wait_el(By.CSS_SELECTOR, "span.select2-container--open input.select2-search__field",
                                     appear=False):
                raise ValueError("Время ожидания истекло: Должность не выбрана")

    def _fill_rezhim(self, root, rezhim, stavka=False):
        self.anchor.click()

        select_rezhim = Select(root.find_element(By.XPATH, ".//select[@name='workingHours']"))
        select_rezhim.select_by_visible_text(rezhim)

        if stavka:
            self._sel_wait_el(By.XPATH, ".//input[@name='tariffRate']")
            input_stavka = root.find_element(By.XPATH, ".//input[@name='tariffRate']")
            input_stavka.send_keys(stavka)

        print("self._fill_rejim > done")

    def _fill_dog(self, data):
        driver = self.driver
        driver.get(self.url["add"])
        if self._sel_wait_el(By.XPATH, "//h3[text()='Добавление договора']"):
            self.anchor = driver.find_element(By.XPATH, "//h3[text()='Добавление договора']")
            root = driver.find_element(By.CSS_SELECTOR, ".content")

            iin_iin = data["ИИН"]
            self._fill_iin(root, iin_iin)

            string_num = data["Номер договора"]
            string_dol = data["Штатная должность"].strip()
            self._fill_string(root, string_num, string_dol)

            select_srok = data["Срок договора"]
            select_vid = data["Вид работы"]
            self._fill_select(root, select_srok, select_vid)

            rezhim = data["Режим рабочего времени"].split(", ")
            rezhim_rezhim = rezhim[0]
            rezhim_stavka = rezhim[1] if len(rezhim) > 1 else False
            self._fill_rezhim(root, rezhim_rezhim, rezhim_stavka)

            date_dogovor = data["Дата заключения договора"]
            date_nachalo = data["Дата начала работы"]
            date_konec = data["Дата окончания действия договора"] if (
                    "на определенный срок не менее одного года" in select_srok or "на время выполнения сезонной работы" in select_srok) else ""
            self._fill_date(root, date_dogovor, date_nachalo, date_konec)

            dol_dol = data["Должность"]
            self._fill_dol(root, dol_dol)

            select_country = Select(root.find_element(By.XPATH, ".//select[@name='workPlaceCountry']"))
            select_country.select_by_visible_text('Казахстан')

            adres_obl = data["Место выполнения работы"]["obl"]
            adres_center = data["Место выполнения работы"]["city"]
            nas_punkt = data["Место выполнения работы"]["nas_punkt"]
            try:
                adres_adres = data["Место выполнения работы"]["street"]
            except:
                adres_adres = data["Место выполнения работы"]["nas_punkt"]

            self._fill_adres(root, adres_obl, adres_center, adres_adres, nas_punkt)

            # try:
            #     select_military = Select(root.find_element(By.XPATH, ".//select[@name='army']"))
            #     select_military.select_by_visible_text(data["Военная обязанность"])
            # except:
            #     print("Нет Военной обязанности")

            if add_ecp:
                submit = root.find_element(By.XPATH, ".//input[@value='Сохранить']")
                submit.click()
        else:
            raise ValueError("Время ожидания истекло: https://www.enbek.kz/ru/cabinet/dogovor/add")
        print("self._fill_dog > done")

    def check_in_all_section(self, root, iin):
        self.anchor.click()

        input_iin = root.find_element(By.XPATH, ".//input[@name='iin']")
        input_iin.send_keys(iin)
        button_iin = root.find_element(By.XPATH,
                                       './/button[@type="submit" and text()="Найти"]')
        button_iin.click()

        if self._sel_wait_el(By.XPATH, '//strong[text()="Пусто..."]', 5):
            return True
        else:
            return False

    def _search_iin_create(self, root, iin):
        self.anchor.click()
        input_iin = root.find_element(By.XPATH, ".//input[@name='iin']")
        input_iin.send_keys(iin)
        button_iin = root.find_element(By.XPATH,
                                       './/button[@type="submit" and text()="Найти"]')
        button_iin.click()

        if self._sel_wait_el(By.XPATH, '//strong[text()="Пусто..."]', 5):
            return False
        else:
            return True

    def _check_dogovor_create(self, data):
        driver = self.driver
        driver.get(self.url["all"])
        if self._sel_wait_el(By.XPATH, "//a[text()[contains(., 'Добавить')]]"):
            self.anchor = driver.find_element(By.XPATH, '//strong[text()="Договоры"]')
            root = driver.find_element(By.CSS_SELECTOR, ".content")
            iin_iin = data["ИИН"]
            iin_valid = self.check_in_all_section(root, iin_iin)
            if not iin_valid:
                raise ValueError("Tricky dogovor!")

            # ---------- End of checking in all section --------------#
            driver.get(self.url["list"])
            self.anchor = driver.find_element(By.XPATH, '//strong[text()="Договоры"]')
            root = driver.find_element(By.CSS_SELECTOR, ".content")
            iin_exist = self._search_iin_create(root, iin_iin)
            print(iin_exist)
        else:
            raise ValueError("Время ожидания истекло: https://www.enbek.kz/ru/cabinet/dogovor/list")
        print("self._check_dogovor > done")
        return iin_exist

    def _to_ecp_priem(self):
        driver = self.driver
        if self._sel_wait_el(By.XPATH, '//*[text()[contains(., "Ошибка добавления! У вас уже есть действующий договор по данному ИИН")]]', sec=5):
            raise ValueError("Clash in dogovor")
        WebDriverWait(driver, 10).until(
            ec.presence_of_element_located((By.XPATH, '//button[text()="Подписать договор и отправить"]')))

        driver.find_element(By.XPATH, '//button[text()="Подписать договор и отправить"]').click()

    # Public
    def create_dogovor(self, data):
        if not self.driver:
            self._sel_init()

        iin_exist = self._check_dogovor_create(data)  # Проверка есть ли данный ИИН
        if not iin_exist:
            self._fill_dog(data)
            # time.sleep(300)

            if add_ecp:
                self._to_ecp_priem()

                self._apply_dogovor()
        else:
            raise ValueError("No dogovor")
        time.sleep(3)

    def _search_iin_create1(self, root, iin, driver, data):
        self.anchor.click()

        input_iin = root.find_element(By.XPATH, ".//input[@name='iin']")
        input_iin.send_keys(iin)
        button_iin = root.find_element(By.XPATH,
                                       './/button[@type="submit" and text()="Найти"]')
        button_iin.click()

        if self._sel_wait_el(By.XPATH, '//strong[text()="Пусто..." ]', 3):
            raise ValueError("No dogovor")

        dog = driver.find_element_by_class_name("item-list")

        dog_i = dog.find_element_by_xpath('div/div/a')
        link = dog_i.get_attribute('href')
        print(link)
        dog_i.click()

        dop_dr = self.driver
        dop_dr.get(link)
        elements = dop_dr.find_elements(By.XPATH, "//div[@class = 'bordered-box item']")
        if len(elements) > 1:
            copy_element_check = []
            if self._sel_wait_el(By.XPATH, "//*[contains(text(), '№ Дополнительного соглашения')]", sec=5):
                temp_list = dop_dr.find_element(By.XPATH,
                                                "//*[contains(text(), '№ Дополнительного соглашения')]").find_element(
                    By.XPATH, "..").find_element(By.XPATH, "..").find_element(By.XPATH, "..").find_element(By.TAG_NAME,
                                                                                                           "tbody").find_elements(
                    By.TAG_NAME, "tr")
                for i in range(len(temp_list)):
                    copy_element_check.append(temp_list[i].find_elements(By.XPATH, "td")[
                                                  1].text)
                    copy_element_check.append(temp_list[i].find_elements(By.XPATH, "td")[
                                                  2].text)
                    copy_element_check.append(temp_list[i].find_elements(By.XPATH, "td")[
                                                  4].text)
                    copy_element_check.append(temp_list[i].find_elements(By.XPATH, "td")[
                                                  5].text)

                    if str(copy_element_check[0]).lower() == str(data["Номер доп соглашения"]).lower() and \
                            copy_element_check[1] == data[
                        "Дата начала действия доп соглашения"] and copy_element_check[2] == data["Должность"] \
                            and str(copy_element_check[3]).lower() == str(data["Штатная должность"]).lower():
                        raise ValueError("Доп. соглашение уже создано")
                    copy_element_check.clear()

        dop = WebDriverWait(dop_dr, 3).until(ec.presence_of_element_located((By.LINK_TEXT, 'Добавить доп. соглашение')))
        link1 = dop.get_attribute('href')
        print(link1)
        dop.click()
        return link1

    def _check_dop_sogl(self, data):
        driver = self.driver
        driver.get(self.url["list"])
        print(self.url["list"])

        if self._sel_wait_el(By.XPATH, "//a[text()[contains(., 'Добавить')]]"):
            self.anchor = driver.find_element(By.XPATH, '//strong[text()="Договоры"]')
            root = driver.find_element(By.CSS_SELECTOR, ".content")

            iin_iin = data["ИИН"]
            link = self._search_iin_create1(root, iin_iin, driver, data)

            print("self._check_dogovor > done")
            return link
        else:
            print("self._check_dogovor > lost")
            raise ValueError("Время ожидания истекло: https://www.enbek.kz/ru/cabinet/dogovor/list")

    def _data_append(self, data, link):
        root = self.driver
        root.get(link)
        no_dop_sogl = root.find_element_by_xpath('//input[@name="numDogovor"]')
        no_dop_sogl.send_keys(data["Номер доп соглашения"])

        select_rezhim = Select(root.find_element(By.XPATH, "//select[@name='workingHours']"))
        select_rezhim.select_by_visible_text(data["Режим рабочего времени"])

        date_nachalo_dop_sogl = root.find_element_by_xpath("//input[@name='dateBegDogovor']")
        date_nachalo_dop_sogl.click()
        date_nachalo_dop_sogl.send_keys(data["Дата начала действия доп соглашения"])

        date_zakl_dop_sogl = root.find_element_by_xpath("//input[@name='dateZakDogovor']")
        date_zakl_dop_sogl.click()
        date_zakl_dop_sogl.send_keys(data["Дата заключения дополнительного соглашения"])

        select_srok = Select(root.find_element(By.XPATH, "//select[@name='srokdop']"))
        select_srok.select_by_visible_text(data['Срок договора'])
        select_srok = data['Срок договора']

        select_vid = Select(root.find_element(By.XPATH, ".//select[@name='partTime']"))
        select_vid.select_by_visible_text(data["Вид работы"])

        date_konec = data["Дата окончания действия трудового договора"] if (
                "на определенный срок не менее одного года" in select_srok or "на время выполнения сезонной работы" in select_srok) else ""

        if len(date_konec):
            input_konec = root.find_element(By.XPATH, ".//input[@name='endDateTD']")
            input_konec.click()
            input_konec.send_keys(date_konec)

        dol = data['Должность']
        span_dogovor = root.find_element(By.XPATH, "//span[@class='selection']")
        span_dogovor = span_dogovor.find_element(By.XPATH, 'span')
        span_dogovor.click()
        # Error between 697-699
        root_dogovor = root.find_element(By.XPATH,
                                         "//span[@class='select2-container select2-container--default select2-container--open']")
        input_dogovor = root_dogovor.find_element(By.XPATH, ".//input[@class='select2-search__field']")
        input_dogovor.send_keys(dol)

        if self._sel_wait_el(By.XPATH,
                             "//li[@class='select2-results__option select2-results__option--highlighted' and text()='" + dol + "']",
                             sec=5):
            li_dogovor = root.find_element(By.XPATH,
                                           "//li[@class='select2-results__option select2-results__option--highlighted' and text()='" + dol + "']")
        elif self._sel_wait_el(By.XPATH, "//li[@class='select2-results__option' and text()='" + dol + "']", sec=5):
            li_dogovor = root.find_element(By.XPATH,
                                           "//li[@class='select2-results__option' and text()='" + dol + "']")
        else:
            raise ValueError("Время ожидания истекло: Должность не найдена")

        li_dogovor.click()
        if not self._sel_wait_el(By.CSS_SELECTOR, "span.select2-container--open input.select2-search__field",
                                 appear=False):
            raise ValueError("Время ожидания истекло: Должность не выбрана")

        input_dol = root.find_element(By.XPATH, "//input[@name='shtatDolj']")
        input_dol.send_keys(data['Штатная должность'])

        obl = data["Место выполнения работы"]["obl"]
        center = data["Место выполнения работы"]["city"]
        nas_punkt = data["Место выполнения работы"]["nas_punkt"]
        try:
            adres = data["Место выполнения работы"]["street"]
        except:
            adres = data["Место выполнения работы"]["nas_punkt"]
        select_country = Select(root.find_element(By.XPATH, "//select[@name='workPlaceCountry']"))
        select_country.select_by_visible_text('Казахстан')
        button_obl = root.find_element(By.XPATH, "//Button[text()='Выбрать']")
        button_obl.click()

        self._sel_wait_el(By.XPATH, "//div[@class='modal-content' and //h4[text()='Справочник регионов']]")
        root_adres = root.find_element(By.XPATH,
                                       ".//div[@class='modal-content' and //h4[text()='Справочник регионов']]")
        li_obl = root_adres.find_element(By.XPATH, ".//li[span[text()='" + obl + "']]")
        li_obl.click()
        self._sel_wait_el(By.XPATH, "//li[span[text()='" + center + "']]")
        time.sleep(1)
        li_center = root_adres.find_element(By.XPATH, "//li[span[text()='" + center + "']]")
        time.sleep(1)
        li_center.click()

        button_adres = root_adres.find_element(By.XPATH, ".//button[text()='Выбор']")
        button_adres.click()
        if not self._sel_wait_el(By.XPATH, "//div[@class='modal-content' and //h4[text()='Справочник регионов']]",
                                 appear=False):
            raise ValueError("Время ожидания истекло: Адрес не выбран")

        input_adres = root.find_element(By.XPATH, "//input[@name='workPlace']")
        input_adres.send_keys(adres)

        span_nas = root.find_element(By.XPATH,
                                     ".//label[text()='Населённый пункт ']/parent::div//span[@class='selection']")
        span_nas = span_nas.find_element(By.XPATH, 'span')
        span_nas.click()
        root_dogovor = root.find_element(By.XPATH,
                                         "//span[@class='select2-container select2-container--default select2-container--open']")
        input_dogovor = root_dogovor.find_element(By.XPATH, ".//input[@class='select2-search__field']")
        center = nas_punkt
        input_dogovor.send_keys(center)

        if self._sel_wait_el(By.XPATH,
                             "//li[@class='select2-results__option select2-results__option--highlighted' and text()='" + center + "']",
                             sec=5):
            li_center = root.find_element(By.XPATH,
                                          "//li[@class='select2-results__option select2-results__option--highlighted' and text()='" + center + "']")
        elif self._sel_wait_el(By.XPATH, "//li[@class='select2-results__option' and text()='" + center + "']", sec=5):
            li_center = root.find_element(By.XPATH,
                                          "//li[@class='select2-results__option' and text()='" + center + "']")
        else:
            raise ValueError("Время ожидания истекло: Должность не найдена")
        li_center.click()

        if add_ecp:
            apply = root.find_element(By.XPATH, "//input[@value='Сохранить']")
            apply.click()

    def _to_ecp_peremew(self):
        driver = self.driver

        if self._sel_wait_el(By.XPATH, '//*[text()[contains(., "Запись с такими параметрами уже именеентся")]]', sec=5):
            raise ValueError("Доп. соглашение уже создано")

        podpis = driver.find_element(By.XPATH,
                                     "//*[contains(text(), '№ Дополнительного соглашения')]").find_element(
            By.XPATH, "..").find_element(By.XPATH, "..").find_element(By.XPATH, "..").find_element(By.TAG_NAME,
                                                                                                   "tbody").find_elements(
            By.TAG_NAME, "tr")[-1].find_element(By.ID, "dropdownMenuButton")
        podpis.click()
        # time.sleep(2)
        podpis = driver.find_element(By.XPATH,
                                     "//*[contains(text(), '№ Дополнительного соглашения')]").find_element(
            By.XPATH, "..").find_element(By.XPATH, "..").find_element(By.XPATH, "..").find_element(By.TAG_NAME,
                                                                                                   "tbody").find_elements(
            By.TAG_NAME, "tr")[-1].find_element(By.ID, "dropdownMenuButton").find_element(By.XPATH, "..").find_elements(
            By.TAG_NAME, "ul")[0].find_elements(By.TAG_NAME, "li")[2].find_elements(By.TAG_NAME, "a")[0]

        podpis.click()

    def create_dop_sogl(self, data):
        if not self.driver:
            self._sel_init()

        link = self._check_dop_sogl(data)

        self._data_append(data, link)

        if add_ecp:
            self._to_ecp_peremew()

            self._apply_dogovor()

    def _search_iin_create2(self, root, iin, driver):
        self.anchor.click()
        input_iin = root.find_element(By.XPATH, ".//input[@name='iin']")
        input_iin.send_keys(iin)
        button_iin = root.find_element(By.XPATH,
                                       './/button[@type="submit" and text()="Найти"]')

        button_iin.click()
        if self._sel_wait_el(By.XPATH, '//strong[text()="Пусто..." ]', 3):
            raise ValueError("No dogovor")
        dog = driver.find_element_by_class_name("item-list")
        dog_i = dog.find_element_by_xpath('div/div/a')
        link = dog_i.get_attribute('href')
        dog_i.click()

        dop_dr = self.driver
        dop_dr.get(link)
        time.sleep(2)
        rast = WebDriverWait(dop_dr, 3).until(
            ec.presence_of_element_located((By.XPATH, "//button[text()='Расторгнуть']")))
        if add_ecp:
            rast = dop_dr.find_element(By.XPATH, "//button[text()='Расторгнуть']")
            rast.click()
        return dop_dr

    def _check_del_dog(self, data):
        driver = self.driver
        driver.get(self.url["list"])
        if self._sel_wait_el(By.XPATH, "//a[text()[contains(., 'Добавить')]]"):
            self.anchor = driver.find_element(By.XPATH, '//strong[text()="Договоры"]')
            root = driver.find_element(By.CSS_SELECTOR, ".content")

            iin_iin = data["ИИН"]
            driver_common = self._search_iin_create2(root, iin_iin, driver)
            print("self._check_dogovor > done")
            return driver_common
        else:
            raise ValueError("Время ожидания истекло: https://www.enbek.kz/ru/cabinet/dogovor/list")

    def _data_append_del(self, data, root):
        date_rast_dog = WebDriverWait(root, 3).until(
            ec.presence_of_element_located((By.XPATH, "//input[@name='dateCutDogovor']")))
        date_rast_dog.click()
        root.execute_script('document.getElementsByName("dateCutDogovor")[0].removeAttribute("readonly")')
        date_rast_dog.send_keys(data["Дата расторжение"])
        prichina = root.find_element(By.XPATH, "//div[@class='prich']/div/span/span/span")
        prichina.click()
        prich = data["Причина"]
        if prich == "Достижение работником пенсионного возраста":
            prich = "Достижение работником пенсионного возраста, установленного" + "\u00A0" + "пунктом 1 статьи 11" + "\u00A0" + "Закона Республики Казахстан «О пенсионном обеспечении в Республике Казахстан», с правом ежегодного продления срока трудового договора по взаимному согласию сторон"
        input_prichina = root.find_element(By.XPATH, ".//input[@class='select2-search__field']")
        input_prichina.send_keys(prich)
        print(3)
        if self._sel_wait_el(By.XPATH,
                             "//li[@class='select2-results__option select2-results__option--highlighted' and text()='" + prich + "']",
                             sec=5):
            li_prichina = root.find_element(By.XPATH,
                                            "//li[@class='select2-results__option select2-results__option--highlighted' and text()='" + prich + "']")
        elif self._sel_wait_el(By.XPATH, "//li[@class='select2-results__option' and text()='" + prich + "']", sec=5):
            li_prichina = root.find_element(By.XPATH,
                                            "//li[@class='select2-results__option' and text()='" + prich + "']")
        else:
            raise ValueError("Время ожидания истекло: Должность не найдена")
        li_prichina.click()
        if not self._sel_wait_el(By.CSS_SELECTOR, "span.select2-container--open input.select2-search__field",
                                 appear=False):
            raise ValueError("Время ожидания истекло: Должность не выбрана")

        submit = root.find_element(By.XPATH, "//button[text()='Расторгнуть']")
        submit.click()

    def terminate_dog(self, data):
        if not self.driver:
            self._sel_init()

        driver = self._check_del_dog(data)

        if add_ecp:
            self._data_append_del(data, driver)

            self._apply_dogovor()

    def _apply_dogovor(self):
        driver = self.driver
        passw = self.passw
        # WARNING: under 3 line of code are useless
        # podpisat_button = driver.find_element(By.XPATH, "//button[text()='Подписать договор и отправить']")
        # podpisat_button.click()
        # time.sleep(2)
        apply_button = driver.find_element(By.XPATH, "//button[text()='OK']")
        apply_button.click()
        time.sleep(2)
        path_to_ecp = self.path_to_ecp
        time.sleep(1)
        keyboard.write(path_to_ecp)
        keyboard.press('enter')
        time.sleep(2)
        keyboard.write(passw)
        ok_buttons = driver.find_elements(By.XPATH, '//div[@class="modal-footer"]/button[contains(text(), "OK")]')
        for ok_button in ok_buttons:
            if ok_button.is_displayed():
                ok_button.click()
        time.sleep(2)
        if self._sel_wait_el(By.XPATH, '//*[text()[contains(., "Социальный отпуск с указаной датой")]]', sec=5):
            raise ValueError("Soc s datoi uje est and delete sohranenie")
        elif self._sel_wait_el(By.XPATH, '//*[text()[contains(., "не должна быть раньше даты начала договора")]]', sec=5):
            raise ValueError("Soc s datoi posle sohraneniya")
        # time.sleep(2)
        # driver.find_element(By.XPATH, "/html/body/div[3]/div/div[3]/button[3]").click()
        # time.sleep(1)

    #     passw = self.passw
    #     driver = self.driver
    #     time.sleep(2)
    #     apply_button = driver.find_element(By.XPATH, "//button[text()='OK']")
    #     apply_button.click()
    #     time.sleep(2)
    #     print(self.path_to_ecp)
    #     keyboard.write(self.path_to_ecp)
    #     keyboard.press('enter')
    #     time.sleep(2)
    #     print(passw)
    #     keyboard.write(passw)
    #     time.sleep(2)
    #     # driver.find_element(By.XPATH, '//div[@class="modal-footer"]/button').click()
    #     # driver.find_element(By.XPATH, '//div[@class="modal-footer"]/button').click()
    #     # keyboard.press('enter')
    #     ok_buttons = driver.find_elements(By.XPATH, '//div[@class="modal-footer"]/button[contains(text(), "OK")]')
    #     for ok_button in ok_buttons:
    #         ok_button.click()
    #         driver.execute_script("arguments[0].click();", ok_button)
    #     time.sleep(2)

    def quit(self):
        if self.driver:
            self.driver.quit()

    def _search_iin_create_otpusk(self, root, iin, driver, data):
        self.anchor.click()

        input_iin = root.find_element(By.XPATH, ".//input[@name='iin']")
        input_iin.send_keys(iin)
        button_iin = root.find_element(By.XPATH,
                                       './/button[@type="submit" and text()="Найти"]')

        button_iin.click()

        if self._sel_wait_el(By.XPATH, '//strong[text()="Пусто..." ]', 3):
            raise ValueError("No dogovor")

        dog = driver.find_element_by_class_name("item-list")
        dog_i = dog.find_element_by_xpath('div/div/a')
        link = dog_i.get_attribute('href')
        dog_i.click()

        dop_dr = self.driver
        dop_dr.get(link)

        elements = dop_dr.find_elements(By.XPATH, "//div[@class = 'bordered-box item']")
        if len(elements) > 1:
            copy_element_check = []
            if self._sel_wait_el(By.XPATH, "//*[contains(text(), 'Тип отпуска')]", 3):
                temp_list = dop_dr.find_element(By.XPATH,
                                                "//*[contains(text(), 'Тип отпуска')]").find_element(
                    By.XPATH, "..").find_element(By.XPATH, "..").find_element(By.XPATH, "..").find_element(By.TAG_NAME,
                                                                                                           "tbody").find_elements(
                    By.TAG_NAME, "tr")
                for i in range(len(temp_list)):
                    copy_element_check.append(temp_list[i].find_elements(By.XPATH, "td")[
                                                  1].text)
                    copy_element_check.append(temp_list[i].find_elements(By.XPATH, "td")[
                                                  2].text)
                    copy_element_check.append(temp_list[i].find_elements(By.XPATH, "td")[
                                                  3].text)
                    copy_element_check.append(temp_list[i].find_elements(By.XPATH, "td")[
                                                  4].text)

                    if copy_element_check[1] == data[
                        "Дата с"] or copy_element_check[2] == data["Дата по"]:
                        raise ValueError("Soc s datoi uje est")

                    if str(copy_element_check[0]).lower().lstrip() == str(data["Тип отпуска"]).lower().lstrip() and \
                            copy_element_check[1] == data[
                        "Дата с"] and copy_element_check[2] == data["Дата по"] \
                            and copy_element_check[3] == str(data["Номер табеля"]):
                        raise ValueError("Соц. отпуск уже создано")
                    copy_element_check.clear()

        otpusk = WebDriverWait(dop_dr, 3).until(
            ec.presence_of_element_located((By.XPATH, '//a[text()="Добавить соц. отпуск"]')))
        link1 = otpusk.get_attribute('href')
        otpusk.click()
        return link1

    def _check_otpusk(self, data):
        driver = self.driver
        driver.get(self.url["list"])
        if self._sel_wait_el(By.XPATH, "//a[text()[contains(., 'Добавить')]]"):
            self.anchor = driver.find_element(By.XPATH, '//strong[text()="Договоры"]')
            root = driver.find_element(By.CSS_SELECTOR, ".content")

            iin_iin = data["ИИН"]
            link = self._search_iin_create_otpusk(root, iin_iin, driver, data)
            print("self._check_dogovor > done")
            return link
        else:
            print("self._check_dogovor > done")
            raise ValueError("Время ожидания истекло: https://www.enbek.kz/ru/cabinet/dogovor/list")

    def _data_append_otpusk(self, data, link):
        root = self.driver
        root.get(link)

        WebDriverWait(root, 2).until(
            ec.presence_of_element_located((By.XPATH, '//h3[text()="Добавление социального отпуска"]')))

        tip_otpuska = root.find_element(By.XPATH, '//span[@title="Выберите тип отпуска"]')
        tip_otpuska.click()
        print(1)

        tip = data["Тип отпуска"]
        'Без сохранения заработной платы по уходу за ребенком до достижения им возраста 3 лет'
        'Без сохранения заработной платы по уходу за ребенком до достижения им возраста 3 лет'
        if self._sel_wait_el(By.XPATH,
                             "//li[@class='select2-results__option select2-results__option--highlighted' and text()='" + tip + "']",
                             sec=5):
            li_otpusk = root.find_element(By.XPATH,
                                          "//li[@class='select2-results__option select2-results__option--highlighted' and text()='" + tip + "']")
        elif self._sel_wait_el(By.XPATH, "//li[@class='select2-results__option' and text()='" + tip + "']", sec=5):
            li_otpusk = root.find_element(By.XPATH,
                                          "//li[@class='select2-results__option' and text()='" + tip + "']")
        else:
            raise ValueError("Время ожидания истекло: Тип отпуска не найден")
        li_otpusk.click()

        ne_rabotal_s = root.find_element(By.XPATH, '//input[@name="beginDate"]')
        data["Дата с"] = data["Дата с"][3:5] + data["Дата с"][0:2] + data["Дата с"][6:]
        ne_rabotal_s.send_keys(data["Дата с"])

        ne_rabotal_po = root.find_element(By.XPATH, '//input[@name="endDate"]')
        data["Дата по"] = data["Дата по"][3:5] + data["Дата по"][0:2] + data["Дата по"][6:]
        ne_rabotal_po.send_keys(data["Дата по"])

        try:
            no_tabelya = root.find_element(By.XPATH, '//input[@name="sheetNumber"]')
            no_tabelya.send_keys(data["Номер табеля"])
        except:
            print("Другой тип отпуска")

        try:
            first_date = root.find_element(By.XPATH, '//input[@name="workDate"]')
            first_date.send_keys(data["Дата по"])
        except:
            print("Другой тип отпуска")

        if add_ecp:
            apply = root.find_element(By.XPATH, '//input[@value="Сохранить"]')

            apply.click()

    def _to_ecp_otpusk(self):
        driver = self.driver
        # WebDriverWait(driver, 3).until(
        #     ec.presence_of_element_located((By.XPATH, '//strong[text()="Социальные отпуска"]')))
        # temp_list = driver.find_element(By.XPATH,
        #                              "//*[contains(text(), 'Тип отпуска')]").find_element(
        #     By.XPATH, "..").find_element(By.XPATH, "..").find_element(By.XPATH, "..").find_element(By.TAG_NAME,
        #                                                                                            "tbody").find_elements(
        #     By.TAG_NAME, "tr")
        # for i in range(len(temp_list)):
        #     podpis = temp_list[i].find_element(By.ID, "dropdownMenuButton")
        #     podpis.click()
        #     try:
        #         podpis = driver.find_element(By.XPATH,
        #                                      "//*[contains(text(), 'Тип отпуска')]").find_element(
        #             By.XPATH, "..").find_element(By.XPATH, "..").find_element(By.XPATH, "..").find_element(By.TAG_NAME,
        #                                                                                                    "tbody").find_elements(
        #             By.TAG_NAME, "tr")[i].find_element(By.ID, "dropdownMenuButton").find_element(By.XPATH, "..").find_elements(
        #             By.TAG_NAME, "ul")[0].find_elements(By.TAG_NAME, "li")[2].find_elements(By.TAG_NAME, "a")[0]
        #
        #         podpis.click()
        #         passw = self.passw
        #         apply_button = driver.find_element(By.XPATH, "//button[text()='OK']")
        #         apply_button.click()
        #         time.sleep(2)
        #         path_to_ecp = self.path_to_ecp
        #         time.sleep(1)
        #         keyboard.write(path_to_ecp)
        #         keyboard.press('enter')
        #         time.sleep(2)
        #         keyboard.write(passw)
        #         ok_buttons = driver.find_elements(By.XPATH,
        #                                           '//div[@class="modal-footer"]/button[contains(text(), "OK")]')
        #         for ok_button in ok_buttons:
        #             if ok_button.is_displayed():
        #                 ok_button.click()
        #         time.sleep(2)
        #     except:
        #         print("Already podpisan!")
        podpis = driver.find_element(By.XPATH,
                                     "//*[contains(text(), 'Тип отпуска')]").find_element(
            By.XPATH, "..").find_element(By.XPATH, "..").find_element(By.XPATH, "..").find_element(By.TAG_NAME,
                                                                                                   "tbody").find_elements(
            By.TAG_NAME, "tr")[-1].find_element(By.ID, "dropdownMenuButton")
        podpis.click()
        # time.sleep(2)
        podpis = driver.find_element(By.XPATH,
                                     "//*[contains(text(), 'Тип отпуска')]").find_element(
            By.XPATH, "..").find_element(By.XPATH, "..").find_element(By.XPATH, "..").find_element(By.TAG_NAME,
                                                                                                   "tbody").find_elements(
            By.TAG_NAME, "tr")[-1].find_element(By.ID, "dropdownMenuButton").find_element(By.XPATH, "..").find_elements(
            By.TAG_NAME, "ul")[0].find_elements(By.TAG_NAME, "li")[2].find_elements(By.TAG_NAME, "a")[0]

        podpis.click()


        # WebDriverWait(driver, 3).until(
        #     ec.presence_of_element_located((By.XPATH, '//a[text()="Отправить"]')))
        # otpravit = driver.find_element(By.XPATH, '//a[text()="Отправить"]')
        # driver.execute_script("arguments[0].click();", otpravit)


    def to_otpusk(self, data):
        if not self.driver:
            self._sel_init()

        link = self._check_otpusk(data)

        self._data_append_otpusk(data, link)

        if add_ecp:

            self._to_ecp_otpusk()

            self._apply_dogovor()


def test_enbek():
    # VARS
    data_priem = {}

    root_path = str(__file__).split("Sources")[0].replace("/", "\\")

    # OBJECT
    enbek = Enbek(root_path, 'login', 'password', 'path_to_ecp', 'password_ecp')
    try:
        enbek.create_dogovor(data_priem)
    except Exception as e:
        print(e)
        print("FAIL---", data_priem["ИИН"])

    data_dop_sogl = {}
    try:
        enbek.create_dop_sogl(data_dop_sogl)
    except Exception as e:
        print(e)
        print("FAIL---", data_dop_sogl["ИИН"])

    data_del = {}
    try:
        enbek.terminate_dog(data_del)
    except Exception as e:
        print(e)
        print("FAIL---", data_del["ИИН"])
