from fill_up import Create_Data
from scrapping_data import start, end
import config
import agent_initializetion
from exchangelib import Configuration, Account, DELEGATE, Credentials
from result_sending import send_email
import psutil
import time

import json
debug = True

# params = agent_initializetion.get_params()
# agent_initializetion.print_params(params, r"C:\Users\Administrator\Desktop\RPA_Robot\Robot\EnbekRobot\Files\demofile2.txt")
# credentials = params['credentials']
# time.sleep(600)

for config_dict_per_filial in config.filials:
    if debug:
        enbek_account = ""
        enbek_password = ""
        ecp_account = r"C:\Users\Administrator\Desktop\RPA_Robot\RPA_Robot\Robot\EnbekRobot\Files\ЦЖС\ECP.p12"
        ecp_password = ""
        SAP_login = "HRROBOT"
        SAP_password = "SapKtzh*2021"
        SAP_post_code = "5000"
        filial_name = "ЦЖС"
        mail_user = "ktzh\\Ktzhrpa_hr"
        mail_pass = "zxcv159*"
    else:
        #print(config.filials[config_dict_per_filial])

        params = agent_initializetion.get_params()
        agent_initializetion.print_params(params, r"C:\Users\Administrator\Desktop\RPA_Robot\Robot\EnbekRobot\Files\demofile2.txt")
        credentials = params['credentials']
        #print(credentials["CFO_Enbek"]["login"])
        # f = open(r"C:\Users\Administrator\Desktop\RPA_Robot\Robot\EnbekRobot\Files\demofile2.txt", "a")
        # f.write(json.dumps(credentials["CFO_Enbek"]["login"]))
        # f.close()
        enbek_account = credentials[config.filials[config_dict_per_filial]["enbek_cred_name"]]["login"]
        enbek_password = credentials[config.filials[config_dict_per_filial]["enbek_cred_name"]]["password"]
        ecp_account = credentials[config.filials[config_dict_per_filial]["ecp_cred_name"]]["login"]
        ecp_password = credentials[config.filials[config_dict_per_filial]["ecp_cred_name"]]["password"]
        SAP_login = credentials[config.filials[config_dict_per_filial]["SAP_cred_name"]]["login"]
        SAP_password = credentials[config.filials[config_dict_per_filial]["SAP_cred_name"]]["password"]
        SAP_post_code = config.filials[config_dict_per_filial]["SAP_post_code"]
        filial_name = config.filials[config_dict_per_filial]["name"]
        mail_user = "ktzh\\Ktzhrpa_hr"
        mail_pass = "zxcv159*"

    status = start(filial_name, SAP_post_code, SAP_login, SAP_password, "01.01.2020", "none")

    do_enbek = Create_Data(filial_name, enbek_account, enbek_password, ecp_account, ecp_password)

    do_enbek.create_sogl()

    do_enbek.dop_sogl()
    do_enbek.delete_dog()
    do_enbek.otpusk()
    # do_enbek.vacancies()

    for proc in psutil.process_iter():
        if proc.name() == "chromedriver.exe" or proc.name() == "cl.exe" or proc.name() == "chrome.exe":
            proc.kill()

    out_data = r"C:\Users\Administrator\Desktop\RPA_Robot\RPA_Robot\Robot\EnbekRobot\Files\ХОЗУ\RPA_Enbek_03032022_За_янв-фев.xlsx"  #do_enbek.saved_file_data()

    credentials = Credentials(r"ktzh\Ktzhrpa_hr", "zxcv159*") # Do we need to save it in orc. configuration data

    config = Configuration(server="mail-ktzh.ktzh.railways.local", credentials=credentials)
    account = Account(primary_smtp_address="Ktzhrpa_hr@Railways.kz", config=config,
                      autodiscover=False, access_type=DELEGATE)

    attachments = []
    with open(out_data[0], 'rb') as f:
        content = f.read()
    attachments.append((out_data[1], content))

    send_email(account, 'Отчет выполнения работы ' + filial_name, '.', ['ceo@pythonrpa.org', 'almazordiomond2@gmail.com'],
               attachments=attachments)
    end()

    break
