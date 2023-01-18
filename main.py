import datetime
import sqlite3 as sl
import time
import pandas as pd
import pyautogui
import pywhatkit
from pandas import DataFrame
from pynput.keyboard import Key, Controller

keyboard = Controller()

#check phone numder in this def
def send_whatsapp_message(phone, msg):
    now = datetime.datetime.now()
    t = now.strftime("%d-%b %H:%M")
    try:
        pywhatkit.sendwhatmsg_instantly(
            phone_no='+79263165716', # change to phone
            message=msg,
            tab_close=True

        )
        time.sleep(10)
        pyautogui.click()
        time.sleep(2)
        keyboard.press(Key.enter)
        keyboard.release(Key.enter)
        return 'Отправлено ' + t
    except Exception as e:
        return str(e)


def ready_message(d: DataFrame, i):
    return f'Здравствуйте. Сообщение из сервисного центра по квитанции {d["Квитанция"].values[i]}.'"\n" \
           "\n" \
           f'Ваше устройство - {d["Аппарат"].values[i]},' "\n" \
           f'отремонтировано и готово к выдаче.' "\n" \
           "\n" \
           f'Стоимость ремонта (доплата) составляет - {d["Сумма доплаты"].values[i]} руб.'"\n" \
           f'Оплата наличными, не забудьте квитанцию.'"\n" \
           "\n" \
           f'Более детальную информацию Вы можете'"\n" \
           f'уточнить по тел: 8-495-926-72-26.'


def in_work_message(df: DataFrame, i):
    return f'Здравствуйте, сообщение из сервисного центра по квитанции {df["Номер"].values[i]}.'"\n" \
           "\n" \
           f'Информация по диагностике: {df["Аппарат"].values[i]} {df["Модель аппарата"].values[i]},'"\n" \
           f'{df["Выполняемые работы"].values[i].replace("?", "").rstrip()} руб.' "\n" \
           "\n" \
           f'Ремонтируем? Напишите ответ в чат или перезвоните'"\n" \
           f'по тел: 8-995-782-69-55.'"\n" \
           f'С уважением, Александр.'


def end_work_message(df: DataFrame, i):
    return f'Здравствуйте, сообщение из сервисного центра по квитанции {df["Номер"].values[i]}.'"\n" \
           "\n" \
           f'К сожалению Ваше устройство - {df["Аппарат"].values[i]} {df["Модель аппарата"].values[i]},'"\n" \
           f'без возможности ремонта, нет поставки запчастей.'"\n" \
           f'В течение нескольких дней Ваше устройство будет на выдаче.'"\n" \
           "\n" \
           f'Более детальную информацию можно'"\n" \
           f'уточнить по тел: 8-495-926-72-26.'

def phone(d: DataFrame, i):
    num = d["Телефон"].values[i].replace(',', '').replace('.', '').replace('-', '')
    phone = '+7{0}{1}{2}{3}{4}{5}{6}{7}{8}{9}'.format(*[i for i in num if i.isdigit()][1:])

    if phone[2] != '9':
        return phone[1:12]
    else:
        return phone[:12]


def check_create_DataFrame(file_old, file_new, sub):
    con = sl.connect(file_old)
    df_old = pd.read_sql('''SELECT * FROM user''', con)
    df_new: DataFrame = pd.read_excel(file_new, header=0)
    df_new = df_new.dropna().reset_index(drop=True)
    df_new["Статус"] = 'Not sent'
    df = pd.merge(df_old, df_new, how="outer", sort=True)
    df.drop_duplicates(subset=sub, keep=False, inplace=True)
    return df
    # return df[df["Статус"].values == "Not sent"]


def script_run():
    mode = input("Выбери скрипт отправки: готовые(1), согласование(2), без ремонта(3), выход(): ")

    if mode == '1':

        d = check_create_DataFrame('./database/ready_base.db', 'готовые.xlsx', "Квитанция")
        for i in range(len(d.index)):
            if not d["Статус"].values[i].startswith("Отправлено"):
                msg = ready_message(d, i)
                num = phone(d, i)
                print(num, msg)
                res = send_whatsapp_message(num, msg)
                d["Статус"].values[i] = res
                print("Статус =", res)
            else:
                continue

        d.to_html('./report/отправлено по готовым.html')
        df: DataFrame = d[d["Статус"].str.contains('Отправлено') == True]
        con = sl.connect('./database/ready_base.db')
        df.to_sql("user", con=con, if_exists="append", index=False)
        script_run()

    if mode == '2':

        w = check_create_DataFrame('./database/in_work_base.db', 'в работе.xlsx', "Номер")
        sym = ('?')
        for i in range(len(w.index)):
            if w["Статус"].values[i] == "Not sent" and w["Выполняемые работы"].values[i].strip().endswith(sym):
                msg = in_work_message(w, i)
                num = phone(w, i)
                print(num, msg)
                res = send_whatsapp_message(num, msg)
                w["Статус"].loc[w.index[i]] = res
                print("Статус =", res)
            else:
                continue

        df = w[w["Статус"].str.contains('Отправлено') == True]
        df.to_html('./report/отправлено на согласование.html')
        con = sl.connect('./database/in_work_base.db')
        df.to_sql("user", con=con, if_exists="append", index=False)
        script_run()

    if mode == '3':

        g = check_create_DataFrame('./database/in_work_base.db', 'в работе.xlsx', "Номер")
        sym = ('БР', 'бр', 'Ъ')
        for i in range(len(g.index)):
            if g["Статус"].values[i] == "Not sent" and g["Выполняемые работы"].values[i].strip().endswith(sym):
                msg = end_work_message(g, i)
                num = phone(g, i)
                print(num, msg)
                res = send_whatsapp_message(num, msg)
                g["Статус"].loc[g.index[i]] = res
                print("Статус =", res)

        df = g[g["Статус"].str.contains('Отправлено') == True]
        df.to_html('./report/отправлено по БР.html')
        con = sl.connect('./database/in_work_base.db')
        df.to_sql("user", con=con, if_exists="append", index=False)
        script_run()

    else:
        exit()


def initDB():
    df_r: DataFrame = pd.read_excel('./new/df_r.xlsx', header=0)
    df_w: DataFrame = pd.read_excel('./new/df_w.xlsx', header=0)
    df_r = df_r.dropna()
    df_w = df_w.dropna()
    df_r["Статус"] = 'Not sent'
    df_w["Статус"] = 'Not sent'
    con_r = sl.connect('./database/ready_base.db')
    con_w = sl.connect('./database/in_work_base.db')
    df_r.to_sql("user", con=con_r, if_exists="replace", index=False)
    df_w.to_sql("user", con=con_w, if_exists="replace", index=False)
    print("База данных успешно создана")


if __name__ == '__main__':

    print("Добро пожаловать в скрипт для отправки сообщений чере WhatsApp.")
    choice = input("Создать базу данных(1), продолжить(2), выход()?: ")

    if choice == "1":
        initDB()
        script_run()

    elif choice == "2":
        script_run()

    else:
        exit()
