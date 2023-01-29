import pyttsx3
import speech_recognition as sr
import os
import sys
import random
from fuzzywuzzy import fuzz
import datetime
import win32com.client as wincl
import time
from currency_converter import CurrencyConverter
import requests
import webbrowser
from tkinter import *

opts = {"alias": ('jarvis', 'джарвис'),
        "tbr": ('скажи', 'расскажи', 'покажи', 'сколько', 'произнеси', 'как','сколько','поставь','переведи', "засеки",'запусти','сколько будет'),
        "cmds":
            {"ctime": ('текущее время', 'сейчас времени', 'который час', 'время', 'какое сейчас время'),
             'startStopwatch': ('запусти секундомер', "включи секундомер", "засеки время"),
             'stopStopwatch': ('останови секундомер', "выключи секундомер", "останови"),
             "calc": ('прибавить','умножить','разделить','степень','вычесть','поделить','х','+','-','/'),
             "shutdown": ('выключи', 'выключить', 'отключение', 'отключи', 'выключи компьютер'),
             "conv": ("валюта", "конвертер","доллар",'руб','евро'),
             "internet": ("открой", "вк", "гугл", "сайт", 'вконтакте', "ютуб"),
             "translator": ("переводчик","translate"),
             "game": ("игра"),
             "nastroenie":("классно","отлично","хорошо"),
             "andrey111":("пришёл андрей", "андрей"),
             "mama":("пришла мама","мама зашла в комнату"),
             "pogodka":("погода","какая сегодня погода","скажи погоду"),
             "deals": ("дела","делишки", 'как сам', 'как дела')}}
startTime = 0
speak_engine = pyttsx3.init()
voices = speak_engine.getProperty('voices')
speak_engine.setProperty('voice', voices[0].id)
r = sr.Recognizer()
m = sr.Microphone(device_index=1)
voice = "str"


def speak(what):
    print(what)
    speak = wincl.Dispatch("SAPI.SpVoice")
    speak.Speak(what)




def callback(recognizer, audio):
    try:
        global voice
        voice = recognizer.recognize_google(audio, language="ru-RU").lower()

        print("[log] Распознано: " + voice)

        if voice.startswith(opts["alias"]):
            cmd = voice

            for x in opts['alias']:
                cmd = cmd.replace(x, "").strip()

            for x in opts['tbr']:
                cmd = cmd.replace(x, "").strip()
            voice = cmd
            # распознаем и выполняем команду
            cmd = recognize_cmd(cmd)
            execute_cmd(cmd['cmd'])


    except sr.UnknownValueError:
        print("[log] Голос не распознан!")
    except sr.RequestError as e:
        print("[log] Неизвестная ошибка, проверьте интернет!")
def listen():
    with m as source:
        r .adjust_for_ambient_noise(source)
    stop_listening = r.listen_in_background(m, callback)
    while True: time.sleep(0.1)

















def browser():
    sites = {"https://vk.com":["vk","вк"], 'https://www.youtube.com/':['youtube', 'ютуб'], 'https://ru.wikipedia.org': ["вики", "wiki"], 'https://ru.aliexpress.com':['али', 'ali', 'aliexpress', 'алиэспресс'], 'http://google.com':['гугл','google'], 'https://www.amazon.com':['амазон', 'amazon'], 'https://www.apple.com/ru':['apple','эпл'] }
    site = voice.split()[-1]
    for k, v in sites.items():
        for i in v:
            if i not in site.lower():
                open_tab = None
            else:
                open_tab = webbrowser.open_new_tab(k)
                break

        if open_tab is not None:
            break

def calculator():
    try:
        list_of_nums = voice.split()
        num_1,num_2 = int((list_of_nums[-3]).strip()), int((list_of_nums[-1]).strip())
        opers = [list_of_nums[0].strip(),list_of_nums[-2].strip()]
        for i in opers:
            if 'дел' in i or 'множ' in i or 'лож' in i or 'приба' in i or 'выч' in i or i == 'x' or i == '/' or i =='+' or i == '-' or i == '*':
                oper = i
                break
            else:
                oper = opers[1]
        if oper == "+" or 'слож' in oper:
            ans = num_1 + num_2
        elif oper == "-" or 'выче' in oper:
            ans = num_1 - num_2
        elif oper == "х" or 'множ' in oper:
            ans = num_1 * num_2
        elif oper == "/" or 'дел' in oper:
            if num_2 != 0:
                ans = num_1 / num_2
            else:
                speak("Делить на ноль невозможно")
        elif "степен" in oper:
            ans = num_1 ** num_2
        speak("{0} {1} {2} = {3}".format(list_of_nums[-3], list_of_nums[-2], list_of_nums[-1], ans))
    except:
        speak("Скажите, например: Сколько будет 5+5?")

def convertation():
    class CurrencyError(Exception):
        pass
    c = CurrencyConverter()
    money = None
    from_currency = None
    to_currency = None
    list_of_conv = voice.split()
    if len(list_of_conv) > 4:
        list_of_conv = list_of_conv[1:]
    else:
        print()
    while money is None:
        try:
            money = list_of_conv[0]
        except ValueError:
            speak("Скажите, к примеру: 50 долларов в рубли")
            break
    while from_currency is None:
        try:
            list_of_conv[0] = int(list_of_conv[0])
        except ValueError:
            speak("Скажите, к примеру: 50 долларов в рубли")
            break
        try:
            if "руб" in list_of_conv[1]:
                from_currency = "RUB"
            elif "дол" in list_of_conv[1]:
                from_currency = "USD"
            elif "евр" in list_of_conv[1]:
                from_currency = "EUR"
            if from_currency not in c.currencies:
                raise CurrencyError

        except (CurrencyError, IndexError):
            from_currency = None
            speak("Скажите, например: 50 долларов в рубли")
            break

    while to_currency is None:
        try:
            list_of_conv[0] = int(list_of_conv[0])
        except ValueError:
            return None
        try:
            if "руб" in list_of_conv[3]:
                to_currency = "RUB"
            elif "дол" in list_of_conv[3]:
                to_currency = "USD"
            elif "евр" in list_of_conv[3]:
                to_currency = "EUR"
            if to_currency not in c.currencies:
                raise CurrencyError

        except (CurrencyError, IndexError):
            to_currency = None
            speak("Скажите, например: 50 долларов в рубли")
            break
    while True:
        try:
            speak(f"{money} {from_currency} в {to_currency} - "
                f"{round(c.convert(money, from_currency, to_currency), 2)}")
            break
        except ValueError:
            speak("Скажите, например: 50 долларов в рубли")
            break

def translate():
    url = 'https://translate.yandex.net/api/v1.5/tr.json/translate?'
    key = 'trnsl.1.1.20190227T075339Z.1b02a9ab6d4a47cc.f37d50831b51374ee600fd6aa0259419fd7ecd97'
    text = voice.split()[1:]
    lang = 'en-ru'
    r = requests.post(url, data={'key': key, 'text': text, 'lang': lang}).json()
    try:
        speak(r["text"])
    except:
        speak("Обратитесь к переводчику, начиная со слова 'Переводчик'")

def gamer():
    os.system('D:\\FORZA\\ForzaHorizon4\\ForzaHorizon4.exe')

def pogoda():
    import requests
    s_city = "Ekaterinburg,RU"
    city_id = 1486209
    appid = "69c0941a02d656f5da1cc1bc9921e25d"
    try:
        res = requests.get("http://api.openweathermap.org/data/2.5/weather",
                 params={'id': city_id, 'units': 'metric', 'lang': 'ru', 'APPID': appid})
        data = res.json()
        s1="Состояние погоды:", data['weather'][0]['description']
        s2="Температура равна:", data['main']['temp'], "°Цельсия"
        speak(s1)
        speak(s2)
    except Exception as e:
        print("Exception (weather):", e)
        pass
    #url = 'http://wttr.in/?0T'
    #response = requests.get(url)  # выполните HTTP-запрос
    #print(response.text)  # напечатайте текст HTTP-ответа




def recognize_cmd(cmd):
    RC = {'cmd': '', 'percent': 0}
    for c, v in opts['cmds'].items():
        for x in v:
            vrt = fuzz.ratio(cmd, x)
            if vrt > RC['percent']:
                RC['cmd'] = c
                RC['percent'] = vrt
    return RC
def execute_cmd(cmd):
    global startTime
    if cmd == 'ctime':
        now = datetime.datetime.now()
        speak("Сейчас {0}:{1}".format(str(now.hour), str(now.minute)))
    elif cmd == 'shutdown':
        os.system('shutdown -s')
        speak("Выключаю...")
    elif cmd == 'calc':
        calculator()
    elif cmd == 'conv':
        convertation()
    elif cmd == 'translator':
        translate()
    elif cmd == 'game':
        gamer()
    elif cmd == 'mama':
        speak("Кирилл вас очень сильно любит.")
    elif cmd == 'andrey111':
        speak("Андрей не мешай Кириллу.")
    elif cmd == 'internet':
        browser()
    elif cmd == 'pogodka':
        pogoda()
    elif cmd == 'startStopwatch':
        speak("Секундомер запущен")
        startTime = time.time()
    elif cmd == "stopStopwatch":
        if startTime != 0:
            Time = time.time() - startTime
            speak(f"Прошло {round(Time // 3600)} часов {round(Time // 60)} минут {round(Time % 60, 2)} секунд")
            startTime = 0
        else:
            speak("Секундомер не включен")
    elif cmd == 'deals':
        deals_list=["Пока отлично.","Круто, как никогда.","Пока не родила.","Готов всегда к вашим услугам.","Отлично, только, что освободился.","Хорошо, а у вас как?","Готов к работе, хозяин.","Классно, жду ваших указаний."]
        vibor = random.choice(deals_list)
        speak(vibor)
    elif cmd=="nastroenie":
        speak("Ну вот и хорошо, что всё отлично!")
    else:
        print("Команда не распознана!")









now = datetime.datetime.now()

if now.hour >= 6 and now.hour < 12:
    speak("Доброе утро!")
elif now.hour >= 12 and now.hour < 18:
    speak("Добрый день!")
elif now.hour >= 18 and now.hour < 23:
    speak("Добрый вечер!")
else:
    speak("Доброй ночи!")

listen()

root = Tk()
root.geometry('250x350')
root.configure(bg='gray22')
root.title('Jarvis')
root.resizable(False, False)

lb = Label(root)
lb.configure(bg='gray')
lb.place(x=25, y=25, height=25, width=200)

but1 = Button(root, text='Слушать', command=listen)
but1.configure(bd=1, font=('Castellar', 25), bg='gray')
but1.place(x=50, y=160, height=50, width=150)

but2 = Button(root, text='Выход', command=quit)
but2.configure(bd=1, font=('Castellar', 25), bg='gray')
but2.place(x=50, y=220, height=50, width=150)

root.mainloop()