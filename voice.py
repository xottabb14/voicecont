import speech_recognition as sr #библиотека распознавания речи
import os #библиотека для работы с функциями ОС
import time #библиотека времени
import win32com.client as wincl #подключаем канал для перевода текста в речь средствами Windows
import subprocess #библеотека для запуска параллельных процессов
import random #библиотека рандома

import config #база основных команд
import sinonims #база синонимов к запускаемым программам
import pathscmd #база путей к программам и команды для cmd

import pyautogui #библиотека для горячих клавиш и нажатий
import pywinauto #библиотека для автоматического ввода текста
import webbrowser #библиотека для работы с веб

voicecontr = False #флажок вкл/выкл голосовго управления
writecontr = False #флажок вкл/выкл режима печати
tt = 0 #обнуленный таймер для очистки окна cmd
firstwrite = 1 #проверка первого запуска режима печати
ent = [" Enter"," энтер"," Интер"," интер", " Энтер","Enter","энтер","Интер","интер", "Энтер", "enter"]

#функция для запуска отдельным процессом программ и команд
def startcmd(cmd):
    PIPE = subprocess.PIPE
    p = subprocess.Popen(cmd, shell = True)
    p.poll();

#функция очистки окна cmd
def cleart(tt):
    ttt = time.time();
    ts = ttt-tt;
    if ts > 2000:
        os.system('cls');
        tt = time.time();
    return (tt);

#функция замены символов для голосовой печати
def clean_textn (text):
            if not isinstance(text, str):
                raise TypeError('Это не текст')
            for i in ['\n']:
                text = text.replace(i,'')
            for i in [' ']:
                text = text.replace(i,'{SPACE}')
            for n in ent:
                for i in [n]:
                    text = text.replace(i,'{ENTER}')
            for i in [' точка']:
                text = text.replace(i,'.')
            return text

#функция убирающая слово поиск
def clean_finder (text,w):
            if not isinstance(text, str):
                raise TypeError('Это не текст')
            for i in [w]:
                text = text.replace(i,'')
            for i in ['\n']:
                text = text.replace(i,'')
            return text


#функция озвучивания текста words
def talk(words):
    print(words)
    speak = wincl.Dispatch("SAPI.SpVoice")
    speak.Speak(words)
randhllo = ["Рада Вас видеть.","С возвращением.","Отлично. Вы вернулись.","Доброго дня."]
randnumh = random.randint(0,3)
talk (randhllo[randnumh])

#функция для запуска и закрытия программ onoff = 1 - открыть, 0 - закрыть
def progssin (recog,onoff):
    #Для открытия/закрытия aimp
    for w in sinonims.aimp:
        if w in recog:
            if onoff==1:
                startcmd(pathscmd.aimpon)
                break
            else:
                startcmd(pathscmd.aimpoff)
                break
    #для открытия/закрытия блокнота
    for w in sinonims.notepad:
        if w in recog:
            if onoff==1:
                startcmd(pathscmd.notepadon)
                break
            else:
                startcmd(pathscmd.notepadoff)
                break
    #для открытия/закрытия браузера
    for w in sinonims.browser:
        if w in recog:
            if onoff==1:
                startcmd(pathscmd.browseron)
                break
            else:
                startcmd(pathscmd.browseroff)
                break


#начало основного цикла, где скрипт "слушает"
while True:
    ttt = time.time();
    ts = ttt-tt
    if ts > 60:
        os.system('cls')
        tt = time.time()
    recog = "" #обнуляем переменную строки полученную голосом
    r = sr.Recognizer() #объект записи и распознавания голоса
    with sr.Microphone() as source:
        print("...")
        audio = r.listen(source, phrase_time_limit=3)#полученный аудиофайл голоса в течение 3 сек

    try: #пробуем распознать
        recog = r.recognize_google(audio, language="ru-RU").lower() #переменная строки из голосового файла audio на русском, где все символы строчные
        print("РАСПОЗНАНО: "+recog)

    except sr.UnknownValueError: #если в записи билеберда - говорим что не смогли распознать
        print("!Не распознано!")

    except sr.RequestError as e: #если не получилось отправить для распознавания
        print("Ошибка сервиса; {0}".format(e))

    #проверяем есть ли во фразе активационное "голосовое управление"
    for w in config.voceon:
        if w in recog:
            for w in config.onon: #если включить
                if w in recog:
                    voicecontr = True
                    writecontr = False
                    talk ("Включено голосовое управление")
                    break
            for w in config.ofof: #если выключить
                if w in recog:
                    voicecontr = False
                    talk ("Голосовое управление отключено")
                    break
            break
        else:
            pass
    #проверяем есть ли во фразе активационное "режим печати"
    for w in config.writeon:
        if w in recog:
            for w in config.onon: #если включить
                if w in recog:
                    voicecontr = False
                    writecontr = True
                    talk ("Режим печати активирован")
                    recog = ""
                    break
            for w in config.ofof: #если выключить
                if w in recog:
                    writecontr = False
                    talk ("Деактивация режима печати")
                    break
            break
        else:
            pass

#основная проверка в режиме голосового управления
    if voicecontr == True:
        for w in config.window: # "окно"
            if w in recog:
                for w in config.left: # "лево"
                    if w in recog:
                        pyautogui.hotkey('win','left')
                        time.sleep(1)
                        pyautogui.press('esc')
                        print ('Окно слева')
                        break
                for w in config.right:# "право"
                    if w in recog:
                        pyautogui.hotkey('win','right')
                        time.sleep(1)
                        pyautogui.press('esc')
                        print ('Окно справа')
                        break
                for w in config.ofof:# "закрыть"
                    if w in recog:
                        pyautogui.hotkey('alt','f4')
                        print ('Закрыла окно')
                        break
                for w in config.collapse:# "свернуть"
                    if w in recog:
                        pyautogui.hotkey('win','down')
                        print ('Свернула окно')
                        break
                for w in config.uncollapse:# "развернуть"
                    if w in recog:
                        pyautogui.hotkey('win','up')
                        break
                        print ('Развернула окно')
            else:
                pass
        for w in config.allall: # "все"
            if w in recog:
                for w in config.collapse:# "свернуть"
                    if w in recog:
                        pyautogui.hotkey('win','d')
                        print ('Все свернула')
                        break
                for w in config.uncollapse:# "развернуть"
                    if w in recog:
                        pyautogui.hotkey('win','shift','m')
                        print ('Все развернула')
                        break
            else:
                pass
        for w in config.down: # "дальше"
            if w in recog:
                pyautogui.press('pagedown')
                print ('Дальше')
                break
        for w in config.up: # "вернуть"
            if w in recog:
                pyautogui.press('pageup')
                print ('Вернуть')
                break
        for w in config.back: # "назад"
            if w in recog:
                pyautogui.hotkey('alt','left')
                print ('Назад в браузере')
                break
        for w in config.go:# "вперед"
            if w in recog:
                pyautogui.hotkey('alt','right')
                print ('Вперед в браузере')
                break
        for w in config.play: # "проиграть"
            if w in recog:
                pyautogui.press('playpause')
                print ('Play-pause')
                break
        for w in config.pause: # "стоп"
            if w in recog:
                pyautogui.press('stop')
                print ('Player Stop')
                break
        for w in config.finder: # "поиск"
            if w in recog:
                pyautogui.press('win')
                findstr = clean_finder (recog,w)
                print ('Ищу '+findstr)
                pywinauto.keyboard.SendKeys(findstr)
                break
        for w in config.copier: # "копировать"
            if w in recog:
                pyautogui.hotkey('ctrl','c')
                print ('Скопировала')
                talk ('Скопировала')
                break
        for w in config.cutter: # "вырезать"
            if w in recog:
                pyautogui.hotkey('ctrl','x')
                print ('Вырезала')
                talk ('Вырезала')
                break
        for w in config.inserter: # "вставить"
            if w in recog:
                pyautogui.hotkey('ctrl','v')
                print ('Вставила')
                talk ('Вставила')
                break
        #для открытия сайтов
        for w in config.site:
            if w in recog:
                url = clean_finder (recog,w)
                webbrowser.open(url)
                print ('открываю сайт '+url)
                break

        #проверка для запуска и закрытия приложений
        for w in config.onon:
            if w in recog:
                progssin (recog,1)
                print ('Приложение:'+recog)
                break
        for w in config.ofof:
            if w in recog:
                progssin (recog,0)
                print ('Приложение:'+recog)
                break

#блок режжима печати----------------------
    if writecontr == True:
        print("Режим печати...")
        if recog != "":
            if firstwrite==1: #если первый раз то пробел не ставим вначале
                textrec = clean_textn (recog)
                firstwrite=0
            if firstwrite==0:
                textrec = "{SPACE}"+clean_textn (recog)
                print ('Печатаю:'+textrec)
            pywinauto.keyboard.SendKeys(textrec) #печатаем с клавиатуры
        else:
            pass
