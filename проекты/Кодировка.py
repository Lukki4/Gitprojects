import configparser



def cod():
    global codec
    codec = ''
    codecs = ['cp1251', 'cp1252', 'utf-8']
    fig = configparser.ConfigParser()
    for j in codecs:
        try:
            fig.read('path.ini', encoding=j)
            directory = fig.get('con', 'directory') + '\\'  # директория где файлы лежат
            if 'Р' in directory or '°' in directory or '±' in directory or '†' in directory:
                continue
            else:
                codec = j
                break
        except UnicodeDecodeError:
            continue
    if codec == '':
        print('Не нашлось изестных кодировок. Принтскрин ошибки пришлите на электроннyю почту: ProhorenkoSV@nesk.ru')


cod()




