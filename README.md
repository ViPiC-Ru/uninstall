# Описание
`JScript` для **удаления** или **инвентаризации** приложений через `WMI` на **локальном** или **удалённом** компьютере в сети. Полученный список приложений можно **экспортировать** в файл, поддерживается несколько форматов. Поддерживаются фильтры по **названию**, **автору** и **версии** приложения. Для команды на удаление можно передавать дополнительные параметры. Доступны ограничения по **области установки** и **типу** (разрядности) приложения.

# Использование
В командной строке **Windows** введите следующую команду. Если необходимо скрыть отображение окна консоли, то вместо `cscript` можно использовать `wscript`.
```bat
cscript uninstall.min.js [\\<context>] [<output>] [<type>] [<scope>] [<option>...]
                         <author> <name> <version>
                         [<argument>...]
```
- `<context>` - В контексте какого компьютера выполнить действия.
- `<output>` - Формат текстовых данных стандартного потока вывода для **экспорта** списка приложений.
    - **txt** - Отправляет в поток данные со списком приложений `txt` формате.
    - **csv** - Отправляет данные в `csv` формате (заглавное написание добавляет ещё и заголовок).
    - **tsv** - Отправляет данные в `tsv` формате (заглавное написание добавляет ещё и заголовок).
- `<type>` - Тип приложений, участвующих в проверке (регистр не важен).
    - **native** - Разрядность проверяемых приложений совпадает с **системой**.
    - **x64** - Только **64 разрядные** приложения участвуют в проверке.
    - **x86** - Только **32 разрядные** приложения участвуют в проверке.
- `<scope>` - Область установки, участвующая в проверке (регистр не важен).
    - **computer** - Только приложения, установленные для **всех пользователей**.
    - **user** - Только приложения, установленные для **текущего пользователя**.
- `<option>` - Дополнительные опции (может быть несколько, порядок и регистр не важен).
    - **hidden** - Проверять также приложения не отображающиеся в списке системы.
- `<author>` - Фильтр по автору (в формате `VAL|VAL!NOT` регистр не важен).
- `<name>` - Фильтр по названию (в формате `VAL|VAL!NOT` регистр не важен).
- `<version>` - Фильтр по версии (в формате `VAL|VAL!NOT` регистр не важен).
- `<argument>` - Аргументы, добавляемые к команде на удаление (может быть несколько).

Возвращает количество найденных приложений.

# Примеры использования

## Получение данных
Вывести в консоль список установленных приложений.
```bat
cscript uninstall.min.js txt
```
Вывести в консоль список установленных **64 разрядных** приложений от **автора**, содержащего фразу `Microsoft` или `Майкрософт`.
```bat
cscript uninstall.min.js txt x64 "Microsoft|Майкрософт"
```
Вывести в консоль **все** установленные **32 разрядные** приложения для **всех пользователей** от **автора**, содержащего фразу `Microsoft`, и имеющих в своём названии слово `Office` и не содержащих слово `Plugin`. И сделать всё это в контексте компьютера `RUS000WS001`.
```bat
cscript uninstall.min.js txt x86 computer hidden \\RUS000WS001 "Microsoft" "Office!Plugin"
```

## Экспорт данных
Экспортировать список приложений в `csv` файл без заголовка и с кодировкой `UTF-16 LE` в контексте компьютера `RUS000WS001`.
```bat
cscript /nologo /u uninstall.min.js \\RUS000WS001 csv > RUS000WS001.csv
```
Экспортировать список приложений в `csv` файл с заголовком и с кодировкой `UTF-16 LE` в контексте компьютера `RUS000WS001`.
```bat
cscript /nologo /u uninstall.min.js \\RUS000WS001 CSV > RUS000WS001.csv
```

## Удаление приложений
Выполнить тихое удаление приложения `OneDrive`, установленного для **текущего пользователя**.
```bat
wscript uninstall.min.js user "" "OneDrive" "" ""
```
Выполнить удаление приложения `Office` от автора `Microsoft` в контексте компьютера `RUS000WS001` с дополнительными параметрами.
```bat
cscript uninstall.min.js \\RUS000WS001 "Microsoft" "Office" "" /quiet /norestart
```

## Инвентаризация приложений на компьютерах
Загрузить из `txt` файла список компьютеров и сохранить список установленных на них приложений в папке `inventory` в виде `tsv` файлов без заголовков.
```bat
for /f "eol=; tokens=* delims=, " %%i in (list.txt) do (
    cscript /nologo /u uninstall.min.js \\%%i tsv > inventory\%%i.tsv
)
```
Загрузить из `txt` файла список компьютеров и сохранить список установленных на них приложений в один `csv` файл с заголовком.
```bat
cscript /nologo /u uninstall.min.js \\ CSV > inventory.csv
for /f "eol=; tokens=* delims=, " %%i in (list.txt) do (
    cscript /nologo /u uninstall.min.js \\%%i csv >> inventory.csv
)
```