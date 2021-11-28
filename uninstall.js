/* 1.0.0 удаляет или инвентаризирует заданные приложения

cscript uninstall.min.js [\\<context>] [<output>] [<type>] [<scope>] [<option>...] <author> <name> <version> [<argument>...]

<context>       - В контексте какого компьютера выполнить действия.
<output>        - Формат текстовых данных стандартного потока вывода.
    txt         - Вывести список найденных приложений в простом формате.
    csv         - Вывести в csv формате (заглавное написание добавляет заголовок).
    tsv         - Вывести в tsv формате (заглавное написание добавляет заголовок).
<type>          - Тип приложений, участвующих в проверке (регистр не важен).
    native      - Разрядность проверяемых приложений совпадает с системой.
    x64         - Только 64 разрядные приложения участвуют в проверке.
    x86         - Только 32 разрядные приложения участвуют в проверке.
<scope>         - Область установки, участвующая в проверке (регистр не важен).
    computer    - Только приложения, установленные для всех пользователей.
    user        - Только приложения, установленные для текущего пользователя.
<option>        - Дополнительные опции (может быть несколько, порядок и регистр не важен).
    hidden      - Проверять также приложения не отображающиеся в списке системы.
<author>        - Фильтр по автору (в формате VAL|VAL!NOT регистр не важен).
<name>          - Фильтр по названию (в формате VAL|VAL!NOT регистр не важен).
<version>       - Фильтр по версии (в формате VAL|VAL!NOT регистр не важен).
<argument>      - Аргументы, добавляемые к команде на удаление (может быть несколько).
                  Если не переданы аргументы, то удаление не происходит, только поиск.
                  Если передан единственный и пустой аргумент, то для удаления
                  выбирается команда с тихим режимом, если она есть.

Возвращает количество найденных приложений.

*/

var uninstall = new App({
    argWrap: '"',                                       // основное обрамление аргументов
    argDelim: " ",                                      // разделитель значений агрументов
    linDelim: "\r\n",                                   // разделитель строк значений
    keyDelim: "\\",                                     // разделитель ключевых значений
    csvDelim: ";",                                      // разделитель значений для файла выгрузки csv
    tsvDelim: "\t"                                      // разделитель значений для файла выгрузки tsv
});

// подключаем зависимые свойства приложения
(function (wsh, app, undefined) {
    app.lib.extend(app, {
        fun: {// зависимые функции частного назначения
        },
        init: function () {// функция инициализации приложения
            var pid, key, value, index, length, list, item, items, locator, cim, registry, runtime, command,
                name, names, response, method, param, branch, application, filter, data, offset, delim, unit,
                columns, isBreak, isMatch, isAddType, isVisible, branches = [], applications = [],
                host = "", type = "", config = {}, timeout = 1000;

            locator = new ActiveXObject("wbemScripting.Swbemlocator");
            locator.security_.impersonationLevel = 3;// Impersonate
            // получаем основные параметры
            length = wsh.arguments.length;// получаем длину
            for (index = 0; index < length; index++) {// пробигаемся по параметрам
                value = wsh.arguments.item(index);// получаем очередное значение
                // контекст выполнения
                key = "context";// ключ проверяемого параметра
                if (!(key in config)) {// если нет в конфигурации
                    list = value.split(app.val.keyDelim);// вспомогательный список
                    if (3 == list.length && !list[0] && !list[1]) {// если пройдена проверка
                        config[key] = list[2];// задаём значение
                        continue;// переходим к следующему параметру
                    };
                };
                // экспорт данных
                key = "output";// ключ проверяемого параметра
                if (!(key in config)) {// если нет в конфигурации
                    list = ["txt", "csv", "tsv", "CSV", "TSV"];// разрешённые значения
                    if (app.lib.hasValue(list, value, true)) {// если пройдена проверка
                        config[key] = value;// задаём значение
                        continue;// переходим к следующему параметру
                    };
                };
                // тип архитектуры приложения
                key = "type";// ключ проверяемого параметра
                if (!(key in config)) {// если нет в конфигурации
                    list = ["x86", "x64", "native"];// разрешённые значения
                    if (app.lib.hasValue(list, value, false)) {// если пройдена проверка
                        value = value.toLowerCase();
                        config[key] = value;// задаём значение
                        continue;// переходим к следующему параметру
                    };
                };
                // область установки
                key = "scope";// ключ проверяемого параметра
                if (!(key in config)) {// если нет в конфигурации
                    list = ["computer", "user"];// разрешённые значения
                    if (app.lib.hasValue(list, value, false)) {// если пройдена проверка
                        value = value.toLowerCase();
                        config[key] = value;// задаём значение
                        continue;// переходим к следующему параметру
                    };
                };
                // отображение скрытых программ
                key = "hidden";// ключ проверяемого параметра
                if (!(key in config)) {// если нет в конфигурации
                    if (!app.lib.compare(key, value, true)) {// если пройдена основная проверка
                        config[key] = true;// задаём значение
                        continue;// переходим к следующему параметру
                    };
                };
                // если закончились параметры конфигурации
                break;// остававливаем получние параметров
            };
            // получаем фильтрующие параметры
            offset = index;// запоминаем смещение по параметрам
            for (index = offset; index < length; index++) {// пробигаемся по параметрам
                value = wsh.arguments.item(index);// получаем очередное значение
                // автор приложения
                key = "author";// ключ проверяемого параметра
                if (!(key in config)) {// если нет в конфигурации
                    config[key] = value;// задаём значение
                    continue;// переходим к следующему параметру
                };
                // название приложения
                key = "name";// ключ проверяемого параметра
                if (!(key in config)) {// если нет в конфигурации
                    config[key] = value;// задаём значение
                    continue;// переходим к следующему параметру
                };
                // версия приложения
                key = "version";// ключ проверяемого параметра
                if (!(key in config)) {// если нет в конфигурации
                    config[key] = value;// задаём значение
                    continue;// переходим к следующему параметру
                };
                // если закончились параметры конфигурации
                break;// остававливаем получние параметров
            };
            // получаем агрументы для удаления
            offset = index;// запоминаем смещение по параметрам
            for (index = offset; index < length; index++) {// пробигаемся по параметрам
                value = wsh.arguments.item(index);// получаем очередное значение
                key = "arguments";// ключ проверяемого параметра
                if (!(key in config)) {// если нет в конфигурации
                    config[key] = [];// сбрасываем значение
                    if (!value && index == length - 1) {// если единственный и пустой аргумент
                        continue;// переходим к следующему параметру
                    };
                };
                if (!value || -1 != value.indexOf(app.val.argDelim)) {// если есть разделитель
                    value = app.val.argWrap + value + app.val.argWrap;
                };
                config[key].push(value);
            };
            // вносим поправки для конфигурации
            if (!("context" in config)) config.context = ".";
            // создаём служебные объекты
            if (config.context) {// если есть контекст выполнения
                for (index = 1; index; index++) {
                    try {// пробуем подключиться к компьютеру
                        switch (index) {// последовательно создаём объекты
                            case 1: cim = locator.connectServer(config.context, "root\\CIMV2"); break;
                            case 2: runtime = cim.get("Win32_Process"); break;// среда для удалённого запуска процессов
                            case 3: registry = locator.connectServer(config.context, "root\\default").get("stdRegProv"); break;
                            default: index = -1;// завершаем создание
                        };
                    } catch (e) {// при возникновении ошибки
                        switch (index) {// последовательно сбрасываем объекты
                            case 1: cim = null; index = -1; break;// завершаем создание
                            case 2: runtime = null; index = -1; break;// завершаем создание
                            case 3: registry = null; index = -1; break;// завершаем создание
                        };
                    };
                };
            };
            // получаем необходимые данные
            if (cim && runtime && registry) {// если удалось получить доступ к объектам
                // получаем информацию о системе
                response = cim.execQuery(
                    "SELECT dnsHostName, name, systemType" +
                    " FROM Win32_ComputerSystem"
                );
                items = new Enumerator(response);
                while (!items.atEnd()) {// пока не достигнут конец
                    item = items.item();// получаем очередной элимент коллекции
                    items.moveNext();// переходим к следующему элименту
                    // характеристики
                    if (value = item.dnsHostName) host = value;
                    if (value = item.name) if (!host) host = value.toLowerCase();
                    type = item.systemType && item.systemType.indexOf("64") ? "x64" : "x86";
                    // останавливаемся на первом элименте
                    break;
                };
                // формируем список веток реестра для проверки
                if ("computer" == config.scope || !config.scope) {// если компьютер
                    if ("native" == config.type || type == config.type || !config.type) {// если нативные
                        branches.push({// приложения в нативном реестре
                            root: 0x80000002,// HKEY_LOCAL_MACHINE
                            path: "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall",
                            scope: "Computer",
                            type: type
                        });
                    };
                    if ("x64" == type && ("x86" == config.type || !config.type)) {// если не нативные
                        branches.push({// приложения в x86 реестре
                            root: 0x80000002,// HKEY_LOCAL_MACHINE
                            path: "SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall",
                            scope: "Computer",
                            type: "x86"
                        });
                    };
                };
                if ("user" == config.scope || !config.scope) {// если пользователь
                    if ("native" == config.type || !config.type) {// если нативные
                        branches.push({// приложения в реестре пользователя
                            root: 0x80000001,// HKEY_CURRENT_USER
                            path: "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall",
                            scope: "User",
                            type: ""
                        });
                    };
                };
                // формируем список приложений
                data = {// данные для трансформации значений реестра в объект приложения
                    uninstall: { method: "GetStringValue", name: "UninstallString", value: "sValue" },
                    package: { method: "GetDWORDValue", name: "WindowsInstaller", value: "uValue" },
                    сomponent: { method: "GetDWORDValue", name: "SystemComponent", value: "uValue" },
                    parent: { method: "GetStringValue", name: "ParentKeyName", value: "sValue" },
                    name: { method: "GetStringValue", name: "DisplayName", value: "sValue" },
                    author: { method: "GetStringValue", name: "Publisher", value: "sValue" },
                    version: { method: "GetStringValue", name: "DisplayVersion", value: "sValue" },
                    install: { method: "GetStringValue", name: "InstallDate", value: "sValue" },
                    silent: { method: "GetStringValue", name: "QuietUninstallString", value: "sValue" }
                };
                for (var i = 0, iLen = branches.length; i < iLen; i++) {
                    branch = branches[i];// получаем очередной элимент
                    isBreak = false;// нужно ли прикратить обработку
                    // выполняем получение названий дочерних веток
                    if (!isBreak) {// если нужно продолжить обработку
                        method = registry.methods_.item("EnumKey");
                        param = method.inParameters.spawnInstance_();
                        param.hDefKey = branch.root;
                        param.sSubKeyName = branch.path;
                        try {// пробуем выполнить запрос данных
                            item = registry.execMethod_(method.name, param);
                        } catch (e) {// если возникла ошибка
                            isBreak = true;
                        };
                    };
                    // проверяем успешность получения данных
                    if (!isBreak) {// если нужно продолжить обработку
                        if (!item.returnValue) {// если данные получены
                        } else isBreak = true;
                    };
                    // преобразовываем полученные данные
                    if (!isBreak) {// если нужно продолжить обработку
                        try {// пробуем выполнить преобразование данных
                            names = item.sNames.toArray();
                        } catch (e) {// если возникла ошибка
                            isBreak = true;
                        };
                    };
                    // выполняем получение ключей для дочерних веток
                    if (!isBreak) {// если нужно продолжить обработку
                        for (var j = 0, jLen = names.length; j < jLen; j++) {
                            name = names[j];// получаем очередное значение
                            application = {};// сбрасываем значение
                            isMatch = true;// значение соответствует фильтру
                            isVisible = null;// сбрасываем значение
                            // последовательно получаем данные по ключам
                            for (var key in data) {// пробигаемся по ключам
                                // выполняем получение даннных по ключу
                                method = registry.methods_.item(data[key].method);
                                param = method.inParameters.spawnInstance_();
                                param.hDefKey = branch.root;
                                param.sSubKeyName = branch.path + "\\" + name;
                                param.sValueName = data[key].name;
                                item = registry.execMethod_(method.name, param);
                                value = (!item.returnValue ? item[data[key].value] : null) || "";
                                // выполняем проверку соответствия фильтрам
                                switch (key) {// поддерживаемые фильтры
                                    case "uninstall":// команда для удаления
                                    case "package":// установленный пакет
                                        isVisible = isVisible || value;
                                        break;
                                    case "сomponent":// компонент системы
                                    case "parent":// обновление приложения
                                        isVisible = isVisible && !value;
                                        break;
                                    case "name":// название приложения
                                        isMatch = isMatch && value && (isVisible || config.hidden);
                                    case "author":// автор приложения
                                    case "version":// версия приложения
                                        filter = config[key] || "";// фильтр для значения
                                        isMatch = isMatch && app.lib.match(value, filter);
                                        break;
                                };
                                // выполняем преобразование значения
                                switch (key) {// поддерживаемые фильтры
                                    case "install":// дата установки
                                        if (8 == value.length && !isNaN(value)) {
                                            value = new Date(// преобразуем в дату
                                                1 * value.substr(0, 4),
                                                1 * value.substr(4, 2) - 1,
                                                1 * value.substr(6, 2),
                                                0, 0, 0 // 00:00:00
                                            );
                                        } else value = "";
                                        break;
                                };
                                // добавляем значение в объект
                                if (isMatch) application[key] = value;
                                else break;
                            };
                            // добавляем приложение в список
                            if (isMatch) {// если приложение прошло проверку
                                application.scope = branch.scope;
                                application.type = branch.type;
                                applications.push(application);
                            };
                        };
                    };
                };
                // удаляем приложения по списку
                if (config.arguments) {// если передан аргумент для удаления
                    for (var i = 0, iLen = applications.length; i < iLen; i++) {
                        application = applications[i];// получаем очередной элимент
                        // получаем базовую команду на удаление
                        if (config.arguments.length) command = application.uninstall;
                        else command = application.silent || application.uninstall;
                        isBreak = !command;// нужно ли прикратить обработку
                        // поправка для полных путей до файлов без кавычек
                        if (!isBreak) {// если нужно продолжить обработку
                            delim = app.val.keyDelim + app.val.keyDelim;
                            value = command.split(app.val.keyDelim).join(delim);
                            value = value.split("'").join(app.val.keyDelim + "'");
                            response = cim.execQuery(
                                "SELECT name" +
                                " FROM CIM_DataFile" +
                                " WHERE name = '" + value + "'"
                            );
                            items = new Enumerator(response);
                            while (!items.atEnd()) {// пока не достигнут конец
                                item = items.item();// получаем очередной элимент коллекции
                                items.moveNext();// переходим к следующему элименту
                                // характеристики
                                command = app.val.argWrap + command + app.val.argWrap;
                                // останавливаемся на первом элименте
                                break;
                            };
                        };
                        // поправка для msi пакетов приложений
                        if (!isBreak) {// если нужно продолжить обработку
                            if (app.lib.hasValue(command, "MsiExec", false)) {
                                command = command.replace("/I{", "/X{").replace("/i{", "/X{");
                                if (!config.arguments.length) {// если нет аргументов
                                    key = "/quiet";// ключ тихой установки
                                    if (!app.lib.hasValue(command, key, false)) {
                                        command += app.val.argDelim + key;
                                    };
                                };
                            };
                        };
                        // добавляем переданный аргумент в команду
                        if (!isBreak) {// если нужно продолжить обработку
                            if (config.arguments.length) {// если нужно добавить аргумент
                                value = config.arguments.join(app.val.argDelim);
                                command += app.val.argDelim + value;
                            };
                        };
                        // выполняем удалённый вызов команды
                        if (!isBreak) {// если нужно продолжить обработку
                            method = runtime.methods_.item("Create");
                            param = method.inParameters.spawnInstance_();
                            param.CommandLine = command;
                            item = runtime.execMethod_(method.name, param);
                            if (item.processId) {// если получен мдентификатор
                                pid = item.processId;
                            } else isBreak = true;
                        };
                        // ожидаем завершения выполнения
                        while (!isBreak) {// если нужно продолжить обработку
                            isBreak = true;// нужно ли прикратить обработку
                            response = cim.execQuery(
                                "SELECT handle" +
                                " FROM Win32_Process" +
                                " WHERE processId = '" + pid + "'" +
                                " OR parentProcessId = '" + pid + "'"
                            );
                            items = new Enumerator(response);
                            while (!items.atEnd()) {// пока не достигнут конец
                                item = items.item();// получаем очередной элимент коллекции
                                items.moveNext();// переходим к следующему элименту
                                // характеристики
                                isBreak = false;// нужно ли прикратить обработку
                                wsh.sleep(timeout);// деламем паузу между проверками
                                // останавливаемся на первом элименте
                                break;
                            };
                        };
                    };
                };
            };
            // готовим данные в поток вывода
            if (config.output) {// если нужно вывести данные
                delim = "";// сбрасываем разделитель значений
                isAddType = null;// сбрасываем значение
                switch (config.output) {// поддерживаемые служебные параметры
                    case "txt":// значения для формата txt
                        items = [];// сбрасываем значение
                        for (var i = 0, iLen = applications.length; i < iLen; i++) {
                            application = applications[i];// получаем очередной элимент
                            list = [];// сбрасываем значение
                            // характеристики
                            if (value = application.name) list.push(value);
                            if (list.length && !app.lib.hasValue(application.name, application.version, false)) {
                                if (value = application.version) list.push(value);
                            };
                            if (list.length && !app.lib.hasValue(application.name, application.type, false)) {
                                if (value = application.type) list.push(value);
                            };
                            // добавляем массив в список
                            item = list.join(app.val.argDelim);
                            items.push(item);
                        };
                        value = items.join(app.val.linDelim);
                        break;
                    case "TSV":// заголовки для формата tsv
                        if (!delim) delim = app.val.tsvDelim;
                    case "CSV":// заголовки для формата csv
                        if (!delim) delim = app.val.csvDelim;
                        isAddType = false;// добавляем заголовок
                    case "tsv":// значения для формата tsv
                        if (!delim) delim = app.val.tsvDelim;
                    case "csv":// значения для формата csv
                        if (!delim) delim = app.val.csvDelim;
                        items = [];// сбрасываем значение
                        columns = [// список выводимых данных для вывода в одну строку
                            "NET-HOST", "APP-SCOPE", "APP-TYPE", "APP-INSTALL",
                            "APP-AUTHOR", "APP-NAME", "APP-VERSION"
                        ];
                        for (var i = 0, iLen = applications.length; i < iLen; i++) {
                            application = applications[i];// получаем очередной элимент
                            data = {};// сбрасываем значение
                            // характеристики
                            if (value = host) data["NET-HOST"] = value;
                            if (value = application.scope) data["APP-SCOPE"] = value;
                            if (value = application.type) data["APP-TYPE"] = value;
                            if (unit = application.install) data["APP-INSTALL"] = app.lib.date2str(unit, "d.m.Y H:i:s");
                            if (value = application.author) data["APP-AUTHOR"] = value;
                            if (value = application.name) data["APP-NAME"] = value;
                            if (value = application.version) data["APP-VERSION"] = value;
                            // добавляем объект в список
                            items.push(data);
                        };
                        value = app.lib.arr2tsv(items, columns, delim, isAddType);
                        break;
                };
            };
            // отправляем данные с поток вывода
            if (config.output && value) {// если нужно выполнить
                value += app.val.linDelim;
                // отправляем текстовые данные в поток
                try {// пробуем отправить данные
                    wsh.stdOut.write(value);
                } catch (e) { };// игнорируем исключения
            };
            // завершаем сценарий кодом
            wsh.quit(applications.length);
        }
    });
})(WSH, uninstall);
// запускаем инициализацию
uninstall.init();