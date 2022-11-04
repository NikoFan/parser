import requests
from bs4 import BeautifulSoup
import openpyxl as opx

'''pageUrl = "https://www.kids-team.ru/catalog/shkolnaya_odezhda/"

# Применение request на idUrl
pageRsponse = requests.get(pageUrl)
# Получение HTML кода страницы
pageSoup = BeautifulSoup(pageRsponse.content, 'html.parser')

# Поиск всех div с классом item_block col-3
pageFind = pageSoup.find("div", class_ = "nums").text
print(pageFind)
print("---------------------------")

pageArr = []
for i in pageFind:
    if len(i) >= 1 and i != "\n":
        pageArr.append(i)
print(pageArr)'''






class Parser():
    """Описание работы парсера для считывания данных со страницы интернет 
    магазина детской одежды"""
    
    
    def __init__(self, idUrl, itemsUrl, url, idSet,linkItemsSL, idArr, 
                 file, itemSL, sexItemSL, sheet, page, pageN, pageArr, sl, N,
                 fileProduct, sheet_fileProduct, productCount):
        """Инициализируем атрибуты Url 3 видов и добавочные данные"""
        
        
        # Инициализация втрибутов
        self.idUrl = idUrl
        self.itemsUrl = itemsUrl
        self.url = url
        self.idSet = idSet
        self.linkItemsSL = linkItemsSL
        self.idArr = idArr
        self.file = file   
        self.itemSL = itemSL
        self.sexItemSL = sexItemSL
        self.sheet = sheet
        self.pageArr = pageArr
        self.page = page
        self.pageN = pageN
        self.sl = sl
        self.N = N
        self.fileProduct = fileProduct
        self.sheet_fileProduct  =sheet_fileProduct
        self.productCount = productCount
        
        
    def Page(self):
        '''каталог товара на сайте размещен на разных страницах, для их
        парсинга требуется менять URL'''
        
        pageUrl = "https://www.kids-team.ru/catalog/shkolnaya_odezhda/"

        # Применение request на idUrl
        pageRsponse = requests.get(pageUrl)
        # Получение HTML кода страницы
        pageSoup = BeautifulSoup(pageRsponse.content, 'html.parser')

        # Поиск всех div с классом item_block col-3
        pageFind = pageSoup.find("div", class_ = "nums").text

        self.pageArr = []
        for i in pageFind:
            if len(i) >= 1 and i != "\n":
                pageArr.append(i)

        
    def objectID(self):
        """Работа с ID объектов на страничках на основании которых будет 
        проводиться выборка карточек товара для дальнейших действий с ними"""


        idUrl = (
            f"https://www.kids-team.ru/catalog/shkolnaya_odezhda/{self.page}{self.pageN}")
        # Применение request на idUrl
        idRsponse = requests.get(idUrl)
        # Получение HTML кода страницы
        idSoup = BeautifulSoup(idRsponse.content, 'html.parser')

        # Поиск всех div с классом item_block col-3
        IdFind_1 = idSoup.find_all("div", class_ = "item_block col-3")
        # Превращение полученных данных в список, разделяя по пробелам
        IdFind_2 = str(IdFind_1).split(" ")
        
        # Проведениевыборки из массива IdFind_2 основываясь на проверку...
        # ...наличия элементов "id" и "bx" в элементах массива
        self.sl = []
        for Id in IdFind_2:
            if "id" in Id:
                if "bx" in Id:
                    self.sl.append(Id)
        # print(self.sl, " = SL")
        
        # Выборка и запись полученных данных в множество idSet...
        # ...чтобы избежать повторения
        # Что позваоляет получить из таких данных:
            # id="bx_3966226736_52680_sku_tree">
            # <div><div
        # такие:
            # bx_3966226736_52680
        self.idSet = set()
        for i in self.sl:
            trueI = i[4:23]
            if "div" not in trueI:
                self.idSet.add(trueI)
        parser.excelWork1()

                
    def hrefID(self):
        """Узнаем ссылку на объект под определенным ID и считываем с нее даныне
        такие как:
                Пол товара (название) : sex,
                Ссылку на товар : link,
                Вид товара (что это такое, куртка, джинцы и т.д.) : item
                
            После чего добавляем их в списки, с которыми дальше будет 
            проводиться работа по генерации URl карточки товара"""
        
        
        itemsUrl = (
            f"https://www.kids-team.ru/catalog/shkolnaya_odezhda/{self.page}{self.pageN}")
        # Применение request на новый URL - itemUrl
        itemsRsponse = requests.get(itemsUrl)
        # Получение HTML кода страницы
        itemsSoup = BeautifulSoup(itemsRsponse.content, 'html.parser')

        # Создание 3х массивов : Вид одежды; Пол; Ссылка на карточку товара;
        self.itemSL = []
        self.sexItemSL = []
        self.linkItemsSL = []

        # Выделение из общей массы данных нужные нам, такие как:
            # Пол ребенка {Мальчик или Девочка};
            # Вид одежды {джиныс, рубашки и т.д.};
            # Ссылка на товар
        for href in self.idArr:
            # Получение href товара из HTML кода   (*)
            itemNameSoup = itemsSoup.find("div", id = f"{href}")

            # print(itemNameSoup)
            # Получение div с именем товара
            itName = itemNameSoup.find("div", class_ ="item-title")
            # Извлечение имени товара
            itemsHref = itName.find("a")

            # Превращение ссылки на товар (*)  В список разделяя через /
            hrefItems = itemsHref["href"].split("/")
            
            # Получение навзания товара : 
                # {Брюки для мальчика MAYORAL 41-15 синий, ИСПАНИЯ}
            itemNames = itName.find("span").text

            # Добавление Вида, Пола, Ссылки в отдельные массивы
            self.itemSL.append(hrefItems[3])
            self.sexItemSL.append(hrefItems[2])
            self.linkItemsSL.append(hrefItems[4])
        parser.excelWork2()
        
        
    def parserWork(self):
        """Основная работа парсера. При получении сгенерированного URl карточки
        товара начинается обработки страницы сайта и получение данных:
                Имя товара : nsmeData,
                Цену товара : price,
                Валюту : currency,
                Артикул товара : article
                
            После чего выводим их в строгой последовательности"""


        position = 1
        count = 2
        
        for id1 in self.idArr:
            # Каждый атрибут ссылки получает свое значение из DB.xlsx
            sex = self.sheet[f"D{count}"].value
            clothes = self.sheet[f"C{count}"].value
            link = self.sheet[f"A{count}"].value
            # Конечный сгенерированные URL *открытой карточки товара
            url = (
                f"https://www.kids-team.ru/catalog/{sex}/{clothes}/{link}/")

            # Применение request для сгенерированного URL
            response = requests.get(url)    
            # получение обработанного html кода страницы
            soup = BeautifulSoup(response.content, 'html.parser') 
            
            # Поиск названия товара
            nameData = soup.find("h1").text
            print()
            # Вывод ссылки страницы на которой назодится Парсер
            print(f"Ссылка на товар для поискка в Интернете:\n {url}", "\n")
            
            # Для удобства получения Артикула - разбиваем строку на элементы
            articlePosition = nameData.split(" ")
            # Находим в получившемся массиве индекс Артикула
            article = articlePosition[-3]
            
            # Выводим необходимые данные
            print(
                f"Наименование товара: {' '.join(articlePosition[:-3])}", "\n")
            print(
                f"Страна производитель: {articlePosition[-1]}", "\n")
            print(
                f"Артикул товара: {article}", "\n")

            # price
            '''На открытой карточке товара ID товара меняется,
            но меняется не полностью, что позволяет найти закономерность
            и применить для генерации 2 ID для открытой карточки товара'''
            
            
            # Занесение повторяющейся части ID в константу
            CONST_ID = "117848907_"
            # Получение ценника товара
            priceData = soup.find(
                "div", id =f"{id1[:3] + CONST_ID + id1[14:]}_price")
            
            # Выделение цены из HTML кода
            price = priceData.find("span", class_ = "price_value").text
            
            # Выделение валюты из HTML кода
            currency = priceData.find("span", class_ = "price_currency").text
            
            # Вывод цены товара
            print(f"Цена товара: {price}{currency}", "\n")
            count += 1
            
            # Вывод позиции товара на страничке
            print(f"Номер товара в списке: {position}", "\n")
            # Добавляем в Excel таблицу : Название, Артикул, Страна,Цена
            self.sheet_fileProduct[f"A{self.productCount}"] = (
                ' '.join(articlePosition[:-3]))
            self.sheet_fileProduct[f"B{self.productCount}"] = article
            self.sheet_fileProduct[f"C{self.productCount}"] = (
                articlePosition[-1])
            self.sheet_fileProduct[f"D{self.productCount}"] = (
                f"{price}{currency}")
            # Номер строки в Excel
            self.productCount += 1
            # Позиуция товара на страничке
            position += 1
        # Сохраняем изменения в файле
        fileProduct.save("products _nfo.xlsx")
        
        # Проверяем : Если номер на страничке равне длинне массива с ID
        # Значит что вся страничка спаршена, надо переходть на следующую
        if position - 1 == len(self.idSet):
            # Попытка поменять страничку
            try:
                # Обновляем данные
                self.idSet = set()
                self.page = "page"
                self.N += 1
                # Выводим номер провереной странички
                print("номер страницы :", self.N)

                self.pageN = pageArr[self.N] + "/"
                self.itemSL = []
                self.sexItemSL = []
                self.linkItemsSL = []
                self.idArr  = []
                self.sl = []
                parser.objectID()
            except IndexError:
                # если больше страниц нет - заканчиваем
                print("Parser закончил свою работу!!!")
                print("Exit")
            
        # Excel close
        self.file.close()
            
          
    def excelWork1(self):
        """Описание работы Excel по добавлению и обработке данных. Открываем 
        файл и сразу создаем 2 коолонки SEX и TYPE OF CLOTHES (ПОЛребенка и ТИП
                                                               ОДЕЖДы)"""


        # открытие Excel файла
        file = opx.open("DDB.xlsx")
        # Назначение переменной ЛИСТ
        sheet = file.active
        # Создание в таблице колонок ПОЛ и ТИП ОДЕЖДЫ
        self.sheet[f"C1"] = "Type of clothes"
        self.sheet[f"D1"] = "sex"
        
        # Excel Добавление в БД ID карточек товара
        
        # Создание колонки ID товара
        self.sheet["B1"] = "ID"
        # Создание массива под ID товара формата СПИСОК (Прошлый был множество)
        self.idArr = []
        
        # Перенос данных из прошлого массива в новый
        for i in self.idSet:
                self.idArr.append(i)
        
        # Добавление в таблицу ID каждог отовара
        for ID in range(len(self.idSet)):
            Id = ID + 2
            self.sheet[f"B{Id}"] = self.idArr[ID]

        
        # Сохранение изменений в фалйе
        file.save("DB.xlsx")
        parser.hrefID()
        
       
    def excelWork2(self):
        """Добавление в таблицу объектов 'types' : Тип одежды,
                                            'sex' : пол ребенка"""
            
         
        count = 2
        # Занесение в таблицу ТИПОВ ОДЕЖДЫ
        for types in self.itemSL:
            self.sheet[f"C{count}"] = types
            count += 1

        count = 2
        # Занесение в таблицу ПОЛА 
        for sex in self.sexItemSL:
            self.sheet[f"D{count}"] = sex
            count += 1
        self.file.save("DDB.xlsx")
        parser.excelWork3()
        
        
    def excelWork3(self):
        """Добавление ссылок в таблицу, то из чего в преимуществе будут 
        создаваться конечные URL"""
        
        
        # Создание колонки Ссылка в таблицы
        sheet["A1"] = "Link"
        
        # Занесение в таблицу данных по ссылкам
        for lin in range(len(self.idSet)):
            LINK = lin + 2
            self.sheet[f"A{LINK}"] = self.linkItemsSL[lin]
        
        # Сохранение изменений в таблице
        self.file.save("DDB.xlsx")
        parser.excelProduct()
        
    
    def excelProduct(self):
        '''Процесс создания колонок в Excel таблице для записи туда параметров
        товара таких как : 
            Название продукта,
            Артикул,
            Страну производителя,
            Цена'''
            
            
        self.sheet_fileProduct["A1"] = "Product Name"
        self.sheet_fileProduct["B1"] = "Article"
        self.sheet_fileProduct["C1"] = "Country"
        self.sheet_fileProduct["D1"] = "Price"
        fileProduct.save("products _nfo.xlsx")
        parser.parserWork()
        
        
        
# Шаблоны данных для класса

idUrl = "https://www.kids-team.ru/catalog/shkolnaya_odezhda/"
itemsUrl = "https://www.kids-team.ru/catalog/shkolnaya_odezhda/"
url = ""
idSet = set()
linkItemsSL = []
idArr = []
file = opx.open("DDB.xlsx")  
itemSL = []
sexItemSL = []
sheet = file.active
page = ""
pageN = ""
pageArr = []
sl = []
N = 0

fileProduct = opx.open("products _nfo.xlsx")
sheet_fileProduct = fileProduct.active
productCount = 2

# Вызов класса
parser = Parser(idUrl, itemsUrl, url, idSet, linkItemsSL, idArr, file, itemSL, 
                sexItemSL, sheet, page, pageN, pageArr, sl, N, fileProduct,
                sheet_fileProduct, productCount)

# Вызов элементов класса

parser.Page()
parser.objectID()


# Закрытие файла DB.xlsx и fileProduct.xlsx
file.close()
fileProduct.close()
