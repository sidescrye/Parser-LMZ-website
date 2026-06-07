import pandas as pd

# === 1. Общая информация ===
company = pd.DataFrame([{
    "Название": "ООО ЛМЗ СТАРТ",
    "Полное название": "Арзамасский литейно-механический завод",
    "Телефон": "+7 (83147) 6-04-96",
    "Email": "info@lmzstart.ru",
    "Адрес": "г. Арзамас, ул. Льва Толстого, 61",
    "Опыт": "Более 35 лет (с 1991 г.)",
    "Площадь": "7000 кв.м (+ 2700 кв.м в стройке)",
    "Сертификация": "ИСО ГОСТ Р 9001-2015",
    "Сайт": "https://lmzstart.ru/"
}])

# === 2. Услуги ===
BASE = "https://lmzstart.ru"
services = pd.DataFrame([
    ("Литье цветных металлов", "/uslugi/lite-metallov.html"),
    ("Литье пластмасс под давлением", "/uslugi/lite-plastmass.html"),
    ("Фрезерная обработка", "/uslugi/frezernaya-obrabotka.html"),
    ("Токарная обработка", "/uslugi/tokarnaya-obrabotka.html"),
    ("Штамповка листового металла", "/uslugi/shtampovka-listovogo-metalla.html"),
    ("Вакуумная пропитка", "/uslugi/vakuumnaya-propitka.html"),
    ("Гидроабразивная резка", "/uslugi/gidroabrazivnaya-rezka.html"),
    ("Дробеструйная установка", "/uslugi/drobestrujnaya-ustanovka.html"),
    ("Электроэрозионная проволочно-вырезная обработка",
     "/uslugi/elektroerozionnaya-provolochno-vyireznaya-obrabotka.html"),
], columns=["Услуга", "URL"])
services["URL"] = BASE + services["URL"]

# === 3. Преимущества ===
advantages = pd.DataFrame([
    ("01", "Устойчивое развитие"),
    ("02", "Инновации и автоматизация"),
    ("03", "Бережливое производство"),
    ("04", "Высокая квалификация"),
    ("05", "Клиентоориентированность"),
], columns=["№", "Преимущество"])

# === 4. Этапы работ ===
stages = pd.DataFrame([
    ("01", "Проектирование и заключение договора",
     "Технический аудит, разработка ТЗ или анализ КД, согласование параметров заказа."),
    ("02", "Производство и контроль качества",
     "Изготовление продукции на современном оборудовании с контролем качества на каждом этапе."),
    ("03", "Поставка и гарантия",
     "Отгрузка в срок, транспортировка до объекта заказчика с документами и сертификатами."),
], columns=["Этап", "Название", "Описание"])

# === 5. Партнёры ===
partners = pd.DataFrame([
    ("НПЦ КЗ", "https://npckz.ru/"),
    ("Asahi (АДС)", "https://asahi.ru/"),
    ("АМЗ", ""),
    ("Лада-Имидж", "https://www.lada-image.ru/"),
    ("НАЗ", "https://naz.ru/"),
    ("Техномер", "https://tehnomer.ru/ru/"),
    ("Tolvo", "https://tolvo.llc/"),
    ("УралАЗ", "https://uralaz.ru/"),
    ("АвтоВАЗ", "https://info.avtovaz.ru/"),
    ("ЯМЗ Мотор", "https://www.ymzmotor.ru/"),
    ("УАЗ", "https://www.uaz.ru/"),
    ("АПЗ", "https://aoapz.ru/"),
    ("Автопровод", "https://www.avtoprovod.ru/"),
    ("Красное Знамя", "https://www.kznamya.ru/"),
    ("ПАЗ", "https://paz-bus.ru/"),
    ("Темп-Авиа", "https://www.temp-avia.ru/"),
], columns=["Партнёр", "Сайт"])

# === 6. Меню навигации ===
menu = pd.DataFrame([
    ("Главная", "/"),
    ("Услуги", "/uslugi/"),
    ("О предприятии", "/o-prepriyatii.html"),
    ("Качество", "/kachestvo.html"),
    ("Новости", "/novosti/"),
    ("Карьера", "/karera/"),
    ("Вакансии", "/karera/vakansii.html"),
    ("Контакты", "/kontaktyi.html"),
], columns=["Раздел", "URL"])
menu["URL"] = BASE + menu["URL"]

# === Сохраняем в один Excel-файл со множеством листов ===
with pd.ExcelWriter("lmzstart_data.xlsx", engine="openpyxl") as writer:
    company.to_excel(writer, sheet_name="Компания", index=False)
    services.to_excel(writer, sheet_name="Услуги", index=False)
    advantages.to_excel(writer, sheet_name="Преимущества", index=False)
    stages.to_excel(writer, sheet_name="Этапы работ", index=False)
    partners.to_excel(writer, sheet_name="Партнёры", index=False)
    menu.to_excel(writer, sheet_name="Меню сайта", index=False)

# === И отдельные CSV-файлы ===
company.to_csv("company.csv", index=False, encoding="utf-8-sig")
services.to_csv("services.csv", index=False, encoding="utf-8-sig")
advantages.to_csv("advantages.csv", index=False, encoding="utf-8-sig")
stages.to_csv("stages.csv", index=False, encoding="utf-8-sig")
partners.to_csv("partners.csv", index=False, encoding="utf-8-sig")
menu.to_csv("menu.csv", index=False, encoding="utf-8-sig")

print("✅ Готово!")
print("📂 Создан Excel-файл: lmzstart_data.xlsx (6 листов)")
print("📂 Созданы CSV-файлы: company / services / advantages / stages / partners / menu")