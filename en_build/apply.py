# -*- coding: utf-8 -*-
"""Translate every <a:t> text node per slide RU->EN by exact match.
Reports any text node not covered by the map (misses) so coverage is 100%."""
import re, os

SLIDES = r"C:\Users\Aset\Downloads\Telegram Desktop\efes (5)\efes\en_build\unpacked\ppt\slides"

T = {}

T["slide1.xml"] = {
 'EFES KAZAKHSTAN': 'EFES KAZAKHSTAN',
 'Отдел продаж': 'Sales Department',
 'БИЗНЕС-КЕЙС · АВТОМАТИЗАЦИЯ': 'BUSINESS CASE · AUTOMATION',
 'Efes Nexus': 'Efes Nexus',
 'Менеджер промо-активностей': 'Promo Activity Manager',
 'Единая платформа, которая заменяет ручную работу в Excel и Outlook сквозными цифровыми процессами, авто-уведомлениями и AI-проверкой запросов партнёров.':
   'A single platform that replaces manual work in Excel and Outlook with end-to-end digital processes, automated notifications and AI verification of partner requests.',
 'Промо-активности': 'Promo activities',
 '3 бизнес-процесса': '3 business processes',
 'Efes Kazakhstan': 'Efes Kazakhstan',
 'Efes Nexus · бизнес-кейс для руководителя продаж': 'Efes Nexus · business case for the Head of Sales',
}

T["slide2.xml"] = {
 'BPMN · ПРОЦЕССЫ 01–03': 'BPMN · PROCESSES 01–03',
 'Жизненный цикл промо-активности': 'Promo activity lifecycle',
 'Шаги BPMN-диаграммы. ': 'Steps of the BPMN diagram. ',
 'Синим': 'Blue',
 ' выделено то, что выполняет система; ': ' marks what the system does; ',
 'серым': 'gray',
 ' — действия и решения человека.': ' — human actions and decisions.',
 '01': '01',
 'Создание промо-активности': 'Creating a promo activity',
 '1': '1',
 'TRADE-МАРКЕТИНГ': 'TRADE MARKETING',
 'Канал, партнёр, продукты, ': 'Channel, partner, products, ',
 'квот': 'quot',
 'а': 'a',
 'человек': 'human',
 '→': '→',
 '2': '2',
 'NEXUS': 'NEXUS',
 'Проверка наличия продукта ': 'Checking product availability ',
 'на': 'in',
 ' ': ' ',
 'складе': "the distributor's",
 ' дистрибьютора': ' warehouse',
 'авто': 'auto',
 '3': '3',
 'Присвоение кода PR-##### ': 'Assigning PR-##### code ',
 'и проверка на дубли': 'and checking for duplicates',
 '4': '4',
 'Составление Excel-файла промо': 'Generating the promo Excel file',
 '5': '5',
 'CRM-СПЕЦИАЛИСТ': 'CRM SPECIALIST',
 'Создание промо-активности в Panorama': 'Creating the promo activity in Panorama',
 '6': '6',
 'Авто-уведомление партнёров письмом': 'Auto-notifying partners by email',
 '02': '02',
 'Изменение промо-активности': 'Modifying a promo activity',
 'TM / ON-TRADE': 'TM / ON-TRADE',
 'Запрос: квота, период или продукт': 'Request: quota, period or product',
 'Обновление Excel-файла с изменениями': 'Updating the Excel file with changes',
 'Правка промо-активности в Panorama': 'Editing the promo activity in Panorama',
 'Синхронизация и уведомление партнёров': 'Synchronization and partner notification',
 '03': '03',
 'Закрытие месяца и расчёт': 'Month-end closing and calculation',
 'ПАРТНЁР': 'PARTNER',
 '«Запрос»: проданные литры по промо': '«Request»: liters sold under the promo',
 'Расчёт ∑(литры × цена периода)': 'Calculating ∑(liters × period price)',
 'Верификация «': 'Verification of «',
 'Запрос': 'Request',
 'ов': 's',
 '» и ': '» and ',
 'создание отчета': 'report generation',
 'Автоматизировано системой': 'Automated by the system',
 'Действие / решение человека': 'Human action / decision',
 '8 из 13': '8 of 13',
 '  шагов выполняет система — людям остаются ввод промо и его создание в Panorama':
   '  steps are done by the system — people only handle promo entry and its creation in Panorama',
 'Efes Kazakhstan': 'Efes Kazakhstan',
 'Efes Nexus · BPMN': 'Efes Nexus · BPMN',
}

T["slide3.xml"] = {
 'ОХВАТ': 'SCOPE',
 'Функционал системы – N+1, N+KZT': 'System functionality – N+1, N+KZT',
 '→': '→',
 'Как было:': 'Before:',
 'Разрозненные Excel-файлы и коммуникация через почту': 'Scattered Excel files and email communication',
 '(Отсутствие единого источника данных и ведение коммуникации в письмах → Специалист трейд маркетинга – SSC специалист – CRM специалист - SSC специалист – Партнеры)':
   '(No single source of data, with communication run over email → Trade-marketing specialist – SSC specialist – CRM specialist – SSC specialist – Partners)',
 'Отчеты Promo Status создаются вручную': 'Promo Status reports are created manually',
 '(Существенные временные затраты сотрудника)': '(A significant time burden on the employee)',
 'Присвоение кодов, расчет компенсаций и верификация большого объема запросов в неавтоматизированном процессе':
   'Code assignment, compensation calculation and verification of a large volume of requests in a non-automated process',
 '(Риск дублей, переплат и искажений в отчётах)': '(Risk of duplicates, overpayments and distortions in reports)',
 'Зависимость от человеческого фактора': 'Dependence on the human factor',
 '(300+ промо в месяц)': '(300+ promos per month)',
 'Как будет:': 'After:',
 'Единая система управления промо-процессами': 'A single system to manage promo processes',
 '(Единый реестр и вся коммуникация в одном пространстве → полная прозрачность, мгновенная синхронизация и отсутствие потерь информации)':
   '(A single registry and all communication in one space → full transparency, instant synchronization and no loss of information)',
 'Автоматизированная отчётность': 'Automated reporting',
 '(Формирование и отправка отчетов за секунды → оперативная аналитика и быстрые решения)':
   '(Reports generated and sent in seconds → real-time analytics and fast decisions)',
 'Умные алгоритмы расчётов и кодирования': 'Smart calculation and coding algorithms',
 '(Система сама присваивает код, проверяя при этом дубли, наличие продукта на складе дистрибьютора и осуществляет расчет компенсации, сверяя с запросами партнеров)':
   "(The system assigns the code itself, checking for duplicates and product availability in the distributor's warehouse, and calculates the compensation by reconciling it against partner requests)",
 'Минимизация человеческого фактора': 'Minimizing the human factor',
 '(Стандартизированный и прозрачный процесс → предсказуемый результат и снижение рисков)':
   '(A standardized and transparent process → predictable results and lower risks)',
 'Efes Kazakhstan': 'Efes Kazakhstan',
 'Efes Nexus': 'Efes Nexus',
}

T["slide4.xml"] = {
 'АРХИТЕКТУРА ПРОЦЕССОВ': 'PROCESS ARCHITECTURE',
 'Три бизнес-процесса — в одной системе': 'Three business processes — in one system',
 'Каждый процесс описан в нотации BPMN. Следующий слайд показывает, какие шаги берёт на себя система.':
   'Each process is described in BPMN notation. The diagram shows which steps the system takes on.',
 '01': '01',
 'Промо': 'Promo',
 'Создание промо-активности': 'Creating a promo activity',
 'Запрос от Trade-маркетинга или On-trade → присвоение кода → запись в Panorama → уведомление партнёров.':
   'Request from Trade marketing or On-trade → code assignment → record in Panorama → partner notification.',
 '02': '02',
 'Изменение промо-активности': 'Modifying a promo activity',
 'Смена квоты, периода, добавление продукта или точки → правка в Panorama → синхронизация и уведомление.':
   'Changing the quota or period, adding a product or outlet → edit in Panorama → synchronization and notification.',
 '03': '03',
 'Закрытие': 'Closing',
 'Закрытие месяца и расчёт': 'Month-end closing and calculation',
 'Партнёр присылает «Запрос» → расчёт ∑(литры × цена периода) → AI-проверка → передача в Контроллинг.':
   'Partner sends a «Request» → calculation ∑(liters × period price) → AI check → handover to Controlling.',
 'Создание': 'Creation',
 '→': '→',
 'Активно &amp; изменения': 'Active &amp; changes',
 'Закрытие месяца': 'Month-end closing',
 'Efes Kazakhstan': 'Efes Kazakhstan',
 'Efes Nexus': 'Efes Nexus',
}

T["slide5.xml"] = {
 'МАСШТАБ': 'SCALE',
 'Объём растёт быстрее, чем успевают руки': 'Volume grows faster than hands can keep up',
 'За три года число промо-активностей выросло почти в 5 раз, а годовой бюджет промо — в 2,4 раза. При почти неизменном числе партнёров объём растёт быстрее, чем справляются ручные процессы.':
   'Over three years the number of promo activities grew almost 5×, while the annual promo budget grew 2.4×. With an almost unchanged number of partners, volume is growing faster than manual processes can handle.',
 'Промо-активности по годам': 'Promo activities by year',
 'Количество заведённых акций за год': 'Number of promos created per year',
 '2024 → 2025': '2024 → 2025',
 '2 084': '2,084',
 '2023': '2023',
 '2 187': '2,187',
 '2024': '2024',
 '10 112': '10,112',
 '2025': '2025',
 '×4,6': '×4.6',
 '+362 %': '+362%',
 'ПАРТНЁРЫ': 'PARTNERS',
 'за 6 мес': 'in 6 mo',
 '→': '→',
 '31': '31',
 '33': '33',
 '+6,5 %': '+6.5%',
 'ПРОДУКТЫ · SKU': 'PRODUCTS · SKU',
 '78': '78',
 '105': '105',
 '+35 %': '+35%',
 'СУММА ПРОМО / ГОД · МЛРД ₸': 'PROMO TOTAL / YEAR · BLN ₸',
 '2023 → 2025': '2023 → 2025',
 '1,00': '1.00',
 '2,41': '2.41',
 '+140 %': '+140%',
 'Efes Kazakhstan': 'Efes Kazakhstan',
 'Efes Nexus': 'Efes Nexus',
}

T["slide6.xml"] = {
 'ПЕРСПЕКТИВА': 'OUTLOOK',
 'Чем больше объём — тем выше автоматизация': 'The larger the volume — the higher the automation',
 'Автоматизированные шаги — уведомления, расчёты, Excel, AI-проверка, синхронизация — выполняются почти без затрат времени на каждую новую промо. Ручные шаги фиксированы. Поэтому с ростом объёма доля автоматизации растёт к структурному потолку 8 из 13 шагов (≈ 62 %).':
   'Automated steps — notifications, calculations, Excel, AI checks, synchronization — take almost no time per new promo. Manual steps are fixed. So as volume grows, the share of automation rises toward the structural ceiling of 8 of 13 steps (≈ 62%).',
 'Год 1': 'Year 1',
 '28 %': '28%',
 'автоматизации': 'automation',
 '158 ч/мес  ': '158 h/mo  ',
 'Запуск платформы. Единый источник данных вместо разрозненных Excel-файлов.':
   'Platform launch. A single source of data instead of scattered Excel files.',
 'Накопленный эффект  ч/год': 'Cumulative effect  h/year',
 '≈ 7 000 ч': '≈ 7,000 h',
 'Команда сфокусирована на аналитике и создании новых идей вместо ручной рутины':
   'The team is focused on analytics and generating new ideas instead of manual routine',
 'Технологичный бренд - сильный PR и возможность масштабирования системы на все департаменты ':
   'A tech-forward brand — strong PR and the ability to scale the system across all departments ',
 'Отрыв от конкурентов по скорости, точности и масштабу процессов':
   'A lead over competitors in speed, accuracy and process scale',
 'Efes Kazakhstan': 'Efes Kazakhstan',
 'Efes Nexus': 'Efes Nexus',
}

T["slide7.xml"] = {
 'МАСШТАБИРОВАНИЕ': 'SCALING',
 'Следующий шаг — ': 'Next step — ',
 'масштабирование': 'scaling',
 ' ': ' ',
 'на другие отделы': 'to other departments',
 'Департамент сбыта уже перешёл на новые правила компенсаций: ретро-скидки (RI/BO), разовые начисления (SS/OC), программы лояльности Efes Bonus Club (LP), Rozliv+ (EP) и Kega Boom (KB). Nexus ':
   'The sales department has already moved to new compensation rules: retro-discounts (RI/BO), one-time accruals (SS/OC), Efes Bonus Club loyalty programs (LP), Rozliv+ (EP) and Kega Boom (KB). Nexus ',
 'возьмёт': 'will take',
 'проверку': 'the verification of',
 ' накладных 1С': ' 1C invoices',
 ' и ': ' and ',
 'коммуникацию': 'communication',
 'с партнерами ': 'with partners ',
 'на': 'upon',
 'себя': 'itself',
 '.': '.',
 'Без Nexus': 'Without Nexus',
 '193': '193',
 ' ч': ' h',
 'ручной сверки в месяц на специалиста': 'of manual reconciliation per month per specialist',
 '≈ ': '≈ ',
 '28': '28',
 ' % рабочего времени': '% of working time',
 '33 партнёра × 10 мин сверки 1С и Panorama + ~': '33 partners × 10 min of reconciliation in 1C and Panorama + ~',
 '1': '1',
 ' ч переписки об ошибках ': ' h of correspondence about ',
 'формата': 'format errors',
 'Контроль накладных по 10 000+ торговых точек': 'Invoice control across 10,000+ outlets',
 '100 % накладных': '100% of invoices',
 'вместо выборки до 10 ТТ — выше integrity, у партнёров нет возможности манипулировать суммами':
   'instead of sampling up to 10 outlets — higher integrity, and partners cannot manipulate the amounts',
 'Авто-проверка выгрузок 1С от 33 партнёров: коды ТТ, суммы, разбивка по типам скидок RI/BO · SS/OC · LP · EP · KB':
   'Auto-check of 1C exports from 33 partners: outlet codes, amounts, breakdown by discount type RI/BO · SS/OC · LP · EP · KB',
 'Некорректный формат (код и название ТТ в одной ячейке) система находит сама и запрашивает исправление у партнёра':
   'The system detects an incorrect format (outlet code and name in one cell) on its own and requests a fix from the partner',
 'Экономия ≈ ': 'Savings ≈ ',
 '2316': '2,316',
 ' ч/год на команду из 5 специалистов — без обучения ручной сверке и переписки с партнёрами':
   ' h/year for a team of 5 specialists — no training in manual reconciliation or correspondence with partners',
 'Efes Kazakhstan': 'Efes Kazakhstan',
 'Efes Nexus': 'Efes Nexus',
}

pat = re.compile(r"<a:t>(.*?)</a:t>", re.DOTALL)
total_miss = 0
for fname in sorted(T.keys()):
    path = os.path.join(SLIDES, fname)
    xml = open(path, encoding="utf-8").read()
    table = T[fname]
    misses = []
    repl_count = [0]
    def rep(m):
        inner = m.group(1)
        if inner in table:
            repl_count[0] += 1
            return "<a:t>" + table[inner] + "</a:t>"
        misses.append(inner)
        return m.group(0)
    new = pat.sub(rep, xml)
    open(path, "w", encoding="utf-8").write(new)
    print("%s: replaced %d nodes, misses=%d" % (fname, repl_count[0], len(misses)))
    for mm in misses:
        print("    MISS: " + ascii(mm))
    total_miss += len(misses)

print("TOTAL MISSES:", total_miss)
