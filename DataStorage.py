class DataStorage:
    """
    Класс для хранения всех данных,
    используемых в программе
    """

    def __init__(self):
        self.table = '\\\\192.168.2.10\\xchg\\Производство\\Работа с МСЛ\\Регистрация МСЛ.xlsx'

        #self.table = "C:\\Users\\EDN\\PycharmProjects\\AutoInsert\\Templates\\table.xlsx"

        self.first: int = 0

        self.amount: int = 0

        self.per_one_msl: int = 0

        self.device_name: str = ''

        self.master_name: str = ''

        self.contract: str = ''

        self.msl_number: int = 1

        self.multiplier: int = 1

        self.devices: list = ['ПИ СПЛР.469559.026-02',
                              'ПИ СПЛР.469555.007',
                              'ПИ СПЛР.469555.006',
                              'УСФВИ СПЛР.467669.001',
                              'Жгут СПЛР.685666.003',
                              'Датчик глаза СПЛР.469639.011',
                              'ВИП СПЛР.563344.001',
                              'Жгут СПЛР.685666.003-01',
                              'ПИ СПЛР.469559.026-04',
                              'МСЛ на платы для УСФВИ',
                              'МСЛ Пл. кнопки',
                              'МСЛ Пл. индик.',
                              'МСЛ Пл. упр.',
                              'МСЛ Жгут',
                              'МСЛ ББ',
                              'МСЛ ИП',
                              'МСЛ Платы ИП',
                              'МСЛ на корпус ВИП',
                              'Табличка СПЛР.741121.036',
                              'МСЛ Корпус ИП',
                              'ПИ СПЛР.469555.004']

        self.masters: list = ['Соловьев Евгений',
                              'Коваленко Владимир']

    def refresh_data(self):
        self.first = 0
        self.amount = 0
        self.per_one_msl = 0
        self.device_name = ''
        self.master_name = ''
        self.contract = ''
        self.msl_number = 1
        self.multiplier = 1
