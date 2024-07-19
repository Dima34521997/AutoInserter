class DataStorage:
    """
    Класс для хранения всех данных,
    используемых в программе
    """
    def __init__(self):

        self.table = '\\\\192.168.2.10\\xchg\\Производство\\Работа с МСЛ\\Регистрация МСЛ.xlsx'

        self.first: int = 0

        self.amount: int = 0

        self.num_list: list = []

        self.per_one_msl: int = 0

        self.device_name: str = ''

        self.master_name: str = ''

        self.contract: str = ''

        self.msl_number: int = 1
