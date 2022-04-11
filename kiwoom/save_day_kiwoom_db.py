def save_day_kiwoom_db(self, stock_code=None):
    stock_name = self.dynamicCall("GetMasterCodeName(QString)", stock_code)
    # 숫자나 특수문자 (")가 들어가면 OptionError가 발생
    # 테이블에 들어갈 때 따옴표는 빠진 상태로 저장됨
    table_name = "\"" + stock_name + "\""

    # DB에 있는 테이블의 목록을 뽑아내는 것
    query = "SELECT name FROM sqlite_master WHERE type='table'"
    self.cursor.execute(query)

    for row in self.cursor.fetchall():
        if row[0] == stock_name:
            return

    query = "CREATE TABLE IF NOT EXISTS {} \
        (current_price integer, volume integer, trade_price integer, \
            date integer PRIMARY KEY, start_price integer, high_price integer, low_price integer)".format(table_name)
    self.cursor.execute(query)

    for item in self.calculator_list:
        calculator_tuple = tuple(item[1:-1])
        query = "INSERT INTO {} (current_price, volume, trade_price, date, \
            start_price, high_price, low_price) VALUES(?, ?, ?, ?, ?, ?, ?)".format(table_name)
        self.cursor.execute(query, calculator_tuple)