def granvile_theory(self):
    query = "SELECT name FROM sqlite_master WHERE type='table'"
    self.cursor.execute(query)

    for row in self.cursor.fetchall():
        table_name = "\"" + row[0] + "\""
        query = "SELECT * from {}".format(table_name)
        self.cursor.execute(query)
        calculator_list = []
        for item in self.cursor.fetchall():
            itemList = list(item)
            itemList.insert(0, '')
            itemList.insert(len(itemList), '')
            calculator_list.append(itemList)
        # DB에는 날짜를 기준으로 내림차순이 아니라 오름차순 정렬되어있는 상태이기때문에 이를 뒤집어 주어야 함
        calculator_list.reverse()
        self.calculator(calculator_list, self.kosdaq_dict[row[0]])