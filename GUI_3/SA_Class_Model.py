class Container():
    def __init__(self, con_id, con_type, con_pos):
        self.con_id = con_id
        self.con_type = con_type
        self.con_pos = con_pos  ###[船/堆场,贝位,层,列]  ###[船/堆场,箱区，贝位，层，列]    船舶默认为1箱区


class Ship_Pos():
    def __init__(self, ship_con_status, ship_con_pre_type, ship_pos, ship_con, load_flag):
        self.ship_con_status = ship_con_status
        self.ship_con_pre_type = ship_con_pre_type
        self.ship_pos = ship_pos ###[箱区，贝位，层，列]
        self.ship_con = ship_con
        self.load_flag = load_flag


class Yard_Pos():
    def __init__(self, yard_con_status, yard_pos, yard_con):
        self.yard_con_status = yard_con_status
        self.yard_pos = yard_pos ###[箱区，贝位，层，列]
        self.yard_con = yard_con
