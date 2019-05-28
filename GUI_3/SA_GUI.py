import tkinter as tk
import win32ui
from tkinter import scrolledtext

import random
import math
import time
import copy
import matplotlib.pyplot as plt
import numpy as np
import xlrd
import xlwt
import xlutils.copy
from decimal import *
import threading

import SA_Thread_Model
import SA_Write_Excel_Model
from SA_Class_Model import *

###########################################################################################################
########模拟退火算法参数设计#######################################################################
T0 = 1e2
T = 10
T_min = 0.5
alpha_rate = 0.4
loop_in = 90
xls_yard_pos_read = ""
xls_ship_pos_read = ""

yard_plan_list_best = []
ship_load_strategy_list_best = []

strategy_best = []

show_strategy_index = []
show_score = []
show_score_history = []

Yard_Box,Yard_Bay,Yard_Tier,Yard_Row = 6,90,5,10
Ship_Bay, Ship_Tier, Ship_Row, Ship_Down_Tier, Ship_Up_Tier = 51,14,20,6,8
strategy_best_history = []
ship_load_strategy_list_best_history = []
yard_plan_list_best_history = []


Yard_Pos_List = []
Ship_Pos_List = []
Yard_Common_Container_Dictionary = dict()
Ship_Load_Container_List = []
Restow_Max = 10

SA_Thread = ""

###############################################################################################
def Create_Yard_Pos_List():
    global Yard_Box,Yard_Bay, Yard_Tier, Yard_Row, Yard_Pos_List,xls_yard_pos_read
    ##########配置excel表格
    #xls_yard_pos_read = "yard.xls"
    book_read = xlrd.open_workbook(xls_yard_pos_read)


    read_excel_yard_tier = 0

    read_excel_yard_con_type_row = 2
    read_excel_yard_id_row = 10
    read_excel_yard_status_row = 1

    for b in range(Yard_Box):
        sh_read = book_read.sheet_by_index(b)
        Yard_Box_List = []
        for i in range(Yard_Bay):
            Yard_Bay_List = []
            for j in range(Yard_Tier):
                Yard_Tier_List = []
                for k in range(Yard_Row):
                    read_excel_yard_tier = 50*i+10*j+k+1
                    ###先创建堆场集装箱
                    #print(read_excel_yard_tier,"***",b)
                    con_id = sh_read.cell_value(read_excel_yard_tier, read_excel_yard_id_row)
                    con_type = sh_read.cell_value(read_excel_yard_tier, read_excel_yard_con_type_row)

                    con_pos = ["yard",b+1 ,i + 1, j + 1, k + 1]
                    container = Container(con_id, con_type, con_pos)

                    yard_con_status = int(sh_read.cell_value(read_excel_yard_tier, read_excel_yard_status_row))
                    yard_pos_pos = [b+1,i + 1, j + 1, k + 1]
                    yard_con = container
                    yard_pos = Yard_Pos(yard_con_status, yard_pos_pos, yard_con)

                    Yard_Tier_List.append(yard_pos)

                Yard_Bay_List.append(Yard_Tier_List)
            Yard_Box_List.append(Yard_Bay_List)
        Yard_Pos_List.append(Yard_Box_List)

    print("Create_Yard_Pos_List OK!")


def Creat_Ship_Pos_List():
    global Ship_Bay, Ship_Tier, Ship_Row, Ship_Pos_List,xls_ship_pos_read
    ##########配置excel表格
    #xls_ship_pos_read = "ship.xls"
    book_read = xlrd.open_workbook(xls_ship_pos_read)
    sh_read = book_read.sheet_by_index(0)

    read_excel_ship_tier = 0

    read_excel_ship_con_pre_type_row = 3
    read_excel_ship_load_flag_row = 2
    read_excel_ship_status_row = 1
    read_excel_ship_con_id_row = 10

    for i in range(Ship_Bay):
        Ship_Bay_List = []
        for j in range(Ship_Tier):
            Ship_Tier_List = []
            for k in range(Ship_Row):
                read_excel_ship_tier += 1
                
                # 先创建集装箱
                con_id = sh_read.cell_value(read_excel_ship_tier, read_excel_ship_con_id_row)
                con_type = sh_read.cell_value(read_excel_ship_tier, read_excel_ship_con_pre_type_row)
                con_pos = ["ship",1 ,i + 1, j + 1, k + 1]
                container = Container(con_id, con_type, con_pos)

                ship_con_status = sh_read.cell_value(read_excel_ship_tier, read_excel_ship_status_row)
                ship_con_pre_type = sh_read.cell_value(read_excel_ship_tier, read_excel_ship_con_pre_type_row)
                ship_pos_pos = [1,i + 1, j + 1, k + 1]
                #print(read_excel_ship_tier,ship_pos_pos)
                ship_con = container
                load_flag = sh_read.cell_value(read_excel_ship_tier, read_excel_ship_load_flag_row)
                ship_pos = Ship_Pos(ship_con_status, ship_con_pre_type, ship_pos_pos, ship_con, load_flag)

                Ship_Tier_List.append(ship_pos)

            Ship_Bay_List.append(Ship_Tier_List)
        Ship_Pos_List.append(Ship_Bay_List)
    print("Creat_Ship_Pos_List OK!")


def Creat_Add_Yard_Restow_List():
    global Yard_Box, Yard_Bay, Yard_Tier, Yard_Row, Yard_Pos_List, Restow_Max
    Yard_Tier += Restow_Max

    for box in range(Yard_Box):
        for bay in range(Yard_Bay):
            bay_list = Yard_Pos_List[box][bay]
            for tier in range(Restow_Max):
                Yard_Tier_List = []
                for row in range(Yard_Row):
                    con_id = ""
                    con_type = ""
                    con_pos = "AAA"
                    container = Container(con_id, con_type, con_pos)
                    yard_con_status = 0
                    yard_pos_pos = [box + 1, bay + 1, tier + 5 + 1, row + 1]
                    yard_con = container
                    yard_pos = Yard_Pos(yard_con_status, yard_pos_pos, yard_con)

                    Yard_Tier_List.append(yard_pos)
                bay_list.append(Yard_Tier_List)

    print("Creat_Add_Yard_Restow_List OK!")


###这样就把空间矩阵呢搭建好了

###接下来是将堆场中相同复合属性且待装的集装箱放在一起，
def Create_Yard_Common_Container_Dictionary():
    global Yard_Box, Yard_Bay, Yard_Tier, Yard_Row, Yard_Pos_List, Restow_Max, Yard_Common_Container_Dictionary,xls_yard_pos_read

    ###配置excel
    #xls_pos_read = "yard.xls"
    book_read = xlrd.open_workbook(xls_yard_pos_read)

    read_excel_yard_tier = 0

    read_excel_yard_con_type_row = 2
    read_excel_yard_bay_change_row = 6
    read_excel_yard_tier_change_row = 7
    read_excel_yard_row_change_row = 8
    read_excel_yard_box_row = 9

    None_Restow_Tier = Yard_Tier - Restow_Max

    for b in range(Yard_Box):
        sh_read = book_read.sheet_by_index(b)
        for i in range(Yard_Bay):
            for j in range(None_Restow_Tier):
                for k in range(Yard_Row):
                    read_excel_yard_tier = 50*i+10*j+k+1
                    ###获得 复合属性2非E 获得换算贝位6层7列8
                    #print(read_excel_yard_tier)
                    con_type = sh_read.cell_value(read_excel_yard_tier, read_excel_yard_con_type_row)
                    con_bay = int(sh_read.cell_value(read_excel_yard_tier, read_excel_yard_bay_change_row))
                    con_tier = int(sh_read.cell_value(read_excel_yard_tier, read_excel_yard_tier_change_row))
                    con_row = int(sh_read.cell_value(read_excel_yard_tier, read_excel_yard_row_change_row))
                    con_box = int(sh_read.cell_value(read_excel_yard_tier,read_excel_yard_box_row))

                    common_con_pos = [con_box, con_bay, con_tier, con_row]
                    if con_type != "ELSE"  and con_type != "" or con_type != "E"  and con_type != "" :
                        if con_type in Yard_Common_Container_Dictionary:
                            Yard_Common_Container_Dictionary[con_type].append(common_con_pos)
                        else:
                            #print(con_type)
                            Yard_Common_Container_Dictionary.setdefault(con_type, [])
                            Yard_Common_Container_Dictionary[con_type].append(common_con_pos)
    print("Create_Yard_Common_Container_Dictionary OK!")
    #print(Yard_Common_Container_Dictionary)


# 得出船舶的装船序列
def Get_Ship_Load_Container_List():
    global Yard_Box,Ship_Bay, Ship_Tier, Ship_Row, Yard_Pos_List, Restow_Max, Yard_Common_Container_Dictionary, Ship_Load_Container_List,xls_ship_pos_read
    ###配置excel
    #xls_pos_read = "ship.xls"
    book_read = xlrd.open_workbook(xls_ship_pos_read)
    sh_read = book_read.sheet_by_index(0)
    ###洋山港装箱信号2 复合属性3 贝位7 层8 列9
    read_excel_ship_tier = 0

    read_excel_ship_load_flag_row = 2
    read_excel_ship_con_pre_type_row = 3
    read_excel_ship_bay_change_row = 7
    read_excel_ship_tier_change_row = 8
    read_excel_ship_row_change_row = 9
    for i in range(Ship_Bay):
        for j in range(Ship_Tier):
            for k in range(Ship_Row):
                read_excel_ship_tier += 1
                ship_load_flag = sh_read.cell_value(read_excel_ship_tier, read_excel_ship_load_flag_row)
                # ship_con_pre_type = sh_read.cell_value(read_excel_ship_tier,read_excel_ship_con_pre_type_row)
                ship_bay = int(sh_read.cell_value(read_excel_ship_tier, read_excel_ship_bay_change_row))
                ship_tier = int(sh_read.cell_value(read_excel_ship_tier, read_excel_ship_tier_change_row))
                ship_row = int(sh_read.cell_value(read_excel_ship_tier, read_excel_ship_row_change_row))

                ship_pos = [1,ship_bay, ship_tier, ship_row]
                if ship_load_flag == "有":
                    Ship_Load_Container_List.append(ship_pos)
                else:
                    pass
    print("Get_Ship_Load_Container_List OK!")


def Create_Strategy():
    global Yard_Pos_List, Ship_Pos_List, Yard_Common_Container_Dictionary, Ship_Load_Container_List
    Strategy_List = []

    ###依照顺序读出Ship_Load_Container_List中每个复合属性 然后在Yard_Common_Container_Dictionary中查找
    Yard_Common_Container_Dictionary_Test = copy.deepcopy(Yard_Common_Container_Dictionary)
    Ship_Load_Container_List_Test = copy.deepcopy(Ship_Load_Container_List)

    #print(Ship_Load_Container_List_Test)

    for Load_Container_Pos in Ship_Load_Container_List_Test:
        Strategy = []
        ship_box,ship_bay, ship_tier, ship_row = Load_Container_Pos[0], Load_Container_Pos[1], Load_Container_Pos[2],Load_Container_Pos[3]
        container_type = Ship_Pos_List[ship_bay - 1][ship_tier - 1][ship_row - 1].ship_con_pre_type

        yard_common_container_list = Yard_Common_Container_Dictionary_Test[container_type]
        Strategy.append(yard_common_container_list.pop(random.randint(0, len(yard_common_container_list) - 1)))
        Strategy.append(Load_Container_Pos)
        Strategy.append(0)  # 未完成的意思

        Strategy_List.append(Strategy)

    #print("Create_Strategy  OK!")
    #print(Strategy_List)

    return Strategy_List


def Calculate_Score(Strategy_List):
    global Yard_Pos_List, Ship_Pos_List, Yard_Common_Container_Dictionary, Ship_Load_Container_List, Yard_Box, Yard_Bay, Yard_Tier, Yard_Row
    Yard_Pos_List_Test = copy.deepcopy(Yard_Pos_List)
    Ship_Pos_List_Test = copy.deepcopy(Ship_Pos_List)

    score = 0

    yard_plan_list_new = []
    yard_plan_index = 0
    yard_plan_str = ""

    for Strategy in Strategy_List:
        #####
        #print(Strategy)

        yard_con_box = Strategy[0][0]
        yard_con_bay = Strategy[0][1]
        yard_con_tier = Strategy[0][2]
        yard_con_row = Strategy[0][3]

        ship_con_box = Strategy[1][0]
        ship_con_bay = Strategy[1][1]
        ship_con_tier = Strategy[1][2]
        ship_con_row = Strategy[1][3]

        yard_con_object = Yard_Pos_List_Test[yard_con_box-1][yard_con_bay - 1][yard_con_tier - 1][yard_con_row - 1].yard_con
        yard_con_pos_object = Yard_Pos_List_Test[yard_con_box-1][yard_con_bay - 1][yard_con_tier - 1][yard_con_row - 1]

        ################################################################################################

        ship_con_pos_object = Ship_Pos_List_Test[ship_con_bay - 1][ship_con_tier - 1][ship_con_row - 1]

        if yard_con_tier == Yard_Tier:

            yard_plan_index += 1
            yard_plan_str = "将堆场"+str(yard_con_box)+"箱区，第"+ str(yard_con_bay) + "贝位" + str(yard_con_tier) + "层" + str(
                yard_con_row) + "列的集装箱，移至集卡"

            yard_plan_list_new.append([yard_plan_index, yard_plan_str,[yard_con_box,yard_con_bay,yard_con_tier,yard_con_row],"集卡",Strategy[1]])

            yard_con_pos_object.yard_con_status = 0
            yard_con_pos_object.yard_con = ""

            yard_con_object.con_pos = ["ship", ship_con_box,ship_con_bay, ship_con_tier, ship_con_row]

            ship_con_pos_object.ship_con_status = 1
            ship_con_pos_object.ship_con = yard_con_object

            Strategy[2] = 1
            ###这就是在顶层，直接将它移到集卡
            ###要修正的属性   堆场 集装箱 船舶 策略
        else:
            if Yard_Pos_List_Test[yard_con_box-1][yard_con_bay - 1][yard_con_tier][yard_con_row - 1].yard_con_status == 1:
                # 需要倒箱 判断倒箱量的多少
                max_tier = 0
                for i in range(Yard_Tier):
                    if Yard_Pos_List_Test[yard_con_box-1][yard_con_bay - 1][i][yard_con_row - 1].yard_con_status == 0:
                        max_tier = i
                        break
                    if i == (Yard_Tier - 1):
                        max_tier = Yard_Tier
                        break
                    ##得到了最大层数 max_tier       它是层数，不是索引！
                # print("倒箱量="+str(max_tier-yard_con_tier)+"*"+str(Strategy))
                # for i in range(15):
                #     print_line = []
                #     for j in range(10):
                #         print_line.append(Yard_Pos_List_Test[yard_con_bay - 1][14-i][j].yard_con_status)
                #     print(print_line)
                score = score + max_tier - yard_con_tier
                restow_list = []
                for i in range(yard_con_tier, max_tier):
                    restow_list.append([yard_con_box,yard_con_bay, max_tier - i + yard_con_tier, yard_con_row])
                # print(restow_list)
                for restow_con in restow_list:
                    restow_box = restow_con[0]
                    restow_bay = restow_con[1]
                    restow_tier = restow_con[2]
                    restow_row = restow_con[3]
                    # 找一个落箱箱位随机找一个落箱箱位，但是有一些条件限制   最高层不能超过7层
                    ## 首先是除去本列外，就是要获得表层元素
                    surface_con_list = []

                    for i in range(Yard_Row):
                        for j in range(Yard_Tier):
                            if Yard_Pos_List_Test[yard_con_box-1][yard_con_bay - 1][j][i].yard_con_status == 0 and (
                                    i + 1) != yard_con_row:
                                surface_con_list.append([yard_con_box,yard_con_bay, j + 1, i + 1])
                                break
                    # print(surface_con_list)

                    # 落箱箱位
                    restow_con_obj_add = surface_con_list[random.randint(0, len(surface_con_list) - 1)]
                    restow_con_obj = Yard_Pos_List_Test[restow_con_obj_add[0] - 1][restow_con_obj_add[1] - 1][
                        restow_con_obj_add[2] - 1][restow_con_obj_add[3]-1]

                    # 倒箱箱位
                    restow_add = Yard_Pos_List_Test[restow_box-1][restow_bay - 1][restow_tier - 1][restow_row - 1]

                    # 倒箱箱位的集装箱
                    restow_con = restow_add.yard_con

                    ###开始倒箱
                    ###就是要改变一些性质
                    ###集装箱的性质 pos  落箱箱位的性质status con     倒箱箱位的性质status  con
                    yard_plan_index += 1
                    box, bay, tier, row = restow_add.yard_pos[0], restow_add.yard_pos[1], restow_add.yard_pos[2], restow_add.yard_pos[3]
                    box1,bay1, tier1, row1 = restow_con_obj_add[0],restow_con_obj_add[1],restow_con_obj_add[2],restow_con_obj_add[3]
                    yard_plan_str = "将堆场"+str(box)+"箱区，第"+ str(bay) + "贝位" + str(tier) + "层" + str(row) + "列的集装箱，移至堆场" + str(box1)+"箱区"+str(
                        bay1) + "贝位" + str(tier1) + "层" + str(row1) + "列"
                    yard_plan_list_new.append([yard_plan_index, yard_plan_str,[box,bay*2-1,tier,row],[box1,bay1*2-1,tier1,row1],"***"])
                    

                    restow_con.con_pos = ["yard", restow_con_obj_add[0], restow_con_obj_add[1], restow_con_obj_add[2],restow_con_obj_add[3]]
                    

                    restow_con_obj.yard_con_status = 1
                    restow_con_obj.yard_con = restow_con

                    restow_add.yard_con_status = 0
                    restow_add.yard_con = ""

                    ###在这里有一个东西忽略掉了，由于倒箱操作改变了原来集装箱策略的的位置
                    ###因此在这里要对反走的集装箱的策略做改变strategy_list
                    for strategy_change in Strategy_List:
                        if strategy_change[0] == [yard_con_box,yard_con_bay, restow_tier, restow_row] and strategy_change[2] == 0:
                            #print(strategy_change[0],restow_con_obj.yard_pos)
                            strategy_change[0] = [restow_con_obj.yard_pos[0], restow_con_obj.yard_pos[1],
                                                  restow_con_obj.yard_pos[2],restow_con_obj.yard_pos[3]]
                            break

                yard_plan_index += 1
                yard_plan_str = "将堆场" +str(yard_con_box)+"箱区，第"+ str(yard_con_bay) + "贝位" + str(yard_con_tier) + "层" + str(
                    yard_con_row) + "列的集装箱，移至集卡"
                yard_plan_list_new.append([yard_plan_index, yard_plan_str,[yard_con_box,yard_con_bay,yard_con_tier,yard_con_row],"集卡",Strategy[1]])

                yard_con_pos_object.yard_con_status = 0
                yard_con_pos_object.yard_con = ""

                yard_con_object.con_pos = ["ship", ship_con_box,ship_con_bay, ship_con_tier, ship_con_row]

                ship_con_pos_object.ship_con_status = 1
                ship_con_pos_object.ship_con = yard_con_object

                Strategy[2] = 1
                ###这就是在顶层，直接将它移到集卡
                ###要修正的属性   堆场 集装箱 船舶 策略

            else:
                yard_plan_index += 1
                yard_plan_str = "将堆场"+str(yard_con_box)+"箱区，第" + str(yard_con_bay) + "贝位" + str(yard_con_tier) + "层" + str(
                    yard_con_row) + "列的集装箱，移至集卡"
                yard_plan_list_new.append([yard_plan_index, yard_plan_str,[yard_con_box,yard_con_bay,yard_con_tier,yard_con_row],"集卡",Strategy[1]])
                yard_con_pos_object.yard_con_status = 0
                yard_con_pos_object.yard_con = ""

                #print(Strategy[0])
                yard_con_object.con_pos = ["ship", ship_con_box,ship_con_bay, ship_con_tier, ship_con_row]

                ship_con_pos_object.ship_con_status = 1
                ship_con_pos_object.ship_con = yard_con_object

                Strategy[2] = 1
                ###这就是在顶层，直接将它移到集卡
                ###要修正的属性   堆场 集装箱 船舶 策略
    # print("Calculate_Score OK!")

    return score, Ship_Pos_List_Test, yard_plan_list_new

###############################################################################################
###############################################################################################




###############################################################################################
###############################################################################################
def accept_probability(score_best, score_new, T):
    de = score_new - score_best

    T = Decimal(T)
    de = Decimal(de)
    probability_stand = round(math.exp(-de / T), 4)
    probability_rand = random.randint(0, 10000) / 10000
    if probability_rand < probability_stand:
        # 接受
        return 1
    else:
        # 不接受
        return 0


def Run_SA(self):
    
    global strategy_best,ship_load_strategy_list_best,yard_plan_list_best,show_strategy_index,show_score,show_score_history
    global strategy_best_history,ship_load_strategy_list_best_history,yard_plan_list_best_history
    global SA_Thread
    global Ship_Bay,Ship_Tier,Ship_Row,Ship_Down_Tier,Ship_Up_Tier

     # 计次
    index = 1

    start_time = time.time()

    Creat_Ship_Pos_List()
    Create_Yard_Pos_List()
    Creat_Add_Yard_Restow_List()
    Create_Yard_Common_Container_Dictionary()
    Get_Ship_Load_Container_List()
    # Strategy_List = Create_Strategy()
    # Calculate_Score(Strategy_List)

    # 生成初始解（best）
    strategy_best = Create_Strategy()
    # 计算得分（best）
    score_best, ship_load_strategy_list_best, yard_plan_list_best = Calculate_Score(strategy_best)

    #历史最优解
    score_best_history = score_best
    strategy_best_history = strategy_best
    ship_load_strategy_list_best_history = ship_load_strategy_list_best
    yard_plan_list_best_history = yard_plan_list_best

    show_strategy_index.append(index)
    show_score.append(score_best)
    show_score_history.append(score_best_history)
    
    run_state = "已运行:" + str(index) + "次,耗时:" + str(round(time.time() - start_time, 2)) + "秒," + "当前倒箱次数:" + str(score_best) +\
                        ",历史最优解倒箱次数,"+str(score_best_history)+"次\n"
            
    print(run_state)
    self.lb.insert(tk.END,run_state)
    

    # 绘图准备
    
   
    global T,T_min

    while T > T_min:
        for i in range(loop_in):
            # 得到新的解
            strategy_new = Create_Strategy()
            strategy_new_test = copy.deepcopy(strategy_new)
            # 计算得分
            score_new, ship_load_strategy_list_new, yard_plan_list_new = Calculate_Score(strategy_new_test)
            #和历史最优解对比：
            if score_best_history>score_new:
                
                strategy_best_history = strategy_new
                ship_load_strategy_list_best_history = ship_load_strategy_list_new
                yard_plan_list_best_history = yard_plan_list_new
                score_best_history = score_new
                
            # 和最优解做对比，考虑概率接收
            flag = accept_probability(score_best, score_new, T)
            if flag == 1:
                strategy_best = strategy_new
                score_best = score_new
                yard_plan_list_best = yard_plan_list_new
                ship_load_strategy_list = ship_load_strategy_list_new
            else:
                pass
            index += 1
            
            show_strategy_index.append(index)
            show_score.append(score_best)
            show_score_history.append(score_best_history)

            run_state = "已运行:" + str(index) + "次,耗时:" + str(round(time.time() - start_time, 2)) + "秒," + "当前倒箱次数:" + str(score_best) +\
                        ",历史最优解倒箱次数,"+str(score_best_history)+"次\n"
            
            print(run_state)
            self.lb.insert(tk.END,run_state)
           
        # 内循环完毕，更新温度
        T = T * alpha_rate
        # T = T0*math.exp(-0.01*index)

    # 打印集装箱装船顺序
    SA_Write_Excel_Model.write_container_strategy(strategy_best_history,Yard_Pos_List)
    ####打印船舶配载策略 Ship_Bay, Ship_Tier, Ship_Row ship_load_strategy_list
    SA_Write_Excel_Model.write_ship_load_container(ship_load_strategy_list_best,Ship_Bay,Ship_Tier,Ship_Row,Ship_Down_Tier,Ship_Up_Tier)
    ####场桥调度
    SA_Write_Excel_Model.write_yard_strategy(yard_plan_list_best_history)

    print("******************************OK!******************************")
    print("集装箱发箱顺序，船舶配载，场桥调度策略已生成完毕，请转至桌面查看^0^")
    self.lb.insert(tk.END,"******************************OK!******************************\n")
    self.lb.insert(tk.END,"集装箱发箱顺序，船舶配载，场桥调度策略已生成完毕，请转至桌面查看^0^\n")

    min_score = min(show_score)-20
    max_score = max(show_score)+20

    my_y_ticks = np.arange(min_score, max_score, 5)
    plt.title("模拟退火算法演进", fontproperties='SimHei', fontsize=15)
    plt.xlabel(u'迭代次数', fontproperties='SimHei', fontsize=12)
    plt.ylabel(u'倒箱量', fontproperties='SimHei', fontsize=12)

    plt.plot(show_strategy_index, show_score,linestyle="--",label="Score")
    plt.plot(show_strategy_index, show_score_history,linestyle="-.",label="Score_History")
    plt.legend() 
    plt.show()
    print("准备杀死线程")
    SA_Thread_Model.stop_thread(SA_Thread)


def Interrupt_Output(self):
     # 打印集装箱装船顺序
    global strategy_best,ship_load_strategy_list_best,yard_plan_list_best,show_strategy_index,show_score,show_score_history,SA_Thread
    global strategy_best_history,ship_load_strategy_list_best_history,yard_plan_list_best_history
    global Ship_Bay,Ship_Tier,Ship_Row,Ship_Down_Tier,Ship_Up_Tier

    if SA_Thread == "":
        self.lb.insert(tk.END,"无算法运行，无法中断\n")
        SA_Thread = ""
        print("无算法运行，无法中断\n")
        return 
    
    SA_Thread_Model.stop_thread(SA_Thread)
    SA_Thread = ""
    
    SA_Write_Excel_Model.write_container_strategy(strategy_best_history,Yard_Pos_List)
    ####打印船舶配载策略 Ship_Bay, Ship_Tier, Ship_Row ship_load_strategy_list
    SA_Write_Excel_Model.write_ship_load_container(ship_load_strategy_list_best,Ship_Bay,Ship_Tier,Ship_Row,Ship_Down_Tier,Ship_Up_Tier)
    ####场桥调度
    SA_Write_Excel_Model.write_yard_strategy(yard_plan_list_best_history)

    print("******************************OK!******************************")
    print("集装箱发箱顺序，船舶配载，场桥调度策略已生成完毕，请转至桌面查看^0^")
    self.lb.insert(tk.END,"******************************OK!******************************\n")
    self.lb.insert(tk.END,"集装箱发箱顺序，船舶配载，场桥调度策略已生成完毕，请转至桌面查看^0^\n")

    min_score = min(show_score)-20
    max_score = max(show_score)+20

    my_y_ticks = np.arange(min_score, max_score, 5)
    plt.yticks(my_y_ticks)
    plt.title("模拟退火算法演进", fontproperties='SimHei', fontsize=15)
    plt.xlabel(u'迭代次数', fontproperties='SimHei', fontsize=12)
    plt.ylabel(u'倒箱量', fontproperties='SimHei', fontsize=12)

    plt.plot(show_strategy_index, show_score,linestyle="--",label="Score")
    plt.plot(show_strategy_index, show_score_history,linestyle="-.",label="Score_History")
    #plt.scatter(show_strategy_index, show_score,marker="*",label="Score")
    #plt.scatter(show_strategy_index, show_score_history,marker=".",label="Score_History")
    plt.legend()
    plt.show()
    
###########################################################################################################



class App_SA:
    def __init__(self,app):
    
        self.app = app
        app.title("基于模拟退火算法的船舶配载，集装箱调度决策方法")
        app.geometry("570x670")
           
        tk.Label(app,text="船舶结构输入：").place(x=20,y=0)

        tk.Label(app,text="船舶贝位数量：").place(x=20,y=30)
        tk.Label(app,text="船舶甲板上层：").place(x=20,y=60)
        tk.Label(app,text="船舶甲板下层：").place(x=20,y=90)
        tk.Label(app,text="船舶列位数量：").place(x=20,y=120)

        self.v_ship_bay = tk.IntVar()
        self.v_ship_tier_up = tk.IntVar()
        self.v_ship_tier_down = tk.IntVar()
        self.v_ship_row = tk.IntVar()        
        self.v_ship_excel_pos = tk.StringVar()

        self.v_ship_bay.set(21)
        self.v_ship_tier_up.set(4)
        self.v_ship_tier_down.set(6)
        self.v_ship_row.set(12)
        
        tk.Entry(app,textvariable = self.v_ship_bay).place(x=120,y=30)
        tk.Entry(app,textvariable = self.v_ship_tier_up).place(x=120,y=60)
        tk.Entry(app,textvariable = self.v_ship_tier_down).place(x=120,y=90)
        tk.Entry(app,textvariable = self.v_ship_row).place(x=120,y=120)

        tk.Button(app,text = "导入船舶信息",command = self.open_ship).place(x=20,y=150)
        tk.Entry(app,textvariable = self.v_ship_excel_pos,state="readonly").place(x=120,y=150)
         
###################################################################################

        tk.Label(app,text="堆场信息输入：").place(x=300,y=10)

        tk.Label(app,text="堆场箱区数量：").place(x=300,y=30)
        tk.Label(app,text="堆场贝位数量：").place(x=300,y=60)
        tk.Label(app,text="堆场层位数量：").place(x=300,y=90)
        tk.Label(app,text="堆场列位数量：").place(x=300,y=120)

        self.v_yard_box = tk.IntVar()
        self.v_yard_bay = tk.IntVar()
        self.v_yard_tier = tk.IntVar()
        self.v_yard_row = tk.IntVar()
        self.v_yard_excel_pos = tk.StringVar()
        
        self.v_yard_box.set(1)
        self.v_yard_bay.set(30)
        self.v_yard_tier.set(5)
        self.v_yard_row.set(10)

        
        tk.Entry(app,textvariable = self.v_yard_box).place(x=400,y=30)
        tk.Entry(app,textvariable = self.v_yard_bay).place(x=400,y=60)
        tk.Entry(app,textvariable = self.v_yard_tier).place(x=400,y=90)
        tk.Entry(app,textvariable = self.v_yard_row).place(x=400,y=120)

        tk.Button(app,text = "导入堆场信息",command = self.open_yard).place(x=300,y=150)
        tk.Entry(app,textvariable = self.v_yard_excel_pos,state="readonly").place(x=400,y=150)

###################################################################################
        tk.Label(app,text="算法参数输入：").place(x=20,y=200)

        tk.Label(app,text="初始温度：").place(x=20,y=230)
        tk.Label(app,text="终止温度：").place(x=300,y=230)
        tk.Label(app,text="内循环次数：").place(x=20,y=260)
        tk.Label(app,text="alpha_rate：").place(x=300,y=260)

        self.v_T = tk.IntVar()
        self.v_T_min = tk.DoubleVar()
        self.v_loop_in = tk.IntVar()
        self.v_alpha_rate = tk.DoubleVar()
        
        self.v_T.set("100")
        self.v_T_min.set("0.01")
        self.v_loop_in.set("50")
        self.v_alpha_rate.set("0.50")

        tk.OptionMenu(app,self.v_T,"10","20","30","40","50","60","70","80","90","100").place(x=120,y=230)
        tk.OptionMenu(app,self.v_T_min,"0.01","0.05","0.10","0.50","1.00").place(x=400,y=230)
        tk.OptionMenu(app,self.v_loop_in,"10","20","30","40","50","60","70","80","90","100").place(x=120,y=260)
        tk.OptionMenu(app,self.v_alpha_rate,"0.10","0.20","0.30","0.40","0.50","0.60","0.70","0.80","0.90","0.99").place(x=400,y=260)
        

###################################################################################
        tk.Button(app,text = "运行",command = self.Run_SA,width=10).place(x=20,y=300)
        tk.Button(app,text = "中断",command = lambda :Interrupt_Output(self),width=10).place(x=120,y=300)
###################################################################################
        frame_run_state = tk.Frame(app,width=570,height=330)
        frame_run_state.pack(side=tk.BOTTOM)
        
        tk.Label(frame_run_state,text="模拟退火算法演进过程").place(x=220,y=0)
       # v_run_infor = tk.StringVar()
       # v_run_infor.set()
        
        #self.lb = tk.Listbox(frame_run_state,listvariable=v_run_infor,width=69,height=15)
        self.lb = tk.scrolledtext.ScrolledText(frame_run_state, width=69, height=20)
        self.lb.insert(tk.END,"*************************模拟退火算法演进过程************************")
        self.lb.place(x=39,y=30)
    

    def open_ship(self):        
        dlg = win32ui.CreateFileDialog(1)  # 1表示打开文件对话框
        dlg.SetOFNInitialDir("C:\\Users\\Administrator\\Desktop\\答辩准备")  # 设置打开文件对话框中的初始显示目录
        dlg.DoModal()

        filename = dlg.GetPathName()  # 获取选择的文件名称
        self.v_ship_excel_pos.set(filename)
                
    def open_yard(self):
     
        dlg = win32ui.CreateFileDialog(1)  # 1表示打开文件对话框
        dlg.SetOFNInitialDir("C:\\Users\\Administrator\\Desktop\\答辩准备")  # 设置打开文件对话框中的初始显示目录
        dlg.DoModal()

        filename = dlg.GetPathName()  # 获取选择的文件名称       
        self.v_yard_excel_pos.set(filename)

    def Run_SA(self):

        self.lb.delete(1.0,tk.END)
        self.lb.insert(tk.END,"*************************模拟退火算法演进过程************************")
        
        global Yard_Box,Yard_Bay,Yard_Tier,Yard_Row
        global Ship_Bay,Ship_Row,Ship_Down_Tier,Ship_Up_Tier,Ship_Tier
        global T,T_min,alpha_rate,loop_in
        global xls_ship_pos_read,xls_yard_pos_read
        global Yard_Pos_List,Ship_Pos_List,Yard_Common_Container_Dictionary,Ship_Load_Container_List,Restow_Max
        global show_strategy_index,show_score,show_score_history
        global strategy_best_history,ship_load_strategy_list_best_history,yard_plan_list_best_history
        
        yard_plan_list_best = []
        ship_load_strategy_list_best = []
        strategy_best = []
       

        Yard_Pos_List = []
        Ship_Pos_List = []
        Yard_Common_Container_Dictionary = dict()
        Ship_Load_Container_List = []
        Restow_Max = 10

        show_strategy_index = []
        show_score = []
        show_score_history = []
        

        strategy_best_history = []
        ship_load_strategy_list_best_history = []
        yard_plan_list_best_history = []
        score_best_history = 10000
      
        Yard_Box,Yard_Bay,Yard_Tier,Yard_Row = self.v_yard_box.get(),self.v_yard_bay.get(),self.v_yard_tier.get(),self.v_yard_row.get()
        Ship_Bay = self.v_ship_bay.get()
        Ship_Row = self.v_ship_row.get()
        Ship_Down_Tier = self.v_ship_tier_down.get()
        Ship_Up_Tier = self.v_ship_tier_up.get()
        Ship_Tier = Ship_Down_Tier+Ship_Up_Tier

        T = self.v_T.get()
        T_min = self.v_T_min.get()
        alpha_rate = self.v_alpha_rate.get()
        loop_in = self.v_loop_in.get()

        xls_ship_pos_read = self.v_ship_excel_pos.get()
        xls_yard_pos_read = self.v_yard_excel_pos.get()

        global SA_Thread
        
        SA_Thread = threading.Thread(target = Run_SA,args = (self,))
        SA_Thread.start()


        

app = tk.Tk()
app.iconbitmap("C:\\Users\\Administrator\\Desktop\\图片\\port.ico")
App_SA(app)

app.mainloop()






