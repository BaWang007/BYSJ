import xlrd
import xlwt
import xlutils.copy
import copy
import time
import math


def write_ship_load_container(ship_load_strategy_list_best,Ship_Bay,Ship_Tier,Ship_Row,Ship_Down_Tier,Ship_Up_Tier):
    time_str = time.strftime("%Y-%m-%d %H-%M-%S",time.localtime((time.time())))
    xls_pos = "C:/Users/Administrator/Desktop/strategy_ship"+time_str+".xls"

    wb = xlwt.Workbook()
    sh = wb.add_sheet('strategy_ship', cell_overwrite_ok=True)

    al = xlwt.Alignment()
    al.horz = 0x02  # 设置水平居中
    al.vert = 0x01  # 设置垂直居中

    style = xlwt.XFStyle()
    pattern = xlwt.Pattern()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    style.alignment = al
    pattern.pattern_fore_colour = xlwt.Style.colour_map['yellow']
    style.pattern = pattern

    style_al = xlwt.XFStyle()  # 创建一个样式对象，初始化样式
    style_al.alignment = al

    style_mark = xlwt.XFStyle()  # 创建一个样式对象，初始化样式
    style_mark.alignment = al

    borders = xlwt.Borders()  # Create Borders

    # DASHED虚线
    # NO_LINE没有
    # THIN实线

    # NO_LINE, THIN, MEDIUM, DASHED, DOTTED, THICK, DOUBLE, HAIR, MEDIUM_DASHED,
    # THIN_DASH_DOTTED, MEDIUM_DASH_DOTTED, THIN_DASH_DOT_DOTTED,
    # MEDIUM_DASH_DOT_DOTTED, SLANTED_MEDIUM_DASH_DOTTED, or 0x00 through 0x0D.
    borders.left = xlwt.Borders.MEDIUM
    borders.right = xlwt.Borders.MEDIUM
    borders.top = xlwt.Borders.MEDIUM
    borders.bottom = xlwt.Borders.MEDIUM
    borders.left_colour = 0x40
    borders.right_colour = 0x40
    borders.top_colour = 0x40
    borders.bottom_colour = 0x40

    style.borders = borders
    style_al.borders = borders

    bay_cycle = 3 + Ship_Row

    for i in range(Ship_Bay):
        for j in range(Ship_Tier):
            for k in range(Ship_Row):

                con_id = ship_load_strategy_list_best[i][j][k].ship_con.con_id

                bay = i + 1
                tier = j + 1
                row = k + 1

                if tier <= Ship_Down_Tier:
                    ## 15 = 3 + 12     bay_cycle = 3+Ship_Row
                    ## 7 = Ship_Up_Tier+3+(Ship_Down_Tier+1-tier)
                    write_tier = bay_cycle * (math.ceil(bay / 3) - 1) + Ship_Up_Tier + 3 + (Ship_Down_Tier + 1 - tier)
                else:
                    ## (11-tier)+2 = Ship_Tier + 1 -tier +2
                    write_tier = bay_cycle * (math.ceil(bay / 3) - 1) + Ship_Tier + 1 - tier + 2

                bay_list = [3, 1, 2]
                write_row = bay_cycle * (bay_list[bay % 3] - 1) + 2 + row

                sh.col(write_row - 1).width = 4000

                if con_id:
                    if "SMU" in con_id:
                        ###黄色 对齐 边框
                        sh.write(write_tier - 1, write_row - 1, con_id, style=style)
                    else:
                        ### 对齐 边框
                        sh.write(write_tier - 1, write_row - 1, con_id, style_al)
                else:
                    sh.write(write_tier - 1, write_row - 1, con_id)

                tier_mark_row = bay_cycle * (bay_list[bay % 3] - 1) + 2

                row_mark_tier = bay_cycle * (math.ceil(bay / 3) - 1) + 2

                sh.write(write_tier - 1, tier_mark_row - 1, tier, style_mark)
                sh.write(row_mark_tier - 1, write_row - 1, row, style_mark)

        # 14 = 10 + 1 + 2    -> 10 + 3  = Ship_Tier+1+2
        write_tier = bay_cycle * (math.ceil(bay / 3) - 1) + Ship_Tier + 1 + 2 + 1
        write_row_min = bay_cycle * (bay_list[bay % 3] - 1) + 2 + 1
        write_row_max = bay_cycle * (bay_list[bay % 3] - 1) + 2 + Ship_Row
        bay_str = "第" + str(i + 1) + "贝位，即Bay" + str(2 * i + 1)
        sh.write_merge(write_tier - 1, write_tier - 1, write_row_min - 1, write_row_max - 1, bay_str, style_al)

    wb.save(xls_pos)


def write_yard_strategy(yard_plan_list_best):
    time_str = time.strftime("%Y-%m-%d %H-%M-%S",time.localtime((time.time())))
    xls_pos = "C:/Users/Administrator/Desktop/strategy_yard"+time_str+".xls"

    wb = xlwt.Workbook()
    sh = wb.add_sheet('strategy_yard', cell_overwrite_ok=True)

    sh.write(0, 0, "步骤")
    sh.write(0, 1, "操作内容")
    sh.write(0,2,"起点")
    sh.write(0,3,"终点")
    sh.write(0,4,"集装箱船舶位置")

    tier_index = 0

    row_step = 0
    row_plan = 1
    row_start = 2
    row_end = 3
    row_con_ship = 4

    for plan in yard_plan_list_best:
        tier_index += 1

        sh.write(tier_index, row_step, plan[0])
        sh.write(tier_index, row_plan, plan[1])
        sh.write(tier_index, row_start, str(plan[2]))
        sh.write(tier_index, row_end, str(plan[3]))
        sh.write(tier_index, row_con_ship, str(plan[4]))

    wb.save(xls_pos)


def write_container_strategy(strategy_best,Yard_Pos_List):
    time_str = time.strftime("%Y-%m-%d %H-%M-%S",time.localtime((time.time())))
    xls_pos = "C:/Users/Administrator/Desktop/strategy_container"+time_str+".xls"

    wb = xlwt.Workbook()
    sh = wb.add_sheet('strategy_container', cell_overwrite_ok=True)

    sh.write(0, 0, "顺序")
    sh.write(0, 1, "集装箱号")

    tier_index = 0

    row_step = 0
    row_con_id = 1

    # print("装船顺序：")
    print_line_list = []
    for strate in strategy_best:
        tier_index += 1

        con_step = tier_index
        con_id = Yard_Pos_List[strate[0][0] - 1][strate[0][1] - 1][strate[0][2] - 1][strate[0][3] - 1].yard_con.con_id

        sh.write(tier_index, row_step, con_step)
        sh.write(tier_index, row_con_id, con_id)

    wb.save(xls_pos)
