from Notify import Notify

from datetime import timedelta
import matplotlib.pyplot as plt
import pandas as pd
import matplotlib
import requests
import datetime
import time
import xlrd

def test_main_1():
    while True:
        now = datetime.datetime.now()
        minute = now.minute
        seven_hour = now.hour + 8
        seven_minute = now.minute

        # 1.整點10分出前後3小時
        # if minute == 10:
        if True:
            today_date = (datetime.datetime.now() + timedelta(hours=8)).strftime('%Y_%m_%d')
            file_name = f'{today_date}_update.xls'
            response = requests.get(f'https://www.taoyuan-airport.com/uploads/fos/{file_name}')

            if response.status_code == 200:
                with open(file_name, 'wb') as file:
                    file.write(response.content)

                now_hour = str(datetime.datetime.now().hour + 8) + ':00'

                workbook = xlrd.open_workbook(filename=file_name)
                worksheet = workbook.sheet_by_index(0)

                def __cell_value(col_num):
                    for row in range(start_row, end_row + 1):
                        h_value = worksheet.cell_value(row, col_num)
                        if now_hour in h_value:
                            def _cell_value(row_num):
                                data_values = []
                                for col in range(col_num, col_num+3):
                                    try:
                                        cell_value = int(worksheet.cell_value(row-row_num, col))
                                    except:
                                        cell_value = worksheet.cell_value(row-row_num, col)
                                    if '小計' in str(cell_value) or '' == str(cell_value) or '時間區間' == str(cell_value):return
                                    data_values.append(cell_value)
                                for col in range(col_num+1, col_num+3):
                                    try:
                                        cell_value = int(worksheet.cell_value(row-row_num, col+7))
                                    except:
                                        print(row-row_num, col+7)
                                        cell_value = worksheet.cell_value(row-row_num, col+7)
                                    if '小計' in str(cell_value) or '' == str(cell_value) or '時間區間' == str(cell_value):return
                                    data_values.append(cell_value)
                                all_data.append(data_values)

                            for ii in [3, 2 ,1, 0, -1, -2, -3]:
                                _cell_value(ii)
                            break
                try:
                    start_row = 3
                    end_row = 26
                    all_data = []
                    __cell_value(7)
                except:
                    start_row = 3
                    end_row = 26
                    all_data = []
                    __cell_value(10)
                for row in all_data:
                    row[2], row[3] = row[3], row[2]

                total = [0] * (len(all_data[0])-1)
                for item in all_data:
                    for i, value in enumerate(item[1:]):
                        total[i] += value
                all_data.append(['Total'] + total)
                print(all_data)

                # columns=['時間區間', '入境桃園(一)', '入境桃園(二)', '出境桃園(一)', '出境桃園(二)']
                columns=['Time', 'Entering Taoyuan (1)', 'Entering Taoyuan (2)', 'Departing Taoyuan (1)', 'Depart Taoyuan (2)']
                df1 = pd.DataFrame(all_data, columns=columns)

                # font_path = '微软正黑体.ttf'
                # matplotlib.rcParams['font.family'] = matplotlib.font_manager.FontProperties(fname=font_path).get_name()

                fig, ax = plt.subplots(figsize=(2, 2))
                ax.axis('tight')
                ax.axis('off')
                col_colors = ['#cce5ff']*5
                cellColours = []
                for i in range(len(all_data)):
                    if i%2 == 0:
                        cellColours+=[['white']*5]
                    else:
                        cellColours+=[['#cce5ff']*5]
                table=ax.table(cellText=all_data, colLabels=columns, cellLoc='center', loc='center', colColours=col_colors, cellColours=cellColours)

                for cell in table._cells:
                    cell_text = table._cells[cell].get_text()
                    cell_text.set_fontsize(12)
                table.auto_set_column_width(col=list(range(df1.shape[1])))
                for key, cell in table.get_celld().items():
                    cell.set_height(0.2)

                plt.savefig('data1_table.png', bbox_inches='tight', pad_inches=0.2)
                plt.close()

                Notify('整點10分-提醒', 'data1_table.png')
                time.sleep(2340)

        # 2.晚上7:00出明天的
        if seven_hour == 19 and seven_minute in [0, 1, 2, 3, 4 , 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18 ,19, 20]:
        # if True:
            next_date = (datetime.datetime.now().date() + datetime.timedelta(hours=32)).strftime('%Y_%m_%d')
            file_name = f'{next_date}.xls'
            url = f'https://www.taoyuan-airport.com/uploads/fos/{file_name}'
            response = requests.get(url)

            if response.status_code == 200:
                with open(file_name, 'wb') as file:
                    file.write(response.content)

                workbook = xlrd.open_workbook(filename=file_name)
                worksheet = workbook.sheet_by_index(0)

                def __cell_value(col_num):
                    for row in range(start_row, end_row + 1):
                        h_value = worksheet.cell_value(row, col_num)
                        def _cell_value():
                            data_values = []
                            for col in range(col_num, col_num+3):
                                try:
                                    cell_value = int(worksheet.cell_value(row, col))
                                except:
                                    cell_value = worksheet.cell_value(row, col)
                                if '小計' in str(cell_value) or '' == str(cell_value) or '時間區間' == str(cell_value):return
                                data_values.append(cell_value)
                            for col in range(col_num+1, col_num+3):
                                try:
                                    cell_value = int(worksheet.cell_value(row, col+7))
                                except:
                                    print(row, col+7)
                                    cell_value = worksheet.cell_value(row, col+7)
                                if '小計' in str(cell_value) or '' == str(cell_value) or '時間區間' == str(cell_value):return
                                data_values.append(cell_value)
                            all_data.append(data_values)
                        _cell_value()
                try:
                    start_row = 3
                    end_row = 26
                    all_data = []
                    __cell_value(7)
                except:
                    start_row = 3
                    end_row = 26
                    all_data = []
                    __cell_value(10)
                for row in all_data:
                    row[2], row[3] = row[3], row[2]

                total = [0] * (len(all_data[0])-1)
                for item in all_data:
                    for i, value in enumerate(item[1:]):
                        total[i] += value
                all_data.append(['Total'] + total)
                print(all_data)

                # columns=['時間區間', '入境桃園(一)', '入境桃園(二)', '出境桃園(一)', '出境桃園(二)']
                columns=['Time', 'Entering Taoyuan (1)', 'Entering Taoyuan (2)', 'Departing Taoyuan (1)', 'Depart Taoyuan (2)']
                df1 = pd.DataFrame(all_data, columns=columns)

                # font_path = '微软正黑体.ttf'
                # matplotlib.rcParams['font.family'] = matplotlib.font_manager.FontProperties(fname=font_path).get_name()

                fig, ax = plt.subplots(figsize=(2, 2))
                ax.axis('tight')
                ax.axis('off')
                col_colors = ['#cce5ff']*5
                cellColours=[['white']*5, ['#cce5ff']*5]*12+[['white']*5]
                table=ax.table(cellText=all_data, colLabels=columns, cellLoc='center', loc='center', colColours=col_colors, cellColours=cellColours)

                for cell in table._cells:
                    cell_text = table._cells[cell].get_text()
                    cell_text.set_fontsize(12)
                table.auto_set_column_width(col=list(range(df1.shape[1])))
                for key, cell in table.get_celld().items():
                    cell.set_height(0.2)

                plt.savefig('data2_table.png', bbox_inches='tight', pad_inches=0.2)
                plt.close()

                Notify('晚上7:00-提醒', 'data2_table.png')
                time.sleep(540)
