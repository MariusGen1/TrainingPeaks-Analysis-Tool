import pandas as pd
import datetime, collections, os, subprocess, platform
from math import isnan
from tkinter import Tk, filedialog

root = Tk() # pointing root to Tk() to use it as Tk() in program.
root.withdraw() # Hides small tkinter window.
root.attributes('-topmost', True)

open_file = filedialog.askopenfilenames(filetypes=[("CSV", ".csv")])

data = pd.read_csv(open_file[0])
# data = pd.read_csv('workouts.csv')

months={1: 'Januar', 2: 'Februar', 3: 'Mars', 4: 'April', 5: 'Mai', 6: 'Juni', 7: 'Juli', 8: 'August', 9: 'September', 10: 'Oktober', 11: 'November', 12: 'Desember'}

def getTimePeriod(row):
    date = datetime.datetime.strptime(row['WorkoutDay'], "%Y-%m-%d")
    timePeriod = months[date.month] + ' ' + str(date.year)
    return timePeriod

def getData(dataframe, name):
    values=collections.defaultdict(float)
    for index, row in dataframe.iterrows():
        if name in row and not isnan(row[name]):
            values[getTimePeriod(row)] = float(values[getTimePeriod(row)]) + row[name]
        else:
            values[getTimePeriod(row)] = float(values[getTimePeriod(row)])

    values['Sum'] = sum(values.values())
    return values


def format(values, exponent, roundto, unit):
    newValues = []
    for value in values:
        if unit=='':
            newValues.append(round(value*(10**exponent), roundto))
        else:
            newValues.append((str(round(value*(10**exponent), roundto))+unit))
    return newValues


def getDataFromDataframe(dataframe):
    formatted = {
        'Tidsperiode': getData(dataframe, 'TimeTotalInHours').keys(),
        'Tid': format(getData(dataframe, 'TimeTotalInHours').values(), 0, 1, ''),
        'TSS': format(getData(dataframe, 'TSS').values(), 0, 0, ''),

        'Totaldistanse': format(getData(dataframe, 'DistanceInMeters').values(), -3, 0, ''),
        'Tid i i1 (puls)': format(getData(dataframe, 'HRZone1Minutes').values(), -1.77815, 1, ''),
        'Tid i i2 (puls)': format(getData(dataframe, 'HRZone2Minutes').values(), -1.77815, 1, ''),
        'Tid i i3 (puls)': format(getData(dataframe, 'HRZone3Minutes').values(), -1.77815, 1, ''),
        'Tid i i4 (puls)': format(getData(dataframe, 'HRZone4Minutes').values(), -1.77815, 1, ''),
        'Tid i i5 (puls)': format(getData(dataframe, 'HRZone5Minutes').values(), -1.77815, 1, ''),
        'Tid i i6 (puls)': format(getData(dataframe, 'HRZone6Minutes').values(), -1.77815, 1, ''),

        'Tid i i1 (watt)': format(getData(dataframe, 'PWRZone1Minutes').values(), -1.77815, 1, ''),
        'Tid i i2 (watt)': format(getData(dataframe, 'PWRZone2Minutes').values(), -1.77815, 1, ''),
        'Tid i i3 (watt)': format(getData(dataframe, 'PWRZone3Minutes').values(), -1.77815, 1, ''),
        'Tid i i4 (watt)': format(getData(dataframe, 'PWRZone4Minutes').values(), -1.77815, 1, ''),
        'Tid i i5 (watt)': format(getData(dataframe, 'PWRZone5Minutes').values(), -1.77815, 1, ''),
        'Tid i i6 (watt)': format(getData(dataframe, 'PWRZone6Minutes').values(), -1.77815, 1, ''),
    }
    return pd.DataFrame.from_dict(formatted)


def columnChart(workbook, worksheet, x_axis, title, categories_cells, values_cells, position):
    chart = workbook.add_chart({'type': 'column'})
    chart.set_title({'name': title})
    chart.add_series({'categories': categories_cells, 'values': values_cells, 'gap': 10})
    chart.set_y_axis({'name': x_axis, 'major_gridlines': {'visible': True}})
    chart.set_legend({'position': 'none'})
    worksheet.insert_chart(position, chart)

def pieChart(workbook, worksheet, title, categories_cells, values_cells, position, colors):
    chart = workbook.add_chart({'type': 'pie'})
    chart.set_title({'name': title})
    chart.add_series({'categories': categories_cells, 'values': values_cells, 'points': colors})
    worksheet.insert_chart(position, chart)


def sheet_from_dataframe(dataframe, sheet_name):
    getDataFromDataframe(dataframe).to_excel(writer, sheet_name=sheet_name)

    workbook = writer.book
    worksheet = writer.sheets[sheet_name]

    icolors = [{'fill': {'color': '#999999'}},{'fill': {'color': '#3399ff'}},{'fill': {'color': '#00ff00'}},{'fill': {'color': '#ffff00'}},{'fill': {'color': '#ff3300'}},{'fill': {'color': '#990000'}}]

    maxRow=1+len(getData(dataframe, 'TimeTotalInHours').keys())

    pieChart(workbook, worksheet, 'Tid i pulssoner', '='+sheet_name+'!F1:K1', '='+sheet_name+'!F'+str(maxRow)+':K'+str(maxRow), 'B'+str(maxRow+2), icolors)

    pieChart(workbook, worksheet, 'Tid i wattsoner', '='+sheet_name+'!L1:Q1', '='+sheet_name+'!L'+str(maxRow)+':Q'+str(maxRow), 'J'+str(maxRow+2), icolors)

    columnChart(workbook, worksheet, 'Tid', 'Totaltid', '='+sheet_name+'!B2:B'+str(maxRow), '='+sheet_name+'!C2:C'+str(maxRow-1), 'B'+str(maxRow+18))

    columnChart(workbook, worksheet, 'TSS', 'TSS', '='+sheet_name+'!B2:B'+str(maxRow), '='+sheet_name+'!D2:D'+str(maxRow-1), 'J'+str(maxRow+18))

    columnChart(workbook, worksheet, 'Distanse (km)', 'Totaldistanse', '='+sheet_name+'!B2:B'+str(maxRow), '='+sheet_name+'!E2:E'+str(maxRow-1), 'B'+str(maxRow+34))

    # set_column(first_col, last_col, width, cell_format, options)
    worksheet.set_column(1, 1, 15)



writer = pd.ExcelWriter('Trening.xlsx', engine='xlsxwriter')

sheet_from_dataframe(data, 'Oversikt')
for sport in data['WorkoutType'].unique():
    sport_data = data[data['WorkoutType']==sport]

    sheet_from_dataframe(sport_data, sport)

data.groupby(pd.to_datetime(data['WorkoutDay'],format="%Y-%m-%d"))

data.to_excel(writer, sheet_name='All data')

writer.save()
subprocess.call(('open', 'Trening.xlsx'))















