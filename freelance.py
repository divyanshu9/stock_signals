import math
import csv
import xlrd
import numpy


def calculate_standard_deviation(arr_values):
    return numpy.std(arr_values)


location = ("1650107370-2--Sample----with-criteria.xlsx")
wb = xlrd.open_workbook(location)
s1 = wb.sheet_by_index(0)
s1.cell_value(0, 0)  # initializing cell from the cell position

print(" No. of rows: ", s1.nrows)
print(" No. of columns: ", s1.ncols)
stock_data = list(s1.get_rows())[2:]
len_stock_data = len(stock_data)
forward_column_map = {"1-Day Return": -1, "10-Day Return": 10, "25-Day Return": 25}
new_column_map_results = {"criteria-90": [None] * 10}
first, last = 0, 0
for key in forward_column_map:
    forward_index = forward_column_map[key]
    if key not in new_column_map_results:
        new_column_map_results[key] = []
    for index in range(0, len_stock_data):
        try:
            stat_value = stock_data[index][4].value
            last_value = (stock_data[index + forward_index][4]).value
            if ((index + forward_index) < 0) or ((index + forward_index) > len_stock_data):
                result = -1
            else:
                if forward_index > 0:
                    result = ((last_value - stat_value) / (stat_value * 1.0)) * 100.0
                else:
                    result = ((stat_value - last_value) / (last_value * 1.0)) * 100.0
            new_column_map_results[key].append(result)
        except:
            # noinspection PyTypeChecker
            new_column_map_results[key].append(-1)

last_new_column_map = {"V-10": 10, "V-90": 90}
for key in last_new_column_map:
    last = last_new_column_map[key]
    if key not in new_column_map_results:
        new_column_map_results[key] = [0] * last
    for index in range(last, len_stock_data):
        last = last + 1
        start_index = last - last_new_column_map[key]
        values = new_column_map_results["1-Day Return"][start_index:last]
        std_dev = calculate_standard_deviation([i for i in values if i != -1])
        new_column_map_results[key].append(std_dev * math.sqrt(365))
hvg = []
for index in range(0, len_stock_data):
    try:
        hvg.append(new_column_map_results["V-10"][index] / new_column_map_results["V-90"][index])
    except:
        hvg.append(-1)

criteria_map = {"criteria-10": 10}

for key in criteria_map:
    last = criteria_map[key]
    if key not in new_column_map_results:
        new_column_map_results[key] = [None] * 90
    for index in range(90, len_stock_data):
        last = index
        start_index = last - criteria_map[key]
        values = hvg[start_index + 1:last + 1]
        max_value = max([i for i in values if i != math.inf])
        max_value_index = [index for index in range(0, len(values)) if max_value == values[index]][-1]
        main_index = (last - 10) + max_value_index + 1
        close_value = stock_data[len_stock_data - 1 if main_index == len_stock_data else main_index][4].value
        divide_value = ((stock_data[len_stock_data - 1 if last == len_stock_data else last][
                             4].value * 1.0) / close_value) / (hvg[index] / max_value)
        new_column_map_results[key].append(divide_value * 100.0)

for index in range(10, len_stock_data):
    last = index
    start_index = last - 10
    values_2 = new_column_map_results["V-10"][start_index + 1:last + 1]
    # noinspection PyTypeChecker
    max_value_2 = max([i for i in values_2 if i != math.inf])
    max_value_index_2 = [index for index in range(0, len(values_2)) if max_value_2 == values_2[index]][-1]
    main_index_2 = (last - 10) + max_value_index_2 + 1
    close_value_2 = stock_data[len_stock_data - 1 if main_index_2 == len_stock_data else main_index_2][4].value
    divide_value_2 = ((stock_data[len_stock_data - 1 if last == len_stock_data else last][
                           4].value * 1.0) / close_value_2) / (new_column_map_results["V-10"][index] / max_value_2)
    new_column_map_results["criteria-90"].append(divide_value_2 * 100.0)

t1, t2, n = 200, 200, 3
check_mapping = {str(n) + "days": n}
c1_check_array, check_one_passed = [], False
c1, c2, cons_check_days_count = False, False, 0

for index in range(0, len_stock_data):
    c1, c2 = False, False
    # noinspection PyTypeChecker
    if new_column_map_results["criteria-90"][index] is not None and float(
            new_column_map_results["criteria-90"][index]) > t1:
        c1 = True
    elif new_column_map_results["criteria-10"][index] is not None and float(
            new_column_map_results["criteria-10"][index]) > t2:
        c2 = True
    if c1 or c2:
        cons_check_days_count += 1
        c1_check_array.append(index)
    else:
        if check_one_passed:
            break
        cons_check_days_count = 0
        c1_check_array = []
    if cons_check_days_count >= n:
        check_one_passed = True

if not check_one_passed:
    print("Check one failed")
    exit()
print(new_column_map_results)
