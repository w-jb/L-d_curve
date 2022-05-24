import pandas as pd
import matplotlib.pyplot as plt
import xlsxwriter as xwr


# show menu
def show_menu():
    # message to display
    message = """Press:
    [1] - read data file
    [2] - analyze provided data
    [3] - generate report
    [4] -exit"""
    print(message)
    # read user input
    # check if integer
    try:
        selected_option = int(input())
    # not proper input
        if selected_option not in [1, 2, 3, 4]:
            print("Invalid input - integer out of range 1-4")
            show_menu()
    except ValueError:
        print("Invalid input - not an integer")
        show_menu()
    match selected_option:
        case 1:
            open_file()
        case 2:
            # depends on from file type
            if txt_file:
                # data from *.txt shall be preprocessed
                data_process()
            else:
                # data from excel might be processed directly
                stiffness_calculation()
                stiffness_curve()
                hysteresis_curve()
        case 3:
            report()
        case 4:
            exit()


# open data file
def open_file():
    # data handling
    data_set = []
    # get file name
    print("Enter file name")
    file_name = input()
    try:
        # open file, read only
        with open(file_name, "r") as file:
            # separate end of the file name
            file_ext = file_name[-4:]
            # check file type
            if file_ext == ".txt":
                # read lines of text
                data = file.readlines()
                # for each line check if contains data, or is empty
                for line in data:
                    if line == '\n' or line == '\t\n':
                        # empty line means end of data set
                        loaded_data_table.append(list(data_set))
                        data_set.clear()
                    else:
                        # store point data in set
                        data_set.append(line)
                # store set of data
                loaded_data_table.append(list(data_set))
                # report readiness
                print(file_name + " opened")
                show_menu()

            elif file_ext == ".xls" or file_ext == "xlsx":
                global txt_file
                txt_file = False
                excel_df = pd.read_excel(file_name)
                excel_data = pd.DataFrame(excel_df).values.tolist()
                for cell in excel_data:
                    if pd.isnull(cell).any():
                        results_table.append(list(data_set))
                        data_set.clear()
                    else:
                        data_set.append(tuple(cell))
                results_table.append(list(data_set))
                print(file_name + " opened")
                show_menu()

    except FileNotFoundError:
        print("File not found")
        show_menu()


# data processing
def data_process():
    # verify if data available
    if len(loaded_data_table) == 0:
        print("Data not loaded")
        show_menu()
    else:
        # check number of elements
        data_sets = len(loaded_data_table)
        for element in range(data_sets):
            # list for data storage
            test_point = list()
            # for each line
            for data_block in loaded_data_table[element]:
                # separate load and displacement separated by sequence:"' "
                item_list = data_block.split("\t")
                # store displacement cast as float
                first_parameter = float(item_list[0].replace(",", "."))
                # store load cast as float
                second_parameter = float(item_list[1].replace(",", "."))
                # combine load and displacement int tuple
                point = (first_parameter, second_parameter)
                # add single tuple to data set
                test_point.append(point)
            # store data set
            results_table.append(list(test_point))

    # proceed with stiffness calculation and plotting
    stiffness_calculation()
    stiffness_curve()
    hysteresis_curve()


# stiffness curve plot
def stiffness_curve():
    # global bool to set appropriate units
    global kN

    # process data storage
    xs = list()
    ys = list()

    # processed curve
    graphs_number = 0

    # plot curve for each data set
    for curve in results_table:
        # load values set to determine units
        y_range = set()

        # add points for plotting
        for point in curve:
            xs.append(point[0])
            ys.append(point[1])
            y_range.add(point[1])

        # check if N or kN shall be used
        y_max = max(y_range)
        if y_max > 50:
            unit = '[N]'
        else:
            unit = '[kN]'
            kN = True

        # use calculated stiffness for description
        stiffness = slope[graphs_number]
        # plot curve
        plt.figure(graphs_number)
        plt.xlabel('displacement [mm]')
        plt.ylabel('force ' + unit)
        plt.suptitle(f'Graph number: {graphs_number}, stiffness: {stiffness}')
        plt.plot(xs, ys)
        plt.savefig(f'stiffness{graphs_number}.png')
        graphs_number += 1
    # plot
    plt.show()


# calculate and plot hysteresis curve
def hysteresis_curve():
    # storage for load and displacement
    xs = list()
    ys = list()
    # rounded values
    xs_round = list()
    ys_round = list()
    # number of currently processed measurement data set
    graphs_number = 0

    # calculate for each data set separately
    for curve in results_table:
        # sets to store ranges of displacement and load
        y_range = set()
        x_range = set()
        # lists for hysteresis and round values
        y_hysteresis = list()
        yr_hysteresis = list()

        # prepare points data for process
        for point in curve:
            xs.append(point[0])
            ys.append(point[1])

        # round and add points to data set - load
        for each in xs:
            xs_round.append(round(each, 3))
            x_range.add(round(each, 3))

        # find extreams
        xr_min = min(xs_round)
        xr_max = max(xs_round)

        # round and add points to data set - travel
        for each in ys:
            ys_round.append(round(each, 3))
            y_range.add(round(each, 3))

        # support list for doubling the data set
        xs_cut = xs_round
        ys_cut = ys_round

        # eliminate first element before joining
        xs_cut.pop(0)
        ys_cut.pop(0)

        # join main and supporting list
        xs_round.extend(xs_cut)
        ys_round.extend(ys_cut)

        # find extreams position
        first_xrmax = xs_round.index(xr_max)
        first_xrmin = xs_round.index(xr_min)

        # cut from data set points before slope begin
        xr_downslope = xs_round[first_xrmax:]
        # remove points after slope ends
        del xr_downslope[first_xrmin-(first_xrmax-1):]

        # repeat above for second slope
        xr_upslope = xs_round[first_xrmin:]
        second_xrmax = xr_upslope.index(xr_max)
        del xr_upslope[(second_xrmax+1):]

        # prepare load data based on displacement input for both slopes
        yr_downslope = ys_round[first_xrmax:]
        del yr_downslope[first_xrmin-(first_xrmax-1):]
        yr_upslope = ys_round[first_xrmin:]
        del yr_upslope[(second_xrmax+1):]

        # change order of the data set to correspond with rising displacement
        yr_upslope.reverse()

        # calculate difference in force value between slopes
        for (item1, item2) in zip(yr_downslope, yr_upslope):
            y_hysteresis.append(item1 - item2)

        # round calculated values
        for each in y_hysteresis:
            yr_hysteresis.append(round(each, 3))

        # find extreams
        hyst_max = max(yr_hysteresis)
        hyst_min = min(yr_hysteresis)

        # store extreams with corresponding displacement
        for (position, difference) in zip(xr_upslope, yr_hysteresis):
            if difference == hyst_max:
                hysteresis.append((position, difference))
            if difference == hyst_min:
                hysteresis.append((position, difference))

        # graph plot
        plt.figure(graphs_number)
        plt.suptitle(f'Hysteresis number: {graphs_number}')
        plt.xlabel('displacement [mm]')
        plt.ylabel('difference in force')
        plt.plot(xr_upslope, yr_hysteresis)
        plt.savefig(f'hysteresis{graphs_number}.png')
        graphs_number += 1
    # graph display and return to main menu
    plt.show()
    show_menu()


# menu for stiffness calculation
# user input of range to calculate
def stiffness_calculation():
    global selected_option
    # message to display
    message = """Select range:
        [1] - +/- 0.5mm
        [2] - user selected
        [3] - back"""
    print(message)
    # read user input
    # check if correct
    try:
        selected_option = int(input())
        # not proper input
        if selected_option not in [1, 2, 3]:
            print("Invalid input - integer out of range 1-3")
            stiffness_calculation()
    except ValueError:
        print("Invalid input - not an integer")
        stiffness_calculation()
    match selected_option:
        case 1:
            calculation(0.5)
        case 2:
            enter_range()
        case 3:
            show_menu()


# calculate stiffness in requested range
def calculation(slope_range):
    # lists to store slopes data
    # displacements and force
    xs1 = list()
    ys1 = list()
    xs2 = list()
    ys2 = list()
    # lists for sum elements of calculated slope
    slope_numerator_portion = list()
    slope_denominator_portion = list()

    # bool to separate the rising and falling slope
    first_slope = False

    # calculate stiffness for each data sequence separately
    for curve in results_table:
        # end of curve points
        max_point = 0.0
        min_point = 0.0

        # find max and min in displacement (end points)
        for point in curve:
            if point[0] < min_point:
                min_point = point[0]
            if point[0] > max_point:
                max_point = point[0]

        # add points in range to storage lists
        for point in curve:
            if -1 * slope_range <= point[0] <= slope_range:
                if first_slope:
                    xs1.append(point[0])
                    ys1.append(point[1])
                else:
                    xs2.append(point[0])
                    ys2.append(point[1])
            # change slope when end of curve is reached
            if point[0] == max_point or point[0] == min_point:
                first_slope = not first_slope

        # average calculation for each slope (load and displacement)
        xs1_avg = average(xs1)
        ys1_avg = average(ys1)
        xs2_avg = average(xs2)
        ys2_avg = average(ys2)

        # calculate first slope
        for element in range(len(xs1)):
            slope_numerator_portion.append((xs1[element] - xs1_avg)*(ys1[element] - ys1_avg))
            slope_denominator_portion.append((xs1[element] - xs1_avg) ** 2)
        half_slope1 = (sum(slope_numerator_portion) / sum(slope_denominator_portion))

        # calculate second slope
        for element in range(len(xs2)):
            slope_numerator_portion.append((xs2[element] - xs2_avg)*(ys2[element] - ys2_avg))
            slope_denominator_portion.append((xs2[element] - xs2_avg) ** 2)
        half_slope2 = (sum(slope_numerator_portion) / sum(slope_denominator_portion))

        # add average of two slopes to storage list
        slope.append((half_slope1 + half_slope2) / 2)


# calculate average of the list
def average(var):
    var_number = len(var)
    var_sum = sum(var)
    return var_sum / var_number


def enter_range():
    message = """Enter range [+/-]: """
    print(message)
    try:
        user_input = float(input())
        calculation(user_input)
    except ValueError:
        print("Invalid input - not an number")
        stiffness_calculation()


# store calculated data in excel file
# input: data from lists: slope, hysteresis
# output: excel file
def report():
    # initiate file
    workbook = xwr.Workbook('report.xlsx')
    worksheet = workbook.add_worksheet()

    # write header
    worksheet.write('A1', 'Report')
    worksheet.write('A2', 'Test run:')
    worksheet.write('B2', 'Stiffness')
    worksheet.write('C2', 'Hysteresis max')
    worksheet.write('D2', '@mm')
    worksheet.write('E2', 'Hysteresis min')
    worksheet.write('F2', '@mm')

    # used unit type for the stiffness
    if kN:
        unit = "kN"
    else:
        unit = "N"

    # write data
    row = 3
    for each in range(len(slope)):
        worksheet.write(f'A{row}', f'{each}')
        worksheet.write(f'B{row}', f'{round(slope[each], 2)} {unit}')
        worksheet.write(f'C{row}', f'{round(hysteresis[each][1], 2)}')
        worksheet.write(f'D{row}', f'{round(hysteresis[each][0], 2)}')
        worksheet.write(f'E{row}', f'{round(hysteresis[each+1][1], 2)}')
        worksheet.write(f'F{row}', f'{round(hysteresis[each+1][0], 2)}')
        row += 1

    # finalize file
    workbook.close()


# globals to transfer data
# bool True if data provided in *.txt; False if excel file
txt_file = True
# bool True if force measurement in kN; if False, force in N
kN = False
# handling data list for *.txt source
loaded_data_table = list()
# handling list for preprocessed data
results_table = list()
# calculated stiffness storage
slope = list()
# calculated hysteresis storage - tuples(force delta, @displacement
hysteresis = list()
# main menu - start function
show_menu()
