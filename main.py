# variable to communicate with menu
selected_option = 0


def show_menu():
    global selected_option
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
    # not improper input
        if selected_option not in [1, 2, 3, 4]:
            print("Invalid input - integer out of range 1-4")
            show_menu()
    except ValueError:
        print("Invalid input - not an integer")
        show_menu()


show_menu()
print(selected_option)
match selected_option:
    case 1:
        pass
    case 2:
        pass
    case 3:
        pass
    case 4:
        exit()
