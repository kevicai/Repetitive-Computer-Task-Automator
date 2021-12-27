import pyautogui
import time
import xlrd
import pyperclip


def mouse_click(click_times: int, l_or_right: str, img: str, num_of_repeats: int) -> None:
    # if command is to repeat forever
    if num_of_repeats == -1:
        while True:
            location = pyautogui.locateCenterOnScreen(img)
            if location is not None:
                pyautogui.click(location.x, location.y, clicks=click_times, interval=0.2,
                                duration=0.2, button=l_or_right)
            time.sleep(0.1)
    elif num_of_repeats >= 1:
        for _ in range(num_of_repeats):
            location_invalid = True
            location = None
            while location_invalid:
                location = pyautogui.locateCenterOnScreen(img)
                location_invalid = location is None
                if location_invalid:
                    print("Couldn't find image, rescanning")
                    time.sleep(0.1)

            pyautogui.click(location.x, location.y, clicks=click_times, interval=0.2,
                            duration=0.2, button=l_or_right)
            time.sleep(0.1)


def read_tasks(commands) -> None:
    row = 1
    while row < commands.nrows:
        # get commands from the xlsx file
        command = commands.row(row)[0]
        # 1 is single left click
        if command.value == 1.0:
            # get image name
            img = commands.row(row)[1].value
            # TODO check if image is valid
            num_of_repeats = 1
            if commands.row(row)[2].ctype == 2 and commands.row(row)[2].value != 0:
                num_of_repeats = commands.row(row)[2].value
            mouse_click(1, "left", img, num_of_repeats)
            print("left clicked ", img)
        # 2 is double left click
        elif command.value == 2.0:
            img = commands.row(row)[1].value
            num_of_repeats = 1
            if commands.row(row)[2].ctype == 2 and commands.row(row)[2].value != 0:
                num_of_repeats = commands.row(row)[2].value
            mouse_click(2, "left", img, num_of_repeats)
            print("double left clicked ", img)
        # 3 is right click
        elif command.value == 3.0:
            img = commands.row(row)[1].value
            num_of_repeats = 1
            if commands.row(row)[2].ctype == 2 and commands.row(row)[2].value != 0:
                num_of_repeats = commands.row(row)[2].value
            mouse_click(1, "right", img, num_of_repeats)
            print("right clicked ", img)
            # 4 is type words
        elif command.value == 4.0:
            input_value = commands.row(row)[1].value
            # copy and paste the input values
            pyperclip.copy(input_value)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.5)
            print("inputted: ", input_value)
            # 5 is to wait
        elif command.value == 5.0:
            wait_time = commands.row(row)[1].value
            time.sleep(wait_time)
            print("waited ", wait_time, "s")
        # 6 is to scroll
        elif command.value == 6.0:
            scroll = commands.row(row)[1].value
            pyautogui.scroll(int(scroll))
            print("scrolled ", int(scroll))
        row += 1


if __name__ == '__main__':
    # Open xls file
    file = 'cmd.xlsx'
    workbook = xlrd.open_workbook(filename=file)

    # Search for the workbook
    key = int(input('Which workbook to work with?'))
    # TODO add workbook index check
    sheet = workbook.sheet_by_index(key - 1)

    # Begin program
    print("Welcome to using 'Simple Computer Task Automation'")
    key = input("Please enter an option: 1. Do task once 2. Do task 'n' times \n")
    if key == '1':
        read_tasks(sheet)
    elif key == '2':
        num_of_times = int(input("Please enter how many times"))
        for i in range(num_of_times):
            print("Beginning cycle " + str(i + 1))
            read_tasks(sheet)
            print("Cycle " + str(i + 1) + " completed!")
            time.sleep(0.1)
