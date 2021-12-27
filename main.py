import pyautogui
import time
import xlrd
import pyperclip


def mouse_click(click_times: int, l_or_right: str, img: str, num_of_repeats: int) -> None:
    ## if command is to repeat forever
    if num_of_repeats == -1:
        while True:
            location = pyautogui.locateCenterOnScreen(img)
            if location is not None:
                pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2,
                                duration=0.2, button=lOrR)
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

            pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2,
                                duration=0.2, button=lOrR)
            time.sleep(0.1)



def read_tasks(img: str) -> None:
    row = 1
    while row < sheet1.nrows:
        # 取本行指令的操作类型
        command = sheet1.row(row)[0]
        if command.value == 1.0:
            # 取图片名称
            img = sheet1.row(row)[1].value
            num_of_repeats = 1
            if sheet1.row(row)[2].ctype == 2 and sheet1.row(row)[2].value != 0:
                num_of_repeats = sheet1.row(row)[2].value
            mouse_click(1, "left", img, num_of_repeats)
            print("单击左键", img)
        # 2代表双击左键
        elif command.value == 2.0:
            # 取图片名称
            img = sheet1.row(row)[1].value
            # 取重试次数
            num_of_repeats = 1
            if sheet1.row(row)[2].ctype == 2 and sheet1.row(row)[2].value != 0:
                num_of_repeats = sheet1.row(row)[2].value
            mouse_click(2, "left", img, num_of_repeats)
            print("双击左键", img)
        # 3代表右键
        elif command.value == 3.0:
            # 取图片名称
            img = sheet1.row(row)[1].value
            # 取重试次数
            num_of_repeats = 1
            if sheet1.row(row)[2].ctype == 2 and sheet1.row(row)[2].value != 0:
                num_of_repeats = sheet1.row(row)[2].value
            mouse_click(1, "right", img, num_of_repeats)
            print("右键", img)
            # 4代表输入
        elif command.value == 4.0:
            input_value = sheet1.row(row)[1].value
            pyperclip.copy(input_value)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.5)
            print("输入:", input_value)
            # 5代表等待
        elif command.value == 5.0:
            # 取图片名称
            wait_time = sheet1.row(row)[1].value
            time.sleep(wait_time)
            print("等待", wait_time, "秒")
        # 6代表滚轮
        elif command.value == 6.0:
            # 取图片名称
            scroll = sheet1.row(row)[1].value
            pyautogui.scroll(int(scroll))
            print("滚轮滑动", int(scroll), "距离")
        row += 1


if __name__ == '__main__':
    # Open xls file
    file = 'cmd.xls'
    workbook = xlrd.open_workbook(filename=file)

    # Search for the workbook
    key = int(input('Which workbook to work with?'))
    # TODO add workbook index check
    sheet = workbook.sheet_by_index(key - 1)

    # Begin program
    print("Welcome to using 'Simple Computer Task Automation'")
    key = input("Please enter an option: 1. Do task once 2. Do task 'n' times \n")
    if key == '1':
        # 循环拿出每一行指令
        read_tasks(sheet)
    elif key == '2':
        num_of_times = int(input("Please enter how many times"))
        for i in range(num_of_times):
            print("Beginning cycle " + str(i + 1))
            read_tasks(sheet)
            print("Cycle " + str(i + 1) + " completed!")
            time.sleep(0.1)
