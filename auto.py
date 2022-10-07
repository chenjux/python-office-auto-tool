import pyautogui
import time
import xlrd
import pyperclip


# Define how to use your mouse

# pyautogui other usage https://blog.csdn.net/qingfengxd1/article/details/108270159

def mouseClick(clickTimes, lOrR, img, reTry):
    if reTry == 1:
        while True:
            location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
            if location is not None:
                pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2, button=lOrR)
                break
            print("No matching image found, try again in 0.1 seconds")
            time.sleep(0.1)
    elif reTry == -1:
        while True:
            location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
            if location is not None:
                pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2, button=lOrR)
            time.sleep(0.1)
    elif reTry > 1:
        i = 1
        while i < reTry + 1:
            location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
            if location is not None:
                pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2, button=lOrR)
                print("Repeat")
                i += 1
            time.sleep(0.1)


# data check
# cmdType.value 1.0 left click 2.0 left double click 3.0 right click 4.0 enter 5.0 wait 6.0 scroll wheel
# ctype empty: 0
# string: 1
# number: 2
# Date: 3
# boolean: 4
# error: 5
def dataCheck(sheet1):
    checkCmd = True
    # 行数检查
    if sheet1.nrows < 2:
        print("No data")
        checkCmd = False
    # 每行数据检查
    i = 1
    while i < sheet1.nrows:
        # 第1列 操作类型检查
        cmdType = sheet1.row(i)[0]
        if cmdType.ctype != 2 or (cmdType.value != 1.0 and cmdType.value != 2.0 and cmdType.value != 3.0
                                  and cmdType.value != 4.0 and cmdType.value != 5.0 and cmdType.value != 6.0
                                  and cmdType.value != 7.0 and cmdType.value != 8.0):
            print(' ', i + 1, "row, column 1 data error")
            checkCmd = False
        # 第2列 内容检查
        cmdValue = sheet1.row(i)[1]
        # 读图点击类型指令，内容必须为字符串类型
        if cmdType.value == 1.0 or cmdType.value == 2.0 or cmdType.value == 3.0:
            if cmdValue.ctype != 1:
                print('第', i + 1, "row, column 2 data error")
                checkCmd = False
        # 输入类型，内容不能为空
        if cmdType.value == 4.0:
            if cmdValue.ctype == 0:
                print('第', i + 1, "row, column 2 data error")
                checkCmd = False
        # 等待类型，内容必须为数字
        if cmdType.value == 5.0:
            if cmdValue.ctype != 2:
                print('第', i + 1, "row, column 2 data error")
                checkCmd = False
        # 滚轮事件，内容必须为数字
        if cmdType.value == 6.0:
            if cmdValue.ctype != 2:
                print('第', i + 1, "row, column 2 data error")
                checkCmd = False
        # 移动事件，默认为向右移动，内容为数字
        if cmdType.value == 7.0:
            if cmdValue.ctype != 2:
                print('第', i + 1, "row, column 2 data error")
                checkCmd = False
        # 移动事件，默认为向右移动，内容为数字
        if cmdType.value == 8.0:
            if cmdValue.ctype != 2:
                print('第', i + 1, "row, column 2 data error")
                checkCmd = False
        i += 1
    return checkCmd


# task
def mainWork(img):
    i = 1
    while i < sheet1.nrows:
        # Get the operation type of this line of instructions
        cmdType = sheet1.row(i)[0]
        if cmdType.value == 1.0:
            # get image name
            img = sheet1.row(i)[1].value
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(1, "left", img, reTry)
            print("left click", img)
        # 2 means double-clicking the left button
        elif cmdType.value == 2.0:
            # get image name
            img = sheet1.row(i)[1].value
            # Get the number of retries
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(2, "left", img, reTry)
            print("Double left  click", img)
        # 3 means right click
        elif cmdType.value == 3.0:
            # get image name
            img = sheet1.row(i)[1].value
            # Get the number of retries
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(1, "right", img, reTry)
            print("Right click", img)
            # 4 means Enter
        elif cmdType.value == 4.0:
            inputValue = sheet1.row(i)[1].value
            print('input value', inputValue)
            print('input type', type(inputValue))

            if type(inputValue) == float:
                inputValue = str(inputValue)[:-2]

            pyperclip.copy(inputValue)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.5)
            print("Enter:", inputValue)
            # 5 means wait
        elif cmdType.value == 5.0:
            # Get image's name
            waitTime = sheet1.row(i)[1].value
            time.sleep(waitTime)
            print("Wait", waitTime, "second")
        # 6 control scroll up/down
        elif cmdType.value == 6.0:
            scroll = sheet1.row(i)[1].value
            pyautogui.scroll(int(scroll))
            print("Scroll", int(scroll), " distance")
        # 7 means move to right, add an negative number to move left
        elif cmdType.value == 7.0:
            move_to_right = sheet1.row(i)[1].value
            pyautogui.moveRel(int(move_to_right), 0, duration=0.1)
            print("move to right", int(move_to_right), " distance")
            pyautogui.click()

        # 8 means move to downwards;
        elif cmdType.value == 8.0:
            # 取图片名称
            move_to_down = sheet1.row(i)[1].value
            pyautogui.moveRel(0, int(move_to_down), duration=0.1)
            print("Move Downward", int(move_to_down), " distance")
            pyautogui.click()

        i += 1


if __name__ == '__main__':
    file = 'cmd.xls'
    # open file
    wb = xlrd.open_workbook(filename=file)

    # 通过索引获取表格sheet页
    sheet1 = wb.sheet_by_index(0)
    print('Welcome')
    # data integration check
    checkCmd = dataCheck(sheet1)
    if checkCmd:

        mainWork(sheet1)

    else:
        print('Enter error')
