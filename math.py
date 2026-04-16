# *题库插入图片*
import pyautogui
import time

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 1

amount = eval(input("输入图片数："))
location = input("输入开始插入的位置：")
print("程序即将运行，请打开Excel")
time.sleep(5)  # 窗口时间
pyautogui.press("shiftleft")  # 把excel的输入法，调成英文

for i in range(amount):
    # 确定插入位置
    pyautogui.moveTo(65, 245, duration=1)
    pyautogui.click()
    if i == 0:
        pyautogui.typewrite(location)
    else:
        location = location[:1]+str(eval(location[1:])+18)
        pyautogui.typewrite(location)
    time.sleep(1)
    pyautogui.press("enter")

    # 插入图片
    pyautogui.press("altleft")
    pyautogui.press("n")
    pyautogui.press("p")
    pyautogui.press("o")
    pyautogui.press("d")
    pyautogui.moveTo(285, 236, duration=1)
    if i == 0:
        pyautogui.doubleClick()  # 进入图片文件夹
    pyautogui.click()
    if i > 0:
        for j in range(i):
            time.sleep(0.5)
            pyautogui.press("right")
    pyautogui.moveTo(734, 552, duration=1)
    pyautogui.click()

    # 调整图片大小
    pyautogui.press("altleft")
    pyautogui.hotkey('j','p')
    pyautogui.press('w')
    pyautogui.typewrite('15')
    pyautogui.press("enter")

    # 调整图片边框
    pyautogui.press("altleft")
    pyautogui.hotkey('j','p')
    pyautogui.hotkey('s','o')
    pyautogui.press('a')

print(pyautogui.alert(text="程序已完成！", title="数二"))


'''
# *数二笔记复习*
import pyautogui
import time

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 1

review_time = 2*60*60  # 复习时间：2时->秒
print("三部分：函数-微分方程、二元函数-公式表、行列式-二次型")
print("程序即将运行，请打开Word")
time.sleep(4)

start_time = time.time()
stop_time = 0
while (stop_time-start_time) <= review_time:
    pyautogui.press("down")
    time.sleep(10)  # 每行复习时间
    stop_time = time.time()

print(pyautogui.alert(text="程序已完成！", title="数二"))

# 说明
# 1、无
'''
