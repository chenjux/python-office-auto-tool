# python-office-auto-tool
Python auto click tool can help you to simplify duplicate job such as fill in information on websites, manage workflows, and etc. 
The basic idea is that it utilize pyautogui this python package to control your mouse mainly and keyboard such as copy paste to complete your tasks. Another basic idea is that it utilize the screenshot to help itself to locate the position. For example, the internet sign screenshot can help it move the cursor to that position
python packages:
1.	pip install pyperclip
2.	pip install xlrd
3.	pip install pyautogui==0.9.50
4.	pip install opencv-python 
5.	pip install pillow
6.	pip install pandas

This software consists of two parts:
1.	Command generation: cmd_generator.py
cmd_generator.py will utilize the cmd_template command template converts the data in data.csv into a new file new_cmd.csv

2.	Execute Command: auto.py
auto.py is the command executor. If you do not want repeat tasks, you can directly revise new_cmd.csv

Operation steps, fill in the information to be filled in data.csv and save it 
3.	Before use, delete new_cmd.csv and then open cmd_generator.py and run, it will generate a new new_cmd.csv. It contains the instructions required for the next step, copy the content in new_cmd.csv, just copy the content that will be used 
4.	Open our auto.py and run it, quickly minimize the program, keep the web page maximized, there is a five-second waiting time, the software will automatically run
![image](https://user-images.githubusercontent.com/114720922/194473166-98d50b5e-d008-482e-8101-a5dd034b9af5.png)
