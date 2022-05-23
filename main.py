from cmath import log
from tkinter import ttk, filedialog
from tkinter.filedialog import askopenfile
from convert import *
from tkinter import *
import os

def select_file():
    global f
    global t
    global file_name
    file = filedialog.askopenfile(mode='r', filetypes=[('Excel Files', '*.xlsx')])
    result.config(text="")
    error_label.config(text="")
    if file:
        filepath = os.path.abspath(file.name)
        file_label.config(text = f"선택된 파일 : {str(filepath)}")
        file_name = str(filepath).split('/')[-1].strip('.xlsx')
        file_name = str(file_name).split('\\')[-1].strip('.xlsx')
        f = str(filepath)
        if t:
            t = selected + "/" + file_name + "변환" + ".xlsx"
            path_label.config(text = f"선택된 위치 : {selected}/{file_name}-변환.xlsx")

def select_folder():
    global t
    global selected
    selected = filedialog.askdirectory()
    t = selected + "/" + file_name + "변환" + ".xlsx"
    path_label.config(text = f"선택된 위치 : {selected}/{file_name}-변환.xlsx")
    result.config(text="")
    error_label.config(text="")

def do_convert():
    global f
    global t
    result.config(text="")
    error_label.config(text="")
    if not f:
        return result.config(text="변환할 파일을 선택해주세요.")
    if not t:
        return result.config(text="저장할 위치를 선택해주세요.")
    
    
    converter(f, t)
    return result.config(text="변환이 완료되었습니다.")
    # except Exception as log:
    #     result.config(text="알 수 없는 오류가 발생했습니다. 아래 로그를 캡처하여 제작자에게 문의해주세요.")
    #     error_label.config(text=log)

f = str()
t = str()

root = Tk()
root.title("Inventory Manager") # GUI 제목 지정
root.geometry("640x480") # 가로 x 세로 크기 지정
# root.resizable(False, False) # x너비, y높이 값 변경 불가
file_label = Label(root, text="선택된 파일 : ")
path_label = Label(root, text="선택된 위치 : ")
file_name = ""
selected = ""
result = Label(root, text="")
error_label = Label(root, text="")

if __name__ == '__main__':

    Label(root, text="변환할 파일을 선택하세요").pack(pady=2)
    ttk.Button(root, text="찾아보기", command=select_file).pack(pady=2)
    file_label.pack(pady=(2,30))

    Label(root, text="저장할 위치를 선택하세요").pack(pady=2)
    ttk.Button(root, text="찾아보기", command=select_folder).pack(pady=2)
    path_label.pack(pady=(2,30))

    ttk.Button(root, text="변환하기", command=do_convert).pack(pady=2)
    result.pack(pady=(2,30))

    error_label.pack(pady=(2,20))

    root.mainloop()