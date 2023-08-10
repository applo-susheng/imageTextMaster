# -*- coding: utf-8 -*-
import pytesseract
import os
from pdf2image import convert_from_path
import tkinter as tk
from PIL import Image, ImageEnhance
from tkinter import ttk, filedialog
import threading
import pandas as pd
from glob import glob
from datetime import datetime
result_dir=""
file_paths=None
excel_file=None
folder_path=None
log=None
pytesseract.pytesseract.tesseract_cmd=r".\tesseract\tesseract.exe"
def handle_selected_option():
    global selected_option
    selected_option = option.get()  # 获取选择的选项
def handle_execution():
    global selected_option
    selected_option = option.get()  # 获取选择的选项

    # 根据选择的选项切换显示的小部件
    if selected_option == "excel":
        excel_frame.pack()
        paste_frame.pack_forget()
        folder_frame.pack_forget()
    elif selected_option == "paste":
        excel_frame.pack_forget()
        paste_frame.pack()
        folder_frame.pack_forget()
    elif selected_option == "folder":
        excel_frame.pack_forget()
        paste_frame.pack_forget()
        folder_frame.pack()
#保存位置的选择文件夹，并将结果赋值给result_dir
def choose_save_directory():
    global result_dir
    save_dir = filedialog.askdirectory()  # 唤醒系统文件夹选择保存路径
    if save_dir:
        entry_save_location.delete(0, tk.END)  # 清空文本框内容
        entry_save_location.insert(0, save_dir)  # 将选择的文件夹路径插入到文本框中

# 判断文件是否存在，并返回扩展名
def get_file_type(file_path):
    # 判断文件是否存在
    if not os.path.exists(file_path):
        return "文件不存在"

    # 使用os模块获取文件扩展名
    file_extension = os.path.splitext(file_path)[1]
    return file_extension
#获取选项信息，并将其赋值给selected_option
def insert_red_text1(text):
    # 设置标签样式（红色文本）
    paste_print.tag_config("red", foreground="red")

    # 插入带有红色标签的文本
    paste_print.insert("end", text, "red")

    # 刷新文本框并滚动到最后一行
    paste_print.update()
    paste_print.see("end")
def insert_red_text2(text):
    # 设置标签样式（红色文本）
    excel_print.tag_config("red", foreground="red")

    # 插入带有红色标签的文本
    excel_print.insert("end", text, "red")

    # 刷新文本框并滚动到最后一行
    excel_print.update()
    excel_print.see("end")
def insert_red_text3(text):
    # 设置标签样式（红色文本）
    folder_print.tag_config("red", foreground="red")

    # 插入带有红色标签的文本
    folder_print.insert("end", text, "red")

    # 刷新文本框并滚动到最后一行
    folder_print.update()
    folder_print.see("end")

#进度条
#粘贴窗口
def paste_text(event=None):
    global file_paths
    global log
    text = text_pad.get("1.0", tk.END).strip()  # 获取输入框中的文本
    text_pad.tag_configure("normal", foreground="black")
    text_pad.delete("1.0", tk.END)
    text_pad.insert(tk.END, text, "normal")
    # print(text)
    # # 将输入的行转换为列表
    # text = r'{}'.format(text)
    lines = text.strip().split("\n")
    # file_paths = [r'{}'.format() for line in lines]
    # print(lines)
    file_paths = [line.strip().strip('"') for line in lines]
    file_paths = [os.path.normpath(line) for line in file_paths]
    log=pd.DataFrame(columns=["输入文件路径","输出文件路径"])
    log['输入文件路径']=file_paths

def on_click(event):
    if text_pad.get("1.0", "end-1c") == placeholder_text:
        text_pad.delete("1.0", tk.END)
        text_pad.configure(fg="black")
def shift1(file_paths):  #将pdf和图片转为txt
    global log
    count=0

    for file_path in file_paths:
        # print(file_path)
        try:
            txt_file = os.path.join(result_dir, file_path.split('\\')[-1] + '.txt')
            file_extension = get_file_type(file_path)
            if file_extension == ".jpg" or file_extension == ".png":
                paste_label.config(text=f"进度: {0:.1f}%")  # 更新标签的文本内容
                root.update()  # 更新界面
                root.after(1)  # 等待1毫秒
                image = Image.open(file_path)
                enhancer = ImageEnhance.Contrast(image)
                image = enhancer.enhance(2.0)
                text = pytesseract.image_to_string(image, lang='chi_sim')
                # 将提取的文本添加到列表中
                txt_f = open(txt_file, 'w')
                txt_f.write(text)
                #将文件路径写在log里
                log.iloc[count,1]=txt_file
                count+=1
                paste_print.insert("end", f"{count}\n")
                paste_print.update()  # 刷新文本框
                paste_print.see("end")
                # percentage = j / page_count * 100  # 根据实际进度计算百分比
                paste_label.config(text=f"进度: {100:.1f}%")  # 更新标签的文本内容
                root.update()  # 更新界面
                root.after(1)  # 等待1毫秒

            elif file_extension == ".pdf":
                with open(file_path, 'rb') as f:
                    images = convert_from_path(file_path)
                    page_count = len(images)
                    text_pages = []
                    # 遍历图像列表并进行处理
                    for i, image in enumerate(images):

                        image_path = f'./temp/page_{i + 1}.jpg'

                        image.save(image_path, 'JPEG')
                        # 使用 pytesseract 对图像进行 OCR
                        text = pytesseract.image_to_string(image_path, lang='chi_sim')
                        # 将提取的文本添加到列表中
                        j=i+1
                        text_pages.append(text)
                        percentage = j / page_count * 100  # 根据实际进度计算百分比
                        paste_label.config(text=f"进度: {percentage:.1f}%")  # 更新标签的文本内容
                        root.update()  # 更新界面
                        root.after(1)  # 等待1毫秒
                    # 将所有文本合并为一个字符串
                    full_text = '\n'.join(text_pages)
                    txt_f = open(txt_file, 'w')
                    txt_f.write(full_text)
                    #添加log文件

                    log.iloc[count, 1] = txt_file
                    count += 1
                    #打印窗口
                    paste_print.insert("end", f"{count}\n")
                    paste_print.update()  # 刷新文本框
                    paste_print.see("end")
            else:

                log.iloc[count,1]=file_extension
                count+=1
                error_text = f"{file_extension}{file_path}\n"
                insert_red_text1(error_text)#将pdf和图片转为txt
        except:
            a = "发生错误:"
            log.iloc[count, 1] = a
            count += 1
            error_text = f"{a}{file_path}\n"
            insert_red_text1(error_text)  # 将pdf和图片转为txt
    g="...执行完成..."
    paste_print.insert("end", f"{g}\n")
    paste_print.update()  # 刷新文本框
    paste_print.see("end")


def run1():
    global result_dir
    result_dir = entry_save_location.get()  # 获取文本框中的内容，并赋值给result变量
    paste_text()
    # start_processing()
    # shift1(file_paths)
    thread = threading.Thread(target=shift1, args=(file_paths,))
    thread.start()



#excel文件窗口
def choose_excel_directory():
    global excel_file
    save_dir =filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
    if save_dir:
        entry_excel_location.delete(0, tk.END)  # 清空文本框内容
        entry_excel_location.insert(0, save_dir)  # 将选择的文件夹路径插入到文本框中
def Processing_Path(excel_file):
    global file_paths
    global log
    excel_file=excel_file.strip().strip('"')

    df = pd.read_excel(excel_file,sheet_name=0)
    column_1 = df.iloc[:, 0].tolist()
    log = pd.DataFrame(columns=["输入文件路径", "输出文件路径"])
    log['输入文件路径'] = column_1

    file_paths = [line.strip().strip('"') for line in column_1]
    file_paths = [os.path.normpath(line) for line in file_paths]

def shift2(file_paths):  #将pdf和图片转为txt
    global log
    file_extension=None
    count=0
    for file_path in file_paths:
        # print(file_path)
        #打印窗口
        try:
            txt_file = os.path.join(result_dir, file_path.split('\\')[-1] + '.txt')
            file_extension = get_file_type(file_path)
            if file_extension == ".jpg" or file_extension == ".png":
                excel_label.config(text=f"进度: {0:.1f}%")  # 更新标签的文本内容
                root.update()  # 更新界面
                root.after(1)  # 等待1毫秒
                image = Image.open(file_path)
                enhancer = ImageEnhance.Contrast(image)
                image = enhancer.enhance(2.0)
                text = pytesseract.image_to_string(image, lang='chi_sim')
                txt_f = open(txt_file, 'w')
                txt_f.write(text)

                #将结果添加到log
                log.iloc[count, 1] = txt_file
                count += 1
                #打印
                excel_print.insert("end", f"{count}\n")
                excel_print.update()  # 刷新文本框
                excel_print.see("end")
                # percentage = j / page_count * 100  # 根据实际进度计算百分比
                excel_label.config(text=f"进度: {100:.1f}%")  # 更新标签的文本内容
                root.update()  # 更新界面
                root.after(1)  # 等待1毫秒

            elif file_extension == ".pdf":
                with open(file_path, 'rb') as f:
                    images = convert_from_path(file_path)
                    page_count = len(images)
                    text_pages = []
                    # 遍历图像列表并进行处理
                    for i, image in enumerate(images):
                        #进度条

                        # start_processing(0,page_count)

                        image_path = f'./temp/page_{i + 1}.jpg'

                        image.save(image_path, 'JPEG')
                        # 使用 pytesseract 对图像进行 OCR
                        text = pytesseract.image_to_string(image_path, lang='chi_sim')
                        # 将提取的文本添加到列表中
                        j=i+1
                        text_pages.append(text)
                        percentage = j / page_count * 100  # 根据实际进度计算百分比
                        excel_label.config(text=f"进度: {percentage:.1f}%")  # 更新标签的文本内容
                        root.update()  # 更新界面
                        root.after(1)  # 等待1毫秒
                    # 将所有文本合并为一个字符串
                    full_text = '\n'.join(text_pages)
                    txt_f = open(txt_file, 'w')
                    txt_f.write(full_text)
                    log.iloc[count,1]=txt_file
                    count+=1
                    #打印窗口
                    excel_print.insert("end", f"{count}\n")
                    excel_print.update()  # 刷新文本框
                    excel_print.see("end")
            else:
                log.iloc[count, 1] = file_extension
                count += 1
                error_text = f"{file_extension}{file_path}\n"
                insert_red_text2(error_text)#将pdf和图片转为txt
        except:
            a="发生错误"
            log.iloc[count, 1] = a
            count += 1
            error_text = f"{a}{file_path}\n"
            insert_red_text2(error_text)  # 将pdf和图片转为txt

    g="...执行完成..."
    excel_print.insert("end", f"{g}\n")
    excel_print.update()  # 刷新文本框
    excel_print.see("end")
def run2():
    global result_dir
    global excel_file
    result_dir = entry_save_location.get()  # 获取文本框中的内容，并赋值给result变量
    excel_file = entry_excel_location.get()  # 获取文本框中的内容
    Processing_Path(excel_file)

    # start_processing()
    # shift(file_paths)
    thread = threading.Thread(target=shift2, args=(file_paths,))
    thread.start()
#文件夹窗口
def choose_folder_directory():
    global folder_path
    save_dir = filedialog.askdirectory()  # 唤醒系统文件夹选择保存路径
    if save_dir:
        entry_folder_location.delete(0, tk.END)  # 清空文本框内容
        entry_folder_location.insert(0, save_dir)  # 将选择的文件夹路径插入到文本框中

def deal_folder_path(folder_path):
    global file_paths
    global log
    # image_files = glob('./test_images/*.*')
    file_paths = glob(folder_path + '/*.*')
    # print(file_paths)
    file_paths = [line.strip().strip('"') for line in file_paths]
    file_paths = [os.path.normpath(line) for line in file_paths]
    log=pd.DataFrame(columns=["输入文件路径","输出文件路径"])
    log['输入文件路径']=file_paths

def shift3(file_paths):  #将pdf和图片转为txt
    global log
    file_extension=None
    count=0
    count=0
    for file_path in file_paths:
        # print(file_path)
        #打印窗口
        try:
            txt_file = os.path.join(result_dir, file_path.split('\\')[-1] + '.txt')
            file_extension = get_file_type(file_path)
            if file_extension == ".jpg" or file_extension == ".png":
                folder_label.config(text=f"进度: {0:.1f}%")  # 更新标签的文本内容
                root.update()  # 更新界面
                root.after(1)  # 等待1毫秒
                text_pages = []
                image = Image.open(file_path)
                enhancer = ImageEnhance.Contrast(image)
                image = enhancer.enhance(2.0)
                text = pytesseract.image_to_string(image, lang='chi_sim')
                # 将提取的文本添加到列表中
                txt_f = open(txt_file, 'w')
                txt_f.write(text)
                log.iloc[count, 1] = txt_file
                count += 1
                #打印
                folder_print.insert("end", f"{count}\n")
                folder_print.update()  # 刷新文本框
                folder_print.see("end")
                # percentage = j / page_count * 100  # 根据实际进度计算百分比
                folder_label.config(text=f"进度: {100:.1f}%")  # 更新标签的文本内容
                root.update()  # 更新界面
                root.after(1)  # 等待1毫秒

            elif file_extension == ".pdf":

                with open(file_path, 'rb') as f:
                    images = convert_from_path(file_path)
                    page_count = len(images)
                    text_pages = []
                    # 遍历图像列表并进行处理
                    for i, image in enumerate(images):
                        #进度条

                        # start_processing(0,page_count)

                        image_path = f'./temp/page_{i + 1}.jpg'

                        image.save(image_path, 'JPEG')
                        # 使用 pytesseract 对图像进行 OCR
                        text = pytesseract.image_to_string(image_path, lang='chi_sim')
                        # 将提取的文本添加到列表中
                        j=i+1
                        text_pages.append(text)
                        percentage = j / page_count * 100  # 根据实际进度计算百分比
                        folder_label.config(text=f"进度: {percentage:.1f}%")  # 更新标签的文本内容
                        root.update()  # 更新界面
                        root.after(1)  # 等待1毫秒
                    # 将所有文本合并为一个字符串
                    full_text = '\n'.join(text_pages)
                    txt_f = open(txt_file, 'w')
                    txt_f.write(full_text)
                    log.iloc[count, 1] = txt_file
                    count += 1
                    #打印窗口
                    folder_print.insert("end", f"{count}\n")
                    folder_print.update()  # 刷新文本框
                    folder_print.see("end")
            else:
                log.iloc[count, 1] = file_extension
                count += 1
                error_text = f"{file_extension}{file_path}\n"
                insert_red_text3(error_text)#将pdf和图片转为txt
        except:
            a="发生错误:"
            log.iloc[count, 1] = a
            count += 1
            error_text = f"{a}{file_path}\n"
            insert_red_text3(error_text)  # 将pdf和图片转为txt

    g="...执行完成..."
    folder_print.insert("end", f"{g}\n")
    folder_print.update()  # 刷新文本框
    folder_print.see("end")
def run3():
    global result_dir
    global folder_path
    result_dir = entry_save_location.get()  # 获取文本框中的内容，并赋值给result变量
    folder_path = entry_folder_location.get()  # 获取文本框中的内容，并赋值给result变量
    deal_folder_path(folder_path)
    thread = threading.Thread(target=shift3, args=(file_paths,))
    thread.start()

root = tk.Tk()
root.title('Image To TextMaster')
root.geometry('800x400+400+100')

left_frame = tk.Frame(root)
left_frame.pack(side=tk.LEFT, anchor=tk.N, padx=5, pady=5)

way_frame = tk.LabelFrame(left_frame, text='选择输入文件路径的方式', padx=5, pady=5)
way_frame.pack(fill=tk.X, pady=(0, 5))

option = tk.StringVar()

r1 = tk.Radiobutton(way_frame, text="excel文件", variable=option, value="excel", command=handle_execution)
r1.pack(anchor=tk.W)

r2 = tk.Radiobutton(way_frame, text="直接粘贴文件路径", variable=option, value="paste", command=handle_execution)
r2.pack(anchor=tk.W)

r3 = tk.Radiobutton(way_frame, text="所在文件夹", variable=option, value="folder", command=handle_execution)
r3.pack(anchor=tk.W)

#保存位置
save_location_frame = tk.LabelFrame(left_frame, text='保存位置', padx=5, pady=5)
save_location_frame.pack(side=tk.TOP, anchor=tk.N, fill=tk.X)

entry_save_location = ttk.Entry(save_location_frame)
entry_save_location.pack(side=tk.LEFT, fill=tk.X, expand=True,padx=(0, 5), ipadx=5)

button_choose_directory = ttk.Button(save_location_frame, text='选择文件夹', command=choose_save_directory,width=10)
button_choose_directory.pack(side=tk.RIGHT)

right_frame = tk.Frame(root)
right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

#excel小部件窗口
excel_frame = tk.Frame(right_frame)
tk.Label(excel_frame, text='Excel文件选项').pack()
save_location_frame = tk.LabelFrame(excel_frame, text='excel位置', padx=5, pady=5)
save_location_frame.pack(side=tk.TOP, anchor=tk.W, fill=tk.X)

entry_excel_location = ttk.Entry(save_location_frame)
entry_excel_location.pack(side=tk.LEFT, fill=tk.X, expand=True,padx=(0, 5), ipadx=5)

button_choose_directory = ttk.Button(save_location_frame, text='选择excel文件', command=choose_excel_directory,width=15)
button_choose_directory.pack(side=tk.RIGHT,fill=tk.X)
#进度条
excel_label = tk.Label(excel_frame, text="进度: 0%")
excel_label.pack(pady=10)


#打印窗口
excel_print = tk.Text(excel_frame, width=58, height=6)
excel_print.pack(side=tk.LEFT)
send_button = tk.Button(excel_frame, text='执行', width=5,command=run2)
send_button.pack(side=tk.RIGHT, fill=tk.Y)



#粘贴窗口
paste_frame = tk.Frame(right_frame)
# 直接粘贴文件路径选项的小部件
tk.Label(paste_frame, text='直接粘贴文件路径选项').pack()

info_frame = tk.Frame(paste_frame)
info_frame.pack()
tk.Label(info_frame, text='粘贴文件路径').pack(anchor=tk.W)
#
text_pad = tk.Text(info_frame, width=62,height=15)
text_pad.pack(side=tk.LEFT, fill=tk.X)

placeholder_text = "在此处粘贴文本... "
text_pad.insert(tk.END, placeholder_text)
text_pad.tag_configure("placeholder", foreground="gray")
text_pad.tag_add("placeholder", "1.0", "end-1c")
text_pad.bind("<Button-1>", on_click)
text_pad.bind("<FocusIn>", on_click)

send_text_bar = tk.Scrollbar(info_frame)
send_text_bar.pack(side=tk.RIGHT, fill=tk.Y)
#进度条
paste_label = tk.Label(paste_frame, text="进度: 0%")
paste_label.pack(pady=10)
#打印窗口
paste_print = tk.Text(paste_frame, width=58, height=6)
paste_print.pack(side=tk.LEFT)
send_button = tk.Button(paste_frame, text='执行', width=5,command=run1)
send_button.pack(side=tk.RIGHT, fill=tk.Y)

#文件夹窗口
folder_frame = tk.Frame(right_frame)
tk.Label(folder_frame, text='所在文件夹选项').pack()
save_location_frame = tk.LabelFrame(folder_frame, text='所在文件夹路径', padx=5, pady=5)
save_location_frame.pack(side=tk.TOP, anchor=tk.W, fill=tk.X)

entry_folder_location = ttk.Entry(save_location_frame)
entry_folder_location.pack(side=tk.LEFT, fill=tk.X, expand=True,padx=(0, 5), ipadx=5)

button_choose_directory = ttk.Button(save_location_frame, text='选择文件夹', command=choose_folder_directory,width=15)
button_choose_directory.pack(side=tk.RIGHT,fill=tk.X)
#进度条
folder_label = tk.Label(folder_frame, text="进度: 0%")
folder_label.pack(pady=10)

#打印窗口
folder_print = tk.Text(folder_frame, width=58, height=6)
folder_print.pack(side=tk.LEFT)
send_button = tk.Button(folder_frame, text='执行', width=5,command=run3)
send_button.pack(side=tk.RIGHT, fill=tk.Y)

root.mainloop()
#输出log文件
current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
log_location = './log/'
log.to_excel(f'{log_location}{current_time}.xlsx', index=False)
temp_folder_path = r"./temp"  # 替换为你要删除的文件夹路径

# 获取文件夹下的所有文件路径
# file_paths = glob.glob(os.path.join(folder_path, "*"))
temp_file_paths = glob(temp_folder_path + '/*.*')
# 删除文件夹下的所有文件
for file_path in temp_file_paths:
    os.remove(file_path)