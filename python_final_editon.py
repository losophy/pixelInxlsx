#coding=utf-8
from win32com import client
import numpy as np
from PIL import Image
import sys
#coding = utf-8
import tkinter as tk
from tkinter import filedialog
import os
import  time
from tkinter import messagebox
def jpg_image_to_array(image):
    """
    Loads JPEG image into 3D Numpy array of shape 
    (width, height, channels)
    """
    im_arr = np.fromstring(image.tobytes(), dtype=np.uint8)
    im_arr = im_arr.reshape((image.size[1], image.size[0], 3))
    return im_arr

def rgbToInt(rgb):
    colorInt = rgb[0] + (rgb[1] * 256) + (rgb[2] * 256 * 256)
    return colorInt

def get_txt():
    my_filetypes = [('all files', '.*'), ('text files', '.txt')]
    application_window = tk.Tk()
    answer = messagebox.askyesno("关闭Excel", "是否关闭Excel?", parent=application_window)
    if not answer:
        "如果不关闭Excel,直接退出程序"
        sys.exit()
    file_name_abs = filedialog.askopenfilename(parent=application_window,
                                        initialdir=os.getcwd(),
                                        title="Please select a file:",
                                        filetypes=my_filetypes)
    path =os.path.join(file_name_abs)
    application_window.destroy()
    return path
start = time.clock()
path = get_txt()
f = Image.open(path)
f1 = f.convert("P").convert("RGB")
arr = jpg_image_to_array(f1)
arr2 = np.apply_along_axis(rgbToInt,2,arr)

#------------call the excel com -------------
width = arr2.shape[1]
height = arr2.shape[0]
print(width,height)
E = client.Dispatch("Excel.Application")
E.visible = False
E.DisplayAlerts = False
WB = E.Workbooks.Add()
wt = WB.Worksheets("Sheet1")
rng = wt.Range(wt.cells(1,1),wt.cells(height,width))
rng.Columns.ColumnWidth = 1.75
rng.Rows.RowHeight = 14.25
i = 0
size = height * width
print("欢迎使用本软件，tc竭诚为您服务！","\n"
      "在开始前，请关闭所有的Excel！")
for j in rng:
        j.Interior.Color = int(arr2[j.Row-1,j.Column-1])
        i = i + 1
        p = i / size
        now = time.clock()
        inteval = int(now - start)
        if inteval % 10 ==0:
            print("已经完成" + str(p),"请继续等待并不要打开Excel")

final_path = path.split(".")[0] + ".xlsx"
WB.SaveAs(Filename = final_path.replace("/","\\"))
WB.Close(SaveChanges=0)
del E
