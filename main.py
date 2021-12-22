import pandas as pd
from tkinter.filedialog import askopenfilename
from tkinter import *
from tkinter import ttk
import tkinter.filedialog as fd

# Create an instance of tkinter frame or window
win = Tk()

# Set the geometry of tkinter frame
win.geometry("700x350")


def open_file1():
    file1 = fd.askopenfilename(parent=win, title='Select the Excel File with RH and Temp Data')
    global Database
    print(file1)
    Database = pd.read_excel(file1, sheet_name='All', header=1, engine='openpyxl')
    print('Database loaded')


def open_file2():
    file2 = fd.askopenfilename(parent=win, title='Choose a .inp file as the template')
    global inptemplate
    print(file2)
    inptemplate = pd.read_csv(file2, header=0)
    print('template loaded')


# Add a Label widget
label = Label(win, text="Select the Button to Open the File", font='Aerial 11')
label.pack(pady=30)

# Add a Button Widget
ttk.Button(win, text="Select the Excel File with RH and Temp Data", command=open_file1).pack()
ttk.Button(win, text="Choose a .inp file as the template", command=open_file2).pack()

win.mainloop()

print(inptemplate)  # line 6 (title), 19(speed),20(RH),21(Temp)
print(Database)
df = Database

cases = ['RH_Lowest', 'RH_Average', 'TEMP_Average', 'TEMP_Lowest']
months = ['Jan', 'Feb']

for month in months:
    for case in cases:
        col_name = str(case)+"_"+str(month)
        if df.columns.str.contains(col_name).any():
            print(df.columns[df.columns.str.contains(col_name)][0]) #hour 1-24-> row 0-23
        else:
            print('No')



#extract value

#create txt file and replace certain lines # line 6 (title<-filename), 19(speed<-constant),20(RH<-from DF),21(Temp<-from DF)