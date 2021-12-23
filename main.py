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

template = inptemplate.values.tolist()
# print(type(inptemplate))  # line 6 (title), 19(speed),20(RH),21(Temp)
# print(type(template))
# print(template)
# print(Database)
df = Database

speeds = ['15. 16. 17. 18. 19. 20. 21. 22. 23. 24. 25. 26. 27. 28. 29. 30. 31. 32. 33. 34. 35. 36. 37. 38.',
          '39. 40. 41. 42. 43. 44. 45. 46. 47. 48. 49. 50. 51. 52. 53. 54. 55. 56. 57. 58. 59. 60. 61. 62.',
          '63. 64. 65. 66. 67. 68. 69. 70. 71. 72. 73. 74. 75. 76. 77. 78. 79. 80.']

cases = ['RH_Lowest', 'RH_Average'] #'TEMP_Average', 'TEMP_Lowest'
months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
hours = list(range(1,25))   #1-24

for month in months:
    for case in cases:
        col_name = str(case)+"_"+str(month)
        for i in range(24):
            if df.columns.str.contains(col_name).any():
                col = (df.columns[df.columns.str.contains(col_name)][0]) #obtained column name
                row = i
                print(df.loc[row, col])
                if 'Lowest' in col_name:
                    RH = df.loc[row, col]
                    temp_col = 'TEMP_Lowest'+ "_" + str(month)
                    TEMP = df.loc[row, temp_col]
                    index = 0
                    # create txt file and replace certain lines # line 6 (title<-filename), 19(speed<-constant),20(RH<-from DF),21(Temp<-from DF)
                    for index, speed in enumerate(speeds, start=1):
                        filename = "2023_Lowest" + "_" + str(month) + "_hr_" + str(i+1) + "_" + str(index)
                        textfile = open(str(filename), "w+")
                        template[6] = ['    Title ' + filename, "fake"]  # 6
                        template[19] = ['    Emfac-Speed ' + speed, "fake"]  # 19
                        template[20] = ['Emfac - RH ' + str(int(RH)), "fake"]  # 20
                        template[21] = ['Emfac - Temp ' + str(int(TEMP)), "fake"]  # 21
                        print(filename)
                        print(template)
                        print(type(template))
                        for x in range(22):
                            textfile.write(str(template[x][0]) + "\n")
                        textfile.close()
                elif 'Average' in col_name:
                    RH = df.loc[row, col]
                    temp_col = 'TEMP_Average'+ "_" + str(month)
                    TEMP = df.loc[row, temp_col]
                    # create txt file and replace certain lines # line 6 (title<-filename), 19(speed<-constant),20(RH<-from DF),21(Temp<-from DF)
                    index = 0
                    for index, speed in enumerate(speeds, start=1):
                        filename = "2023_Average" + "_" + str(month) + "_hr_" + str(i+1) + "_" + str(index)
                        textfile = open(str(filename), "w+")
                        template[6] = ['    Title ' + filename, "fake"]  # 6
                        template[19] = ['    Emfac-Speed ' + speed, "fake"]  # 19
                        template[20] = ['Emfac - RH ' + str(int(RH)), "fake"]  # 20
                        template[21] = ['Emfac - Temp ' + str(int(TEMP)), "fake"]  # 21
                        print(filename)
                        print(template)
                        print(type(template))
                        for x in range(22):
                            textfile.write(str(template[x][0]) + "\n")
                        textfile.close()
            else:
                print('Something is wrong')
