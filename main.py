import pandas as pd
from tkinter.filedialog import askopenfilename
from tkinter import *
from tkinter import ttk
import tkinter.filedialog as fd
from tkinter import messagebox
print(ttk.__version__)
print(pd.__version__)

# Create an instance of tkinter frame or window
win = Tk()

# Set the geometry of tkinter frame
win.geometry("700x500")


def open_file1():
    global Database
    file1 = fd.askopenfilename(parent=win, title='Select the Excel File with RH and Temp Data')
    # print(file1)
    Database = pd.read_excel(file1, sheet_name='All', header=1, engine='openpyxl')
    # print('Database loaded')
    # Add a Label widget to display file inputted
    label1 = Label(win, text="import", font='Aerial 11')
    label1.pack(side= TOP)
    label1.config(text="Database loaded: "+file1)

def open_file2():
    global inptemplate
    file2 = fd.askopenfilename(parent=win, title='Choose a .inp file as the template')
    # print(file2)
    inptemplate = pd.read_csv(file2, header=None)
    # print('template loaded')
    # Add a Label widget to display file inputted
    label2 = Label(win, text="import", font='Aerial 11')
    label2.pack(side= TOP)
    label2.config(text="Template loaded: "+file2)

def Close():
    if messagebox.askokcancel("Quit", "Do you want to quit?"):
        win.destroy()

def year_limit(year):
    """ Determine if inp string is a valid integer (or empty) and is no more
        than MAX_DIGITS long."""
    MAX_DIGITS = 4
    try:
        int(year)  # Valid integer?
    except ValueError:
        valid = (year == '')  # Invalid unless it's just empty.
    else:
        valid = (len(year) <= MAX_DIGITS)  # OK unless it's too long.

    if not valid:
        messagebox.showinfo('Entry error',
                                'Invalid input (should be {} digits)'.format(MAX_DIGITS),
                                icon=messagebox.WARNING)
    return valid


def main():
    template = inptemplate.values.tolist()
    # print(type(inptemplate))  # line 7 (title), 20(speed),21(RH),22(Temp)
    print(type(template))
    print(template)
    # print(Database)
    df = Database

    speeds = ['15. 16. 17. 18. 19. 20. 21. 22. 23. 24. 25. 26. 27. 28. 29. 30. 31. 32. 33. 34. 35. 36. 37. 38.',
              '39. 40. 41. 42. 43. 44. 45. 46. 47. 48. 49. 50. 51. 52. 53. 54. 55. 56. 57. 58. 59. 60. 61. 62.',
              '63. 64. 65. 66. 67. 68. 69. 70. 71. 72. 73. 74. 75. 76. 77. 78. 79. 80.']

    cases = ['RH_Lowest', 'RH_Average']  # 'TEMP_Average', 'TEMP_Lowest'
    months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    hours = list(range(1, 25))  # 1-24

    for month in months:
        for case in cases:
            col_name = str(case) + "_" + str(month)
            for i in range(24):
                if df.columns.str.contains(col_name).any():
                    col = (df.columns[df.columns.str.contains(col_name)][0])  # obtained column name
                    row = i
                    print(df.loc[row, col])
                    if 'Lowest' in col_name:
                        RH = df.loc[row, col]
                        temp_col = 'TEMP_Lowest' + "_" + str(month)
                        TEMP = df.loc[row, temp_col]
                        index = 0
                        # create txt file and replace certain lines # line 7 (title<-filename), 20(speed<-constant),21(RH<-from DF),22(Temp<-from DF)
                        for index, speed in enumerate(speeds, start=1):
                            filename = str(year.get()) + "_Lowest" + "_" + str(month) + "_hr_" + str(i + 1) + "_" + str(
                                index)
                            textfile = open(str(filename) + ".inp", "w+")
                            template[7] = ['    Title ' + filename, "fake"]  # 6+1

                            template[14] = ['    CYr ' + str(year.get()), "fake"]  # 13+1
                            template[20] = ['    Emfac-Speed ' + speed, "fake"]  # 19+1
                            template[21] = ['    Emfac-RH ' + str(round(RH)) + ".", "fake"]  # 20+1
                            template[22] = ['    Emfac-Temp ' + str(round(TEMP)) + ".", "fake"]  # 21+1
                            print(filename)
                            print(template)
                            print(type(template))
                            for x in range(len(template)):
                                textfile.write(str(template[x][0]) + "\n")
                            textfile.close()
                    elif 'Average' in col_name:
                        RH = df.loc[row, col]
                        temp_col = 'TEMP_Average' + "_" + str(month)
                        TEMP = df.loc[row, temp_col]
                        # create txt file and replace certain lines # line 6 (title<-filename), 19(speed<-constant),20(RH<-from DF),21(Temp<-from DF)
                        index = 0
                        for index, speed in enumerate(speeds, start=1):
                            filename = str(year.get()) + "_Average" + "_" + str(month) + "_hr_" + str(
                                i + 1) + "_" + str(index)
                            textfile = open(str(filename) + ".inp", "w+")
                            template[7] = ['    Title ' + filename, "fake"]  # 6+1

                            template[14] = ['    CYr ' + str(year.get()), "fake"]  # 13+1
                            template[20] = ['    Emfac-Speed ' + speed, "fake"]  # 19+1
                            template[21] = ['    Emfac-RH ' + str(round(RH)) + ".", "fake"]  # 20+1
                            template[22] = ['    Emfac-Temp ' + str(round(TEMP)) + ".", "fake"]  # 21+1
                            print(filename)
                            print(template)
                            print(type(template))
                            for x in range(len(template)):
                                textfile.write(str(template[x][0]) + "\n")
                            textfile.close()
                else:
                    print('Something is wrong')


# store user input
year = StringVar()
# Enter frame
enter = ttk.Frame(win)
enter.pack(padx=40, pady=40, fill='x', expand=False)

# register year entry constrains
reg = win.register(year_limit)  # Register Entry validation function.

# year entry
year_label = ttk.Label(enter, text="Run Year:")
year_label.pack(fill=None, expand=False)

year_entry = ttk.Entry(enter, textvariable=year, validate='key', validatecommand=(reg, '%P'))  # text variable is stored in variable 'year', with constraints to ensure 4 digit number is entered
year_entry.pack(fill=None, expand=False)
year_entry.focus()

# Destroy window when click cross
win.protocol("WM_DELETE_WINDOW", Close)

# Add a Label widget
label = Label(win, text="Select the Button to Open the File", font='Aerial 11')
label.pack(pady=5)


# Add a Button Widget
ttk.Button(win, text="Select the Excel File with RH and Temp Data", command=open_file1).pack(side= TOP, pady=10)
ttk.Button(win, text="Choose a .inp file as the template", command=open_file2).pack(side= TOP, pady=20)
ttk.Button(win, text="RUN", command=main).pack(side= TOP, pady=10, ipady=20)

win.mainloop()

