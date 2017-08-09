import sys
from tkinter import *
from tkinter import filedialog
import re




def myOpen():
    myOpen = filedialog.askopenfile()
    filename=re.findall("\'.*?\'",str(myOpen))
    return filename[0].replace("\'","")


myApp = Tk()
# create a string variable
ment = StringVar()

# set the size of window
myApp.geometry('450x450+200+200')

myApp.title('location geocode')

mLabel = Label(myApp, text='choose excel file').pack()

mButton = Button(myApp, text='choose', command=myOpen).pack()


print mButton

myApp.mainloop()