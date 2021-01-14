import datetime
import os
import subprocess
import sys
import tkinter as tk
from tkinter import font
import tkinter.font

from PIL import Image, ImageTk

import Popbox



window = None
frameEntry = None
popboxUrlEntry = None

EANField = None
PPGField = None
SalesOrgField = None
BaseSKUField = None
projectDescField = None


class State:
    scrapedData = None
    projectDescription = None
    EANValue = None
    PPGValue = None
    SalesOrgValue = None
    BaseSKUValue = None
    counter = 0


def startSession():
    os.system('Main.py')


def initWindow():
    window = tk.Tk()
    window.rowconfigure(100, minsize=100)
    window.geometry("600x800")
    window.title("SKU Creator")
    window.config(background='ghost white')
    window.iconbitmap(r"C:\Users\Paula Matosek\Popbox\toothbrush.ico")
    return window



def replace_line():
    lines = open('sessionIdAndProduct.txt', 'r').readlines()
    lines[2] = str(popboxUrlEntry.get())
    out = open('sessionIdAndProduct.txt', 'w')
    out.writelines(lines)
    out.close()


def append_line():
    with open('sessionIdAndProduct.txt', 'a') as file:
        file.write('\n' + str(popboxUrlEntry.get()))


def scrapeData():
    file = 'sessionIdAndProduct.txt'
    lines = open('sessionIdAndProduct.txt', 'r').readlines()
    print(str(popboxUrlEntry.get()))
    if (len((lines)) == 2):
        append_line()
    else:
        replace_line()
    runPopboxScrapping()


def runPopboxScrapping():
    data = Popbox.run()
    print(data)
    State.scrapedData = data
    State.projectDescription = data.getProjectDescription()
    State.EANValue = data.EANStrategy
    State.PPGValue = data.PPG
    State.SalesOrgValue = data.SalesOrg
    State.BaseSKUValue = data.BaseSKU


def runSap():
    p = subprocess.Popen([sys.executable, 'SAP.py'],
                         stdout=subprocess.PIPE,
                         stderr=subprocess.STDOUT)


def initButtons(window):
    scrapButton = tk.Button(
        master=window,
        text="Scrap data",
        font=('Adobe Garamond Pro',9,'bold'),
        bg='DodgerBlue2',
        fg="white",
        relief = 'flat',
        command=scrapeData)
    scrapButton.grid(row=3, column=2, pady=10)


    skuCreateButton = tk.Button(
        master=window,
        text="Run SKU creation",
        font=('Adobe Garamond Pro',9, 'bold'),
        bg='DodgerBlue2',
        fg="white",
        relief = 'flat',
        command=runSap)
    skuCreateButton.grid(row=11, column=2, pady=10)

    return (scrapButton, skuCreateButton)

def w1():
    image = Image.open("toothbrush.pgm")
    new_image = image.resize((300,300))
    photo = ImageTk.PhotoImage(new_image)

    w1 = tk.Label(window, image=photo)
    w1.grid(row =13, column = 2)


def EANField():
    EANfield = tk.Text(window, fg='black', width=45, height = 3)
    EANfield.grid(row=4, column=2, pady=10)
    return EANfield


def EANLabel():
    EANLabel = tk.Label(window, text='EANs', font = ('Adobe Garamond Pro',9),background='ghost white')
    EANLabel.grid(row=4, column=1, pady=10)

def ProjectDescField():
    projectDescField = tk.Text(window, fg='black', width=45, height=10)
    projectDescField.grid(row=6, column=2, pady=10)
    return projectDescField

def ProjectDescLabel():
    ProjectDescLabel = tk.Label(window, text='Project description', font = ('Adobe Garamond Pro',9),background='ghost white')
    ProjectDescLabel.grid(row=6, column=1, pady=10)


def PPGField():
    PPGField = tk.Entry(window, state='readonly', readonlybackground='white', fg='black', width=60)
    PPGField.grid(row=8, column=2, pady=10)
    return PPGField


def PPGFieldLabel():
    PPGLabel = tk.Label(window, text='PPG',font = ('Adobe Garamond Pro',9),background='ghost white')
    PPGLabel.grid(row=8, column=1, pady=10)


def SalesOrgField():
    SalesOrgField = tk.Entry(window, state='readonly', readonlybackground='white', fg='black', width=60)
    SalesOrgField.grid(row=9, column=2, pady=10)
    return SalesOrgField


def SalesOrgLabel():
    SalesOrgLabel = tk.Label(window, text='Sales orgs',font = ('Adobe Garamond Pro',9),background='ghost white')
    SalesOrgLabel.grid(row=9, column=1, pady=10)


def BaseSKUField():
    BaseSKUField = tk.Entry(window, state='readonly', readonlybackground='white', fg='black', width=60)
    BaseSKUField.grid(row=10, column=2, pady=10)
    return BaseSKUField


def BaseSKULabel():
    BaseSKULabel = tk.Label(window, text='Base SKU',font = ('Adobe Garamond Pro',9),background='ghost white')
    BaseSKULabel.grid(row=10, column=1, pady=10)

startSession()
window = initWindow()

frameEntry = tk.Frame(master=window)
frameEntry.grid(row=1, column=2, padx=10)

scrapButton, skuCreateButton = initButtons(window)

popboxUrlEntry = tk.Entry(master=frameEntry, width=60)
popboxUrlEntry.grid(row=2, column=0)


def popboxLabel():
    popboxLabel = tk.Label(window, text='Insert link to Popbox Brief',font = ('Adobe Garamond Pro',8),background='ghost white')
    popboxLabel.grid(row=1, column=1, pady=10)


eanField = EANField()
eanLabel = EANLabel()
projectDescriptionField = ProjectDescField()
projectDescLabel = ProjectDescLabel()
ppgField = PPGField()
ppgLabel = PPGFieldLabel()
salesOrgLabel = SalesOrgLabel()
salesOrgField = SalesOrgField()
baseSKULabel = BaseSKULabel()
baseSKUField = BaseSKUField()
popboxLabel = popboxLabel()
w1 = w1()


def clock():
    eanVar = tk.StringVar()
    eanVar.set(State.EANValue)
    projectDescVar = tk.StringVar()
    projectDescVar.set(State.projectDescription)
    ppgVar = tk.StringVar()
    ppgVar.set(State.PPGValue)
    salesOrgsVar = tk.StringVar()
    salesOrgsVar.set(State.SalesOrgValue)
    baseSkuVar = tk.StringVar()
    baseSkuVar.set(State.BaseSKUValue)


    ppgField.config(textvariable=ppgVar, relief='flat')
    salesOrgField.config(textvariable=salesOrgsVar, relief='flat')
    baseSKUField.config(textvariable=baseSkuVar, relief='flat')

    if (State.projectDescription != None):
        projectDescriptionField.delete("1.0", tk.END)
        projectDescriptionField.insert(tk.END, State.projectDescription)

    if (State.EANValue != None):
        eanField.delete("1.0", tk.END)
        eanField.insert(tk.END, State.EANValue)

    window.after(1000, clock)  # run itself again after 1000 ms


# run first time
clock()

window.mainloop()
