import tkinter as ui
import os



def openExcel():
    window.destroy()
    os.startfile("Klasu grafiks 2023_2024.xlsx")
    global temp
    temp = 1

def closeWindow():
    window.destroy()

def continueInput():
    name = str(teacher_entry.get())
    lesson = str(lesson_entry.get())
    place = str(place_entry.get())
    window.destroy()
    global temp

window = ui.Tk() 
window.geometry("400x400")

start_menu = ui.LabelFrame(window, text = "Bungu skolas grafiks")
start_menu.grid(row = 0, column = 0)
start_menu.pack()

temp = 0
button_frame = ui.LabelFrame(window)
button_frame.pack()
excel_button = ui.Button(button_frame, text = "Atvērt tabulu", command = openExcel)
excel_button.grid(row = 1, column = 1)
data_button = ui.Button(button_frame, text = "ievadīt datus", command = closeWindow)
data_button.grid(row = 2, column = 1)


window.mainloop()

if temp == 0:
    window = ui.Tk() 
    frame = ui.Frame(window)
    frame.pack()    

    data_menu = ui.LabelFrame(frame, text = "Datu ievade")
    data_menu.grid(row = 0, column = 0)
    data_menu.pack()

    teacher_info = ui.Label(data_menu, text = "Pasniedzēja vārds")
    teacher_info.grid(row = 0, column = 0)
    teacher_entry = ui.Entry(data_menu)
    teacher_entry.grid(row = 0 , column = 1)

    lesson_info = ui.Label(data_menu, text = "Nodarbības veids")
    lesson_info.grid(row = 1, column = 0)
    lesson_entry = ui.Spinbox(data_menu, values = ["Individuālā", "Grupu"])
    lesson_entry.grid(row = 1 , column = 1)

    place_info = ui.Label(data_menu, text = "Filiāle")
    place_info.grid(row = 2, column = 0)
    place_entry = ui.Spinbox(data_menu, values = ["Rīga", "Mārupe", "Ādaži"])
    place_entry.grid(row = 2 , column = 1) 

    button_frame = ui.LabelFrame(frame)
    button_frame.pack()
    accept_button = ui.Button(button_frame, text = "Apstiprināt", command = continueInput)   
    accept_button.grid(row = 1, column = 1)

    window.mainloop()

