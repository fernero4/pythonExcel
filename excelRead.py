from tkinter import *
from tkinter import ttk, filedialog
from tkinter import messagebox
import numpy
import pandas as pd
import pathlib


root = Tk()
root.title('Excel read')
root.geometry('1100x400+200+200')
root.configure(bg='#0078d7')
root.resizable(False, False)

IMAG_PATH = 'C:/Users/ferna/Downloads/python/pythonExcel/Images/'

def Open():
    filename=filedialog.askopenfilename(title="Open a file", filetype=(("xlxs files", ".*xlsx"),("All files", "*.*")))

    if filename:
        try:
            filename= r"{}".format(filename)
            df=pd.read_excel(filename)

        except:
            messagebox.showerror("Error", "You can't access this file!")

    tree.delete()
    tree['column'] = list(df.columns)
    tree['show'] = "headings"

    #heading title
    for col in tree['column']:
        tree.heading(col, text=col) 

    #data
    df_rows=df.to_numpy().tolist()
    for row in df_rows:
        tree.insert("", "end", values=row)


# icon
icon_image = PhotoImage(file=IMAG_PATH+'iconExcel.png')
root.iconphoto(False, icon_image)

# heading
frame=Frame(root, bg="white")
frame.pack(padx=10, pady=30)

#three view
tree=ttk.Treeview(frame)
tree.pack()

#button
button=Button(root, width=60,height=2, text="Open", font=30, fg="white", bg="blue", command=Open)
button.pack(padx=10, pady=40)


root.mainloop()