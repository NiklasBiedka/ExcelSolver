from tkinter import ttk
from tkinter import *
from tkinter.ttk import *
import sqlite3
#import xlwings as xw

root_width = 543
root_height = 350

# Connection to Excel Worksheet
#wb = xw.Book()
#wbTest = xw.Book('test2.xlsx')

# Database
conn = sqlite3.connect('data_sets.db')
c = conn.cursor()

'''
# Create Table
c.execute("""CREATE TABLE datasets(
        name text,
        cell_range text,
        index_ranges text
)""")
'''

root = Tk()
root.title("Data Editor")

def show_datasets():
    global name_entry
    global cr_entry
    global ir_entry
    global list_tree
    conn = sqlite3.connect('data_sets.db')
    
    c = conn.cursor()
    c.execute("SELECT *,oid FROM datasets")
    
    records = c.fetchall()

    #print(records)

    for i in list_tree.get_children():
        	list_tree.delete(i)

    list_tree.insert("", 1, text="<Add a New Data Set>")

    for record in records:
        list_tree.insert("", 10, text=record[0], values=(record[1], record[2], record[3]))

    conn.commit()
    conn.close()


# define functions for the widgets
def add_data_set():
    global name_entry
    global cr_entry
    global ir_entry
    global list_tree
    global error_label
    
    name_double = 0

    error_label.config(text="")
    
    conn = sqlite3.connect('data_sets.db')
    c = conn.cursor()
    c.execute("SELECT * FROM datasets")
    records = c.fetchall()

    for record in records:
        if record[0] == name_entry.get():
            name_double = 1

    if len(name_entry.get()) == 0:
        error_label.config(text="Name hat keinen Wert")
    elif len(cr_entry.get()) == 0:
        error_label.config(text="Cell range hat keinen Wert")
    elif name_double == 1:
        error_label.config(text="Dieser Name existiert bereits")
    else:
        if len(ir_entry.get()) != 0:
            c.execute("INSERT INTO datasets VALUES (:name, :cell_range, :index_ranges)",
            {
                'name': name_entry.get(),
                'cell_range': cr_entry.get(),
                'index_ranges': ir_entry.get()
            })
        else:
            c.execute("INSERT INTO datasets VALUES (:name, :cell_range, :index_ranges)",
            {
                'name': name_entry.get(),
                'cell_range': cr_entry.get(),
                'index_ranges': None,
            })   
        conn.commit()
        conn.close()
        show_datasets()
        

# Delete Entry from DB
def delete_dataset():
    global delete_entry
    global error_label

    error_label.config(text="")

    conn = sqlite3.connect('data_sets.db')
    c = conn.cursor()

    c.execute("DELETE FROM datasets WHERE oid=?", (delete_entry.get(),))
    conn.commit()

    conn.close()
    show_datasets()

def clear_entries():
    global name_entry
    global cr_entry
    global ir_entry
    global error_label

    error_label.config(text="")

    name_entry.delete(0, END)
    name_entry.insert(0, "")

    cr_entry.delete(0, END)
    cr_entry.insert(0, "")

    ir_entry.delete(0, END)
    ir_entry.insert(0, "")


# create all of the main containers
top_frame = Frame(root, width=root_width, height=100)
center_frame = Frame(root, width=root_width, height = 100)
btm_frame = Frame(root, width=root_width, height=50)

# layout all of the main containers
top_frame.grid(row=1, sticky="nsew")
center_frame.grid(row=2, sticky="ew")
btm_frame.grid(row=3, sticky="ew")

# create the widgets for the top frame
list_tree = ttk.Treeview(top_frame)
list_tree["columns"] = ("cr", "ir","#")
list_tree.column("#0", stretch=NO)
list_tree.column("cr", stretch=NO)
list_tree.column("ir",  stretch=NO)
list_tree.column("#", stretch=NO, width=20)

list_tree.heading("#0", text="Name:", anchor=W)
list_tree.heading("cr", text="Cell Range:", anchor=W)
list_tree.heading("ir", text="Index Range(s):", anchor=W)
list_tree.heading("#", text="#", anchor=W)

list_tree.insert("", 0, text="<Add New Data Item>")

# layout the widgets in the top frame
list_tree.pack()

# create the widgets for the center frame
btm_name_label = Label(center_frame, text="Name:", width=15)
btm_cr_label = Label(center_frame, text="Cell Range", width=15)
btm_ir_label = Label(center_frame, text="Insert Range(s)", width=15)
error_label = Label(center_frame)

name_entry = Entry(center_frame)
cr_entry = Entry(center_frame, width=10)
ir_entry = Entry(center_frame, width=10)
delete_entry = Entry(center_frame, width=10)

delete_button = Button(center_frame, text="Delete Data Item", command=delete_dataset)
add_button = Button(center_frame, text="Add Data Item", command=add_data_set)
cancel_button = Button(center_frame, text="Clear", command=clear_entries)

# layout the widgets in the center frame
btm_name_label.grid(row=0, column=0)
btm_cr_label.grid(row=0, column=1)
btm_ir_label.grid(row=0, column=2, columnspan=1)
error_label.grid(row=0, column=4, columnspan=3, rowspan=2)


name_entry.grid(row=1, column=0)
cr_entry.grid(row=1, column=1)
ir_entry.grid(row=1, column=2)
delete_entry.grid(row=1, column=3)

delete_button.grid(row=3, column=3)
add_button.grid(row=3, column=2)
cancel_button.grid(row=3, column=4)

# create the widgets for the bottom frame
close_button = Button(btm_frame, text="Close", command=root.quit)

# layour the widgets in the bottom frame
close_button.pack(anchor='e')

show_datasets()

root.mainloop()