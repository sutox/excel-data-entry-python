import tkinter as tk
from tkinter import ttk
import openpyxl
import re


def load_data():
    global workbook, sheet, list_values,path
    path = "المنتجات.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    list_values = list(sheet.values)
    for col_name in list_values[0]:
        treeView.heading(col_name, text=col_name)
    for value_tuple in list_values[1:]:
        treeView.insert("", tk.END, values=value_tuple)


def insert_row():
    product = product_entry.get()
    price1 = int(price1_entry.get())
    price2 = int(price2_entry.get())
    print(product)
    path = "المنتجات.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    row_values = [product, price1, price2]
    sheet.append(row_values)
    workbook.save(path)
    treeView.insert("", tk.END, values=row_values,)

    product_entry.delete(0, "end")
    product_entry.insert(0, "اسم المنتج")
    price1_entry.delete(0, "end")
    price1_entry.insert(0, "سعر الجملة")
    price2_entry.delete(0, "end")
    price2_entry.insert(0, "سعر البيع")


def clear():
    product_entry.delete(0, "end")
    product_entry.insert(0, "اسم المنتج")
    price1_entry.delete(0, "end")
    price1_entry.insert(0, "سعر الجملة")
    price2_entry.delete(0, "end")
    price2_entry.insert(0, "سعر البيع")


def search():
    searched_value = search_entry.get().strip()
    for product in treeView.get_children():
        treeView.delete(product)
    for value_tuple in list_values[1:]:
        if re.search(searched_value, value_tuple[0]):
            treeView.insert("", tk.END, values=value_tuple)


def back_treeview():
    for products in treeView.get_children():
        treeView.delete(products)

    for col_name in list_values[0]:
        treeView.heading(col_name, text=col_name)

    for value_tuple in list_values[1:]:
        treeView.insert("", tk.END, values=value_tuple)

    search_entry.delete(0, "end")
    search_entry.insert(0, "البحث")


def remove_one():

    selcted_item=treeView.focus()
    details = treeView.item(selcted_item)
    product_name= details.get("values")[0]
    products = [] 
    index=0
    for row in sheet: 
     
        name = row[0].value 

        products.append(name)
    
    for i in products :
        index+=1
        if product_name == i :
           sheet.delete_rows(index)
           workbook.save(path)
    treeView.delete(selcted_item)


root = tk.Tk()
w = root.winfo_screenwidth()
h = root.winfo_screenheight()
#root.geometry("%dx%d" %(w,h))


root.title("ياس للأنشائيه")
root.rowconfigure(0,weight=1)
root.rowconfigure(1,weight=1)
root.columnconfigure(0,weight=1)
root.columnconfigure(0,weight=1)


style = ttk.Style(root)
root.tk.call("source", "forest-light.tcl")
style.theme_use("forest-light")


frame = ttk.Frame(root)
frame.pack()

widgets_frame = ttk.Frame(frame,)

widgets_frame.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")

search_entry = ttk.Entry(widgets_frame,width=8,justify="center")
search_entry.insert(0, "البحث")
search_entry.bind("<FocusIn>", lambda e: search_entry.delete("0", "end"))
search_entry.grid(row=0, column=0, pady=5, sticky="nsew")

button_frame = ttk.Frame(widgets_frame)
button_frame.grid(row=1,column=0,pady=5, sticky="nsew",)
button_search = ttk.Button(button_frame, text="أبحث", width=50, command=search)
button_search.grid(row=0, column=0,padx=20, sticky="e")

button_back = ttk.Button(button_frame, text="الرجوع", width=50, command=back_treeview)
button_back.grid(row=0, column=1, sticky="e",padx=20)


product_entry = ttk.Entry(widgets_frame,justify="center")
product_entry.insert(0, "اسم المنتج")
product_entry.bind("<FocusIn>", lambda e: product_entry.delete("0", "end"))
product_entry.grid(row=2, column=0, padx=5, pady=5, sticky="nsew")

price1_entry = ttk.Entry(widgets_frame,justify="center")
price1_entry.insert(0, "سعر الجملة")
price1_entry.bind("<FocusIn>", lambda e: price1_entry.delete("0", "end"))
price1_entry.grid(row=3, column=0, padx=5, pady=5, sticky="nsew")

price2_entry = ttk.Entry(widgets_frame,justify="center")
price2_entry.insert(0, "سعر البيع")
price2_entry.bind("<FocusIn>", lambda e: price2_entry.delete("0", "end"))
price2_entry.grid(row=4, column=0, padx=5, pady=5, sticky="nsew")

button = ttk.Button(widgets_frame, text="ضف", command=insert_row)
button.grid(row=5, column=0, padx=5, pady=10, sticky="nsew")
clear_button = ttk.Button(widgets_frame, text="امحي", command=clear)
clear_button.grid(row=6, column=0, padx=5, pady=10, sticky="nsew")

remove_button = ttk.Button(widgets_frame, text="مسح", command=remove_one)
remove_button.grid(row=7, column=0, padx=5, pady=10, sticky="nsew")

treeFrame = ttk.Frame(frame)
treeFrame.rowconfigure(0,weight=1)
treeFrame.columnconfigure(0, weight=1)
treeFrame.rowconfigure(0, weight=1)
treeFrame.grid(row=0, column=0, pady=20, sticky="nsew")
treeScroll = ttk.Scrollbar(treeFrame)
treeScroll.pack(side="left", fill="y")

cols = ("المنتج", "سعر الجملة", "سعر البيع")
treeView = ttk.Treeview(
    treeFrame, show="headings", yscrollcommand=treeScroll.set, columns=cols, height=15,
)
treeView.column("المنتج", width=150,anchor='center')
treeView.column("سعر الجملة", width=100,anchor='center')
treeView.column("سعر البيع", width=100,anchor='center')
treeView.pack()
treeScroll.config(command=treeView.yview)
load_data()
root.mainloop()
