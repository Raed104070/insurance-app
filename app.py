import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog

# تحميل ملف Excel
def load_excel_file():
    filepath = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
    if filepath:
        df = pd.read_excel(filepath)
        display_all_data(df)
        global current_df
        current_df = df

# عرض البيانات في الجدول
def display_all_data(df):
    for widget in table_frame.winfo_children():
        widget.destroy()

    style = ttk.Style()
    style.theme_use("default")

    # تنسيق الرأس والصفوف
    style.configure("Treeview",
        background="white",
        foreground="black",
        rowheight=28,
        fieldbackground="white",
        font=("Segoe UI", 10)
    )

    style.configure("Treeview.Heading",
        background="#003366",  # أزرق غامق
        foreground="white",
        font=("Segoe UI", 11, "bold")
    )

    style.map("Treeview",
        background=[("selected", "#b30000")],  # أحمر داكن عند التحديد
        foreground=[("selected", "white")]
    )

    tree = ttk.Treeview(table_frame, columns=list(df.columns), show='headings')
    tree.pack(expand=True, fill='both')

    for col in df.columns:
        tree.heading(col, text=col)
        tree.column(col, width=150, anchor='center')

    for _, row in df.iterrows():
        tree.insert('', 'end', values=list(row))

    global current_tree
    current_tree = tree

# البحث بحسب اسم الشركة و/أو التغطية
def search_filtered():
    company_query = company_entry.get().strip().lower()
    coverage_query = coverage_entry.get().strip().lower()

    df = current_df.copy()

    if company_query:
        df = df[df.iloc[:, 0].astype(str).str.lower().str.contains(company_query)]

    if coverage_query:
        matched_cols = [col for col in df.columns if coverage_query in str(col).lower()]
        if matched_cols:
            columns_to_show = [df.columns[0]] + matched_cols  # إبقاء اسم الشركة
            columns_to_show = list(dict.fromkeys(columns_to_show))  # إزالة التكرار إن وجد
            df = df[columns_to_show]
        else:
            df = df[[df.columns[0]]]  # فقط اسم الشركة

    display_all_data(df)

# إعداد الواجهة
root = tk.Tk()
root.title("تغطيات وثائق التأمين الصحي - شركة تمكين للتأمين")
root.geometry("1200x700")
root.configure(bg="#f4f4f4")

# عنوان البرنامج
title_label = tk.Label(root, text="📋 تغطيات وثائق التأمين الصحي - شركة تمكين للتأمين", font=("Segoe UI", 18, "bold"), bg="#f4f4f4", fg="#003366")
title_label.pack(pady=10)

# زر التحميل
load_button = tk.Button(root, text="📂 تحميل ملف Excel", font=("Segoe UI", 12), command=load_excel_file, bg="#4CAF50", fg="white", padx=20, pady=5)
load_button.pack(pady=5)

# إطار البحث
search_frame = tk.Frame(root, bg="#f4f4f4")
search_frame.pack(pady=10)

tk.Label(search_frame, text="🔍 اسم الشركة:", font=("Segoe UI", 11), bg="#f4f4f4").grid(row=0, column=0, padx=10)
company_entry = tk.Entry(search_frame, width=30, font=("Segoe UI", 10))
company_entry.grid(row=0, column=1, padx=10)

tk.Label(search_frame, text="🔍 اسم التغطية:", font=("Segoe UI", 11), bg="#f4f4f4").grid(row=0, column=2, padx=10)
coverage_entry = tk.Entry(search_frame, width=30, font=("Segoe UI", 10))
coverage_entry.grid(row=0, column=3, padx=10)

search_button = tk.Button(search_frame, text="بحث", command=search_filtered, font=("Segoe UI", 11, "bold"), bg="#2196F3", fg="white", padx=15)
search_button.grid(row=0, column=4, padx=20)

# إطار الجدول
table_frame = tk.Frame(root, bg="#f4f4f4")
table_frame.pack(expand=True, fill='both', padx=20, pady=10)

# بيانات مؤقتة
current_df = pd.DataFrame()
current_tree = None

# تشغيل البرنامج
root.mainloop()

