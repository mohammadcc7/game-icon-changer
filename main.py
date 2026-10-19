import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
import win32com.client

def update_icons(shortcuts_dir, icons_dir):
    if not os.path.exists(shortcuts_dir) or not os.path.exists(icons_dir):
        messagebox.showerror("خطأ", "يرجى التنسيق واختيار المسارات الصحيحة.")
        return

    shell = win32com.client.Dispatch("WScript.Shell")
    updated_count = 0

    for filename in os.listdir(shortcuts_dir):
        if filename.endswith(".lnk"):
            game_name = os.path.splitext(filename)[0]
            icon_path = os.path.join(icons_dir, f"{game_name}.ico")

            if os.path.exists(icon_path):
                shortcut_path = os.path.join(shortcuts_dir, filename)
                shortcut = shell.CreateShortCut(shortcut_path)
                shortcut.IconLocation = f"{icon_path}, 0"
                shortcut.Save()
                updated_count += 1

    messagebox.showinfo("تم بنجاح", f"تم تحديث أيقونات {updated_count} اختصار بنجاح!")

def browse_shortcuts():
    path = filedialog.askdirectory(title="اختر مجلد الاختصارات (.lnk)")
    if path:
        entry_shortcuts.delete(0, tk.END)
        entry_shortcuts.insert(0, path)

def browse_icons():
    path = filedialog.askdirectory(title="اختر مجلد الأيقونات (.ico)")
    if path:
        entry_icons.delete(0, tk.END)
        entry_icons.insert(0, path)

def start_process():
    update_icons(entry_shortcuts.get(), entry_icons.get())

# إنشاء النافذة الرئيسية
root = tk.Tk()
root.title("مغير أيقونات الألعاب - Game Icon Changer")
root.geometry("480x220")
root.resizable(False, False)

# عناصر الواجهة
tk.Label(root, text=":مجلد اختصارات الألعاب", font=("Tahoma", 10)).pack(anchor="e", padx=15, pady=(15, 2))
frame1 = tk.Frame(root)
frame1.pack(fill="x", padx=15)
entry_shortcuts = tk.Entry(frame1, justify="right")
entry_shortcuts.pack(side="left", expand=True, fill="x", padx=(0, 5))
tk.Button(frame1, text="تحديد...", command=browse_shortcuts).pack(side="right")

tk.Label(root, text=":مجلد الأيقونات (.ico)", font=("Tahoma", 10)).pack(anchor="e", padx=15, pady=(10, 2))
frame2 = tk.Frame(root)
frame2.pack(fill="x", padx=15)
entry_icons = tk.Entry(frame2, justify="right")
entry_icons.pack(side="left", expand=True, fill="x", padx=(0, 5))
tk.Button(frame2, text="تحديد...", command=browse_icons).pack(side="right")

btn_start = tk.Button(root, text="تطبيق وتغيير الأيقونات", bg="#27ae60", fg="white", font=("Tahoma", 11, "bold"), command=start_process)
btn_start.pack(pady=20, fill="x", padx=15)

root.mainloop()
