import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
import win32com.client

def clean_name(name):
    """دالة لتنظيف الأسماء وتوحيدها لتجنب مشاكل المسافات والأحرف الكبيرة والصغيرة"""
    return name.strip().lower()

def update_icons(shortcuts_dir, icons_dir):
    if not os.path.exists(shortcuts_dir) or not os.path.exists(icons_dir):
        messagebox.showerror("خطأ", "يرجى التحقق من صحة المسارات المدخلة.")
        return

    shell = win32com.client.Dispatch("WScript.Shell")
    
    # قراءة محتويات مجلد الأيقونات وتكوين خريطة (معيارية) للبحث السريع
    icons_map = {}
    for icon_file in os.listdir(icons_dir):
        if icon_file.lower().endswith(".ico"):
            base_name = os.path.splitext(icon_file)[0]
            icons_map[clean_name(base_name)] = os.path.join(icons_dir, icon_file)

    updated_count = 0
    failed_games = []

    # المرور على جميع الاختصارات
    for filename in os.listdir(shortcuts_dir):
        if filename.endswith(".lnk"):
            game_name = os.path.splitext(filename)[0]
            cleaned_game_name = clean_name(game_name)
            
            # البحث عن مطابقة دقيقة أو مطابقة بعد تنظيف المسافات والأحرف
            if cleaned_game_name in icons_map:
                icon_path = icons_map[cleaned_game_name]
                shortcut_path = os.path.join(shortcuts_dir, filename)
                try:
                    shortcut = shell.CreateShortCut(shortcut_path)
                    shortcut.IconLocation = f"{icon_path}, 0"
                    shortcut.Save()
                    updated_count += 1
                except Exception as e:
                    failed_games.append(f"{game_name} (خطأ في الحفظ)")
            else:
                failed_games.append(game_name)

    # عرض تقرير نهائي مفصل بالنتيجة
    report = f"تم تحديث أيقونات {updated_count} اختصار بنجاح!"
    if failed_games:
        report += f"\n\nلم يتم العثور على أيقونات مطابقة لـ {len(failed_games)} لعبة:\n" + "\n".join(failed_games[:10])
        if len(failed_games) > 10:
            ...
    
    messagebox.showinfo("تقرير التحديث الجذري", report)

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
root.title("مغير أيقونات الألعاب - الإصدار المطور والجذري")
root.geometry("500x250")
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

btn_start = tk.Button(root, text="تطبيق التحديث الجذري", bg="#27ae60", fg="white", font=("Tahoma", 11, "bold"), command=start_process)
btn_start.pack(pady=20, fill="x", padx=15)

root.mainloop()
