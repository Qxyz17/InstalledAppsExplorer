import os
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox
import winreg

class InstalledAppsViewer:
    def __init__(self, root):
        self.root = root
        self.root.title("已安装程序列表")
        self.root.geometry("600x400")

        # 创建一个树形控件
        self.tree = ttk.Treeview(root, columns=("Name", "Version", "Publisher"), show="headings")
        self.tree.heading("Name", text="程序名称")
        self.tree.heading("Version", text="版本")
        self.tree.heading("Publisher", text="发布者")
        self.tree.pack(fill="both", expand=True)

        # 添加右键菜单
        self.popup_menu = tk.Menu(root, tearoff=0)
        self.popup_menu.add_command(label="打开文件所在位置", command=self.open_file_location)
        self.popup_menu.add_command(label="卸载", command=self.uninstall_app)
        self.popup_menu.add_command(label="打开卸载程序所在位置", command=self.open_uninstall_location)

        # 绑定右键事件
        self.tree.bind("<Button-3>", self.show_popup_menu)

        # 添加按钮菜单
        self.button_frame = tk.Frame(root)
        self.button_frame.pack(fill="x")

        self.open_location_button = tk.Button(self.button_frame, text="打开文件所在位置", command=self.open_file_location)
        self.open_location_button.pack(side="left", padx=5, pady=5)

        self.uninstall_button = tk.Button(self.button_frame, text="卸载", command=self.uninstall_app)
        self.uninstall_button.pack(side="left", padx=5, pady=5)

        self.open_uninstall_location_button = tk.Button(self.button_frame, text="打开卸载程序所在位置", command=self.open_uninstall_location)
        self.open_uninstall_location_button.pack(side="left", padx=5, pady=5)

        # 加载已安装程序列表
        self.load_installed_apps()

    def show_popup_menu(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.popup_menu.post(event.x_root, event.y_root)

    def load_installed_apps(self):
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall") as key:
                for i in range(winreg.QueryInfoKey(key)[0]):
                    subkey_name = winreg.EnumKey(key, i)
                    with winreg.OpenKey(key, subkey_name) as subkey:
                        try:
                            name = winreg.QueryValueEx(subkey, "DisplayName")[0]
                            version = winreg.QueryValueEx(subkey, "DisplayVersion")[0]
                            publisher = winreg.QueryValueEx(subkey, "Publisher")[0]
                            uninstall_string = winreg.QueryValueEx(subkey, "UninstallString")[0]
                            install_location = winreg.QueryValueEx(subkey, "InstallLocation")[0]
                            self.tree.insert("", "end", values=(name, version, publisher), tags=(uninstall_string, install_location))
                        except FileNotFoundError:
                            pass
        except Exception as e:
            messagebox.showerror("错误", f"加载已安装程序列表时出错: {e}")

    def open_file_location(self):
        selected_item = self.tree.selection()
        if selected_item:
            _, install_location = self.tree.item(selected_item, "tags")
            if install_location and os.path.exists(install_location):
                try:
                    # 构造命令行字符串，打开文件夹并选中文件
                    folder_path = os.path.dirname(install_location)
                    file_path = os.path.basename(install_location)
                    subprocess.run(f'explorer /select,"{file_path}" "{folder_path}"', shell=True)
                except Exception as e:
                    messagebox.showerror("错误", f"打开文件所在位置时出错: {e}")
            else:
                messagebox.showinfo("信息", "未找到安装位置")
        else:
            messagebox.showinfo("信息", "请先选择一个程序")

    def uninstall_app(self):
        selected_item = self.tree.selection()
        if selected_item:
            uninstall_string = self.tree.item(selected_item, "tags")[0]
            try:
                subprocess.run(uninstall_string, shell=True)
            except Exception as e:
                messagebox.showerror("错误", f"卸载程序时出错: {e}")
        else:
            messagebox.showinfo("信息", "请先选择一个程序")

    def open_uninstall_location(self):
        selected_item = self.tree.selection()
        if selected_item:
            uninstall_string = self.tree.item(selected_item, "tags")[0]
            try:
                uninstall_path = os.path.dirname(uninstall_string)
                if os.path.exists(uninstall_path):
                    subprocess.run(["explorer.exe", uninstall_path])
                else:
                    messagebox.showinfo("信息", "未找到卸载程序所在位置")
            except Exception as e:
                messagebox.showerror("错误", f"打开卸载程序所在位置时出错: {e}")
        else:
            messagebox.showinfo("信息", "请先选择一个程序")

if __name__ == "__main__":
    root = tk.Tk()
    app = InstalledAppsViewer(root)
    root.mainloop()