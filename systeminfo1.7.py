import win32com.client
import os
import requests
import math
import tkinter as tk
from tkinter import ttk, messagebox

# 创建WMI对象
wmi = win32com.client.GetObject("winmgmts:")

# 获取计算机硬件信息的函数
def get_system_info():
    info = {}
    try:
        info["login_username"] = os.getlogin()

        # 获取计算机名称
        for item in wmi.InstancesOf("Win32_ComputerSystem"):
            computer_name = item.Properties_("Name").Value
            manufacturer = item.Properties_("Manufacturer").Value
            model = item.Properties_("Model").Value
            total_memory = float(item.Properties_("TotalPhysicalMemory").Value) / (1024 ** 3)  # 转换为GB

            info["computer_name"] = computer_name
            info["brand"] = manufacturer
            info["model"] = model
            info["memory_gb"] = math.ceil(total_memory)

        # 获取显卡信息
        try:
            # 查询所有显卡信息
            graphics_cards = wmi.ExecQuery("SELECT * FROM Win32_VideoController")
            gpu_names = ", ".join([gpu.Caption for gpu in graphics_cards])
            info["graphics_cards"] = gpu_names
        except Exception as e:
            info["graphics_cards"] = "获取显卡信息时出错:" + str(e)

        # 获取操作系统名称
        for item in wmi.ExecQuery("SELECT * FROM Win32_OperatingSystem"):
            os_name = item.Caption
            info["os_name"] = os_name

        # 获取所有物理网卡的MAC地址
        wired_mac_addresses = []
        wireless_mac_addresses = []
        virtual_keywords = ["virtual", "vmware", "virtualbox", "hyper-v", "pseudo", "vpn", "loopback"]
        for item in wmi.InstancesOf("Win32_NetworkAdapter"):
            if not any(keyword in item.Description.lower() for keyword in virtual_keywords) and item.NetConnectionID:
                mac_address = item.Properties_("MACAddress").Value
                if mac_address:
                    formatted_mac = mac_address.replace(":", "-")  # 替换分隔符
                    if "wireless" in item.Description.lower() or "wi-fi" in item.Description.lower():
                        wireless_mac_addresses.append(formatted_mac)
                    else:
                        wired_mac_addresses.append(formatted_mac)
        info["wired_mac"] = ', '.join(wired_mac_addresses)
        info["wireless_mac"] = ', '.join(wireless_mac_addresses)

        # 获取主板SN号
        for item in wmi.InstancesOf("Win32_BaseBoard"):
            motherboard_sn = item.Properties_("SerialNumber").Value
            info["motherboard_sn"] = motherboard_sn

        # 获取BIOS信息
        for item in wmi.InstancesOf("Win32_BIOS"):
            sn = item.Properties_("SerialNumber").Value
            info["bios_sn"] = sn

        # 获取CPU物理核心数
        for item in wmi.InstancesOf("Win32_Processor"):
            num_cores = item.Properties_("NumberOfCores").Value
            cpu_name = item.Properties_("Name").Value
            clock_speed = float(item.Properties_("MaxClockSpeed").Value) / 1000  # 转换为GHz

            info["cpu_model"] = cpu_name
            info["cpu_frequency"] = round(clock_speed, 2)
            info["cpu_cores"] = num_cores

    except Exception as e:
        info["错误"] = str(e)

    return info

# 创建图形界面的函数
def create_gui():
    root = tk.Tk()
    root.title("计算机固定资产信息收集")

    # 映射API字段名称到中文标签
    labels = {
        "login_username": "计算机使用人工号\n(公用电脑请输入使用部门任意员工工号)",
        "computer_name": "计算机名称",
        "os_name": "操作系统名称",
        "brand": "品牌",
        "model": "型号",
        "bios_sn": "BIOS SN号",
        "motherboard_sn": "主板SN号",
        "cpu_cores": "CPU物理核心数",
        "cpu_model": "CPU型号",
        "cpu_frequency": "CPU主频 (GHz)",
        "memory_gb": "物理内存大小 (GB)",
        "wired_mac": "所有物理有线网卡的MAC地址",
        "wireless_mac": "所有物理无线网卡的MAC地址",
        "graphics_cards": "显卡名称",
        "os_name": "操作系统名称"
    }

    info = get_system_info()

    for index, (key, label) in enumerate(labels.items()):
        ttk.Label(root, text=label).grid(row=index, column=0, padx=10, pady=5, sticky=tk.W)
        if key == "login_username":
            # 使用可编辑的文本框显示当前登录用户名
            entry = ttk.Entry(root)
            entry.insert(0, info[key])
            entry.grid(row=index, column=1, padx=10, pady=5, sticky=tk.W)
        else:
            ttk.Label(root, text=info[key]).grid(row=index, column=1, padx=10, pady=5, sticky=tk.W)
        ttk.Button(root, text="复制", command=lambda value=info[key]: root.clipboard_clear() or root.clipboard_append(value)).grid(row=index, column=2, padx=10, pady=5)

    # 上传信息到服务器的函数
    def upload_info():
        try:
            data = get_system_info()
            data["employee_id"] = entry.get()  # 从文本框获取值     
            response = requests.post("http://systeminfo.leadchina.cn/gather_computer_info", json=data)
            if response.status_code == 200:
                messagebox.showinfo("成功", "数据上传成功！")
            else:
                response_text = response.json().get('message')
                messagebox.showerror("错误", f"上传数据失败。服务器响应为: {response_text}")
        except Exception as e:
            messagebox.showerror("错误", str(e))
    # 添加上传按钮
    upload_btn = ttk.Button(root, text="上报计算机信息", command=upload_info)
    upload_btn.grid(row=len(labels), column=1, pady=10)

    root.mainloop()

if __name__ == "__main__":
    create_gui()
