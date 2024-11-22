import win32com.client
import os
import requests
import math
import tkinter as tk
from tkinter import ttk, messagebox

# 创建WMI对象，用于访问Windows的管理信息，例如硬件和操作系统信息
wmi = win32com.client.GetObject("winmgmts:")

# 获取计算机硬件信息的函数
def get_system_info():
    # 用于存储所有获取到的硬件信息
    info = {}
    try:
        # 获取当前登录的用户名
        info["login_username"] = os.getlogin()

        # 获取计算机名称、品牌和型号
        for item in wmi.InstancesOf("Win32_ComputerSystem"):
            computer_name = item.Properties_("Name").Value  # 计算机名称
            manufacturer = item.Properties_("Manufacturer").Value  # 品牌
            model = item.Properties_("Model").Value  # 型号

            info["computer_name"] = computer_name
            info["brand"] = manufacturer
            info["model"] = model

        # 获取内存插槽信息，使用 Win32_PhysicalMemory 类获取每个插槽的容量
        memory_slots = []
        for item in wmi.InstancesOf("Win32_PhysicalMemory"):
            capacity_gb = float(item.Properties_("Capacity").Value) / (1024 ** 3)  # 转换为GB
            memory_slots.append(f"{math.ceil(capacity_gb)}GB")  # 以GB为单位存储
        info["memory_slots"] = ', '.join(memory_slots)  # 将每个插槽的内存大小连接成字符串

        # 获取显卡信息
        try:
            # 查询所有显卡的信息
            graphics_cards = wmi.ExecQuery("SELECT * FROM Win32_VideoController")
            gpu_names = ", ".join([gpu.Caption for gpu in graphics_cards])  # 获取显卡名称
            info["graphics_cards"] = gpu_names
        except Exception as e:
            info["graphics_cards"] = "获取显卡信息时出错:" + str(e)

        # 获取操作系统名称
        for item in wmi.ExecQuery("SELECT * FROM Win32_OperatingSystem"):
            os_name = item.Caption  # 操作系统名称
            info["os_name"] = os_name

        # 获取所有物理网卡的MAC地址
        wired_mac_addresses = []  # 存储有线网卡的MAC地址
        wireless_mac_addresses = []  # 存储无线网卡的MAC地址
        virtual_keywords = ["virtual", "vmware", "virtualbox", "hyper-v", "pseudo", "vpn", "loopback"]
        for item in wmi.InstancesOf("Win32_NetworkAdapter"):
            if not any(keyword in item.Description.lower() for keyword in virtual_keywords) and item.NetConnectionID:
                mac_address = item.Properties_("MACAddress").Value  # 获取MAC地址
                if mac_address:
                    formatted_mac = mac_address.replace(":", "-")  # 替换分隔符
                    # 判断是有线网卡还是无线网卡
                    if "wireless" in item.Description.lower() or "wi-fi" in item.Description.lower():
                        wireless_mac_addresses.append(formatted_mac)
                    else:
                        wired_mac_addresses.append(formatted_mac)
        info["wired_mac"] = ', '.join(wired_mac_addresses)
        info["wireless_mac"] = ', '.join(wireless_mac_addresses)

        # 获取主板的序列号
        for item in wmi.InstancesOf("Win32_BaseBoard"):
            motherboard_sn = item.Properties_("SerialNumber").Value  # 主板序列号
            info["motherboard_sn"] = motherboard_sn

        # 获取BIOS信息
        for item in wmi.InstancesOf("Win32_BIOS"):
            sn = item.Properties_("SerialNumber").Value  # BIOS序列号
            info["bios_sn"] = sn

        # 获取CPU物理核心数、型号和频率
        for item in wmi.InstancesOf("Win32_Processor"):
            num_cores = item.Properties_("NumberOfCores").Value  # CPU物理核心数
            cpu_name = item.Properties_("Name").Value  # CPU型号
            clock_speed = float(item.Properties_("MaxClockSpeed").Value) / 1000  # 转换为GHz

            info["cpu_model"] = cpu_name
            info["cpu_frequency"] = round(clock_speed, 2)  # 保留两位小数
            info["cpu_cores"] = num_cores

        # 获取硬盘信息
        disk_info = []
        for item in wmi.ExecQuery("SELECT * FROM Win32_DiskDrive"):
            size_gb = float(item.Properties_("Size").Value) / (1024 ** 3)  # 硬盘容量转换为GB
            interface_type = item.Properties_("InterfaceType").Value  # 硬盘接口类型（如SATA, SSD等）
            model = item.Properties_("Model").Value  # 硬盘型号
            disk_info.append(f"{model} ({interface_type}), {math.ceil(size_gb)}GB")  # 存储硬盘信息
        info["disk_info"] = ', '.join(disk_info)

    except Exception as e:
        info["错误"] = str(e)  # 捕获任何错误并存储到info中

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
        "memory_slots": "内存插槽和大小",
        "wired_mac": "所有物理有线网卡的MAC地址",
        "wireless_mac": "所有物理无线网卡的MAC地址",
        "graphics_cards": "显卡名称",
        "disk_info": "硬盘信息",
    }

    # 获取系统信息
    info = get_system_info()

    # 创建界面元素
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
            # 获取系统信息
            data = get_system_info()
            # 从文本框获取员工工号
            data["employee_id"] = entry.get()
            # 将数据POST到服务器
            #response = requests.post("http://systeminfo.leadchina.cn/gather_computer_info", json=data)
            response = requests.post("http://10.30.162.153:5000/gather_computer_info", json=data)
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
