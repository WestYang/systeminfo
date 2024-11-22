import win32com.client
import os
import requests
import math

# 创建WMI对象，用于访问Windows的管理信息，例如硬件和操作系统信息
wmi = win32com.client.GetObject("winmgmts:")

# 获取计算机硬件信息的函数
def get_system_info():
    info = {}
    try:
        info["login_username"] = os.getlogin()

        for item in wmi.InstancesOf("Win32_ComputerSystem"):
            computer_name = item.Properties_("Name").Value
            manufacturer = item.Properties_("Manufacturer").Value
            model = item.Properties_("Model").Value

            info["computer_name"] = computer_name
            info["brand"] = manufacturer
            info["model"] = model

        memory_slots = []
        for item in wmi.InstancesOf("Win32_PhysicalMemory"):
            capacity_gb = float(item.Properties_("Capacity").Value) / (1024 ** 3)
            memory_slots.append(f"{math.ceil(capacity_gb)}GB")
        info["memory_slots"] = ', '.join(memory_slots)

        try:
            graphics_cards = wmi.ExecQuery("SELECT * FROM Win32_VideoController")
            gpu_names = ", ".join([gpu.Caption for gpu in graphics_cards])
            info["graphics_cards"] = gpu_names
        except Exception as e:
            info["graphics_cards"] = "获取显卡信息时出错:" + str(e)

        for item in wmi.ExecQuery("SELECT * FROM Win32_OperatingSystem"):
            os_name = item.Caption
            info["os_name"] = os_name

        wired_mac_addresses = []
        wireless_mac_addresses = []
        virtual_keywords = ["virtual", "vmware", "virtualbox", "hyper-v", "pseudo", "vpn", "loopback"]
        for item in wmi.InstancesOf("Win32_NetworkAdapter"):
            if not any(keyword in item.Description.lower() for keyword in virtual_keywords) and item.NetConnectionID:
                mac_address = item.Properties_("MACAddress").Value
                if mac_address:
                    formatted_mac = mac_address.replace(":", "-")
                    if "wireless" in item.Description.lower() or "wi-fi" in item.Description.lower():
                        wireless_mac_addresses.append(formatted_mac)
                    else:
                        wired_mac_addresses.append(formatted_mac)
        info["wired_mac"] = ', '.join(wired_mac_addresses)
        info["wireless_mac"] = ', '.join(wireless_mac_addresses)

        for item in wmi.InstancesOf("Win32_BaseBoard"):
            motherboard_sn = item.Properties_("SerialNumber").Value
            info["motherboard_sn"] = motherboard_sn

        for item in wmi.InstancesOf("Win32_BIOS"):
            sn = item.Properties_("SerialNumber").Value
            info["bios_sn"] = sn

        for item in wmi.InstancesOf("Win32_Processor"):
            num_cores = item.Properties_("NumberOfCores").Value
            cpu_name = item.Properties_("Name").Value
            clock_speed = float(item.Properties_("MaxClockSpeed").Value) / 1000

            info["cpu_model"] = cpu_name
            info["cpu_frequency"] = round(clock_speed, 2)
            info["cpu_cores"] = num_cores

        disk_info = []
        for item in wmi.ExecQuery("SELECT * FROM Win32_DiskDrive"):
            size_gb = float(item.Properties_("Size").Value) / (1024 ** 3)
            interface_type = item.Properties_("InterfaceType").Value
            model = item.Properties_("Model").Value
            disk_info.append(f"{model} ({interface_type}), {math.ceil(size_gb)}GB")
        info["disk_info"] = ', '.join(disk_info)

    except Exception as e:
        info["错误"] = str(e)

    return info

# 上传信息到服务器的函数
def upload_info():
    try:
        data = get_system_info()
        data["employee_id"] = os.getlogin()  # 直接从登录用户获取员工工号
        headers = {'Content-Type': 'application/json; charset=utf-8'}
        requests.post("http://systeminfo.leadchina.cn:80/gather_computer_info", json=data, headers=headers)

    except Exception as e:
        pass  # 忽略错误，保证无感知

# 主函数，自动上传信息
if __name__ == "__main__":
    upload_info()
