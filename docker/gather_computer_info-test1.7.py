import json
import pymssql
import requests
from flask import Flask, request, jsonify

app = Flask(__name__)

# 数据库连接配置，根据您的SQL Server配置进行调整
DB_CONFIG = {
    "server": "10.30.162.74",
    "user": "ops",
    "password": "Qwerty123456",
    "database": "OPSDB"
}


def query_employee_by_job_no(job_no):
    # 定义API接口URL
    url = 'https://it.leadchina.cn/proxy/api/ehr/query_employee_by_job_no'
    # 使用form-data格式提交'job_no'参数
    headers = {}
    files = []
    data = {
        'job_no': job_no,
        # 'domain': 'leadchina.cn'
    }

    try:
        # 使用POST方式调用API接口
        response = requests.request("POST", url, headers=headers, data=data, files=files)

        # 解析返回的JSON数据
        result = response.json()

        # 检查'errcode'字段是否为'0'，并返回相应的值
        if result.get('errcode') == '0':
            return True, result.get('data').get('accounts')[0]
        else:
            print(result)
            return False, None

    # 捕获API调用过程中的异常
    except requests.RequestException as e:
        print(f"API call error: {e}")
        return False, None

    except json.JSONDecodeError as e:
        print(f"Error decoding JSON: {e}")
        return False, None


@app.route('/gather_computer_info', methods=['POST'])
def gather_computer_info():
    data = request.json

    # 数据校验  // employee_id login_username
    required_fields = ["employee_id", "computer_name", "bios_sn", "login_username"]
    for field in required_fields:
        if field not in data or not data[field]:
            return jsonify({"message": u"使用人工号字段是必填字段，不能为空!"}), 400

    try:
        # 使用工号查询所属部门，同时校验工号是否合法
        employee_id = data.get('employee_id')
        login_username = data.get('login_username')
        # # 工号与Windows系统登录名一致
        # if employee_id == login_username:
        #     pass
        # else:
        #     # 公共账号登录，使用所属部门人员工号
        #     pass
        is_success, employee_info = query_employee_by_job_no(job_no=data.get('employee_id'))
        level2_deptName = employee_info.get('level2_deptName', '') if is_success else ''
        if not is_success:
            return jsonify({"message": u"计算机使用人工号不存在，请重新输入在职员工工号!"}), 400

        # 连接到SQL Server数据库
        with pymssql.connect(**DB_CONFIG) as conn:
            cursor = conn.cursor()

            # 创建数据表（如果不存在）
            cursor.execute('''
            IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='ComputerInfo' AND xtype='U')
            CREATE TABLE ComputerInfo (
                ID INT PRIMARY KEY IDENTITY(1,1),
                EmployeeId  NVARCHAR(255),
                LoginUsername NVARCHAR(255),
                ComputerName NVARCHAR(255),
                Brand NVARCHAR(255),
                Model NVARCHAR(255),
                BIOS_SN NVARCHAR(255),
                Motherboard_SN NVARCHAR(255),
                CPU_Cores INT,
                CPU_Model NVARCHAR(255),
                CPU_Frequency FLOAT,
                Memory_GB FLOAT,
                Wired_MAC NVARCHAR(255),
                Wireless_MAC NVARCHAR(255),
                level2_deptName NVARCHAR(255),
                GraphicsCards NVARCHAR(1000),  -- 新增用于存储显卡信息的字段
		        Os_Name NVARCHAR(1000),  -- 新增用于系统名称的字段
                CreateTime DATETIME DEFAULT GETDATE(),
                UpdateTime DATETIME
            )
            ''')

            # 添加唯一键约束
            cursor.execute('''
            IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name='UQ_ComputerInfo' AND object_id = OBJECT_ID('ComputerInfo'))
            ALTER TABLE ComputerInfo
            ADD CONSTRAINT UQ_ComputerInfo UNIQUE (ComputerName, BIOS_SN)
            ''')

            # 检查是否存在匹配的记录
            cursor.execute("SELECT COUNT(*) FROM ComputerInfo WHERE ComputerName=%s AND BIOS_SN=%s",
                           (data['computer_name'], data['bios_sn']))
            count = cursor.fetchone()[0]

            if count:
                # 更新记录
                cursor.execute('''
                UPDATE ComputerInfo
                SET EmployeeId=%s, LoginUsername=%s, Brand=%s, Model=%s, Motherboard_SN=%s, CPU_Cores=%s, CPU_Model=%s,
                CPU_Frequency=%s, Memory_GB=%s, Wired_MAC=%s, Wireless_MAC=%s, level2_deptName=%s, GraphicsCards=%s, Os_Name=%s, UpdateTime=GETDATE()
                WHERE ComputerName=%s AND BIOS_SN=%s
                ''', (data.get('employee_id'), data['login_username'], data.get('brand', None), data.get('model', None),
                      data.get('motherboard_sn', None), data.get('cpu_cores', None), data.get('cpu_model', None),
                      data.get('cpu_frequency', None), data.get('memory_gb', None), data.get('wired_mac', None),
                      data.get('wireless_mac', None), level2_deptName, data.get('graphics_cards', None), data.get('os_name', None), data['computer_name'], data['bios_sn']))
            else:
                # 插入新记录
                cursor.execute('''
                INSERT INTO ComputerInfo (EmployeeId, LoginUsername, ComputerName, Brand, Model, BIOS_SN, Motherboard_SN
                , CPU_Cores, CPU_Model, CPU_Frequency, Memory_GB, Wired_MAC, Wireless_MAC, level2_deptName, GraphicsCards, Os_Name)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                ''', (data.get('employee_id'), data['login_username'], data['computer_name'], data.get('brand', None),
                      data.get('model', None),
                      data['bios_sn'], data.get('motherboard_sn', None), data.get('cpu_cores', None),
                      data.get('cpu_model', None), data.get('cpu_frequency', None), data.get('memory_gb', None),
                      data.get('wired_mac', None), data.get('wireless_mac', None), level2_deptName, data.get('graphics_cards', None), data.get('os_name', None)))

            conn.commit()

        return jsonify({"message": "Data stored successfully!"}), 200

    except pymssql.DatabaseError as e:
        return jsonify({"message": f"Database error: {str(e)}"}), 500
    except Exception as e:
        return jsonify({"message": str(e)}), 500


if __name__ == '__main__':
    app.run(host='0.0.0.0', debug=True)
