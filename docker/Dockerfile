# 使用一个基础的Python镜像
FROM python:3.7
#FROM python:3.7-alpine


# 设置工作目录
WORKDIR /home

# 复制应用程序代码到容器中
#COPY . /home
ADD . .

# 安装依赖
#RUN pip install -r requirements.txt
RUN pip install -i https://pypi.tuna.tsinghua.edu.cn/simple -r requirements.txt
# 暴露程序运行的端口（如果需要的话）
EXPOSE 5000

# 启动应用程序
CMD ["python", "./gather_computer_in-1.7.py"]
