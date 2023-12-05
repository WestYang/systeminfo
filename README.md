部署流程
1.打包前端软件

```
pyinstaller.exe --noconsole --icon .\2.ico --upx-dir=C:\upx --onefile .\systeminfo-test1.7.py
```

2.打包后端docker镜像

```
docker build -t gather-computer-info-in:1.7 .
```

3.部署docker容器

```
docker run -d -p 5000:5000 gather-computer-info-in:1.7
```

