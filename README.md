# TLC_Browers
一个开源的vb6基于webview2支持h5的浏览器
环境配置:
1.找到dlls文件夹
2.使用regsvr32.exe注册file_controlv2.dll，mathv3.dll，RC6.dll，Windows_FormApi.dll
3.运行runtime_install.exe在线安装webview runtime

具体介绍:
kernel:
控制webview2内核的代码
client:
主程序代码，将kernel变为自己的主窗体
进程通信方式:共享文件


