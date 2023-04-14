# TLC_Browers
一个开源的vb6基于webview2支持h5的浏览器
环境配置:
1.找到dlls文件夹
2.使用regsvr32.exe注册file_controlv2.dll，mathv3.dll，RC6.dll，Windows_FormApi.dll
3.运行runtime_install.exe在线安装webview runtime

具体介绍:
kernel:控制webview2内核的代码
client:主程序代码，将kernel变为自己的主窗体，进程通信方式:共享文件
release:正式版包

项目名称:TLC浏览器(TLC_Nlp机器人的附属产品)
技术架构:在vb6使用webview2 runtime、以及windows api对窗体的操作
项目特点:使用vb6开发，可实现多标签的效果,接入nlp大模型(实现中)
参考资料:https://www.vbforums.com/showthread.php?889202-VB6-WebView2-Binding-(Edge-Chromium)
power by 福州机电工程职业技术学校 wh
命令通过共享文件形式，也可以适配成winsock有时候不太稳定，所以采用共享文件
注册windows7及以上系统

目前此项目仅开发了3个小时，肯定是存在很多问题的...
