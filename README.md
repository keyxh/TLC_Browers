项目名称:TLC浏览器(TLC_NLP机器人的附属产品) 
技术架构:webview2 runtime,可参考链接WebView2 - Microsoft Edge Developer

目录介绍:
kernel:控制webview2内核的代码
client:主程序代码，将kernel变为自己的子窗体，
进程通信方式:共享文件 release:正式版包
特性:抛弃vb6自带的ie7，使用webview2 runtime支持html5，支持多标签



项目的部署:
1.找到dlls目录，分别使用regsvr32.exe注册file_controlv2.dll，mathv3.dll，RC6.dll，Windows_FormApi.dll
2.运行runtime_install.exe在线安装webview runtime，
3.在主程序文件目录创建temp文件夹和logs文件夹(若需要更改kernel或者client也需要在对于工程目录创建这两个文件夹)，即可运行主程序:TLC_Brower.exe



参考资料:https://www.vbforums.com/showthread.php?889202-VB6-WebView2-Binding-(Edge-Chromium) GitHub - sysdzw/WebView2DemoForVb6: WebView2Demo for vb6 
Developed by 福州机电工程职业技术学校 wh
进程间通信通过共享文件形式，也可以适配成winsock传递，但是不太稳定(目前仍然在开发中，后续会出)。
支持windows7及以上系统
目前此项目仅开发了3个小时，肯定是存在很多问题的...，若存在问题可以发lssues
项目地址:https://github.com/keyxh/TLC_Browers/
