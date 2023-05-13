项目名称:TLC浏览器(TLC_NLP机器人的附属产品) 
技术架构:webview2 runtime,可参考链接WebView2 - Microsoft Edge Developer

目录介绍:
kernel:控制webview2内核的代码
client:主程序代码，将kernel变为自己的子窗体，
进程通信方式:共享文件 release:正式版包
特性:抛弃vb6自带的ie7，使用webview2 runtime支持html5，支持多标签



软件使用教程:

窗体置顶方法:双击图标即可窗体置顶

多标签的使用:点击+图标可以添加新标签，双击标签则是删除标签 

标签被选中状态为白色，未被选中状态为蓝色

若出现任何问题可发送至github issue或者发送邮箱到xiaohui032901@foxmail.com



项目的部署:
clone项目后运行一键reg.bat即可


参考资料:https://www.vbforums.com/showthread.php?889202-VB6-WebView2-Binding-(Edge-Chromium) GitHub - sysdzw/WebView2DemoForVb6: WebView2Demo for vb6 
Developed by 福州机电工程职业技术学校 wh
进程间通信目前已完成winsock适配
支持windows7及以上系统
目前此项目已开发100小时+,单仍然可能存在很多问题定是存在很多问题的...，若存在问题可以发lssues
项目地址:https://github.com/keyxh/TLC_Browers/
