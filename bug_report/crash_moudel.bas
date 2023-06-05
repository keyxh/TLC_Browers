Attribute VB_Name = "crash_moudel"
Public crash_path As String

Sub main()
    If App.PrevInstance = True Then End '''监控程序只开一个'''
    crash_path = Environ("temp") + "\tlc_browser"
    If Dir(crash_path, vbDirectory) <> "" Then RmDir crash_path
    MkDir crash_path
    While Dir(crash_path + "\*.*") = ""
        DoEvents
    Wend
    If Dir(App.Path + "\crash.rar") Then Kill (App.Path + "\crash.rar")
    Shell "RAR.exe a -r crash.rar %temp%\tlc_browser\"
    Shell "RAR.exe a -r crash.rar .\logs\"
    crash_form.Show
End Sub

