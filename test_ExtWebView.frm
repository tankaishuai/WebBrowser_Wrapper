VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6795
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11235
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   11235
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_iCount As Long
Private webCtrl As New ExtWebViewLib.WebViewer
Private onEvent As New WebCtrlClass


'
'支持的浏览器内核                 // 需要的运行时库
'
'enum
'{
'    WEBVIEW_TYPE_MINIBLINK = 0, // node.dll
'    WEBVIEW_TYPE_IE = 1,        //
'    WEBVIEW_TYPE_CEF = 2,       // libcef.dll ... etc.
'    WEBVIEW_TYPE_IE_EMBED = 3,  //
'    WEBVIEW_TYPE_WEBKIT = 4,    // wke.dll
'};
'
Const webEngine As Long = 2


Private Sub Form_Load()
    m_iCount = 0
    
    ' 设置回调响应函数
    webCtrl.SetListener "OnWebCtrlEvent", Me
    
    ' 初始化
    cmd = "--parent_wnd=" + Hex(Me.hWnd) + " --tab_rect=0,0,800,600 --url=www.baidu.com"
    webCtrl.InitWebKit cmd, webEngine
    
    ' 注册扩展
    webCtrl.RegisterObject "msg_box", Me
End Sub

Public Function OnWebCtrlEvent(ByVal strEvent As String, ByVal strParam1 As String, ByVal strParam2 As String) As Variant
    'MsgBox strEvent
    If "OnUrlChanged" = strEvent Then
        ' 注册扩展
        webCtrl.RegisterObject "msg_box", Me
    
    ElseIf "OnDocumentReady" = strEvent Then
        If (0 = m_iCount) Then
            ' 再跳转
            webCtrl.LoadUrl App.Path + "\test_ExtWebView.html"
            MsgBox "hwnd = " + Hex(webCtrl.GetHWnd)
            
        ElseIf 1 = m_iCount Then
            ' 执行一个函数, 返回 execute_id
            ret = webCtrl.ExecJScript("getUserAgent()")
            MsgBox "execute_id = " + ret
        End If
        
        m_iCount = m_iCount + 1
        
    ElseIf "OnExecuteCallback" = strEvent Then
        ' 获取执行函数的结果
        ' strParam1 = execute_id
        MsgBox "ret = " + strParam1 + " , " + strParam2
        
    ElseIf "OnExecute" = strEvent Then
        ' 调用函数
        If "msg_box" = strParam1 Then
            MsgBox "call: " + strParam1 + "(" + strParam2 + ")"
            OnWebCtrlEvent = "return from vb6"
            Exit Function
        End If
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    webCtrl.UnInitWebKit
    Set webCtrl = Nothing
End Sub
