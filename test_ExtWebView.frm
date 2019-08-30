VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6792
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   11232
   LinkTopic       =   "Form1"
   ScaleHeight     =   6792
   ScaleWidth      =   11232
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
    
    cmd = "--parent_wnd=" + Hex(Me.hWnd) + " --tab_rect=0,0,800,600 --url=www.baidu.com"
    webCtrl.InitWebKit cmd, webEngine
    
    webCtrl.SetListener "OnWebCtrlEvent", Me      'CreateDispatchInvokerA("OnWebCtrlEvent", Me)
End Sub

Public Function OnWebCtrlEvent(ByVal strEvent As String, ByVal strParam1 As String, ByVal strParam2 As String) As Long
    OnWebCtrlEvent = 0
    MsgBox strEvent
    
    If "OnDocumentReady" = strEvent And (0 = m_iCount) Then
        m_iCount = m_iCount + 1
        
        ' 再跳转到 www.qq.com
        webCtrl.LoadUrl "www.qq.com"
        MsgBox "hwnd = " + Hex(webCtrl.GetHWnd)
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    webCtrl.UnInitWebKit
    Set webCtrl = Nothing
End Sub
