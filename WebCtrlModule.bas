Attribute VB_Name = "WebCtrlModule"

'’‚¿Ô obj ±ÿ–Î byval

Public Declare Function CreateDispatchStd Lib "CreateDispatch.dll" Alias "_CreateDispatchStd@8" ( _
    ByVal str As String, _
    ByVal obj As Object _
    ) As Object
    
Public Declare Function CreateDispatchInvokerA Lib "win32exts_dll.dll" ( _
    ByVal str As String, _
    ByVal obj As Object _
    ) As Object


'=================================
