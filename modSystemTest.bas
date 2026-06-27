Attribute VB_Name = "modSystemTest"
' ==============================================================
' modSystemTest.bas
' Test cac action System: version, info, health, scopes
'
' YEU CAU: modBridge.bas phai co trong cung workbook.
'
' Cach dung:
'   Goi tung ham hoac Test_AllSystem trong Immediate Window
' ==============================================================

Option Explicit

' ==============================================================
' HELPER
' ==============================================================

Private Sub PrintPretty(ByVal sLabel As String, ByVal sJSON As String)
    Debug.Print "------------------------------------------------------------"
    Debug.Print "[" & sLabel & "]"
    Debug.Print PrettyJSON(sJSON)
End Sub

' ==============================================================
' TEST
' ==============================================================

' Test_SystemVersion: phien ban DLL, protocol, Go version
' Ket qua mong doi:
'   {"dllVersion":"1.0.0","protocolVersion":1,"goVersion":"go1.22.x"}
Public Sub Test_SystemVersion()
    PrintPretty "system.version", ExecuteJSON("{""action"":""system.version""}")
End Sub

' Test_SystemInfo: thu muc DLL, OS, CPU, gio hien tai
' Ket qua mong doi:
'   {"dllDir":"D:\...","tokensDir":"...","goos":"windows","numCpu":N}
Public Sub Test_SystemInfo()
    PrintPretty "system.info", ExecuteJSON("{""action"":""system.info""}")
End Sub

' Test_SystemHealth: kiem tra credentials.json, token, tokens/
' Ket qua mong doi:
'   {"healthy":true,"items":[{"name":"credentials.json","ok":true},
'    {"name":"token:default","ok":true,"detail":"hop le..."}]}
Public Sub Test_SystemHealth()
    PrintPretty "system.health", ExecuteJSON("{""action"":""system.health""}")
End Sub

' Test_SystemScopes: liet ke OAuth scope dang cau hinh
' Ket qua mong doi:
'   {"count":6,"scopes":[{"scope":"...","service":"...","restricted":...}]}
Public Sub Test_SystemScopes()
    PrintPretty "system.scopes", ExecuteJSON("{""action"":""system.scopes""}")
End Sub

' Test_AllSystem: chay toan bo system actions
Public Sub Test_AllSystem()
    Debug.Print "============================================================"
    Debug.Print " TEST SYSTEM ACTIONS"
    Debug.Print "============================================================"
    Test_SystemVersion
    Test_SystemInfo
    Test_SystemHealth
    Test_SystemScopes
    Debug.Print "============================================================"
    Debug.Print " HOAN THANH"
    Debug.Print "============================================================"
End Sub
