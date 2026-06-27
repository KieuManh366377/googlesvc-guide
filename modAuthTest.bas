Attribute VB_Name = "modAuthTest"
' ==============================================================
' modAuthTest.bas
' Test cac action Auth: login, logout, status,
' accounts.list, accounts.remove
'
' YEU CAU: modBridge.bas phai co trong cung workbook.
'
' Cach dung:
'   Goi tung ham Test_* trong Immediate Window (Ctrl+G)
' ==============================================================

Option Explicit

' ==============================================================
' HELPER
' ==============================================================

Private Sub PrintPretty(ByVal sLabel As String, ByVal sJSON As String)
    Debug.Print "------------------------------------------------------------"
    Debug.Print "[" & sLabel & "]"
    Debug.Print PrettyJSON(sJSON)   ' PrettyJSON tu modBridge
End Sub

' ==============================================================
' TEST
' ==============================================================

' Test_Version: kiem tra DLL load thanh cong va tra ve version
Public Sub Test_Version()
    Debug.Print "GBridgeVersion = " & GBridgeVersion()
End Sub

' Test_Ping: kiem tra DLL dang chay
' Ket qua mong doi: {"ok":true,"data":{"pong":"googlesvc dang chay"}}
Public Sub Test_Ping()
    PrintPretty "ping", ExecuteJSON("{""action"":""ping""}")
End Sub

' Test_Login: mo browser dang nhap Google (blocking 300s)
' Chi goi 1 lan duy nhat - sau khi co token khong can goi lai.
Public Sub Test_Login()
    PrintPretty "auth.login", ExecuteJSON("{""action"":""auth.login"",""account"":""default""}")
End Sub

' Test_AuthStatus: kiem tra trang thai dang nhap account default
' Ket qua mong doi: {"ok":true,"data":{"loggedIn":true,"tokenValid":true,"expiry":"..."}}
Public Sub Test_AuthStatus()
    PrintPretty "auth.status (default)", _
        ExecuteJSON("{""action"":""auth.status"",""account"":""default""}")
End Sub

' Test_AuthStatus_NotLoggedIn: kiem tra account chua tung dang nhap
' Ket qua mong doi: {"ok":true,"data":{"loggedIn":false,"tokenValid":false,...}}
Public Sub Test_AuthStatus_NotLoggedIn()
    PrintPretty "auth.status (chua dang nhap)", _
        ExecuteJSON("{""action"":""auth.status"",""account"":""khong_ton_tai""}")
End Sub

' Test_AccountsList: liet ke tat ca account da dang nhap
' Ket qua mong doi: {"ok":true,"data":{"accounts":["default"],"count":1}}
Public Sub Test_AccountsList()
    PrintPretty "auth.accounts.list", _
        ExecuteJSON("{""action"":""auth.accounts.list""}")
End Sub

' Test_AccountsRemove: xoa 1 account cu the
' CANH BAO: goi ham nay se xoa token, phai dang nhap lai sau do.
Public Sub Test_AccountsRemove()
    Dim sTarget As String
    sTarget = "test_account"   ' <-- sua thanh account can xoa
    PrintPretty "auth.accounts.remove (" & sTarget & ")", _
        ExecuteJSON("{""action"":""auth.accounts.remove"",""targetAccount"":""" & sTarget & """}")
End Sub

' Test_AccountsRemove_NotExist: xoa account khong ton tai (kiem tra xu ly loi)
' Ket qua mong doi: {"ok":false,"error":"account ... khong ton tai hoac chua dang nhap"}
Public Sub Test_AccountsRemove_NotExist()
    PrintPretty "auth.accounts.remove (khong ton tai)", _
        ExecuteJSON("{""action"":""auth.accounts.remove"",""targetAccount"":""khong_ton_tai""}")
End Sub

' Test_Logout: dang xuat account default
' CANH BAO: phai dang nhap lai bang Test_Login sau khi goi.
Public Sub Test_Logout()
    PrintPretty "auth.logout (default)", _
        ExecuteJSON("{""action"":""auth.logout"",""account"":""default""}")
End Sub

' Test_AllAuth: chay toan bo luong Auth (khong bao gom Login/Logout)
Public Sub Test_AllAuth()
    Debug.Print "============================================================"
    Debug.Print " TEST TOAN BO AUTH ACTIONS"
    Debug.Print "============================================================"
    Test_Version
    Test_Ping
    Test_AuthStatus
    Test_AuthStatus_NotLoggedIn
    Test_AccountsList
    Test_AccountsRemove_NotExist
    Debug.Print "============================================================"
    Debug.Print " HOAN THANH"
    Debug.Print "============================================================"
End Sub
