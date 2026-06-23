Attribute VB_Name = "modGBridgeTest"
Option Explicit
' ==========================================================
' modGBridgeTest.bas
' Giao tiep truc tiep voi GoogleSvc.dll:
'   - Declare ham export cua DLL
'   - Ham wrapper GBridgeCall (gui JSON, nhan JSON)
'   - Ham RPC tung nghiep vu (doc/ghi Sheet, gui mail...)
'   - Ham test chay thu tung buoc
'
' CACH SU DUNG LAN DAU:
'   Buoc 1 - Dam bao cau truc thu muc dung:
'            thu_muc\
'            |-- GoogleSvc.dll
'            |-- credentials.json   <- tai tu Google Cloud Console
'            `-- GoogleSvc.xlsm     <- file Excel nay
'
'   Buoc 2 - Mo Excel, nhan Alt+F11 mo VBE
'   Buoc 3 - Mo Immediate Window: Ctrl+G
'   Buoc 4 - Goi theo thu tu:
'            Test_Ping      <- kiem tra DLL hoat dong
'            Test_Login     <- dang nhap Google (mo browser)
'            Test_SheetsRead <- doc thu (sau khi sua SPREADSHEET_ID)
'
' YEU CAU: Excel 64-bit (DLL chi ho tro 64-bit).
'   Kiem tra: File > Account > About Excel - phai co chu "64-bit".
' ==========================================================

' ----------------------------------------------------------
' PHAN 1 - DECLARE
' Khai bao ham export cua GoogleSvc.dll.
' Signature phai khop chinh xac voi ham Go trong main.go:
'   GBridgeExecute(reqJsonPtr, outBuf *C.ushort, outBufSize int64) int64
'   GBridgeVersion() int64
' LongPtr = con tro 64-bit, LongLong = int64.
' ----------------------------------------------------------
Private Declare PtrSafe Function GBridgeExecute Lib "GoogleSvc.dll" _
    (ByVal reqJsonPtr As LongPtr, _
     ByVal outBuf As LongPtr, _
     ByVal outBufSize As LongLong) As LongLong

Private Declare PtrSafe Function GBridgeVersion Lib "GoogleSvc.dll" () As LongLong

Private Const SPREADSHEET_ID As String = "1punDQ-pRaMsopULsUaNthuHri5pj4ZrhNLFkOD3qkoo"  ' <- SUA O DAY
' ----------------------------------------------------------
' PHAN 2 - HAM WRAPPER CHINH: GBridgeCall
' Gui 1 JSON request, nhan ve JSON response dang String.
' Tu dong xu ly buffer: bat dau 64KB, cap lai neu khong du.
' ----------------------------------------------------------
Public Function GBridgeCall(ByVal reqJson As String) As String
    Const BUF_INIT As Long = 65536   ' 64K ky tu = du cho hau het response
    Dim outBuf()  As Integer          ' Integer 16-bit = uint16 (UTF-16)
    Dim bufChars  As Long
    Dim n         As LongLong

    bufChars = BUF_INIT
    ReDim outBuf(0 To bufChars - 1)

    n = GBridgeExecute(StrPtr(reqJson), VarPtr(outBuf(0)), CLngLng(bufChars))

    ' n < 0: buffer qua nho, |n| la so ky tu can
    If n < 0 Then
        bufChars = CLng(-n)
        ReDim outBuf(0 To bufChars - 1)
        n = GBridgeExecute(StrPtr(reqJson), VarPtr(outBuf(0)), CLngLng(bufChars))
        If n < 0 Then
            GBridgeCall = "{""ok"":false,""error"":""buffer qua nho ngay ca sau khi cap lai""}"
            Exit Function
        End If
    End If

    GBridgeCall = UInt16BufToStr(outBuf, CLng(n))
End Function

' UInt16BufToStr: chuyen mang Integer (UTF-16) thanh String VBA.
' VBA Integer la signed 16-bit nhung String cung luu UTF-16 noi bo,
' dung ChrW de ho tro Unicode day du (tieng Viet co dau).
Private Function UInt16BufToStr(buf() As Integer, ByVal charCount As Long) As String
    Dim s    As String
    Dim i    As Long
    Dim code As Long
    s = String$(charCount, vbNullChar)
    For i = 0 To charCount - 1
        code = buf(i)
        If code < 0 Then code = code + 65536  ' signed -> unsigned
        Mid$(s, i + 1, 1) = ChrW(code)
    Next i
    UInt16BufToStr = s
End Function

' ----------------------------------------------------------
' PHAN 3 - HAM RPC TUNG NGHIEP VU
' Cac ham nay xay dung JSON request roi goi GBridgeCall.
' Day la tang ma modSheetToExcel.bas su dung.
' ----------------------------------------------------------

' GSheetRead_RPC: doc 1 vung Google Sheet.
'   spreadsheetId : ID lay tu URL Sheet (phan giua /d/ va /edit)
'                   hoac URL day du (DLL tu tach ID)
'   rangeA1       : vung doc, vd "Sheet1!A1:C10"
'   account       : "default" neu chi dung 1 tai khoan Google
' Tra ve JSON response day du tu DLL.
Public Function GSheetRead_RPC(ByVal spreadsheetId As String, _
    ByVal rangeA1 As String, _
    Optional ByVal account As String = "default") As String

    Dim req As String
    req = "{""action"":""sheets.read""," & _
          """account"":""" & account & """," & _
          """spreadsheetId"":""" & spreadsheetId & """," & _
          """range"":""" & rangeA1 & """}"
    GSheetRead_RPC = GBridgeCall(req)
End Function

' GSheetWrite_RPC: ghi vao 1 vung Google Sheet (ghi de).
'   valuesJson : mang 2 chieu dang JSON, vd [[""A"",""B""],[1,2]]
' Tra ve JSON response tu DLL.
Public Function GSheetWrite_RPC(ByVal spreadsheetId As String, _
    ByVal rangeA1 As String, ByVal valuesJson As String, _
    Optional ByVal account As String = "default") As String

    Dim req As String
    req = "{""action"":""sheets.write""," & _
          """account"":""" & account & """," & _
          """spreadsheetId"":""" & spreadsheetId & """," & _
          """range"":""" & rangeA1 & """," & _
          """values"":" & valuesJson & "}"
    GSheetWrite_RPC = GBridgeCall(req)
End Function

' GSheetAppend_RPC: ghi THEM DONG moi vao cuoi Sheet (khong ghi de).
' Dung de them ban ghi moi, khong anh huong du lieu hien co.
Public Function GSheetAppend_RPC(ByVal spreadsheetId As String, _
    ByVal rangeA1 As String, ByVal valuesJson As String, _
    Optional ByVal account As String = "default") As String

    Dim req As String
    req = "{""action"":""sheets.append""," & _
          """account"":""" & account & """," & _
          """spreadsheetId"":""" & spreadsheetId & """," & _
          """range"":""" & rangeA1 & """," & _
          """values"":" & valuesJson & "}"
    GSheetAppend_RPC = GBridgeCall(req)
End Function

' GmailSend_RPC: gui email qua Gmail.
'   toAddr  : dia chi nguoi nhan, vd "nguoinhan@gmail.com"
'   subject : tieu de (ho tro tieng Viet co dau)
'   body    : noi dung email (plain text)
Public Function GmailSend_RPC(ByVal toAddr As String, _
    ByVal subject As String, ByVal body As String, _
    Optional ByVal account As String = "default") As String

    ' Xu ly ky tu dac biet trong JSON: chi escape dau " la du voi
    ' du lieu thong thuong. Neu body/subject chua ky tu dac biet phuc
    ' tap khac, nen dung thu vien JSON day du.
    Dim req As String
    req = "{""action"":""gmail.send""," & _
          """account"":""" & account & """," & _
          """to"":""" & toAddr & """," & _
          """subject"":""" & subject & """," & _
          """body"":""" & body & """}"
    GmailSend_RPC = GBridgeCall(req)
End Function

' ----------------------------------------------------------
' PHAN 4 - TEST (chay tung ham tu Immediate Window: Ctrl+G)
' ----------------------------------------------------------

' Test_Ping: kiem tra DLL da load va hoat dong.
' Ket qua mong doi: {"ok":true,"data":{"pong":"gbridge dang chay"}}
Public Sub Test_Ping()
    Debug.Print "=== Test_Ping ==="
    Debug.Print GBridgeCall("{""action"":""ping""}")
End Sub

' Test_Version: lay so phien ban giao thuc noi bo cua DLL.
Public Sub Test_Version()
    Debug.Print "=== Test_Version ==="
    Debug.Print "GBridgeVersion = " & GBridgeVersion()
End Sub

' Test_Login: dang nhap Google 1 lan duy nhat.
' Se tu mo browser -> ban dang nhap -> cap quyen -> dong browser.
' Token duoc luu trong thu muc tokens\ canh DLL, cac lan sau
' khong can dang nhap lai (tu dong refresh).
Public Sub Test_Login()
    Debug.Print "=== Test_Login ==="
    Debug.Print "Dang mo browser de dang nhap Google, vui long cho..."
    Dim resp As String
    resp = GBridgeCall("{""action"":""auth.login"",""account"":""default""}")
    Debug.Print resp
End Sub

' Test_SheetsRead: doc thu Google Sheet.
' SUA SPREADSHEET_ID_CUA_BAN truoc khi chay:
'   Lay tu URL: docs.google.com/spreadsheets/d/{ID O DAY}/edit
Public Sub Test_SheetsRead()
    Debug.Print "=== Test_SheetsRead ==="
    'Const SPREADSHEET_ID As String = "SPREADSHEET_ID_CUA_BAN"  ' <- SUA O DAY
    Dim resp As String
    resp = GSheetRead_RPC(SPREADSHEET_ID, "Sheet1!A1:C5")
    Debug.Print resp
End Sub

' Test_SheetsWrite: ghi thu vao Google Sheet.
' SUA SPREADSHEET_ID_CUA_BAN truoc khi chay.
Public Sub Test_SheetsWrite()
    Debug.Print "=== Test_SheetsWrite ==="
    'Const SPREADSHEET_ID As String = "SPREADSHEET_ID_CUA_BAN"  ' <- SUA O DAY
    Dim resp As String
    resp = GSheetWrite_RPC(SPREADSHEET_ID, "Sheet1!A1:B2", _
           "[[""Xin chao"",""GoogleSvc""],[1,2]]")
    Debug.Print resp
End Sub

' Test_SheetsAppend: ghi them 1 dong moi (khong ghi de dong cu).
Public Sub Test_SheetsAppend()
    Debug.Print "=== Test_SheetsAppend ==="
    'Const SPREADSHEET_ID As String = "SPREADSHEET_ID_CUA_BAN"  ' <- SUA O DAY
    Dim resp As String
    resp = GSheetAppend_RPC(SPREADSHEET_ID, "Sheet1!A:B", _
           "[[""Dong moi"",""Tu Excel""]]")
    Debug.Print resp
End Sub

' Test_GmailSend: gui thu test.
' SUA DIA CHI EMAIL truoc khi chay.
Public Sub Test_GmailSend()
    Debug.Print "=== Test_GmailSend ==="
    Dim resp As String
    resp = GmailSend_RPC( _
        "EMAIL_NGUOI_NHAN@gmail.com", _
        "Test tu GoogleSvc.dll", _
        "Day la mail test gui tu Excel qua Go DLL.")
    Debug.Print resp
End Sub

' Test_All: chay Ping + Version (an toan, khong can dang nhap).
' Cac test khac comment out - bo comment khi can test rieng.
Public Sub Test_All()
    Test_Version
    Test_Ping
    ' Cac test duoi day can dang nhap (Test_Login) va sua ID/email:
    ' Test_SheetsRead
    ' Test_SheetsWrite
    ' Test_SheetsAppend
    ' Test_GmailSend
End Sub
