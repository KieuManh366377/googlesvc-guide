Attribute VB_Name = "modSheetsTest"
' ==============================================================
' modSheetsTest.bas
' Test cac action Sheets: read, write, append, clear,
' getInfo, create, addSheet, rename, delete, batchUpdate
'
' YEU CAU: modBridge.bas phai co trong cung workbook.
'
' Cach dung:
'   1. Sua SHEET_ID thanh Spreadsheet ID thuc te
'   2. Goi tung ham Test_* trong Immediate Window (Ctrl+G)
' ==============================================================

Option Explicit

' ==============================================================
' CONST - sua truoc khi test
' ==============================================================

Private Const SHEET_ID As String = "PUT_YOUR_SPREADSHEET_ID_HERE"

' ==============================================================
' HELPER
' ==============================================================

Private Sub PrintPretty(ByVal sLabel As String, ByVal sJSON As String)
    Debug.Print "------------------------------------------------------------"
    Debug.Print "[" & sLabel & "]"
    Debug.Print PrettyJSON(sJSON)
End Sub

Private Function MakeReq(ByVal sAction As String, ByVal sExtra As String) As String
    If sExtra = "" Then
        MakeReq = "{""action"":""" & sAction & """,""spreadsheetId"":""" & SHEET_ID & """}"
    Else
        MakeReq = "{""action"":""" & sAction & """,""spreadsheetId"":""" & SHEET_ID & """," & sExtra & "}"
    End If
End Function

' ==============================================================
' TEST
' ==============================================================

' Test_SheetsRead: doc du lieu 1 vung
' Ket qua mong doi: {"ok":true,"data":{"values":[...],"range":"..."}}
Public Sub Test_SheetsRead()
    PrintPretty "sheets.read", ExecuteJSON(MakeReq("sheets.read", """range"":""Sheet1!A1:C5"""))
End Sub

' Test_SheetsWrite: ghi du lieu vao 1 vung
' Ket qua mong doi: {"ok":true,"data":{"updatedRange":"...","updatedRows":2,...}}
Public Sub Test_SheetsWrite()
    PrintPretty "sheets.write", ExecuteJSON(MakeReq("sheets.write", _
        """range"":""Sheet1!A1"",""values"":[[""Ten"",""Tuoi""],[""An"",25]]"))
End Sub

' Test_SheetsAppend: them dong moi vao cuoi du lieu
' Ket qua mong doi: {"ok":true,"data":{"updatedRows":1,...}}
Public Sub Test_SheetsAppend()
    PrintPretty "sheets.append", ExecuteJSON(MakeReq("sheets.append", _
        """range"":""Sheet1!A1"",""values"":[[""Binh"",30]]"))
End Sub

' Test_SheetsClear: xoa du lieu trong 1 vung (giu dinh dang)
' Ket qua mong doi: {"ok":true,"data":{"clearedRange":"..."}}
Public Sub Test_SheetsClear()
    PrintPretty "sheets.clear", ExecuteJSON(MakeReq("sheets.clear", """range"":""Sheet1!A1:C10"""))
End Sub

' Test_SheetsGetInfo: lay thong tin spreadsheet va danh sach tab
' Ket qua mong doi: {"ok":true,"data":{"title":"...","sheets":[{"sheetId":0,"title":"Sheet1",...}]}}
Public Sub Test_SheetsGetInfo()
    PrintPretty "sheets.getInfo", ExecuteJSON(MakeReq("sheets.getInfo", ""))
End Sub

' Test_SheetsCreate: tao spreadsheet moi
' Ket qua mong doi: {"ok":true,"data":{"spreadsheetId":"...","title":"...","webViewLink":"..."}}
' SAU KHI CHAY: copy spreadsheetId de dung cho Test_SheetsRename/Delete
Public Sub Test_SheetsCreate()
    PrintPretty "sheets.create", ExecuteJSON("{""action"":""sheets.create"",""title"":""Test GoogleSvc""}")
End Sub

' Test_SheetsAddSheet: them tab moi vao spreadsheet co san
' Ket qua mong doi: {"ok":true,"data":{"sheetId":...,"title":"BaoCao",...}}
Public Sub Test_SheetsAddSheet()
    PrintPretty "sheets.addSheet", ExecuteJSON(MakeReq("sheets.addSheet", _
        """title"":""BaoCao"",""rowCount"":500,""colCount"":10"))
End Sub

' Test_SheetsRename: doi ten 1 tab (sheetId=0 = tab dau tien)
' Ket qua mong doi: {"ok":true,"data":{"status":"doi ten thanh cong",...}}
Public Sub Test_SheetsRename()
    PrintPretty "sheets.rename", ExecuteJSON(MakeReq("sheets.rename", """sheetId"":0,""newTitle"":""DuLieu"""))
End Sub

' Test_SheetsDelete: xoa 1 tab (spreadsheet phai con >= 1 tab)
' CANH BAO: bo comment dong ExecuteJSON de chay that su.
Public Sub Test_SheetsDelete()
    Dim sTmpID As String
    sTmpID = "PUT_NEW_SPREADSHEET_ID_HERE"   ' <-- ID sheet muon xoa tab
    Dim sReq As String
    sReq = "{""action"":""sheets.delete"",""spreadsheetId"":""" & sTmpID & """,""sheetId"":0}"
    ' CANH BAO: bo comment dong duoi de chay that su
    ' PrintPretty "sheets.delete", ExecuteJSON(sReq)
    Debug.Print "[sheets.delete] Bo comment dong PrintPretty de chay that su"
End Sub

' Test_SheetsBatchUpdate: to mau nen xanh + chu trang cho hang dau tien
' Ket qua mong doi: {"ok":true,"data":{"replyCount":1}}
Public Sub Test_SheetsBatchUpdate()
    PrintPretty "sheets.batchUpdate (header format)", ExecuteJSON(MakeReq("sheets.batchUpdate", _
        """requests"":[{""repeatCell"":{" & _
            """range"":{""sheetId"":0,""startRowIndex"":0,""endRowIndex"":1}," & _
            """cell"":{""userEnteredFormat"":{" & _
                """backgroundColor"":{""red"":0.2,""green"":0.5,""blue"":0.9}," & _
                """textFormat"":{""bold"":true,""foregroundColor"":{""red"":1,""green"":1,""blue"":1}}" & _
            "}}," & _
            """fields"":""userEnteredFormat(backgroundColor,textFormat)""" & _
        "}}]"))
End Sub

' Test_AllSheets: chay toan bo luong Sheets (khong bao gom Delete)
Public Sub Test_AllSheets()
    Debug.Print "============================================================"
    Debug.Print " TEST TOAN BO SHEETS ACTIONS"
    Debug.Print "============================================================"
    Test_SheetsGetInfo
    Test_SheetsWrite
    Test_SheetsRead
    Test_SheetsAppend
    Test_SheetsBatchUpdate
    Test_SheetsClear
    Test_SheetsCreate
    Test_SheetsAddSheet
    Debug.Print "============================================================"
    Debug.Print " HOAN THANH - Kiem tra Immediate Window"
    Debug.Print "============================================================"
End Sub
