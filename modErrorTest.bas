Attribute VB_Name = "modErrorTest"
' ==============================================================
' modErrorTest.bas
' Test cac truong hop loi co chu y - kiem tra DLL bao loi
' ro rang va VBA hieu duoc sai o dau.
'
' YEU CAU: modBridge.bas phai co trong cung workbook.
'
' MOI test Expected loi phai tra ve:
'   {"ok":false,"error":"mo ta loi ro rang"}
'
' Cach dung:
'   Goi Test_AllErrors trong Immediate Window
'   Kiem tra moi dong co chu thich "PASS/FAIL" dung khong
' ==============================================================

Option Explicit

' ==============================================================
' HELPER
' ==============================================================

' PrintError in ket qua va kiem tra co phai loi khong.
' PASS = {"ok":false} nhu mong doi
' FAIL = {"ok":true} hoac response bat thuong
Private Sub PrintError(ByVal sLabel As String, ByVal sJSON As String)
    Dim bOK As Boolean
    bOK = (InStr(sJSON, """ok"":true") > 0)
    Debug.Print "------------------------------------------------------------"
    If bOK Then
        Debug.Print "[FAIL - PHAI LA LOI] " & sLabel
    Else
        Debug.Print "[PASS - LOI DUNG] " & sLabel
    End If
    Debug.Print PrettyJSON(sJSON)
End Sub

' ==============================================================
' TEST LOI - NHOM 1: JSON / Action
' ==============================================================

Public Sub Test_Err_InvalidJSON()
    PrintError "JSON rong", ExecuteJSON("")
End Sub

Public Sub Test_Err_NotJSON()
    PrintError "Chuoi bat ky, khong phai JSON", ExecuteJSON("hello world")
End Sub

Public Sub Test_Err_MissingAction()
    PrintError "Thieu field action", ExecuteJSON("{""account"":""default""}")
End Sub

Public Sub Test_Err_UnknownAction()
    PrintError "Action khong ton tai", ExecuteJSON("{""action"":""xyz.abc""}")
End Sub

' ==============================================================
' TEST LOI - NHOM 2: Auth
' ==============================================================

Public Sub Test_Err_Auth_NotLoggedIn()
    PrintError "Sheets.read voi account chua login", ExecuteJSON( _
        "{""action"":""sheets.read"",""account"":""tai_khoan_ao""," & _
        """spreadsheetId"":""abc"",""range"":""Sheet1!A1""}")
End Sub

Public Sub Test_Err_Auth_RemoveNotExist()
    PrintError "accounts.remove account khong ton tai", ExecuteJSON( _
        "{""action"":""auth.accounts.remove"",""targetAccount"":""khong_co""}")
End Sub

' ==============================================================
' TEST LOI - NHOM 3: Sheets
' ==============================================================

Public Sub Test_Err_Sheets_MissingSpreadsheetId()
    PrintError "sheets.read thieu spreadsheetId", ExecuteJSON( _
        "{""action"":""sheets.read"",""range"":""Sheet1!A1""}")
End Sub

Public Sub Test_Err_Sheets_MissingRange()
    PrintError "sheets.read thieu range", ExecuteJSON( _
        "{""action"":""sheets.read"",""spreadsheetId"":""abc123""}")
End Sub

Public Sub Test_Err_Sheets_Create_MissingTitle()
    PrintError "sheets.create thieu title", ExecuteJSON("{""action"":""sheets.create""}")
End Sub

Public Sub Test_Err_Sheets_AddSheet_MissingTitle()
    PrintError "sheets.addSheet thieu title", ExecuteJSON( _
        "{""action"":""sheets.addSheet"",""spreadsheetId"":""abc123""}")
End Sub

' ==============================================================
' TEST LOI - NHOM 4: Drive
' ==============================================================

Public Sub Test_Err_Drive_MissingLocalPath()
    PrintError "drive.upload thieu localPath", ExecuteJSON("{""action"":""drive.upload""}")
End Sub

Public Sub Test_Err_Drive_FileNotFound_Upload()
    PrintError "drive.upload file khong ton tai tren may", ExecuteJSON( _
        "{""action"":""drive.upload"",""localPath"":""C:\\KhongTonTai\\abc.xlsx""}")
End Sub

Public Sub Test_Err_Drive_MissingFileId()
    PrintError "drive.get thieu fileId", ExecuteJSON("{""action"":""drive.get""}")
End Sub

Public Sub Test_Err_Drive_InvalidFileId()
    PrintError "drive.get fileId sai", ExecuteJSON( _
        "{""action"":""drive.get"",""fileId"":""FILE_ID_SAI_HOAN_TOAN""}")
End Sub

Public Sub Test_Err_Drive_Move_MissingNewParent()
    PrintError "drive.move thieu newParentId", ExecuteJSON( _
        "{""action"":""drive.move"",""fileId"":""abc123""}")
End Sub

Public Sub Test_Err_Drive_Trash_Missing()
    PrintError "drive.trash thieu fileId", ExecuteJSON("{""action"":""drive.trash""}")
End Sub

' ==============================================================
' TEST LOI - NHOM 5: Gmail
' ==============================================================

Public Sub Test_Err_Gmail_Send_MissingTo()
    PrintError "gmail.send thieu to", ExecuteJSON( _
        "{""action"":""gmail.send"",""subject"":""test"",""body"":""test""}")
End Sub

Public Sub Test_Err_Gmail_Read_MissingMessageId()
    PrintError "gmail.read thieu messageId", ExecuteJSON("{""action"":""gmail.read""}")
End Sub

Public Sub Test_Err_Gmail_Read_InvalidMessageId()
    PrintError "gmail.read messageId sai", ExecuteJSON( _
        "{""action"":""gmail.read"",""messageId"":""MSG_ID_SAI""}")
End Sub

Public Sub Test_Err_Gmail_Search_MissingQuery()
    PrintError "gmail.search thieu query", ExecuteJSON("{""action"":""gmail.search""}")
End Sub

Public Sub Test_Err_Gmail_Attachments_Missing()
    PrintError "gmail.attachments thieu messageId", ExecuteJSON("{""action"":""gmail.attachments""}")
End Sub

' ==============================================================
' TEST LOI - NHOM 6: Docs
' ==============================================================

Public Sub Test_Err_Docs_MissingDocumentId()
    PrintError "docs.read thieu documentId", ExecuteJSON("{""action"":""docs.read""}")
End Sub

Public Sub Test_Err_Docs_Create_MissingTitle()
    PrintError "docs.create thieu title", ExecuteJSON("{""action"":""docs.create""}")
End Sub

Public Sub Test_Err_Docs_Replace_MissingSearch()
    PrintError "docs.replaceText thieu searchText", ExecuteJSON( _
        "{""action"":""docs.replaceText"",""documentId"":""abc123"",""replaceText"":""xyz""}")
End Sub

' ==============================================================
' TEST LOI - NHOM 7: Calendar
' ==============================================================

Public Sub Test_Err_Calendar_Create_MissingSummary()
    PrintError "calendar.createEvent thieu summary", ExecuteJSON( _
        "{""action"":""calendar.createEvent""," & _
        """start"":""2026-07-01T09:00:00+07:00""," & _
        """end"":""2026-07-01T10:00:00+07:00""}")
End Sub

Public Sub Test_Err_Calendar_Create_MissingTime()
    PrintError "calendar.createEvent thieu start/end", ExecuteJSON( _
        "{""action"":""calendar.createEvent"",""summary"":""Test""}")
End Sub

Public Sub Test_Err_Calendar_Delete_MissingEventId()
    PrintError "calendar.deleteEvent thieu eventId", ExecuteJSON( _
        "{""action"":""calendar.deleteEvent"",""calendarId"":""primary""}")
End Sub

' ==============================================================
' CHAY TOAN BO
' ==============================================================

Public Sub Test_AllErrors()
    Debug.Print "============================================================"
    Debug.Print " TEST LOI CO CHU Y - TAT CA PHAI LA ok:false"
    Debug.Print "============================================================"

    Debug.Print ""
    Debug.Print "--- NHOM 1: JSON / Action ---"
    Test_Err_InvalidJSON
    Test_Err_NotJSON
    Test_Err_MissingAction
    Test_Err_UnknownAction

    Debug.Print ""
    Debug.Print "--- NHOM 2: Auth ---"
    Test_Err_Auth_NotLoggedIn
    Test_Err_Auth_RemoveNotExist

    Debug.Print ""
    Debug.Print "--- NHOM 3: Sheets ---"
    Test_Err_Sheets_MissingSpreadsheetId
    Test_Err_Sheets_MissingRange
    Test_Err_Sheets_Create_MissingTitle
    Test_Err_Sheets_AddSheet_MissingTitle

    Debug.Print ""
    Debug.Print "--- NHOM 4: Drive ---"
    Test_Err_Drive_MissingLocalPath
    Test_Err_Drive_FileNotFound_Upload
    Test_Err_Drive_MissingFileId
    Test_Err_Drive_Move_MissingNewParent
    Test_Err_Drive_Trash_Missing

    Debug.Print ""
    Debug.Print "--- NHOM 5: Gmail ---"
    Test_Err_Gmail_Send_MissingTo
    Test_Err_Gmail_Read_MissingMessageId
    Test_Err_Gmail_Search_MissingQuery
    Test_Err_Gmail_Attachments_Missing

    Debug.Print ""
    Debug.Print "--- NHOM 6: Docs ---"
    Test_Err_Docs_MissingDocumentId
    Test_Err_Docs_Create_MissingTitle
    Test_Err_Docs_Replace_MissingSearch

    Debug.Print ""
    Debug.Print "--- NHOM 7: Calendar ---"
    Test_Err_Calendar_Create_MissingSummary
    Test_Err_Calendar_Create_MissingTime
    Test_Err_Calendar_Delete_MissingEventId

    Debug.Print "============================================================"
    Debug.Print " HOAN THANH - Kiem tra cot [PASS/FAIL] o tren"
    Debug.Print " PASS = DLL bao loi dung nhu mong doi"
    Debug.Print " FAIL = DLL khong bao loi (van tra ok:true) - BUG"
    Debug.Print "============================================================"
End Sub
