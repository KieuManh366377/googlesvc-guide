Attribute VB_Name = "modGmailTest"
' ==============================================================
' modGmailTest.bas
' Test cac action Gmail: send, list, search, read,
' markRead, markUnread, delete, attachments
'
' YEU CAU: modBridge.bas phai co trong cung workbook.
'
' QUAN TRONG - TRUOC KHI TEST:
'   Scope gmail.modify vua duoc them moi. Token cu (chi co
'   gmail.send) CHUA DU QUYEN. Can xoa token cu va dang nhap lai:
'     del tokens\default.json
'     Goi Test_Login trong modAuthTest
'
' THU TU TEST KHUYEN NGHI:
'   Test_GmailSend        -> gui email test
'   Test_GmailList        -> lay messageId
'   Test_GmailSearch      -> tim theo dieu kien
'   Test_GmailRead        -> doc noi dung day du
'   Test_GmailAttachments -> liet ke / download file dinh kem
'   Test_GmailMarkUnread  -> danh dau chua doc
'   Test_GmailMarkRead    -> danh dau da doc
'   Test_GmailDelete      -> chuyen vao Trash (can xac nhan)
' ==============================================================

Option Explicit

' ==============================================================
' CONST - sua truoc khi test
' ==============================================================

' Email nhan thu test
Private Const TEST_EMAIL As String = "kieumanh366377@gmail.com"

' messageId de test read/mark/delete/attachments
' Lay tu ket qua Test_GmailList hoac Test_GmailSearch
Private Const TEST_MSG_ID As String = "PUT_MESSAGE_ID_HERE"

' Thu muc luu file dinh kem khi download
Private Const ATTACH_SAVE_DIR As String = "C:\Temp\attachments\"

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

' Test_GmailSend: gui email test don gian (khong dinh kem)
' Ket qua mong doi: {"ok":true,"data":{"messageId":"...","threadId":"..."}}
Public Sub Test_GmailSend()
    PrintPretty "gmail.send", ExecuteJSON( _
        "{""action"":""gmail.send""," & _
        """to"":""" & TEST_EMAIL & """," & _
        """subject"":""Test GoogleSvc DLL""," & _
        """body"":""Day la email test tu GoogleSvc.dll - Dang ky thanh cong!""}")
End Sub

' Test_GmailSend_WithAttach: gui email co file dinh kem
' attachments: mang [{path, mimeType, name}]
Public Sub Test_GmailSend_WithAttach()
    Dim sAttachPath As String
    sAttachPath = "C:\\Temp\\test_upload.xlsx"   ' <-- sua duong dan file

    PrintPretty "gmail.send (co dinh kem)", ExecuteJSON( _
        "{""action"":""gmail.send""," & _
        """to"":""" & TEST_EMAIL & """," & _
        """subject"":""Test GoogleSvc co dinh kem""," & _
        """body"":""Email test co dinh kem xlsx""," & _
        """attachments"":[{" & _
            """path"":""" & sAttachPath & """," & _
            """mimeType"":""application/vnd.openxmlformats-officedocument.spreadsheetml.sheet""," & _
            """name"":""bao_cao.xlsx""}]}")
End Sub

' Test_GmailList: liet ke 10 email gan nhat
' Ket qua mong doi: [{""id"":""..."",""from"":""..."",""subject"":""...""}]
' SAU KHI CHAY: copy 1 messageId de dung cho Test_GmailRead
Public Sub Test_GmailList()
    PrintPretty "gmail.list (10 emails gan nhat)", _
        ExecuteJSON("{""action"":""gmail.list"",""maxResults"":10}")
End Sub

' Test_GmailList_Unread: chi lay email chua doc
Public Sub Test_GmailList_Unread()
    PrintPretty "gmail.list (unread)", _
        ExecuteJSON("{""action"":""gmail.list"",""query"":""is:unread"",""maxResults"":10}")
End Sub

' Test_GmailSearch: tim email theo tu khoa
Public Sub Test_GmailSearch()
    PrintPretty "gmail.search", ExecuteJSON( _
        "{""action"":""gmail.search""," & _
        """query"":""from:me subject:Test GoogleSvc""," & _
        """maxResults"":5}")
End Sub

' Test_GmailRead: doc noi dung day du 1 email
' Ket qua mong doi: {"id":"...","from":"...","subject":"...","body":"..."}
' Can: lay messageId tu Test_GmailList truoc
Public Sub Test_GmailRead()
    If TEST_MSG_ID = "PUT_MESSAGE_ID_HERE" Then
        Debug.Print "[gmail.read] Chua sua TEST_MSG_ID - lay tu Test_GmailList truoc"
        Exit Sub
    End If
    PrintPretty "gmail.read", ExecuteJSON("{""action"":""gmail.read"",""messageId"":""" & TEST_MSG_ID & """}")
End Sub

' Test_GmailMarkUnread: danh dau email la chua doc
' Ket qua mong doi: {"status":"da danh dau chua doc","messageId":"..."}
Public Sub Test_GmailMarkUnread()
    If TEST_MSG_ID = "PUT_MESSAGE_ID_HERE" Then
        Debug.Print "[gmail.markUnread] Chua sua TEST_MSG_ID"
        Exit Sub
    End If
    PrintPretty "gmail.markUnread", ExecuteJSON("{""action"":""gmail.markUnread"",""messageId"":""" & TEST_MSG_ID & """}")
End Sub

' Test_GmailMarkRead: danh dau email la da doc
' Ket qua mong doi: {"status":"da danh dau da doc","messageId":"..."}
Public Sub Test_GmailMarkRead()
    If TEST_MSG_ID = "PUT_MESSAGE_ID_HERE" Then
        Debug.Print "[gmail.markRead] Chua sua TEST_MSG_ID"
        Exit Sub
    End If
    PrintPretty "gmail.markRead", ExecuteJSON("{""action"":""gmail.markRead"",""messageId"":""" & TEST_MSG_ID & """}")
End Sub

' Test_GmailAttachments_List: liet ke file dinh kem cua 1 email (khong tai)
' Ket qua mong doi:
'   {"ok":true,"data":{"messageId":"...","attachments":[{"filename":"...","size":...}],"count":N}}
' Can: TEST_MSG_ID la email CO dinh kem
Public Sub Test_GmailAttachments_List()
    If TEST_MSG_ID = "PUT_MESSAGE_ID_HERE" Then
        Debug.Print "[gmail.attachments] Chua sua TEST_MSG_ID"
        Debug.Print "Can messageId cua email CO file dinh kem"
        Exit Sub
    End If
    ' savePath rong = chi liet ke, khong tai file
    PrintPretty "gmail.attachments (chi liet ke)", ExecuteJSON( _
        "{""action"":""gmail.attachments""," & _
        """messageId"":""" & TEST_MSG_ID & """}")
End Sub

' Test_GmailAttachments_Download: liet ke VA tai toan bo dinh kem ve may
' Ket qua mong doi:
'   {"ok":true,"data":{"attachments":[{"filename":"...","savedPath":"..."}],"count":N}}
' Can: TEST_MSG_ID la email CO dinh kem + ATTACH_SAVE_DIR phai ton tai
Public Sub Test_GmailAttachments_Download()
    If TEST_MSG_ID = "PUT_MESSAGE_ID_HERE" Then
        Debug.Print "[gmail.attachments] Chua sua TEST_MSG_ID"
        Exit Sub
    End If
    ' savePath co gia tri = tai tat ca dinh kem vao thu muc do
    PrintPretty "gmail.attachments (tai file)", ExecuteJSON( _
        "{""action"":""gmail.attachments""," & _
        """messageId"":""" & TEST_MSG_ID & """," & _
        """savePath"":""" & Replace(ATTACH_SAVE_DIR, "\", "\\") & """}")
End Sub

' Test_GmailDelete: chuyen email vao Trash (co the phuc hoi 30 ngay)
' CANH BAO: bo comment dong ExecuteJSON de chay that su
Public Sub Test_GmailDelete()
    Dim sMsgToDelete As String
    sMsgToDelete = "PUT_MESSAGE_ID_TO_DELETE_HERE"   ' <-- messageId can xoa

    If sMsgToDelete = "PUT_MESSAGE_ID_TO_DELETE_HERE" Then
        Debug.Print "[gmail.delete] Chua sua sMsgToDelete - an toan, bo qua"
        Exit Sub
    End If
    ' CANH BAO: bo comment dong duoi de chay that su
    ' PrintPretty "gmail.delete", ExecuteJSON("{""action"":""gmail.delete"",""messageId"":""" & sMsgToDelete & """}")
    Debug.Print "[gmail.delete] Bo comment dong PrintPretty de chay that su"
End Sub

' Test_AllGmail: chay cac test an toan (khong delete)
Public Sub Test_AllGmail()
    Debug.Print "============================================================"
    Debug.Print " TEST GMAIL ACTIONS (an toan)"
    Debug.Print "============================================================"
    Test_GmailSend
    Test_GmailList
    Test_GmailList_Unread
    Test_GmailSearch
    Debug.Print "============================================================"
    Debug.Print " HOAN THANH"
    Debug.Print " Tiep theo: lay messageId de test Read/Attachments/Mark"
    Debug.Print "============================================================"
End Sub
