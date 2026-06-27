Attribute VB_Name = "modDocsCalendarTest"
' ==============================================================
' modDocsCalendarTest.bas
' Test cac action Docs va Calendar.
'
' YEU CAU: modBridge.bas phai co trong cung workbook.
'
' Docs actions:
'   docs.create      - tao tai lieu moi
'   docs.read        - doc noi dung (text thuan)
'   docs.insertText  - chen text vao vi tri chi dinh
'   docs.replaceText - tim va thay the text trong toan tai lieu
'
' Calendar actions:
'   calendar.list        - liet ke cac calendar cua user
'   calendar.events      - lay su kien trong khoang thoi gian
'   calendar.createEvent - tao su kien moi
'   calendar.updateEvent - cap nhat su kien
'   calendar.deleteEvent - xoa su kien
'
' LUU Y SCOPE:
'   Token cu co the chua co scope "documents" va "calendar".
'   Neu nhan loi 403/insufficient scope, xoa token cu va dang nhap lai:
'     del tokens\default.json  (trong thu muc cung GoogleSvc.dll)
'     Goi Test_Login trong modAuthTest
'
' DINH DANG THOI GIAN:
'   RFC3339: "2026-07-01T09:00:00+07:00"
'   TimeZone: "Asia/Ho_Chi_Minh"
'
' THU TU TEST DOCS:
'   Test_DocsCreate      -> lay documentId
'   Test_DocsInsertText  -> chen text vao doc vua tao
'   Test_DocsReplaceText -> thay the text trong doc
'   Test_DocsRead        -> doc lai noi dung kiem tra
'
' THU TU TEST CALENDAR:
'   Test_CalendarList    -> lay calendarId (thuong = "primary")
'   Test_CalendarEvents  -> xem su kien hien tai
'   Test_CalendarCreate  -> tao su kien moi -> lay eventId
'   Test_CalendarUpdate  -> cap nhat su kien vua tao
'   Test_CalendarDelete  -> xoa su kien vua tao (can xac nhan)
' ==============================================================

Option Explicit

' ==============================================================
' CONST - sua truoc khi test
' ==============================================================

' documentId de test docs.read / insertText / replaceText
' Lay tu ket qua Test_DocsCreate
Private Const TEST_DOC_ID As String = "PUT_DOCUMENT_ID_HERE"

' calendarId de test events/create/update/delete
' "primary" = calendar chinh cua user (thuong dung nay)
Private Const TEST_CALENDAR_ID As String = "primary"

' eventId de test calendar.update / calendar.delete
' Lay tu ket qua Test_CalendarCreate
Private Const TEST_EVENT_ID As String = "PUT_EVENT_ID_HERE"

' ==============================================================
' HELPER
' ==============================================================

Private Sub PrintPretty(ByVal sLabel As String, ByVal sJSON As String)
    Debug.Print "------------------------------------------------------------"
    Debug.Print "[" & sLabel & "]"
    Debug.Print PrettyJSON(sJSON)   ' PrettyJSON tu modBridge
End Sub

' ==============================================================
' TEST - DOCS
' ==============================================================

' Test_DocsCreate: tao tai lieu Google Docs moi
' Ket qua mong doi:
'   {"ok":true,"data":{"documentId":"...","title":"Test GoogleSvc Docs","webViewLink":"..."}}
' SAU KHI CHAY: copy documentId, paste vao hang so TEST_DOC_ID
Public Sub Test_DocsCreate()
    Dim sReq As String
    sReq = "{""action"":""docs.create"",""title"":""Test GoogleSvc Docs""}"
    PrintPretty "docs.create", ExecuteJSON(sReq)
End Sub

' Test_DocsRead: doc noi dung tai lieu (text thuan, strip formatting)
' Ket qua mong doi:
'   {"ok":true,"data":{"documentId":"...","title":"...","text":"...","wordCount":N}}
' Can: TEST_DOC_ID da duoc dat
Public Sub Test_DocsRead()
    If TEST_DOC_ID = "PUT_DOCUMENT_ID_HERE" Then
        Debug.Print "[docs.read] Chua sua TEST_DOC_ID - chay Test_DocsCreate truoc"
        Exit Sub
    End If
    Dim sReq As String
    sReq = "{""action"":""docs.read"",""documentId"":""" & TEST_DOC_ID & """}"
    PrintPretty "docs.read", ExecuteJSON(sReq)
End Sub

' Test_DocsInsertText: chen text vao cuoi tai lieu (index = -1)
' Ket qua mong doi:
'   {"ok":true,"data":{"documentId":"...","index":-1,"textLength":N}}
' Can: TEST_DOC_ID da duoc dat
Public Sub Test_DocsInsertText()
    If TEST_DOC_ID = "PUT_DOCUMENT_ID_HERE" Then
        Debug.Print "[docs.insertText] Chua sua TEST_DOC_ID"
        Exit Sub
    End If
    Dim sReq As String
    ' index = 0 trong JSON -> router.go doi thanh -1 = chen vao cuoi
    sReq = "{""action"":""docs.insertText""," & _
           """documentId"":""" & TEST_DOC_ID & """," & _
           """text"":""Day la dong text tu GoogleSvc.dll\n""," & _
           """index"":0}"
    PrintPretty "docs.insertText (chen vao cuoi)", ExecuteJSON(sReq)
End Sub

' Test_DocsInsertText_AtBegin: chen text vao dau tai lieu (index = 1)
Public Sub Test_DocsInsertText_AtBegin()
    If TEST_DOC_ID = "PUT_DOCUMENT_ID_HERE" Then
        Debug.Print "[docs.insertText] Chua sua TEST_DOC_ID"
        Exit Sub
    End If
    Dim sReq As String
    sReq = "{""action"":""docs.insertText""," & _
           """documentId"":""" & TEST_DOC_ID & """," & _
           """text"":""[HEADER] GoogleSvc Test Doc\n""," & _
           """index"":1}"
    PrintPretty "docs.insertText (index=1, dau tai lieu)", ExecuteJSON(sReq)
End Sub

' Test_DocsReplaceText: tim va thay the text trong toan tai lieu
' Ket qua mong doi:
'   {"ok":true,"data":{"documentId":"...","occurrencesChanged":N}}
' Can: TEST_DOC_ID da co noi dung (chay Test_DocsInsertText truoc)
Public Sub Test_DocsReplaceText()
    If TEST_DOC_ID = "PUT_DOCUMENT_ID_HERE" Then
        Debug.Print "[docs.replaceText] Chua sua TEST_DOC_ID"
        Exit Sub
    End If
    Dim sReq As String
    sReq = "{""action"":""docs.replaceText""," & _
           """documentId"":""" & TEST_DOC_ID & """," & _
           """searchText"":""GoogleSvc.dll""," & _
           """replaceText"":""GoogleSvc v1.0""," & _
           """matchCase"":false}"
    PrintPretty "docs.replaceText", ExecuteJSON(sReq)
End Sub

' Test_AllDocs: chay toan bo luong Docs theo thu tu dung
' LUU Y: can set TEST_DOC_ID truoc khi chay (tru Test_DocsCreate)
Public Sub Test_AllDocs()
    Debug.Print "============================================================"
    Debug.Print " TEST DOCS ACTIONS"
    Debug.Print "============================================================"
    Test_DocsCreate
    Debug.Print ""
    Debug.Print " >>> Copy documentId tu ket qua tren, paste vao TEST_DOC_ID <<<"
    Debug.Print " >>> Sau do goi tung ham Test_Docs* rieng le <<<"
    Debug.Print "============================================================"
    Debug.Print " HOAN THANH BUOC 1 (Create)"
    Debug.Print "============================================================"
End Sub

' Test_AllDocs_WithID: chay het Docs khi da co TEST_DOC_ID
Public Sub Test_AllDocs_WithID()
    If TEST_DOC_ID = "PUT_DOCUMENT_ID_HERE" Then
        Debug.Print "Chua sua TEST_DOC_ID - chay Test_DocsCreate truoc"
        Exit Sub
    End If
    Debug.Print "============================================================"
    Debug.Print " TEST DOCS ACTIONS (co documentId)"
    Debug.Print "============================================================"
    Test_DocsInsertText_AtBegin
    Test_DocsInsertText
    Test_DocsReplaceText
    Test_DocsRead
    Debug.Print "============================================================"
    Debug.Print " HOAN THANH"
    Debug.Print "============================================================"
End Sub

' ==============================================================
' TEST - CALENDAR
' ==============================================================

' Test_CalendarList: liet ke cac calendar cua user
' Ket qua mong doi:
'   {"ok":true,"data":{"calendars":[{"id":"primary","summary":"...","primary":true}],"count":N}}
Public Sub Test_CalendarList()
    Dim sReq As String
    sReq = "{""action"":""calendar.list""}"
    PrintPretty "calendar.list", ExecuteJSON(sReq)
End Sub

' Test_CalendarEvents: lay su kien 30 ngay toi (mac dinh)
' Ket qua mong doi:
'   {"ok":true,"data":{"calendarId":"primary","events":[...],"count":N}}
Public Sub Test_CalendarEvents()
    Dim sReq As String
    sReq = "{""action"":""calendar.events""," & _
           """calendarId"":""" & TEST_CALENDAR_ID & """," & _
           """maxResults"":10}"
    PrintPretty "calendar.events (30 ngay toi)", ExecuteJSON(sReq)
End Sub

' Test_CalendarEvents_Range: su kien trong khoang thoi gian cu the
Public Sub Test_CalendarEvents_Range()
    Dim sReq As String
    ' Lay su kien tu 01/07/2026 den 31/07/2026
    sReq = "{""action"":""calendar.events""," & _
           """calendarId"":""" & TEST_CALENDAR_ID & """," & _
           """timeMin"":""2026-07-01T00:00:00+07:00""," & _
           """timeMax"":""2026-07-31T23:59:59+07:00""," & _
           """maxResults"":20}"
    PrintPretty "calendar.events (thang 7/2026)", ExecuteJSON(sReq)
End Sub

' Test_CalendarCreate: tao su kien moi
' Ket qua mong doi:
'   {"ok":true,"data":{"eventId":"...","summary":"...","start":"...","htmlLink":"..."}}
' SAU KHI CHAY: copy eventId, paste vao TEST_EVENT_ID
Public Sub Test_CalendarCreate()
    Dim sReq As String
    ' Su kien ngay mai 9:00-10:00 (gio Viet Nam)
    Dim dTomorrow As Date
    dTomorrow = Date + 1
    Dim sDate As String
    sDate = Format(Year(dTomorrow), "0000") & "-" & _
            Format(Month(dTomorrow), "00") & "-" & _
            Format(Day(dTomorrow), "00")

    sReq = "{""action"":""calendar.createEvent""," & _
           """calendarId"":""" & TEST_CALENDAR_ID & """," & _
           """summary"":""Test GoogleSvc DLL""," & _
           """description"":""Su kien test tao tu GoogleSvc.dll""," & _
           """location"":""Ho Chi Minh City""," & _
           """start"":""" & sDate & "T09:00:00+07:00""," & _
           """end"":""" & sDate & "T10:00:00+07:00""," & _
           """timeZone"":""Asia/Ho_Chi_Minh""}"
    PrintPretty "calendar.createEvent", ExecuteJSON(sReq)
End Sub

' Test_CalendarCreate_AllDay: tao su kien ca ngay
Public Sub Test_CalendarCreate_AllDay()
    Dim sReq As String
    Dim dTomorrow As Date
    dTomorrow = Date + 2
    Dim sDate As String
    sDate = Format(Year(dTomorrow), "0000") & "-" & _
            Format(Month(dTomorrow), "00") & "-" & _
            Format(Day(dTomorrow), "00")

    sReq = "{""action"":""calendar.createEvent""," & _
           """calendarId"":""" & TEST_CALENDAR_ID & """," & _
           """summary"":""Ngay nghi (test GoogleSvc)""," & _
           """start"":""" & sDate & """," & _
           """end"":""" & sDate & """," & _
           """allDay"":true}"
    PrintPretty "calendar.createEvent (all-day)", ExecuteJSON(sReq)
End Sub

' Test_CalendarUpdate: cap nhat su kien (sua tieu de va them mo ta)
' Can: TEST_EVENT_ID da duoc dat
Public Sub Test_CalendarUpdate()
    If TEST_EVENT_ID = "PUT_EVENT_ID_HERE" Then
        Debug.Print "[calendar.updateEvent] Chua sua TEST_EVENT_ID - chay Test_CalendarCreate truoc"
        Exit Sub
    End If
    Dim sReq As String
    sReq = "{""action"":""calendar.updateEvent""," & _
           """calendarId"":""" & TEST_CALENDAR_ID & """," & _
           """eventId"":""" & TEST_EVENT_ID & """," & _
           """summary"":""Test GoogleSvc DLL [Updated]""," & _
           """description"":""Da cap nhat tu GoogleSvc.dll""}"
    PrintPretty "calendar.updateEvent", ExecuteJSON(sReq)
End Sub

' Test_CalendarDelete: xoa su kien
' CANH BAO: bo comment dong ExecuteJSON de chay that su
Public Sub Test_CalendarDelete()
    If TEST_EVENT_ID = "PUT_EVENT_ID_HERE" Then
        Debug.Print "[calendar.deleteEvent] Chua sua TEST_EVENT_ID"
        Exit Sub
    End If
    Dim sReq As String
    sReq = "{""action"":""calendar.deleteEvent""," & _
           """calendarId"":""" & TEST_CALENDAR_ID & """," & _
           """eventId"":""" & TEST_EVENT_ID & """}"

    ' CANH BAO: bo comment dong duoi de chay that su
    ' PrintPretty "calendar.deleteEvent", ExecuteJSON(sReq)
    Debug.Print "[calendar.deleteEvent] Bo comment dong PrintPretty de chay that su"
    Debug.Print "EventId se xoa: " & TEST_EVENT_ID
End Sub

' Test_AllCalendar: chay cac test an toan (khong delete)
Public Sub Test_AllCalendar()
    Debug.Print "============================================================"
    Debug.Print " TEST CALENDAR ACTIONS"
    Debug.Print "============================================================"
    Test_CalendarList
    Test_CalendarEvents
    Test_CalendarCreate
    Debug.Print ""
    Debug.Print " >>> Copy eventId tu ket qua CalendarCreate <<<"
    Debug.Print " >>> Paste vao TEST_EVENT_ID roi goi Test_CalendarUpdate <<<"
    Debug.Print "============================================================"
    Debug.Print " HOAN THANH (an toan - chua chay Update/Delete)"
    Debug.Print "============================================================"
End Sub

' Test_AllCalendar_WithID: chay het Calendar khi da co TEST_EVENT_ID
Public Sub Test_AllCalendar_WithID()
    If TEST_EVENT_ID = "PUT_EVENT_ID_HERE" Then
        Debug.Print "Chua sua TEST_EVENT_ID - chay Test_CalendarCreate truoc"
        Exit Sub
    End If
    Debug.Print "============================================================"
    Debug.Print " TEST CALENDAR ACTIONS (co eventId)"
    Debug.Print "============================================================"
    Test_CalendarUpdate
    Test_CalendarEvents
    Debug.Print "============================================================"
    Debug.Print " HOAN THANH - Goi Test_CalendarDelete rieng neu can xoa"
    Debug.Print "============================================================"
End Sub
