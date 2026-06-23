Attribute VB_Name = "modSheetToExcel"
Option Explicit
' ==========================================================
' modSheetToExcel.bas
' Doc du lieu tu Google Sheet, do xuong Range Excel.
' Ghi du lieu tu Excel len Google Sheet (append / ghi de).
'
' YEU CAU: modGBridgeTest.bas phai co trong cung project
'   (module nay goi GSheetRead_RPC, GSheetWrite_RPC,
'   GSheetAppend_RPC dinh nghia trong do).
'
' CAU TRUC JSON DLL TRA VE (chi su dung cau truc nay):
'   {"ok":true,"data":{"values":[["a","b"],[1,2]],"range":"..."}}
'   {"ok":false,"error":"mo ta loi"}
'
' LUU Y VE PARSER:
'   Parser trong file nay la parser TOI GIAN, chi doc duoc
'   cau truc "values" co dinh nhu tren. Khong phai parser
'   JSON tong quat. Neu can xu ly JSON phuc tap hon, dung
'   thu vien VBA-JSON cua Tim Hall (github.com/VBA-tools).
' ==========================================================
Private Const SHEET_ID As String = "1punDQ-pRaMsopULsUaNthuHri5pj4ZrhNLFkOD3qkoo"  ' <- SUA O DAY

' ----------------------------------------------------------
' PHAN 1 - DOC GOOGLE SHEET, DO XUONG RANGE EXCEL
' ----------------------------------------------------------

' GSheetToExcelRange: doc 1 vung Google Sheet roi ghi xuong Excel.
'
'   spreadsheetId  : ID hoac URL day du cua Google Sheet
'   sheetRangeA1   : vung doc tren Google Sheet, vd "Sheet1!A1:C10"
'   destCell       : o Excel bat dau ghi, vd "A1" hoac "B3"
'   destSheetName  : ten sheet Excel dich, bo trong = sheet dang mo
'   account        : "default" neu chi dung 1 tai khoan Google
'
' Tra ve: so dong da ghi duoc (0 neu loi hoac Sheet rong).
'
' Vi du goi:
'   n = GSheetToExcelRange("ID_SHEET", "Sheet1!A1:D20", "A1")
'   n = GSheetToExcelRange("ID_SHEET", "DuLieu!A1:C50", "A1", "KetQua")
Public Function GSheetToExcelRange(ByVal spreadsheetId As String, _
    ByVal sheetRangeA1 As String, _
    ByVal destCell As String, _
    Optional ByVal destSheetName As String = "", _
    Optional ByVal account As String = "default") As Long

    ' Goi RPC doc du lieu
    Dim resp As String
    resp = GSheetRead_RPC(spreadsheetId, sheetRangeA1, account)

    ' Kiem tra phan hoi hop le
    If InStr(resp, """ok"":true") = 0 Then
        Debug.Print "[GSheetToExcelRange] Loi: " & resp
        GSheetToExcelRange = 0
        Exit Function
    End If

    ' Parse mang values tu JSON
    Dim rows As Collection
    Set rows = ParseValuesArray(resp)

    If rows Is Nothing Or rows.Count = 0 Then
        Debug.Print "[GSheetToExcelRange] Sheet rong hoac khong co du lieu."
        GSheetToExcelRange = 0
        Exit Function
    End If

    ' Chon sheet Excel dich
    Dim ws As Worksheet
    If destSheetName = "" Then
        Set ws = ActiveSheet
    Else
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(destSheetName)
        On Error GoTo 0
        If ws Is Nothing Then
            Debug.Print "[GSheetToExcelRange] Khong tim thay sheet: " & destSheetName
            GSheetToExcelRange = 0
            Exit Function
        End If
    End If

    ' Ghi tung dong, tung cot
    Dim startCell As Range
    Set startCell = ws.Range(destCell)

    Dim r As Long, c As Long
    Dim oneRow As Collection
    For r = 1 To rows.Count
        Set oneRow = rows(r)
        For c = 1 To oneRow.Count
            startCell.Offset(r - 1, c - 1).Value = oneRow(c)
        Next c
    Next r

    GSheetToExcelRange = rows.Count
    Debug.Print "[GSheetToExcelRange] Da ghi " & rows.Count & " dong xuong " & _
                ws.Name & "!" & destCell
End Function

' ----------------------------------------------------------
' PHAN 2 - GHI DU LIEU TU EXCEL LEN GOOGLE SHEET
' ----------------------------------------------------------

' ExcelRangeToGSheet: doc 1 vung Range Excel, ghi len Google Sheet.
'   Ghi DE (overwrite) bat dau tu goc trai cua destRangeA1.
'
'   srcRange      : Range Excel can ghi len, vd Range("A1:C5")
'   spreadsheetId : ID hoac URL day du cua Google Sheet
'   destRangeA1   : vung dich tren Sheet, vd "Sheet1!A1"
'   account       : "default" neu chi dung 1 tai khoan Google
'
' Tra ve: JSON response tu DLL (kiem tra "ok":true de biet thanh cong).
Public Function ExcelRangeToGSheet(ByVal srcRange As Range, _
    ByVal spreadsheetId As String, _
    ByVal destRangeA1 As String, _
    Optional ByVal account As String = "default") As String

    ' Chuyen Range Excel thanh JSON mang 2 chieu
    Dim valuesJson As String
    valuesJson = RangeToValuesJson(srcRange)

    ExcelRangeToGSheet = GSheetWrite_RPC(spreadsheetId, destRangeA1, valuesJson, account)
End Function

' ExcelRangeAppendToGSheet: doc 1 vung Range Excel, THEM DONG moi
' vao cuoi Google Sheet (khong ghi de du lieu hien co).
'
' Dung de ghi log, them ban ghi moi ma khong lo xoa du lieu cu.
Public Function ExcelRangeAppendToGSheet(ByVal srcRange As Range, _
    ByVal spreadsheetId As String, _
    ByVal destRangeA1 As String, _
    Optional ByVal account As String = "default") As String

    Dim valuesJson As String
    valuesJson = RangeToValuesJson(srcRange)

    ExcelRangeAppendToGSheet = GSheetAppend_RPC(spreadsheetId, destRangeA1, valuesJson, account)
End Function

' RangeToValuesJson: chuyen Range Excel thanh chuoi JSON mang 2 chieu.
' Vd Range A1:B2 co "Xin chao", 123 -> [["Xin chao",123],[...]]
' O rong -> chuoi rong "" trong JSON.
Private Function RangeToValuesJson(ByVal rng As Range) As String
    Dim r      As Long, c As Long
    Dim cellVal As Variant
    Dim rowParts() As String
    Dim rowJsons() As String
    Dim cellStr As String

    ReDim rowJsons(1 To rng.rows.Count)

    For r = 1 To rng.rows.Count
        ReDim rowParts(1 To rng.Columns.Count)
        For c = 1 To rng.Columns.Count
            cellVal = rng.Cells(r, c).Value
            Select Case VarType(cellVal)
                Case vbEmpty, vbNull
                    cellStr = """"""                         ' chuoi rong
                Case vbInteger, vbLong, vbSingle, vbDouble
                    cellStr = CStr(cellVal)                  ' so - khong co dau ""
                Case vbBoolean
                    cellStr = IIf(cellVal, "true", "false")  ' boolean
                Case Else
                    ' Chuoi hoac ngay: escape dau " roi boc trong ""
                    cellStr = """" & Replace(CStr(cellVal), """", "\""") & """"
            End Select
            rowParts(c) = cellStr
        Next c
        rowJsons(r) = "[" & Join(rowParts, ",") & "]"
    Next r

    RangeToValuesJson = "[" & Join(rowJsons, ",") & "]"
End Function

' ----------------------------------------------------------
' PHAN 3 - PARSER JSON NOI BO (chi cho cau truc values co dinh)
' ----------------------------------------------------------

' ParseValuesArray: tach phan "values":[[...]] tu JSON response,
' tra ve Collection cua Collection (moi Collection con = 1 dong).
Private Function ParseValuesArray(ByVal json As String) As Collection
    Dim result As New Collection

    Dim keyPos As Long
    keyPos = InStr(1, json, """values"":[", vbTextCompare)
    If keyPos = 0 Then
        ' Khong co field "values" - Sheet rong, Google khong tra field nay
        Set ParseValuesArray = result
        Exit Function
    End If

    Dim p As Long
    p = keyPos + Len("""values"":[")

    Do While p <= Len(json)
        Dim ch As String
        ch = Mid$(json, p, 1)

        If ch = "]" Then
            Exit Do              ' ket thuc mang ngoai
        ElseIf ch = "[" Then
            Dim rowEnd As Long
            Dim oneRow As Collection
            Set oneRow = ParseOneRow(json, p, rowEnd)
            result.Add oneRow
            p = rowEnd + 1
        Else
            p = p + 1            ' bo qua dau phay, khoang trang
        End If
    Loop

    Set ParseValuesArray = result
End Function

' ParseOneRow: parse 1 dong "[val1,val2,...]" bat dau tai startPos.
' Ghi vi tri dau "]" ket thuc vao endPos (ByRef).
Private Function ParseOneRow(ByVal json As String, ByVal startPos As Long, _
    ByRef endPos As Long) As Collection

    Dim result As New Collection
    Dim p As Long
    p = startPos + 1    ' bo qua "[" dau dong

    Do While p <= Len(json)
        Dim ch As String
        ch = Mid$(json, p, 1)

        Select Case ch
            Case "]"
                endPos = p
                Set ParseOneRow = result
                Exit Function

            Case ",", " "
                p = p + 1

            Case """"
                ' Gia tri kieu chuoi - doc toi dau " dong (xu ly \" escape)
                Dim strEnd As Long
                strEnd = p + 1
                Do While strEnd <= Len(json)
                    If Mid$(json, strEnd, 1) = "\" Then
                        strEnd = strEnd + 2
                    ElseIf Mid$(json, strEnd, 1) = """" Then
                        Exit Do
                    Else
                        strEnd = strEnd + 1
                    End If
                Loop
                result.Add UnescapeJSON(Mid$(json, p + 1, strEnd - p - 1))
                p = strEnd + 1

            Case Else
                ' Gia tri kieu so, hoac true/false/null
                Dim numEnd As Long
                numEnd = p
                Do While numEnd <= Len(json)
                    If Mid$(json, numEnd, 1) = "," Or _
                       Mid$(json, numEnd, 1) = "]" Then Exit Do
                    numEnd = numEnd + 1
                Loop
                Dim raw As String
                raw = Trim$(Mid$(json, p, numEnd - p))
                If IsNumeric(raw) Then
                    result.Add CDbl(raw)
                Else
                    result.Add raw    ' true / false / null -> giu nguyen
                End If
                p = numEnd
        End Select
    Loop

    endPos = p
    Set ParseOneRow = result
End Function

' UnescapeJSON: giai ma escape sequence JSON pho bien.
Private Function UnescapeJSON(ByVal s As String) As String
    s = Replace(s, "\""", """")
    s = Replace(s, "\\", "\")
    s = Replace(s, "\n", vbLf)
    s = Replace(s, "\t", vbTab)
    s = Replace(s, "\/", "/")
    UnescapeJSON = s
End Function

' ----------------------------------------------------------
' PHAN 4 - VI DU SU DUNG
' Chay tung Sub tu Immediate Window (Ctrl+G) de thu.
' SUA SPREADSHEET_ID_CUA_BAN truoc khi chay.
' ----------------------------------------------------------

' Vi du 1: Doc Google Sheet, do xuong sheet dang mo bat dau tu A1.
Public Sub Vidu_DocVeExcel()
    'Const SHEET_ID As String = "SPREADSHEET_ID_CUA_BAN"  ' <- SUA O DAY
    Dim n As Long
    n = GSheetToExcelRange(SHEET_ID, "Sheet1!A1:C10", "A1")
    If n > 0 Then
        Debug.Print "Da doc " & n & " dong tu Google Sheet.", vbInformation
    Else
        MsgBox "Khong co du lieu hoac bi loi.", vbExclamation
    End If
End Sub

' Vi du 2: Doc Google Sheet, do xuong sheet Excel ten "KetQua",
' bat dau tu o B2.
Public Sub Vidu_DocVeSheet_Rieng()
    'Const SHEET_ID As String = "SPREADSHEET_ID_CUA_BAN"  ' <- SUA O DAY
    Dim n As Long
    n = GSheetToExcelRange(SHEET_ID, "Sheet1!A1:D50", "A2", "KetQua")
    Debug.Print "Da ghi " & n & " dong."
End Sub

' Vi du 3: Ghi THEM DONG moi tu Range Excel len Google Sheet.
' Pham vi Range("A1:B1") = dong hien tai muon them.
Public Sub Vidu_ThemDong()
    'Const SHEET_ID As String = "SPREADSHEET_ID_CUA_BAN"  ' <- SUA O DAY
    Dim resp As String
    resp = ExcelRangeAppendToGSheet( _
        ThisWorkbook.Sheets(1).Range("A1:B1"), _
        SHEET_ID, "Sheet1!A:B")
    Debug.Print resp
End Sub

' Vi du 4: Ghi DE (overwrite) vung du lieu tu Excel len Google Sheet.
Public Sub Vidu_GhiDe()
    'Const SHEET_ID As String = "SPREADSHEET_ID_CUA_BAN"  ' <- SUA O DAY
    Dim resp As String
    resp = ExcelRangeToGSheet( _
        ThisWorkbook.Sheets(1).Range("A1:C5"), _
        SHEET_ID, "Sheet1!A1")
    Debug.Print resp
End Sub
