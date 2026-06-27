Attribute VB_Name = "modBridge"
' ==============================================================
' modBridge.bas
' Module dung chung cho toan bo Add-in GoogleSvc.
' Cac module khac (modSheetsTest, modGmailTest...) chi goi ham
' trong module nay, KHONG khai bao Declare rieng.
'
' QUY TAC TYPE:
'   - 64-bit: outBufSize truyen vao DLL phai la LongLong (khop ABI Go int64)
'   - Bien cuc bo nRet, nNeed dung LongLong tren Win64, Long tren Win32
'   - BUF_SIZE la Const Long - can CLngLng() khi truyen vao ham LongLong
'   - StrPtr() tra ve LongPtr - khop ca 32/64 bit
'
' CAU TRUC:
'   PHAN 1 - Declare
'   PHAN 2 - Const
'   PHAN 3 - Wrapper co ban (buf2Str, ExecuteJSON)
'   PHAN 4 - Wrapper Range <-> Sheet
'   PHAN 5 - Wrapper format JSON / tien ich
' ==============================================================

Option Explicit

' ==============================================================
' PHAN 1 - DECLARE
' ==============================================================

#If Win64 Then

Private Declare PtrSafe Function GBridgeExecute Lib "GoogleSvc.dll" _
    (ByVal reqJson As LongPtr, ByVal outBuf As LongPtr, ByVal outBufSize As LongLong) As LongLong

Private Declare PtrSafe Function GBridgeVersion Lib "GoogleSvc.dll" () As LongLong

Private Declare PtrSafe Function GSheetParseID Lib "GoogleSvc.dll" _
    (ByVal urlPtr As LongPtr, ByVal outBuf As LongPtr, ByVal outBufSize As LongLong) As LongLong

Private Declare PtrSafe Function GLogin Lib "GoogleSvc.dll" _
    (ByVal credPath As LongPtr, ByVal account As LongPtr, _
     ByVal outBuf As LongPtr, ByVal outBufSize As LongLong) As LongLong

Private Declare PtrSafe Function GSheetReadCell Lib "GoogleSvc.dll" _
    (ByVal credPath As LongPtr, ByVal account As LongPtr, _
     ByVal ssID As LongPtr, ByVal rangeA1 As LongPtr, _
     ByVal outBuf As LongPtr, ByVal outBufSize As LongLong) As LongLong

Private Declare PtrSafe Function GSheetWriteRange Lib "GoogleSvc.dll" _
    (ByVal credPath As LongPtr, ByVal account As LongPtr, _
     ByVal ssID As LongPtr, ByVal rangeA1 As LongPtr, ByVal valuesJSON As LongPtr, _
     ByVal outBuf As LongPtr, ByVal outBufSize As LongLong) As LongLong

Private Declare PtrSafe Function GDriveUploadFile Lib "GoogleSvc.dll" _
    (ByVal credPath As LongPtr, ByVal account As LongPtr, _
     ByVal localPath As LongPtr, ByVal folderID As LongPtr, _
     ByVal outBuf As LongPtr, ByVal outBufSize As LongLong) As LongLong

Private Declare PtrSafe Function GDriveDownloadFile Lib "GoogleSvc.dll" _
    (ByVal credPath As LongPtr, ByVal account As LongPtr, _
     ByVal fileID As LongPtr, ByVal localPath As LongPtr, _
     ByVal outBuf As LongPtr, ByVal outBufSize As LongLong) As LongLong

Private Declare PtrSafe Function GDriveListFiles Lib "GoogleSvc.dll" _
    (ByVal credPath As LongPtr, ByVal account As LongPtr, _
     ByVal folderID As LongPtr, ByVal maxResults As LongLong, _
     ByVal outBuf As LongPtr, ByVal outBufSize As LongLong) As LongLong

Private Declare PtrSafe Function GSheetRangeToJSON Lib "GoogleSvc.dll" _
    (ByVal inputPtr As LongPtr, _
     ByVal outBuf As LongPtr, ByVal outBufSize As LongLong) As LongLong

Private Declare PtrSafe Function GSheetJSONToRange Lib "GoogleSvc.dll" _
    (ByVal jsonPtr As LongPtr, _
     ByVal outBuf As LongPtr, ByVal outBufSize As LongLong) As LongLong

Private Declare PtrSafe Function GPrettyJSON Lib "GoogleSvc.dll" _
    (ByVal jsonPtr As LongPtr, _
     ByVal outBuf As LongPtr, ByVal outBufSize As LongLong) As LongLong

#Else

Private Declare Function GBridgeExecute Lib "GoogleSvc.dll" _
    (ByVal reqJson As Long, ByVal outBuf As Long, ByVal outBufSize As Long) As Long

Private Declare Function GBridgeVersion Lib "GoogleSvc.dll" () As Long

Private Declare Function GSheetParseID Lib "GoogleSvc.dll" _
    (ByVal urlPtr As Long, ByVal outBuf As Long, ByVal outBufSize As Long) As Long

Private Declare Function GLogin Lib "GoogleSvc.dll" _
    (ByVal credPath As Long, ByVal account As Long, _
     ByVal outBuf As Long, ByVal outBufSize As Long) As Long

Private Declare Function GSheetReadCell Lib "GoogleSvc.dll" _
    (ByVal credPath As Long, ByVal account As Long, _
     ByVal ssID As Long, ByVal rangeA1 As Long, _
     ByVal outBuf As Long, ByVal outBufSize As Long) As Long

Private Declare Function GSheetWriteRange Lib "GoogleSvc.dll" _
    (ByVal credPath As Long, ByVal account As Long, _
     ByVal ssID As Long, ByVal rangeA1 As Long, ByVal valuesJSON As Long, _
     ByVal outBuf As Long, ByVal outBufSize As Long) As Long

Private Declare Function GDriveUploadFile Lib "GoogleSvc.dll" _
    (ByVal credPath As Long, ByVal account As Long, _
     ByVal localPath As Long, ByVal folderID As Long, _
     ByVal outBuf As Long, ByVal outBufSize As Long) As Long

Private Declare Function GDriveDownloadFile Lib "GoogleSvc.dll" _
    (ByVal credPath As Long, ByVal account As Long, _
     ByVal fileID As Long, ByVal localPath As Long, _
     ByVal outBuf As Long, ByVal outBufSize As Long) As Long

Private Declare Function GDriveListFiles Lib "GoogleSvc.dll" _
    (ByVal credPath As Long, ByVal account As Long, _
     ByVal folderID As Long, ByVal maxResults As Long, _
     ByVal outBuf As Long, ByVal outBufSize As Long) As Long

Private Declare Function GSheetRangeToJSON Lib "GoogleSvc.dll" _
    (ByVal inputPtr As Long, _
     ByVal outBuf As Long, ByVal outBufSize As Long) As Long

Private Declare Function GSheetJSONToRange Lib "GoogleSvc.dll" _
    (ByVal jsonPtr As Long, _
     ByVal outBuf As Long, ByVal outBufSize As Long) As Long

Private Declare Function GPrettyJSON Lib "GoogleSvc.dll" _
    (ByVal jsonPtr As Long, _
     ByVal outBuf As Long, ByVal outBufSize As Long) As Long

#End If

' ==============================================================
' PHAN 2 - CONST
' ==============================================================

Public Const BUF_SIZE  As Long = 65536   ' 64KB - du cho phan lon response
Public Const BUF_LARGE As Long = 524288  ' 512KB - khi response co the rat lon

' ==============================================================
' PHAN 3 - WRAPPER CO BAN
' ==============================================================

' buf2Str: cat chuoi tai ky tu NUL dau tien.
' Dung sau moi lan goi DLL nhan out-buffer:
'   Dim buf As String: buf = Space(BUF_SIZE)
'   GBridgeExecute StrPtr(req), StrPtr(buf), CLngLng(BUF_SIZE)
'   Debug.Print buf2Str(buf)
Public Function buf2Str(ByVal s As String) As String
    Dim nNull As Long
    nNull = InStr(s, vbNullChar)
    If nNull > 1 Then
        buf2Str = Left$(s, nNull - 1)
    ElseIf nNull = 1 Then
        buf2Str = ""
    Else
        buf2Str = s
    End If
End Function

' ExecuteJSON: goi GBridgeExecute, tra ve JSON response da cat NUL.
' Tu dong retry voi buffer lon hon neu buffer qua nho (nRet < 0).
Public Function ExecuteJSON(ByVal sRequest As String) As String
    On Error GoTo ErrHandler

    Dim buf As String
    buf = Space(BUF_SIZE)

#If Win64 Then
    Dim nRet As LongLong
    nRet = GBridgeExecute(StrPtr(sRequest), StrPtr(buf), CLngLng(BUF_SIZE))
    If nRet < 0 Then
        Dim nNeed As LongLong
        nNeed = -nRet + 1
        If nNeed < BUF_LARGE Then nNeed = CLngLng(BUF_LARGE)
        buf = Space(CLng(nNeed))
        nRet = GBridgeExecute(StrPtr(sRequest), StrPtr(buf), nNeed)
    End If
#Else
    Dim nRet As Long
    nRet = GBridgeExecute(StrPtr(sRequest), StrPtr(buf), BUF_SIZE)
    If nRet < 0 Then
        Dim nNeed As Long
        nNeed = -nRet + 1
        If nNeed < BUF_LARGE Then nNeed = BUF_LARGE
        buf = Space(nNeed)
        nRet = GBridgeExecute(StrPtr(sRequest), StrPtr(buf), nNeed)
    End If
#End If

    ExecuteJSON = buf2Str(buf)
    Exit Function
ErrHandler:
    ExecuteJSON = "{""ok"":false,""error"":""VBA error: " & Err.Description & """}"
End Function

' ==============================================================
' PHAN 4 - WRAPPER RANGE <-> SHEET
' ==============================================================

' RangeToJSON: chuyen du lieu Range Excel sang JSON [[row]] chuan.
' Vi du:
'   Dim sJSON As String
'   sJSON = RangeToJSON(ws.Range("A1:C5"))
'   ' -> [["Ten","Tuoi","Diem"],["An","25","9.5"],...]
Public Function RangeToJSON(ByVal rng As Range) As String
    On Error GoTo ErrHandler

    ' B1: Tao Tab+NewLine tu Range
    Dim sData As String
    Dim i As Long, j As Long
    Dim nRows As Long: nRows = rng.Rows.Count
    Dim nCols As Long: nCols = rng.Columns.Count
    Dim aParts() As String

    For i = 1 To nRows
        If i > 1 Then sData = sData & Chr(10)
        ReDim aParts(nCols - 1)
        For j = 1 To nCols
            aParts(j - 1) = CStr(rng.Cells(i, j).Value)
        Next j
        sData = sData & Join(aParts, Chr(9))
    Next i

    ' B2: Goi Go chuyen sang JSON
    Dim buf As String: buf = Space(BUF_SIZE)

#If Win64 Then
    Dim nRet As LongLong
    nRet = GSheetRangeToJSON(StrPtr(sData), StrPtr(buf), CLngLng(BUF_SIZE))
    If nRet < 0 Then
        Dim nNeed As LongLong: nNeed = -nRet + 1
        buf = Space(CLng(nNeed))
        nRet = GSheetRangeToJSON(StrPtr(sData), StrPtr(buf), nNeed)
    End If
#Else
    Dim nRet As Long
    nRet = GSheetRangeToJSON(StrPtr(sData), StrPtr(buf), BUF_SIZE)
    If nRet < 0 Then
        Dim nNeed As Long: nNeed = -nRet + 1
        buf = Space(nNeed)
        nRet = GSheetRangeToJSON(StrPtr(sData), StrPtr(buf), nNeed)
    End If
#End If

    RangeToJSON = buf2Str(buf)
    Exit Function
ErrHandler:
    RangeToJSON = "[]"
End Function

' WriteRangeToSheet: ghi JSON [[row]] len Google Sheet.
' Tra ve JSON response {"ok":true/false,...}.
' Vi du:
'   Dim sJSON As String
'   sJSON = RangeToJSON(ws.Range("A1:F50"))
'   Debug.Print PrettyJSON(WriteRangeToSheet(sCred, "", sSheetID, "Sheet1!A1", sJSON))
Public Function WriteRangeToSheet(ByVal sCred As String, ByVal sAccount As String, _
    ByVal sSheetID As String, ByVal sRangeA1 As String, ByVal sJSON As String) As String
    On Error GoTo ErrHandler

    If sAccount = "" Then sAccount = "default"
    Dim buf As String: buf = Space(BUF_SIZE)

#If Win64 Then
    Dim nRet As LongLong
    nRet = GSheetWriteRange(StrPtr(sCred), StrPtr(sAccount), _
                            StrPtr(sSheetID), StrPtr(sRangeA1), StrPtr(sJSON), _
                            StrPtr(buf), CLngLng(BUF_SIZE))
    If nRet < 0 Then
        Dim nNeed As LongLong: nNeed = -nRet + 1
        buf = Space(CLng(nNeed))
        nRet = GSheetWriteRange(StrPtr(sCred), StrPtr(sAccount), _
                                StrPtr(sSheetID), StrPtr(sRangeA1), StrPtr(sJSON), _
                                StrPtr(buf), nNeed)
    End If
#Else
    Dim nRet As Long
    nRet = GSheetWriteRange(StrPtr(sCred), StrPtr(sAccount), _
                            StrPtr(sSheetID), StrPtr(sRangeA1), StrPtr(sJSON), _
                            StrPtr(buf), BUF_SIZE)
    If nRet < 0 Then
        Dim nNeed As Long: nNeed = -nRet + 1
        buf = Space(nNeed)
        nRet = GSheetWriteRange(StrPtr(sCred), StrPtr(sAccount), _
                                StrPtr(sSheetID), StrPtr(sRangeA1), StrPtr(sJSON), _
                                StrPtr(buf), nNeed)
    End If
#End If

    WriteRangeToSheet = buf2Str(buf)
    Exit Function
ErrHandler:
    WriteRangeToSheet = "{""ok"":false,""error"":""VBA error: " & Err.Description & """}"
End Function

' ReadSheetToRange: doc Google Sheet, gan thang vao Range Excel.
' Tra ve True neu thanh cong.
' Vi du:
'   If ReadSheetToRange(sCred, "", sSheetID, "DuLieu!A1:F100", ws.Range("A2")) Then
'     MsgBox "Da doc xong"
'   End If
Public Function ReadSheetToRange(ByVal sCred As String, ByVal sAccount As String, _
    ByVal sSheetID As String, ByVal sRangeA1 As String, ByVal rngDest As Range) As Boolean
    On Error GoTo ErrHandler

    If sAccount = "" Then sAccount = "default"

    ' B1: Doc tu Sheet
    Dim buf As String: buf = Space(BUF_LARGE)

#If Win64 Then
    Dim nRet As LongLong
    nRet = GSheetReadCell(StrPtr(sCred), StrPtr(sAccount), _
                          StrPtr(sSheetID), StrPtr(sRangeA1), _
                          StrPtr(buf), CLngLng(BUF_LARGE))
#Else
    Dim nRet As Long
    nRet = GSheetReadCell(StrPtr(sCred), StrPtr(sAccount), _
                          StrPtr(sSheetID), StrPtr(sRangeA1), _
                          StrPtr(buf), BUF_LARGE)
#End If

    If nRet < 0 Then
        Debug.Print "[ReadSheetToRange] Buffer qua nho, can " & (-nRet) & " ky tu"
        ReadSheetToRange = False
        Exit Function
    End If
    Dim sReadJSON As String: sReadJSON = buf2Str(buf)

    ' B2: Chuyen JSON sang Tab+NewLine
    Dim buf2 As String: buf2 = Space(BUF_LARGE)

#If Win64 Then
    Dim nRet2 As LongLong
    nRet2 = GSheetJSONToRange(StrPtr(sReadJSON), StrPtr(buf2), CLngLng(BUF_LARGE))
#Else
    Dim nRet2 As Long
    nRet2 = GSheetJSONToRange(StrPtr(sReadJSON), StrPtr(buf2), BUF_LARGE)
#End If

    If nRet2 < 0 Then
        Debug.Print "[ReadSheetToRange] Buffer qua nho cho convert, can " & (-nRet2) & " ky tu"
        ReadSheetToRange = False
        Exit Function
    End If
    Dim sTabNL As String: sTabNL = buf2Str(buf2)

    ' Neu Go tra ve JSON loi thi bao loi
    If Left$(sTabNL, 1) = "{" Then
        Debug.Print "[ReadSheetToRange] Loi: " & sTabNL
        ReadSheetToRange = False
        Exit Function
    End If

    ' B3: Split va gan vao Range
    Dim aRows() As String: aRows = Split(sTabNL, Chr(10))
    Dim i As Long, j As Long
    Dim aCols() As String
    For i = 0 To UBound(aRows)
        If aRows(i) <> "" Then
            aCols = Split(aRows(i), Chr(9))
            For j = 0 To UBound(aCols)
                rngDest.Cells(i + 1, j + 1).Value = aCols(j)
            Next j
        End If
    Next i

    ReadSheetToRange = True
    Exit Function
ErrHandler:
    Debug.Print "[ReadSheetToRange] VBA Error: " & Err.Description
    ReadSheetToRange = False
End Function

' ==============================================================
' PHAN 5 - WRAPPER FORMAT / TIEN ICH
' ==============================================================

' PrettyJSON: format JSON dep, indent 4 space (goi GPrettyJSON phia Go).
' Thay the modJsonPretty.bas - co the xoa module do sau khi chuyen sang day.
' Vi du:
'   Debug.Print PrettyJSON(ExecuteJSON(sReq))
Public Function PrettyJSON(ByVal sJSON As String) As String
    On Error GoTo ErrHandler

    Dim buf As String: buf = Space(BUF_LARGE)

#If Win64 Then
    Dim nRet As LongLong
    nRet = GPrettyJSON(StrPtr(sJSON), StrPtr(buf), CLngLng(BUF_LARGE))
#Else
    Dim nRet As Long
    nRet = GPrettyJSON(StrPtr(sJSON), StrPtr(buf), BUF_LARGE)
#End If

    If nRet < 0 Then
        PrettyJSON = sJSON
        Exit Function
    End If
    PrettyJSON = buf2Str(buf)
    Exit Function
ErrHandler:
    PrettyJSON = sJSON
End Function

' ParseID: tach Spreadsheet ID tu URL Google Sheet day du.
' Vi du:
'   sID = ParseID("https://docs.google.com/spreadsheets/d/1abc.../edit")
'   ' -> "1abc..."
Public Function ParseID(ByVal sURLOrID As String) As String
    On Error GoTo ErrHandler

    Dim buf As String: buf = Space(BUF_SIZE)

#If Win64 Then
    Dim nRet As LongLong
    nRet = GSheetParseID(StrPtr(sURLOrID), StrPtr(buf), CLngLng(BUF_SIZE))
#Else
    Dim nRet As Long
    nRet = GSheetParseID(StrPtr(sURLOrID), StrPtr(buf), BUF_SIZE)
#End If

    If nRet < 0 Then
        ParseID = sURLOrID
        Exit Function
    End If
    ParseID = buf2Str(buf)
    Exit Function
ErrHandler:
    ParseID = sURLOrID
End Function
