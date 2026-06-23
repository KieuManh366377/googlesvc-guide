Attribute VB_Name = "modDllHelper"
Option Explicit
' ==========================================================
' modDllHelper.bas
' Tu dong load GoogleSvc.dll khi mo workbook (Auto_Open), va
' giai phong khi dong workbook (Auto_Close).
'
' YEU CAU:
'   GoogleSvc.dll phai nam CUNG THU MUC voi file Excel nay.
'   Vi du:
'     C:\GoogleSvc\
'     |-- GoogleSvc.dll
'     |-- credentials.json
'     `-- GoogleSvc.xlsm  <- file Excel nay
'
' VI SAO CAN LOAD THU CONG (khong dung Declare ... Lib "..."
' thuan tuy):
'   Declare Lib "GoogleSvc.dll" (ten ngan, khong duong dan)
'   khien Windows tim DLL theo PATH he thong, KHONG tu tim
'   trong thu muc chua .xlsm. Phai dung LoadLibraryW voi
'   DUONG DAN DAY DU truoc khi goi bat ky ham Declare nao.
' ==========================================================

' ----------------------------------------------------------
' PHAN 1 - KHAI BAO API WINDOWS (tuong thich 32/64-bit)
' ----------------------------------------------------------
#If VBA7 And Win64 Then
    Public Declare PtrSafe Function LoadLibraryW Lib "kernel32" _
        (ByVal lpLibFileName As LongPtr) As LongPtr
    Public Declare PtrSafe Function FreeLibrary Lib "kernel32" _
        (ByVal hLibModule As LongPtr) As Long
    Public Declare PtrSafe Function SetDllDirectoryW Lib "kernel32" _
        (ByVal lpPathName As LongPtr) As Long
#Else
    Public Declare Function LoadLibraryW Lib "kernel32" _
        (ByVal lpLibFileName As Long) As Long
    Public Declare Function FreeLibrary Lib "kernel32" _
        (ByVal hLibModule As Long) As Long
    Public Declare Function SetDllDirectoryW Lib "kernel32" _
        (ByVal lpPathName As Long) As Long
#End If

' ----------------------------------------------------------
' PHAN 2 - HANDLE LUU LAI DE GIAI PHONG KHI DONG
' ----------------------------------------------------------
#If VBA7 And Win64 Then
    Public hGoogleSvc As LongPtr
#Else
    Public hGoogleSvc As Long
#End If

' ----------------------------------------------------------
' PHAN 3 - TEN FILE DLL (sua o day neu doi ten sau nay)
' ----------------------------------------------------------
Private Const DLL_GOOGLESVC As String = "GoogleSvc.dll"

' ----------------------------------------------------------
' PHAN 4 - LOAD DLL TU THU MUC CHUA FILE EXCEL
' ----------------------------------------------------------

' LoadDllFromWorkbookFolder: load 1 DLL theo ten file, tu ghep
' voi ThisWorkbook.Path de co duong dan day du.
' SetDllDirectoryW duoc goi truoc de Windows biet tim dependency
' trong cung thu muc (quan trong voi DLL build tu Go).
' Tra ve handle (0 neu loi).
#If VBA7 And Win64 Then
Public Function LoadDllFromWorkbookFolder(ByVal DllName As String) As LongPtr
#Else
Public Function LoadDllFromWorkbookFolder(ByVal DllName As String) As Long
#End If
    Dim fullPath As String
    Dim dllDir As String
    fullPath = ThisWorkbook.Path & "\" & DllName
    dllDir = ThisWorkbook.Path

    ' Kiem tra file ton tai truoc khi load
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(fullPath) Then
        MsgBox "Khong tim thay: " & fullPath & Chr(13) & Chr(13) & _
               "Dam bao GoogleSvc.dll nam CUNG THU MUC voi file Excel nay.", _
               vbCritical, "modDllHelper"
        LoadDllFromWorkbookFolder = 0
        Exit Function
    End If

    ' Dat thu muc tim dependency, load DLL, roi reset ngay
    SetDllDirectoryW StrPtr(dllDir)
    LoadDllFromWorkbookFolder = LoadLibraryW(StrPtr(fullPath))
    SetDllDirectoryW 0

    If LoadDllFromWorkbookFolder = 0 Then
        MsgBox "Khong load duoc: " & fullPath & Chr(13) & _
               "Windows error: " & Err.LastDllError, _
               vbCritical, "modDllHelper"
    End If
End Function

' ----------------------------------------------------------
' PHAN 5 - LOAD / FREE
' ----------------------------------------------------------

' LoadGoogleSvc: goi 1 lan (tu Auto_Open). Tra ve True neu thanh cong.
Public Function LoadGoogleSvc() As Boolean
    hGoogleSvc = LoadDllFromWorkbookFolder(DLL_GOOGLESVC)
    LoadGoogleSvc = (hGoogleSvc <> 0)
    If LoadGoogleSvc Then
        Debug.Print "modDllHelper: da load GoogleSvc.dll (handle=0x" & Hex$(hGoogleSvc) & ")"
    End If
End Function

' FreeGoogleSvc: goi 1 lan (tu Auto_Close).
Public Sub FreeGoogleSvc()
    If hGoogleSvc <> 0 Then
        FreeLibrary hGoogleSvc
        hGoogleSvc = 0
        Debug.Print "modDllHelper: da giai phong GoogleSvc.dll"
    End If
End Sub

' ----------------------------------------------------------
' PHAN 6 - AUTO_OPEN / AUTO_CLOSE
' Excel tu goi 2 ham nay khi mo/dong workbook (.xlsm).
' ----------------------------------------------------------

Sub Auto_Open()
    If Not LoadGoogleSvc() Then
        MsgBox "Khong load duoc GoogleSvc.dll." & Chr(13) & _
               "Cac chuc nang Google Sheets/Gmail/Drive se khong hoat dong.", _
               vbExclamation, "modDllHelper"
    End If
End Sub

Sub Auto_Close()
    FreeGoogleSvc
End Sub
