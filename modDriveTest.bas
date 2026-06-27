Attribute VB_Name = "modDriveTest"
' ==============================================================
' modDriveTest.bas
' Test day du 11 action Drive:
'   list, get, search, createFolder, upload, download,
'   copy, move, trash, restore, delete
'
' YEU CAU: modBridge.bas phai co trong cung workbook.
'
' LUONG TEST KHUYEN NGHI (chay theo thu tu):
'   Buoc 1: Test_DriveList         -> xem cac file hien co
'   Buoc 2: Test_DriveCreateFolder -> tao folder test, lay folderId
'   Buoc 3: Test_DriveUpload       -> upload file, lay fileId
'   Buoc 4: Test_DriveGet          -> kiem tra thong tin file vua upload
'   Buoc 5: Test_DriveSearch       -> tim kiem theo type/ten
'   Buoc 6: Test_DriveCopy         -> tao ban sao file
'   Buoc 7: Test_DriveDownload     -> download file ve may
'   Buoc 8: Test_DriveMove         -> di chuyen file sang folder khac
'   Buoc 9: Test_DriveTrash        -> chuyen vao Trash
'  Buoc 10: Test_DriveRestore      -> khoi phuc tu Trash
'  Buoc 11: Test_DriveDelete       -> xoa vinh vien (can bo comment)
'
' CACH NHANH: chay Test_AllDrive_Safe truoc de lay IDs,
' sau do tu dien vao cac Const roi chay tung ham rieng.
' ==============================================================

Option Explicit

' ==============================================================
' CONST - DIEN TRUOC KHI TEST
' Buoc 1: chay Test_DriveList -> lay folderID co san
' Buoc 2: chay Test_DriveCreateFolder -> lay DRIVE_TEST_FOLDER_ID
' Buoc 3: chay Test_DriveUpload -> lay DRIVE_TEST_FILE_ID
' ==============================================================

' File can upload (phai ton tai tren may truoc khi chay Test_DriveUpload)
Private Const LOCAL_UPLOAD_FILE   As String = "C:\Temp\test_upload.xlsx"

' Thu muc download ve (se tu tao neu chua co)
Private Const LOCAL_DOWNLOAD_DIR  As String = "C:\Temp\"

' folderId cua folder test vua tao boi Test_DriveCreateFolder
' -> copy tu ket qua Test_DriveCreateFolder (field "folderId")
Private Const DRIVE_TEST_FOLDER_ID As String = "PUT_FOLDER_ID_HERE"

' fileId cua file vua upload boi Test_DriveUpload
' -> copy tu ket qua Test_DriveUpload (field "fileId")
Private Const DRIVE_TEST_FILE_ID   As String = "PUT_FILE_ID_HERE"

' folderId dich cho Test_DriveMove (folder khac voi DRIVE_TEST_FOLDER_ID)
' -> co the la root My Drive (de rong) hoac folder cu the
Private Const DRIVE_MOVE_DEST_FOLDER_ID As String = ""

' ==============================================================
' HELPER NOI BO
' ==============================================================

Private Sub PrintResult(ByVal sLabel As String, ByVal sJSON As String)
    Debug.Print "------------------------------------------------------------"
    Debug.Print "[" & sLabel & "]"
    Debug.Print PrettyJSON(sJSON)
End Sub

Private Function GuardFileId(ByVal sAction As String) As Boolean
    If DRIVE_TEST_FILE_ID = "PUT_FILE_ID_HERE" Then
        Debug.Print "[" & sAction & "] Chua sua DRIVE_TEST_FILE_ID" & _
                    " - chay Test_DriveUpload truoc roi copy fileId vao Const"
        GuardFileId = False
    Else
        GuardFileId = True
    End If
End Function

Private Function GuardFolderId(ByVal sAction As String) As Boolean
    If DRIVE_TEST_FOLDER_ID = "PUT_FOLDER_ID_HERE" Then
        Debug.Print "[" & sAction & "] Chua sua DRIVE_TEST_FOLDER_ID" & _
                    " - chay Test_DriveCreateFolder truoc roi copy folderId vao Const"
        GuardFolderId = False
    Else
        GuardFolderId = True
    End If
End Function

' ==============================================================
' BUOC 1 - Test_DriveList
' Liet ke file/folder trong DRIVE_TEST_FOLDER_ID (rong = My Drive goc)
' Ket qua mong doi:
'   {"ok":true,"data":[{"id":"...","name":"...","mimeType":"...","size":...},...]}
' ==============================================================
Public Sub Test_DriveList()
    Dim sReq As String
    sReq = "{""action"":""drive.list""," & _
           """folderId"":""" & DRIVE_TEST_FOLDER_ID & """," & _
           """maxResults"":20}"
    PrintResult "drive.list", ExecuteJSON(sReq)
End Sub

' ==============================================================
' BUOC 2 - Test_DriveCreateFolder
' Tao folder ten "GoogleSvc_Test" de dung cho cac test sau.
' Ket qua mong doi:
'   {"ok":true,"data":{"folderId":"...","name":"GoogleSvc_Test","webViewLink":"..."}}
' SAU KHI CHAY: copy folderId -> gan vao DRIVE_TEST_FOLDER_ID
' ==============================================================
Public Sub Test_DriveCreateFolder()
    PrintResult "drive.createFolder", _
        ExecuteJSON("{""action"":""drive.createFolder"",""name"":""GoogleSvc_Test""}")
End Sub

' Test_DriveCreateFolder_WithParent: tao folder con ben trong folder khac
Public Sub Test_DriveCreateFolder_WithParent()
    If Not GuardFolderId("drive.createFolder/withParent") Then Exit Sub
    PrintResult "drive.createFolder (co parentId)", _
        ExecuteJSON("{""action"":""drive.createFolder""," & _
                    """name"":""GoogleSvc_Sub""," & _
                    """parentId"":""" & DRIVE_TEST_FOLDER_ID & """}")
End Sub

' ==============================================================
' BUOC 3 - Test_DriveUpload
' Upload file tu dia len DRIVE_TEST_FOLDER_ID.
' Ket qua mong doi:
'   {"ok":true,"data":{"fileId":"...","fileName":"test_upload.xlsx","webViewLink":"..."}}
' SAU KHI CHAY: copy fileId -> gan vao DRIVE_TEST_FILE_ID
'
' Luu y: LOCAL_UPLOAD_FILE phai ton tai truoc khi chay.
'   Tao nhanh: Workbooks.Add.SaveAs "C:\Temp\test_upload.xlsx"
' ==============================================================
Public Sub Test_DriveUpload()
    ' Kiem tra file co ton tai khong
    If Len(Dir(LOCAL_UPLOAD_FILE)) = 0 Then
        Debug.Print "[drive.upload] File khong ton tai: " & LOCAL_UPLOAD_FILE
        Debug.Print "  -> Tao nhanh bang: Workbooks.Add.SaveAs """ & LOCAL_UPLOAD_FILE & """"
        Exit Sub
    End If

    Dim sPath As String
    sPath = Replace(LOCAL_UPLOAD_FILE, "\", "\\")

    Dim sReq As String
    sReq = "{""action"":""drive.upload""," & _
           """localPath"":""" & sPath & """," & _
           """folderId"":""" & DRIVE_TEST_FOLDER_ID & """}"
    PrintResult "drive.upload", ExecuteJSON(sReq)
End Sub

' ==============================================================
' BUOC 4 - Test_DriveGet
' Lay thong tin chi tiet 1 file theo DRIVE_TEST_FILE_ID.
' Ket qua mong doi:
'   {"ok":true,"data":{"id":"...","name":"...","mimeType":"...","size":...,"webViewLink":"...","parents":[...]}}
' ==============================================================
Public Sub Test_DriveGet()
    If Not GuardFileId("drive.get") Then Exit Sub
    PrintResult "drive.get", _
        ExecuteJSON("{""action"":""drive.get"",""fileId"":""" & DRIVE_TEST_FILE_ID & """}")
End Sub

' ==============================================================
' BUOC 5 - Test_DriveSearch
' Tim kiem theo MIME type hoac ten file.
' ==============================================================

' Tim tat ca file Excel (.xlsx) trong Drive
Public Sub Test_DriveSearch_Excel()
    PrintResult "drive.search (xlsx)", _
        ExecuteJSON("{""action"":""drive.search""," & _
                    """query"":""mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'""," & _
                    """maxResults"":10}")
End Sub

' Tim file theo ten chua "test"
Public Sub Test_DriveSearch_ByName()
    PrintResult "drive.search (name contains test)", _
        ExecuteJSON("{""action"":""drive.search""," & _
                    """query"":""name contains 'test'""," & _
                    """maxResults"":10}")
End Sub

' Tim tat ca folder trong Drive
Public Sub Test_DriveSearch_Folders()
    PrintResult "drive.search (folders only)", _
        ExecuteJSON("{""action"":""drive.search""," & _
                    """query"":""mimeType='application/vnd.google-apps.folder'""," & _
                    """maxResults"":10}")
End Sub

' ==============================================================
' BUOC 6 - Test_DriveCopy
' Tao ban sao DRIVE_TEST_FILE_ID voi ten moi.
' Ket qua mong doi:
'   {"ok":true,"data":{"fileId":"...","name":"test_upload_copy.xlsx","webViewLink":"..."}}
' ==============================================================
Public Sub Test_DriveCopy()
    If Not GuardFileId("drive.copy") Then Exit Sub
    PrintResult "drive.copy", _
        ExecuteJSON("{""action"":""drive.copy""," & _
                    """fileId"":""" & DRIVE_TEST_FILE_ID & """," & _
                    """newName"":""test_upload_copy.xlsx""}")
End Sub

' ==============================================================
' BUOC 7 - Test_DriveDownload
' Download DRIVE_TEST_FILE_ID ve may.
' Ket qua mong doi:
'   {"ok":true,"data":{"status":"download thanh cong","localPath":"C:\\Temp\\downloaded_test.xlsx"}}
' ==============================================================
Public Sub Test_DriveDownload()
    If Not GuardFileId("drive.download") Then Exit Sub

    Dim sLocalPath As String
    sLocalPath = Replace(LOCAL_DOWNLOAD_DIR & "downloaded_test.xlsx", "\", "\\")

    PrintResult "drive.download", _
        ExecuteJSON("{""action"":""drive.download""," & _
                    """fileId"":""" & DRIVE_TEST_FILE_ID & """," & _
                    """localPath"":""" & sLocalPath & """}")
End Sub

' ==============================================================
' BUOC 8 - Test_DriveMove
' Di chuyen DRIVE_TEST_FILE_ID sang DRIVE_MOVE_DEST_FOLDER_ID.
' Ket qua mong doi:
'   {"ok":true,"data":{"fileId":"...","name":"...","newParentId":"..."}}
'
' Luu y:
'   - oldParentId: de rong = Drive tu xac dinh (an toan khi file chi co 1 parent)
'   - DRIVE_MOVE_DEST_FOLDER_ID: de rong "" = move ve My Drive goc
' ==============================================================
Public Sub Test_DriveMove()
    If Not GuardFileId("drive.move") Then Exit Sub

    If DRIVE_MOVE_DEST_FOLDER_ID = "" Then
        Debug.Print "[drive.move] DRIVE_MOVE_DEST_FOLDER_ID dang rong"
        Debug.Print "  -> Move ve My Drive goc (root). Sua Const de chi dinh folder cu the."
    End If

    PrintResult "drive.move", _
        ExecuteJSON("{""action"":""drive.move""," & _
                    """fileId"":""" & DRIVE_TEST_FILE_ID & """," & _
                    """newParentId"":""" & DRIVE_MOVE_DEST_FOLDER_ID & """," & _
                    """oldParentId"":""""}")
End Sub

' ==============================================================
' BUOC 9 - Test_DriveTrash
' Chuyen DRIVE_TEST_FILE_ID vao Trash (co the restore trong 30 ngay).
' Ket qua mong doi:
'   {"ok":true,"data":{"fileId":"...","name":"...","status":"da chuyen vao Trash..."}}
' ==============================================================
Public Sub Test_DriveTrash()
    If Not GuardFileId("drive.trash") Then Exit Sub
    PrintResult "drive.trash", _
        ExecuteJSON("{""action"":""drive.trash"",""fileId"":""" & DRIVE_TEST_FILE_ID & """}")
End Sub

' ==============================================================
' BUOC 10 - Test_DriveRestore
' Khoi phuc DRIVE_TEST_FILE_ID tu Trash (phai chay sau Test_DriveTrash).
' Ket qua mong doi:
'   {"ok":true,"data":{"fileId":"...","name":"...","status":"da khoi phuc khoi Trash"}}
' ==============================================================
Public Sub Test_DriveRestore()
    If Not GuardFileId("drive.restore") Then Exit Sub
    PrintResult "drive.restore", _
        ExecuteJSON("{""action"":""drive.restore"",""fileId"":""" & DRIVE_TEST_FILE_ID & """}")
End Sub

' ==============================================================
' BUOC 11 - Test_DriveDelete
' XOA VINH VIEN - KHONG THE HOAN TAC.
' Bo comment dong PrintResult de chay that su.
' Khuyen nghi: dung fileId cua file copy (tu Test_DriveCopy),
'              KHONG dung DRIVE_TEST_FILE_ID goc.
' ==============================================================
Public Sub Test_DriveDelete()
    Dim sFileToDelete As String
    sFileToDelete = "PUT_COPY_FILE_ID_TO_DELETE_HERE"

    If sFileToDelete = "PUT_COPY_FILE_ID_TO_DELETE_HERE" Then
        Debug.Print "[drive.delete] Chua sua sFileToDelete"
        Debug.Print "  -> Dien fileId cua file ban sao (tu Test_DriveCopy) vao bien sFileToDelete"
        Debug.Print "  -> CANH BAO: xoa vinh vien, KHONG the hoan tac"
        Exit Sub
    End If

    ' Bo comment dong duoi de chay that su:
    ' PrintResult "drive.delete (VINH VIEN)", _
    '     ExecuteJSON("{""action"":""drive.delete"",""fileId"":""" & sFileToDelete & """}")
    Debug.Print "[drive.delete] Bo comment dong PrintResult de chay that su"
End Sub

' ==============================================================
' TEST CHAY NHANH - An toan (khong upload/move/trash/delete)
' ==============================================================

' Test_AllDrive_Safe: chay cac test doc (khong thay doi du lieu)
Public Sub Test_AllDrive_Safe()
    Debug.Print "============================================================"
    Debug.Print " DRIVE TEST - AN TOAN (doc, tim kiem)"
    Debug.Print "============================================================"

    Test_DriveList
    Test_DriveSearch_Excel
    Test_DriveSearch_ByName
    Test_DriveSearch_Folders

    If DRIVE_TEST_FILE_ID <> "PUT_FILE_ID_HERE" Then
        Test_DriveGet
    Else
        Debug.Print "[skip] Test_DriveGet - chua co DRIVE_TEST_FILE_ID"
    End If

    Debug.Print "============================================================"
    Debug.Print " HOAN THANH Test_AllDrive_Safe"
    Debug.Print " Tiep theo:"
    Debug.Print "   1. Chay Test_DriveCreateFolder -> lay folderId"
    Debug.Print "   2. Dien vao DRIVE_TEST_FOLDER_ID"
    Debug.Print "   3. Chay Test_DriveUpload -> lay fileId"
    Debug.Print "   4. Dien vao DRIVE_TEST_FILE_ID"
    Debug.Print "   5. Chay tung Test_DriveGet/Copy/Download/Move/Trash/Restore"
    Debug.Print "============================================================"
End Sub

' Test_AllDrive_Full: chay toan bo luong (can da dien DRIVE_TEST_FILE_ID)
' Khong chay Test_DriveDelete (can xac nhan thu cong)
Public Sub Test_AllDrive_Full()
    If Not GuardFileId("Test_AllDrive_Full") Then
        Debug.Print "-> Chay Test_AllDrive_Safe truoc de lay fileId"
        Exit Sub
    End If

    Debug.Print "============================================================"
    Debug.Print " DRIVE TEST - FULL (can DRIVE_TEST_FILE_ID hop le)"
    Debug.Print "============================================================"

    Test_DriveGet
    Test_DriveCopy
    Test_DriveDownload

    Debug.Print ""
    Debug.Print "--- Cac test thay doi trang thai (trash/restore) ---"
    Test_DriveTrash
    Test_DriveRestore

    Debug.Print ""
    Debug.Print "--- Test_DriveMove (tuy chon) ---"
    Test_DriveMove

    Debug.Print "============================================================"
    Debug.Print " HOAN THANH Test_AllDrive_Full"
    Debug.Print " Test_DriveDelete: chay thu cong, can bo comment trong ham"
    Debug.Print "============================================================"
End Sub
