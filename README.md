# 🔗 GoogleSvc.dll

> **Go DLL** kết nối Excel/VBA với Google Sheets · Gmail · Drive  
> Gọi trực tiếp từ VBA — không cần Python, không cần cài thêm gì.

---

## ✨ Tính năng

| | Dịch vụ | Thao tác |
|---|---|---|
| 📊 | **Google Sheets** | Đọc, ghi, thêm dòng |
| 📧 | **Gmail** | Gửi email (UTF-8, tiếng Việt có dấu) |
| 📁 | **Google Drive** | Upload, download, liệt kê file |

---

## 📦 Nội dung repo

```
releases/
├── amd64/
│   └── GoogleSvc.dll       ← Dùng cho Excel 64-bit (phổ biến)
└── x86/
    └── GoogleSvc.dll       ← Dùng cho Excel 32-bit
```

> **Kiểm tra Excel của bạn:** File → Account → About Excel  
> Thấy chữ **"64-bit"** → dùng thư mục `amd64/`

---

## 🚀 Bắt đầu nhanh

### Bước 1 — Chuẩn bị thư mục

```
📂 Thư mục làm việc\
├── GoogleSvc.dll
├── credentials.json    ← tạo theo hướng dẫn bên dưới
└── YourFile.xlsm
```

### Bước 2 — Tạo credentials.json

> 📖 Xem hướng dẫn chi tiết từng bước (có ảnh minh hoạ):  
> **[Hướng dẫn kết nối Google của bạn — GoogleSvc.dll](https://kieumanh366377.github.io/googlesvc-guide/)**

### Bước 3 — Import module VBA

Mở Excel → `Alt+F11` → **Insert → Module** → paste từng file:

| File | Vai trò |
|---|---|
| `modDllHelper.bas` | Tự động load/free DLL khi mở/đóng file |
| `modGBridgeTest.bas` | Kết nối DLL + hàm gọi Google API |
| `modSheetToExcel.bas` | Đọc Sheet → đổ thẳng xuống Range Excel |

### Bước 4 — Đăng nhập Google (1 lần duy nhất)

```
Ctrl+G  →  gõ:  Test_Login  →  Enter
```

Browser tự mở → đăng nhập → token được lưu tự động.  
Các lần sau **không cần đăng nhập lại**.

---

## 💻 Ví dụ VBA

<details>
<summary><b>📊 Đọc Google Sheet</b></summary>

```vb
' Đọc vùng A1:D10 từ Google Sheet
Dim resp As String
resp = GSheetRead_RPC("SPREADSHEET_ID", "Sheet1!A1:D10")
Debug.Print resp
' → {"ok":true,"data":{"values":[["A","B"],...]}}
```

```vb
' Đọc Sheet → đổ thẳng xuống Excel bắt đầu từ ô A1
Dim n As Long
n = GSheetToExcelRange("SPREADSHEET_ID", "Sheet1!A1:D20", "A1")
Debug.Print "Đã đọc " & n & " dòng"
```

</details>

<details>
<summary><b>✏️ Ghi lên Google Sheet</b></summary>

```vb
' Ghi đè vùng A1:B2
Dim resp As String
resp = GSheetWrite_RPC("SPREADSHEET_ID", "Sheet1!A1:B2", _
    "[[""Tên"",""Điểm""],[""Nguyễn An"",9.5]]")
Debug.Print resp
```

```vb
' Thêm dòng mới vào cuối (không ghi đè)
resp = GSheetAppend_RPC("SPREADSHEET_ID", "Sheet1!A:D", _
    "[[""2024-01-15"",""Đơn hàng mới"",""OK"",150000]]")
```

</details>

<details>
<summary><b>📧 Gửi Gmail</b></summary>

```vb
' Gửi email — hỗ trợ tiếng Việt có dấu trong Subject và Body
Dim resp As String
resp = GmailSend_RPC( _
    "nguoinhan@gmail.com", _
    "Báo cáo tháng 01/2025", _
    "Kính gửi anh/chị, dữ liệu đã được đồng bộ từ Excel.")
Debug.Print resp
```

</details>

<details>
<summary><b>📁 Google Drive</b></summary>

```vb
' Upload file lên Drive
resp = GDriveUpload_RPC("C:\\BaoCao\\thang01.xlsx")

' Liệt kê file trong folder
resp = GDriveList_RPC("FOLDER_ID", 20)

' Download file về máy
resp = GDriveDownload_RPC("FILE_ID", "C:\\TaiVe\\file.xlsx")
```

</details>

---

## 🔧 Lấy Spreadsheet ID

Từ URL Google Sheet:

```
https://docs.google.com/spreadsheets/d/ 1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms /edit
                                        ↑_________________________↑
                                              Spreadsheet ID
```

---

## ❓ Câu hỏi thường gặp

**Q: Chạy xong thấy lỗi "account chua dang nhap"?**  
→ Chạy `Test_Login` một lần để đăng nhập, sau đó thử lại.

**Q: Lỗi "File not found" khi load DLL?**  
→ Kiểm tra `GoogleSvc.dll` và `credentials.json` đặt **cùng thư mục** với file Excel.

**Q: Excel của tôi 32-bit hay 64-bit?**  
→ File → Account → About Excel → tìm chữ "32-bit" hoặc "64-bit".

**Q: Token lưu ở đâu, có an toàn không?**  
→ Token lưu trong thư mục `tokens\` cạnh DLL, được mã hoá bằng **Windows DPAPI** (gắn với máy tính, không dùng được trên máy khác).

---

## 📋 Yêu cầu hệ thống

- Windows 10/11
- Microsoft Excel **2016 trở lên**
- Kết nối internet (để gọi Google API)

---

## 📞 Hỗ trợ

Gặp vấn đề? Liên hệ: **kieumanh366377@gmail.com**  
Hoặc xem hướng dẫn đầy đủ: [googlesvc-guide](https://kieumanh366377.github.io/googlesvc-guide/)

---

<sub>Build với Go · CGO · Windows DPAPI · OAuth2 Google</sub>
