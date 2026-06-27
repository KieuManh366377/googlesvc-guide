# 🔗 GoogleSvc.dll — v1.1.0

> **Go DLL** kết nối Excel/VBA với Google APIs.  
> Gọi trực tiếp từ VBA — không cần Python, không cần cài thêm gì.

---

## ✨ Tính năng — 38 actions

| | Dịch vụ | Actions |
|---|---|---|
| 📊 | **Google Sheets** | read, write, append, clear, getInfo, create, addSheet, rename, delete, batchUpdate |
| 📧 | **Gmail** | send, list, search, read, markRead, markUnread, delete, attachments |
| 📁 | **Google Drive** | upload, download, list, get, search, createFolder, copy, move, trash, restore, delete |
| 📝 | **Google Docs** | create, read, insertText, replaceText |
| 📅 | **Google Calendar** | list, events, createEvent, updateEvent, deleteEvent |
| 🔐 | **Auth** | login, logout, status, accounts.list, accounts.remove |
| ⚙️ | **System** | version, info, health, scopes |

---

## 📦 Nội dung repo

```
googlesvc-guide/
├── GoogleSvc.dll          ← DLL chính (64-bit Windows)
├── GoogleSvc_Demo.xlsb    ← File Excel demo đầy đủ
├── credentials.json       ← Tạo theo hướng dẫn bên dưới (KHÔNG đưa lên git)
├── modBridge.bas          ← Module cầu nối VBA ↔ DLL (bắt buộc)
├── modAuthTest.bas        ← Test Auth actions
├── modSystemTest.bas      ← Test System actions
├── modSheetsTest.bas      ← Test Sheets actions
├── modDriveTest.bas       ← Test Drive actions
├── modGmailTest.bas       ← Test Gmail actions
├── modDocsCalendarTest.bas← Test Docs + Calendar actions
└── modErrorTest.bas       ← Test xử lý lỗi
```

> **Kiểm tra Excel của bạn:** File → Account → About Excel  
> Thấy chữ **"64-bit"** → dùng được `GoogleSvc.dll` trong repo này.

---

## 🚀 Bắt đầu nhanh

### Bước 1 — Chuẩn bị thư mục

```
📂 Thư mục làm việc\
├── GoogleSvc.dll
├── credentials.json
└── YourFile.xlsm
```

Token OAuth sẽ tự động lưu vào:
```
📂 Thư mục làm việc\
└── tokens\
    └── default.json    ← mã hoá bằng Windows DPAPI
```

### Bước 2 — Tạo credentials.json

📖 Xem hướng dẫn chi tiết từng bước (có ảnh minh hoạ):  
**[Hướng dẫn kết nối Google — GoogleSvc.dll](https://kieumanh366377.github.io/googlesvc-guide/)**

Tóm tắt nhanh:
1. Vào [Google Cloud Console](https://console.cloud.google.com/)
2. Tạo project → bật APIs: Sheets, Gmail, Drive, Docs, Calendar
3. Tạo OAuth 2.0 Client ID (Desktop app)
4. Download → đặt tên `credentials.json` cạnh `GoogleSvc.dll`

### Bước 3 — Import module VBA

Mở Excel → `Alt+F11` → **Insert → Module** → import các file `.bas`:

| File | Vai trò |
|---|---|
| `modBridge.bas` | **Bắt buộc** — Declare DLL + hàm ExecuteJSON + helpers |
| `modAuthTest.bas` | Test đăng nhập / quản lý account |
| `modSheetsTest.bas` | Test 10 Sheets actions |
| `modDriveTest.bas` | Test 11 Drive actions |
| `modGmailTest.bas` | Test 8 Gmail actions |
| `modDocsCalendarTest.bas` | Test 4 Docs + 5 Calendar actions |

### Bước 4 — Đăng nhập Google (1 lần duy nhất)

```
Ctrl+G  →  gõ:  Test_Login  →  Enter
```

Browser tự mở → đăng nhập → token được lưu tự động.  
Các lần sau **không cần đăng nhập lại** (token tự refresh).

---

## 💻 Ví dụ VBA

Tất cả actions đều gọi qua `ExecuteJSON` từ `modBridge.bas`, nhận JSON request và trả JSON response.

### 📊 Google Sheets

```vba
' Đọc vùng dữ liệu
Dim sResp As String
sResp = ExecuteJSON("{""action"":""sheets.read"",""spreadsheetId"":""ID"",""range"":""Sheet1!A1:D10""}")
' → {"ok":true,"data":{"values":[["A","B"],...],"range":"Sheet1!A1:D10"}}

' Ghi dữ liệu
sResp = ExecuteJSON("{""action"":""sheets.write"",""spreadsheetId"":""ID"",""range"":""Sheet1!A1""," & _
                    """values"":[[""Ten"",""Diem""],[""An"",9.5]]}")

' Thêm dòng mới (không ghi đè)
sResp = ExecuteJSON("{""action"":""sheets.append"",""spreadsheetId"":""ID"",""range"":""Sheet1!A:D""," & _
                    """values"":[[""2026-06-27"",""Don hang"",""OK"",150000]]}")

' Tạo spreadsheet mới
sResp = ExecuteJSON("{""action"":""sheets.create"",""title"":""Bao Cao Thang 6""}")

' Format màu ô (batchUpdate)
sResp = ExecuteJSON("{""action"":""sheets.batchUpdate"",""spreadsheetId"":""ID"",""requests"":[...]}")
```

### 📧 Gmail

```vba
' Gửi email (hỗ trợ tiếng Việt có dấu)
sResp = ExecuteJSON("{""action"":""gmail.send"",""to"":""nguoinhan@gmail.com""," & _
                    """subject"":""Bao cao thang 6"",""body"":""Kinh gui anh/chi...""}")

' Tìm kiếm email
sResp = ExecuteJSON("{""action"":""gmail.search"",""query"":""from:boss@company.com is:unread"",""maxResults"":10}")

' Danh sách attachment của email
sResp = ExecuteJSON("{""action"":""gmail.attachments"",""messageId"":""MSG_ID"",""action"":""list""}")

' Download attachment
sResp = ExecuteJSON("{""action"":""gmail.attachments"",""messageId"":""MSG_ID""," & _
                    """attachmentId"":""ATT_ID"",""saveDir"":""C:\\Downloads\\""}")
```

### 📁 Google Drive

```vba
' Upload file lên Drive
sResp = ExecuteJSON("{""action"":""drive.upload"",""localPath"":""C:\\BaoCao\\thang6.xlsx"",""folderId"":""FOLDER_ID""}")
' → {"ok":true,"data":{"fileId":"...","fileName":"...","webViewLink":"..."}}

' Download file về máy
sResp = ExecuteJSON("{""action"":""drive.download"",""fileId"":""FILE_ID"",""localPath"":""C:\\TaiVe\\file.xlsx""}")

' Tạo folder
sResp = ExecuteJSON("{""action"":""drive.createFolder"",""name"":""BaoCao_2026"",""parentId"":""PARENT_ID""}")

' Tìm kiếm file theo tên
sResp = ExecuteJSON("{""action"":""drive.search"",""query"":""name contains 'bao cao'"",""maxResults"":20}")

' Di chuyển file (newParentId rỗng = về My Drive gốc)
sResp = ExecuteJSON("{""action"":""drive.move"",""fileId"":""FILE_ID"",""newParentId"":""DEST_FOLDER_ID""}")

' Chuyển vào Trash (có thể restore trong 30 ngày)
sResp = ExecuteJSON("{""action"":""drive.trash"",""fileId"":""FILE_ID""}")
```

### 📝 Google Docs

```vba
' Tạo tài liệu mới
sResp = ExecuteJSON("{""action"":""docs.create"",""title"":""Bao Cao Thang 6""}")
' → {"ok":true,"data":{"documentId":"...","title":"...","webViewLink":"..."}}

' Đọc nội dung (text thuần)
sResp = ExecuteJSON("{""action"":""docs.read"",""documentId"":""DOC_ID""}")
' → {"ok":true,"data":{"text":"...","wordCount":N}}

' Chèn text vào cuối tài liệu (index=-1)
sResp = ExecuteJSON("{""action"":""docs.insertText"",""documentId"":""DOC_ID""," & _
                    """text"":""Noi dung them vao\n"",""index"":-1}")

' Tìm và thay thế text
sResp = ExecuteJSON("{""action"":""docs.replaceText"",""documentId"":""DOC_ID""," & _
                    """searchText"":""cu"",""replaceText"":""moi"",""matchCase"":false}")
' → {"ok":true,"data":{"occurrencesChanged":N}}
```

### 📅 Google Calendar

```vba
' Liệt kê calendars
sResp = ExecuteJSON("{""action"":""calendar.list""}")

' Xem sự kiện 30 ngày tới
sResp = ExecuteJSON("{""action"":""calendar.events"",""calendarId"":""primary"",""maxResults"":20}")

' Tạo sự kiện mới
sResp = ExecuteJSON("{""action"":""calendar.createEvent"",""calendarId"":""primary""," & _
                    """summary"":""Hop noi bo"",""start"":""2026-07-01T09:00:00+07:00""," & _
                    """end"":""2026-07-01T10:00:00+07:00"",""timeZone"":""Asia/Ho_Chi_Minh""}")
' → {"ok":true,"data":{"eventId":"...","summary":"...","htmlLink":"..."}}

' Cập nhật sự kiện
sResp = ExecuteJSON("{""action"":""calendar.updateEvent"",""calendarId"":""primary""," & _
                    """eventId"":""EVENT_ID"",""summary"":""Hop noi bo [Updated]""}")

' Xoá sự kiện
sResp = ExecuteJSON("{""action"":""calendar.deleteEvent"",""calendarId"":""primary"",""eventId"":""EVENT_ID""}")
```

### ⚙️ System

```vba
' Kiểm tra phiên bản DLL
sResp = ExecuteJSON("{""action"":""system.version""}")
' → {"ok":true,"data":{"dllVersion":"1.1.0","buildDate":"2026-06-26","goVersion":"go1.24.x"}}

' Kiểm tra sức khoẻ hệ thống
sResp = ExecuteJSON("{""action"":""system.health""}")
' → {"ok":true,"data":{"healthy":true,"items":[{"name":"credentials.json","ok":true}...]}}
```

---

## 🔑 Lấy Spreadsheet ID

Từ URL Google Sheet:

```
https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms/edit
                                       ↑_______________________________________________↑
                                                      Spreadsheet ID
```

Hoặc dùng hàm VBA helper:
```vba
' parseSpreadsheetIDFromURL tự tách ID từ URL đầy đủ
Dim sID As String
sID = ExecuteJSON("{""action"":""sheets.getInfo"",""spreadsheetId"":""URL_HOAC_ID""}")
```

---

## 🔐 Bảo mật token

- Token lưu trong `tokens\default.json` cạnh DLL
- Mã hoá bằng **Windows DPAPI** (scope MACHINE) — không dùng được trên máy khác
- Token tự động refresh khi hết hạn — không cần đăng nhập lại
- Nhiều account: mỗi account lưu file riêng `tokens\{account}.json`

---

## 👥 Multi-Account

```vba
' Đăng nhập account thứ 2
sResp = ExecuteJSON("{""action"":""auth.login"",""account"":""work""}")

' Gọi API với account cụ thể
sResp = ExecuteJSON("{""action"":""sheets.read"",""account"":""work"",""spreadsheetId"":""ID"",""range"":""A1""}")

' Liệt kê tất cả accounts đang đăng nhập
sResp = ExecuteJSON("{""action"":""auth.accounts.list""}")
```

---

## ❓ Câu hỏi thường gặp

**Q: Lỗi "account chua dang nhap"?**  
→ Chạy `Test_Login` một lần để đăng nhập, sau đó thử lại.

**Q: Lỗi 404 khi thao tác Drive file?**  
→ Drive scope `drive.file` chỉ truy cập được file do app tạo ra trong **session hiện tại**. Đóng Excel hoặc reload DLL → fileId cũ hết hiệu lực. Upload lại để lấy fileId mới.

**Q: Lỗi 403 insufficient scope?**  
→ Xoá file token cũ (`tokens\default.json`) rồi đăng nhập lại để cấp đủ scope Docs/Calendar.

**Q: Excel của tôi 32-bit hay 64-bit?**  
→ File → Account → About Excel → tìm chữ "32-bit" hoặc "64-bit".

**Q: Token lưu ở đâu, có an toàn không?**  
→ Token lưu trong `tokens\` cạnh DLL, mã hoá bằng Windows DPAPI — gắn với máy tính, không dùng được khi copy sang máy khác.

**Q: `drive.search` không tìm thấy file `.xlsx` vừa upload?**  
→ Drive nhận diện file `.xlsx` là `application/zip` (do nội dung là ZIP). Tìm theo tên thay vì MIME type: `name contains 'ten_file'`.

---

## 📋 Yêu cầu hệ thống

- Windows 10/11 (64-bit)
- Microsoft Excel 2016 trở lên (64-bit)
- Kết nối internet

---

## 📞 Hỗ trợ

Gặp vấn đề? Liên hệ: **kieumanh366377@gmail.com**  
Hướng dẫn đầy đủ: [googlesvc-guide](https://kieumanh366377.github.io/googlesvc-guide/)

---

*Build với Go · CGO · Windows DPAPI · OAuth2 Google*  
*v1.1.0 — 2026-06-26 — 38 actions: Sheets · Gmail · Drive · Docs · Calendar · Auth · System*
