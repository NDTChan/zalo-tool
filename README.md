# zalo-tool
zalo tool made by ndtchan
- Chức năng 1: Xuất ra danh sách số điện thoại từ 1 folder (duyệt theo chiều sâu).
- Chức năng 2: Tự động giải nén file nén rồi xuất tiếp ra danh sách số điện thoại.

- Cách dùng: Có tất cả 6 thư mục để đáp ứng điều kiện của phần mềm:
1. import: Thư mục cần duyệt để export file số điện thoại.
2. file-zip: Thư mục chứa file nén (zip hoặc rar) mà hệ thống tìm được từ thư mục import.
3. import-unzip: Các file được giải nén từ thư mục file-zip.
4. exported: Thư mục chứa file số điện thoại được export ra.
5. error: Thư mục chứa các file bị lỗi (không read được).
6. no-phone-column: Thư mục chứa các file có sheet empty hoặc không tìm được column chứa các dấu hiệu nhận biết số điện thoại.
