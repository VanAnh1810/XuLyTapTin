        CÁC KIẾN THỨC CƠ BẢN ĐÃ DÙNG ĐỂ CODE TRONG BT_Xu_li_tap_tin.py
--------------------------------------------------------------------------------

1. KỸ THUẬT XỬ LÝ CHUỖI CĂN BẢN (STRING MANIPULATION):
   - chuoi.strip():
     + Xóa sạch các khoảng trắng dư thừa và ký tự xuống dòng ẩn (\n, \r) ở đầu 
       và cuối chuỗi.
     + Rất quan trọng khi đọc file để dữ liệu không bị thụt dòng hay cách dòng trống.

   - chuoi.split(ky_tu_phan_tach):
     + Cắt một chuỗi lớn thành một Danh sách (List) các chuỗi con.
     + line.split('"') : Tách dòng log dựa trên các dấu ngoặc kép ". 
       Nhờ đó, những đoạn dài nằm trong ngoặc kép (như Request, User-Agent) 
       được tách nguyên vẹn thành từng phần riêng mà không bị cắt vụn.
     + chuoi.split() (không truyền tham số): Tự động tách chuỗi theo khoảng 
       trắng (dấu cách, tab), đồng thời tự gom nhiều dấu cách liên tiếp làm một.

   - chuoi.replace(ky_tu_cu, ky_tu_moi):
     + Thay thế ký tự cũ bằng ký tự mới.
     + Ví dụ: replace("]", "") dùng để gạt bỏ dấu đóng ngoặc vuông khi lấy thời gian.

   - Toán tử ba ngôi (Ternary Operator - if/else trên 1 dòng):
     + Cú pháp: gia_tri_lay = A if dieu_kien else B
     + Ví dụ: method = yeu_cau[0] if len(yeu_cau) > 0 else ""
     + Tác dụng: Giúp code ngắn gọn và chống lỗi văng chương trình (IndexError) 
       khi danh sách bị rỗng.

2. CẤU TRÚC DỮ LIỆU (DATA STRUCTURES):
   - Dictionary (Từ điển - cặp Khóa : Giá trị):
     + Dùng để gom các thông tin của 1 dòng log:
       dong = {"IP": "233.223.117.90", "Phương thức": "DELETE", ...}
     + Ưu điểm: Gọi dữ liệu bằng tên cột rõ ràng (dong["IP"]), không bị nhầm lẫn 
       thứ tự như dùng danh sách số.

   - List (Danh sách):
     + danh_sach_cot.append(dong): Thêm từng dòng đã tách vào danh sách tổng.
     + danh_sach_cot[0]: Lấy dòng đầu tiên để xem mẫu.

3. KỸ THUẬT XỬ LÝ TẬP TIN (FILE HANDLING):
   - Câu lệnh `with open(...) as f:`
     + Tự động quản lý mở và đóng file. Dù chương trình có gặp lỗi giữa chừng, 
       Python vẫn tự động đóng file (f.close()) giải phóng bộ nhớ.
   - Các chế độ và tham số khi mở file:
     + "r" : Mở để đọc (read).
     + "w" : Mở để ghi mới hoặc ghi đè toàn bộ (write).
     + encoding="utf-8" hoặc "utf-8-sig": Đảm bảo hiển thị đúng tiếng Việt 
       và khi mở file CSV bằng Excel không bị lỗi font chữ.
     + errors="ignore": Bỏ qua các ký tự lạ nếu file log có lỗi mã hóa.
   - Đọc từng dòng bằng vòng lặp `for line in f:`
     + ĐÂY LÀ KỸ THUẬT TỐI QUAN TRỌNG: Python sẽ nạp từng dòng một vào bộ nhớ, 
       xử lý xong dòng nào giải phóng dòng đó.
     + Tuyệt đối KHÔNG dùng f.read() hay f.readlines() với file log 241MB vì 
       sẽ tốn hàng trăm MB RAM và dễ làm máy tính bị đơ (treo máy).

4. KỸ THUẬT IN ĐỊNH DẠNG CĂN LỀ BẢNG (F-STRING FORMATTING):
   - Cú pháp: f"{gia_tri:<so_ky_tu>}"
     + `<` : Căn lề bên trái.
     + `18` : Dành đúng 18 khoảng trống ký tự cho cột đó.
     + Ví dụ: f"{'IP':<18} | {'Method':<8}" giúp các dòng in ra thẳng tắp 
       như một bảng kẻ ô trên màn hình terminal.
