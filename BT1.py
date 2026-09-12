# BÀI TẬP: XỬ LÝ TẬP TIN - TÁCH DỮ LIỆU TẬP TIN LOG (logfiles.log) THÀNH CÁC CỘT

# Hàm tách một dòng log thành các thông tin cột riêng biệt
def tach_mot_dong(line):
    # 1. Bỏ khoảng trắng thừa hoặc ký tự xuống dòng ở hai đầu
    line = line.strip()
    if not line:
        return None

    # 2. Dòng log có các phần nằm trong dấu ngoặc kép: "..."
    # Chúng ta dùng hàm split('"') để tách theo dấu ngoặc kép:
    cac_phan = line.split('"')

    # Phần 0: chứa IP và Thời gian nằm trong ngoặc vuông [...]
    # Ví dụ: 233.223.117.90 - - [27/Dec/2037:12:00:00 +0530]
    phan_dau = cac_phan[0].strip().split("[")
    ip = phan_dau[0].split()[0]                       # Lấy địa chỉ IP
    thoi_gian = phan_dau[1].replace("]", "").strip()  # Bỏ dấu ] để lấy thời gian

    # Phần 1: chứa Phương thức, Đường dẫn (URL) và Giao thức
    # Ví dụ: DELETE /usr/admin HTTP/1.0
    yeu_cau = cac_phan[1].split()
    method = yeu_cau[0] if len(yeu_cau) > 0 else ""
    endpoint = yeu_cau[1] if len(yeu_cau) > 1 else ""
    protocol = yeu_cau[2] if len(yeu_cau) > 2 else ""

    # Phần 2: chứa Mã trạng thái (Status) và Kích thước (Size)
    # Ví dụ: 502 4963
    phan_giua = cac_phan[2].split()
    status = phan_giua[0] if len(phan_giua) > 0 else ""
    size = phan_giua[1] if len(phan_giua) > 1 else ""

    # Phần 3: chứa Referer (Trang web nguồn giới thiệu)
    referer = cac_phan[3]

    # Phần 5: chứa User-Agent (Thông tin trình duyệt / thiết bị)
    user_agent = cac_phan[5] if len(cac_phan) > 5 else ""

    # Phần 6: chứa Thời gian phản hồi (Response Time tính bằng ms)
    response_time = cac_phan[6].strip() if len(cac_phan) > 6 else ""

    # Gom tất cả các cột vào một từ điển (dictionary) để dễ quản lý
    return {
        "IP": ip,
        "Thời gian": thoi_gian,
        "Phương thức": method,
        "Đường dẫn": endpoint,
        "Giao thức": protocol,
        "Mã trạng thái": status,
        "Kích thước": size,
        "Nguồn (Referer)": referer,
        "Trình duyệt (User-Agent)": user_agent,
        "Thời gian phản hồi": response_time
    }

# CHƯƠNG TRÌNH CHÍNH

ten_file_log = "logfiles.log"
ten_file_xuat = "logfiles_dang_cot.csv"

# Danh sách chứa các dòng dữ liệu sau khi đã tách cột
danh_sach_cot = []

print("--- ĐANG ĐỌC VÀ TÁCH CỘT TẬP TIN LOG ---")

# Mở tập tin log để đọc từng dòng bằng câu lệnh with quen thuộc
with open(ten_file_log, "r", encoding="utf-8", errors="ignore") as f:
    for line in f:
        cot = tach_mot_dong(line)
        if cot != None:
            danh_sach_cot.append(cot)

        # Lấy thử 10 dòng đầu tiên để xem mẫu (không đọc hết 1 triệu dòng để máy chạy nhanh)
        if len(danh_sach_cot) >= 10:
            break

# 1. In tiêu đề và các cột ra màn hình dạng bảng để quan sát
print("\n" + "=" * 110)
print(f"{'STT':<5} | {'IP':<18} | {'Thời gian':<28} | {'Method':<8} | {'Đường dẫn':<22} | {'Status':<8} | {'Size':<6}")
print("-" * 110)

stt = 1
for dong in danh_sach_cot:
    print(f"{stt:<5} | {dong['IP']:<18} | {dong['Thời gian']:<28} | {dong['Phương thức']:<8} | {dong['Đường dẫn']:<22} | {dong['Mã trạng thái']:<8} | {dong['Kích thước']:<6}")
    stt += 1
print("=" * 110)

# 2. In chi tiết từng cột của 1 dòng đầu tiên cho dễ hiểu
print("\n--- CHI TIẾT ĐẦY ĐỦ CÁC CỘT CỦA DÒNG ĐẦU TIÊN ---")
for ten_cot, gia_tri in danh_sach_cot[0].items():
    print(f"- {ten_cot}: {gia_tri}")

# 3. Ghi dữ liệu đã tách thành các cột ra file mới (ngăn cách bằng dấu phẩy)
with open(ten_file_xuat, "w", encoding="utf-8-sig") as f_out:
    # Ghi dòng tiêu đề cột đầu tiên
    f_out.write("IP,Thời gian,Phương thức,Đường dẫn,Giao thức,Mã trạng thái,Kích thước,Referer,User_Agent,Thời gian phản hồi\n")
    
    # Ghi từng dòng dữ liệu
    for d in danh_sach_cot:
        dong_chuoi = f"{d['IP']},{d['Thời gian']},{d['Phương thức']},{d['Đường dẫn']},{d['Giao thức']},{d['Mã trạng thái']},{d['Kích thước']},{d['Nguồn (Referer)']},\"{d['Trình duyệt (User-Agent)']}\",{d['Thời gian phản hồi']}\n"
        f_out.write(dong_chuoi)

print(f"\nĐã lưu kết quả tách cột vào tập tin: '{ten_file_xuat}'")
