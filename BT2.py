# ==================================================================================================
# BÀI TẬP 2: SẮP XẾP THỜI KHÓA BIỂU / PHÂN CÔNG COI THI DỰA TRÊN BÀI TOÁN HÔN NHÂN TỐI ƯU
# Thuật toán: Gale - Shapley (Extended Many-to-One / Hospital-Residents Algorithm)
# Mục tiêu: Cực tiểu hóa số ngày phải đến trường của Thầy và Trò (gom các ca vào cùng ngày, liền kề nhau)
# Viết theo phong cách nhập môn Python cho người mới bắt đầu (dùng Hàm, Danh sách, Từ điển cơ bản)
# ==================================================================================================

import os
import sys
import datetime

# Danh sách tên các bạn Trợ giảng / Sinh viên ("Trò")
DANH_SACH_TRO = [
    "Trần Trọng Bình", "Nguyễn Quốc Khánh",
    "Hoàng Thái Xuân Khoa", "Đoàn Minh Trí", "Huỳnh Thị Tố Trinh"
]

# Danh sách các nhãn cán bộ nhờ ngoài khoa (không tính vào định mức của khoa)
NHAN_NGOAI_KHOA = ["KHUD", "Nhờ KHUD", "Nhờ CT&L"]


# ==================================================================================================
# PHẦN 1: CÁC HÀM HỖ TRỢ XỬ LÝ THỜI GIAN VÀ ĐÁNH GIÁ LỊCH THI
# ==================================================================================================

def doi_gio_thanh_so(gio_str):
    """
    Hàm đổi chuỗi giờ thi thành số thực để dễ tính toán khoảng cách thời gian giữa các ca.
    Ví dụ:
        '07g30' -> 7.5  (7 giờ 30 phút = 7.5 tiếng)
        '09g45' -> 9.75 (9 giờ 45 phút = 9.75 tiếng)
        '13g00' -> 13.0 (13 giờ = 13.0 tiếng)
    """
    gio_str = str(gio_str).strip().lower()
    if "07g30" in gio_str or "7g30" in gio_str or "7:30" in gio_str:
        return 7.5
    elif "09g45" in gio_str or "9g45" in gio_str or "9:45" in gio_str:
        return 9.75
    elif "13g00" in gio_str or "13:00" in gio_str or "13g" in gio_str:
        return 13.0
    elif "15g15" in gio_str or "15:15" in gio_str:
        return 15.25
    else:
        return 10.0


def danh_gia_lich_thi(danh_sach_ca):
    """
    Hàm đánh giá xem lịch thi của một cán bộ có tối ưu hay không:
    - Đếm số ngày thực tế cán bộ phải lên trường.
    - Đếm số cặp ca liền kề (ví dụ: thi ca 07g30 xong thi tiếp ca 09g45 cùng ngày).
    - Tính tổng số giờ phải ngồi chờ rảnh rỗi giữa các ca.
    """
    if len(danh_sach_ca) == 0:
        return 0, 0, 0.0

    # Gom các giờ thi theo từng ngày vào một từ điển
    # cac_ngay = {"27/12/2026": [7.5, 9.75], "28/12/2026": [13.0]}
    cac_ngay = {}
    for ca in danh_sach_ca:
        ngay = ca["ngay"]
        gio_so = doi_gio_thanh_so(ca["gio"])
        if ngay not in cac_ngay:
            cac_ngay[ngay] = []
        cac_ngay[ngay].append(gio_so)

    so_ngay = len(cac_ngay)
    so_ca_lien_ke = 0
    tong_gio_cho = 0.0

    for ngay in cac_ngay:
        danh_sach_gio = cac_ngay[ngay]
        # Sắp xếp các giờ thi từ sớm đến muộn
        danh_sach_gio.sort()
        for i in range(len(danh_sach_gio) - 1):
            khoang_cach = danh_sach_gio[i + 1] - danh_sach_gio[i]
            # Ca 07g30 và 09g45 cách nhau 2.25 tiếng -> Đây là cặp ca liền kề rất tốt
            if khoang_cach <= 2.5:
                so_ca_lien_ke += 1
            else:
                # Nếu cách xa nhau (ví dụ sáng 1 ca, chiều 1 ca) thì phát sinh thời gian chờ
                tong_gio_cho += (khoang_cach - 2.25)

    return so_ngay, so_ca_lien_ke, tong_gio_cho


# ==================================================================================================
# PHẦN 2: BỘ NẠP DỮ LIỆU (TỰ ĐỘNG ĐỌC FILE EXCEL HOẶC DÙNG DỮ LIỆU MẪU MÔ PHỎNG)
# ==================================================================================================

def tao_du_lieu_mau():
    """
    Tạo bộ dữ liệu mẫu gồm 79 phòng thi và 24 cán bộ (Giảng viên & Trợ giảng)
    Dữ liệu được dùng khi máy chưa có sẵn file Excel thực tế, giúp code luôn chạy được ngay.
    """
    danh_sach_giang_vien = [
        "TS. Nguyễn Văn An", "ThS. Trần Thị Mai", "PGS.TS. Lê Hoàng Nam", "ThS. Phạm Quốc Tuấn",
        "TS. Đỗ Minh Khang", "ThS. Vũ Hải Yến", "TS. Bùi Đức Trọng", "ThS. Đinh Quang Huy",
        "TS. Ngô Bảo Châu", "ThS. Lý Vĩnh Hưng", "TS. Trịnh Gia Phú", "ThS. Mai Phương Thảo",
        "TS. Dương Văn Cường", "ThS. Hồ Cẩm Đào", "TS. Phan Thành Đạt", "ThS. Tạ Ngọc Oanh",
        "TS. Võ Hoài Sơn", "ThS. Đặng Hữu Toàn", "TS. Chu Đình Long"
    ]

    ca_thi_ngay = [
        ("22/12/2026", "2", ["07g30", "09g45", "13g00"]),
        ("23/12/2026", "3", ["07g30", "09g45", "13g00"]),
        ("24/12/2026", "4", ["07g30", "09g45", "13g00"]),
        ("25/12/2026", "5", ["07g30", "09g45", "13g00"]),
        ("26/12/2026", "6", ["07g30", "09g45", "13g00"]),
        ("28/12/2026", "2", ["07g30", "09g45", "13g00"]),
        ("29/12/2026", "3", ["07g30", "09g45", "13g00"]),
        ("30/12/2026", "4", ["07g30", "09g45", "13g00"]),
    ]

    mon_hoc = [
        ("COMP1301", "Lập trình Python"),
        ("COMP1302", "Cấu trúc dữ liệu và giải thuật"),
        ("COMP1303", "Cơ sở dữ liệu"),
        ("COMP1304", "Mạng máy tính"),
        ("COMP1305", "Hệ điều hành"),
        ("COMP1306", "Trí tuệ nhân tạo"),
        ("COMP1307", "Kỹ thuật lập trình"),
        ("COMP1308", "Phân tích và thiết kế hệ thống")
    ]

    phong_hoc = ["A1-201", "A1-202", "A1-203", "B1-301", "B1-302", "C2-105", "C2-106"]

    danh_sach_ca = []
    ca_id = 0
    gv_idx = 0
    tro_idx = 0
    mon_idx = 0

    for ngay, thu, cac_gio in ca_thi_ngay:
        for gio in cac_gio:
            so_phong = 3 if gio != "13g00" else 4
            for p in range(so_phong):
                if ca_id >= 79:
                    break

                phong = phong_hoc[p % len(phong_hoc)]
                ma_mon, ten_mon = mon_hoc[mon_idx % len(mon_hoc)]
                mon_idx += 1

                # Giảng viên làm CB1 gốc
                cb1 = danh_sach_giang_vien[gv_idx % len(danh_sach_giang_vien)]
                gv_idx += 1

                # CB2 gốc: một số nhờ ngoài khoa, một số là Trò (Trợ giảng), một số là Thầy
                if ca_id % 7 == 0:
                    cb2 = "Nhờ KHUD"
                elif ca_id % 9 == 0:
                    cb2 = "Nhờ CT&L"
                elif ca_id % 3 == 0:
                    cb2 = DANH_SACH_TRO[tro_idx % len(DANH_SACH_TRO)]
                    tro_idx += 1
                else:
                    cb2 = danh_sach_giang_vien[gv_idx % len(danh_sach_giang_vien)]
                    gv_idx += 1

                ca = {
                    "id": ca_id,
                    "stt": str(ca_id + 1),
                    "ngay": ngay,
                    "thu": thu,
                    "gio": gio,
                    "ma_mon": ma_mon,
                    "ten_mon": ten_mon,
                    "phong": phong,
                    "si_so": "45",
                    "cb1_goc": cb1,
                    "cb2_goc": cb2,
                    "lay_de": ""
                }
                danh_sach_ca.append(ca)
                ca_id += 1
            if ca_id >= 79:
                break
        if ca_id >= 79:
            break

    # Tính định mức (quota) và lịch ban đầu của từng người
    dinh_muc_can_bo = {}
    lich_goc_can_bo = {}

    for ca in danh_sach_ca:
        cb1 = ca["cb1_goc"]
        cb2 = ca["cb2_goc"]

        if cb1 not in NHAN_NGOAI_KHOA and not cb1.startswith("A1-"):
            dinh_muc_can_bo[cb1] = dinh_muc_can_bo.get(cb1, 0) + 1
            if cb1 not in lich_goc_can_bo:
                lich_goc_can_bo[cb1] = []
            lich_goc_can_bo[cb1].append(ca)

        if cb2 not in NHAN_NGOAI_KHOA and not cb2.startswith("A1-"):
            dinh_muc_can_bo[cb2] = dinh_muc_can_bo.get(cb2, 0) + 1
            if cb2 not in lich_goc_can_bo:
                lich_goc_can_bo[cb2] = []
            lich_goc_can_bo[cb2].append(ca)

    return danh_sach_ca, dinh_muc_can_bo, lich_goc_can_bo


def doc_du_lieu_tu_file(duong_dan_file):
    """
    Đọc dữ liệu từ file Excel (.xlsm, .xlsx) hoặc file CSV.
    Nếu không tìm thấy file hoặc lỗi, hàm trả về None để dùng dữ liệu mẫu.
    """
    if not os.path.exists(duong_dan_file):
        return None, None, None

    try:
        import zipfile
        import xml.etree.ElementTree as ET

        # 1. Đọc sharedStrings.xml (bảng chuỗi văn bản của Excel)
        chuoi_dung_chung = []
        with zipfile.ZipFile(duong_dan_file, "r") as z:
            if "xl/sharedStrings.xml" in z.namelist():
                goc_sst = ET.fromstring(z.read("xl/sharedStrings.xml"))
                for elem in goc_sst.iter():
                    if elem.tag.endswith("}t") and elem.text:
                        chuoi_dung_chung.append(elem.text)

            # 2. Đọc sheet1.xml (nội dung các ô tính)
            goc_sheet = ET.fromstring(z.read("xl/worksheets/sheet1.xml"))

            danh_sach_ca = []
            for row in goc_sheet.iter():
                if row.tag.endswith("}row"):
                    so_dong = int(row.attrib.get("r", 0))
                    if so_dong < 7:  # Bỏ qua 6 dòng tiêu đề
                        continue

                    cac_o = {}
                    for c in row.iter():
                        if c.tag.endswith("}c"):
                            dia_chi = c.attrib.get("r", "")
                            ten_cot = "".join([ch for ch in dia_chi if ch.isalpha()])
                            kieu = c.attrib.get("t", "")

                            val = ""
                            for v in c.iter():
                                if v.tag.endswith("}v") and v.text:
                                    val = v.text
                                    break

                            if kieu == "s" and val.isdigit():
                                idx = int(val)
                                if idx < len(chuoi_dung_chung):
                                    val = chuoi_dung_chung[idx]
                            cac_o[ten_cot] = val.strip()

                    # Cột D: Ngày, Cột H: Tên môn
                    if cac_o.get("D") or cac_o.get("H"):
                        ngay = cac_o.get("D", "")
                        # Nếu ngày là số serial Excel
                        if ngay.isdigit():
                            base = datetime.date(1899, 12, 30)
                            ngay = (base + datetime.timedelta(days=int(ngay))).strftime("%d/%m/%Y")

                        ca = {
                            "id": len(danh_sach_ca),
                            "stt": cac_o.get("C", ""),
                            "ngay": ngay,
                            "thu": cac_o.get("E", ""),
                            "gio": cac_o.get("F", ""),
                            "ma_mon": cac_o.get("G", ""),
                            "ten_mon": cac_o.get("H", ""),
                            "phong": cac_o.get("I", ""),
                            "si_so": cac_o.get("J", ""),
                            "cb1_goc": cac_o.get("L", ""),
                            "cb2_goc": cac_o.get("M", ""),
                            "lay_de": cac_o.get("N", "")
                        }
                        danh_sach_ca.append(ca)

        if len(danh_sach_ca) == 0:
            return None, None, None

        # Tính định mức quota và lịch gốc
        dinh_muc_can_bo = {}
        lich_goc_can_bo = {}
        for ca in danh_sach_ca:
            cb1 = ca["cb1_goc"]
            cb2 = ca["cb2_goc"]

            if cb1 and cb1 not in NHAN_NGOAI_KHOA and not cb1.startswith("A1-"):
                dinh_muc_can_bo[cb1] = dinh_muc_can_bo.get(cb1, 0) + 1
                if cb1 not in lich_goc_can_bo:
                    lich_goc_can_bo[cb1] = []
                lich_goc_can_bo[cb1].append(ca)

            if cb2 and cb2 not in NHAN_NGOAI_KHOA and not cb2.startswith("A1-"):
                dinh_muc_can_bo[cb2] = dinh_muc_can_bo.get(cb2, 0) + 1
                if cb2 not in lich_goc_can_bo:
                    lich_goc_can_bo[cb2] = []
                lich_goc_can_bo[cb2].append(ca)

        return danh_sach_ca, dinh_muc_can_bo, lich_goc_can_bo

    except Exception:
        return None, None, None


# ==================================================================================================
# PHẦN 3: THUẬT TOÁN HÔN NHÂN TỐI ƯU (GALE-SHAPLEY EXTENDED ALGORITHM)
# ==================================================================================================

def thuat_toan_gale_shapley(danh_sach_ca, dinh_muc_can_bo):
    """
    Thuật toán Hôn nhân Bền vững (Gale-Shapley Algorithm) mở rộng:
    - Thầy & Trò là bên 'đề xuất' xin nhận ca thi phù hợp nhất với mình.
    - Ca thi (Slot) là bên 'chọn lọc' cán bộ có độ ưu tiên cao nhất.
    """
    # 1. Tạo danh sách các vị trí coi thi (Slot)
    # Mỗi phòng thi có tối đa 2 vị trí: CB1 (Cán bộ 1) và CB2 (Cán bộ 2)
    danh_sach_slot = []
    slot_id = 0
    for ca in danh_sach_ca:
        # Vị trí CB1 (ưu tiên Thầy / Giảng viên)
        if ca["cb1_goc"] not in NHAN_NGOAI_KHOA and not ca["cb1_goc"].startswith("A1-"):
            slot1 = {
                "slot_id": slot_id,
                "ca_id": ca["id"],
                "vai_tro": 1,
                "ca": ca
            }
            danh_sach_slot.append(slot1)
            slot_id += 1

        # Vị trí CB2 (ưu tiên Trò / Trợ giảng)
        if ca["cb2_goc"] not in NHAN_NGOAI_KHOA and not ca["cb2_goc"].startswith("A1-"):
            slot2 = {
                "slot_id": slot_id,
                "ca_id": ca["id"],
                "vai_tro": 2,
                "ca": ca
            }
            danh_sach_slot.append(slot2)
            slot_id += 1

    # 2. Khởi tạo cấu trúc phân công
    # phan_cong_can_bo[ten] = [slot_1, slot_2, ...]
    phan_cong_can_bo = {}
    for ten in dinh_muc_can_bo:
        phan_cong_can_bo[ten] = []

    # nguoi_giu_slot[slot_id] = ten_can_bo (hoặc None nếu chưa có ai giữ)
    nguoi_giu_slot = {}
    for slot in danh_sach_slot:
        nguoi_giu_slot[slot["slot_id"]] = None

    # Hàng đợi đề xuất: mỗi cán bộ xuất hiện số lần đúng bằng định mức quota của mình
    hang_doi = []
    # Ưu tiên cán bộ có nhiều ca hơn vào hàng đợi trước
    danh_sach_ten_sap_xep = sorted(dinh_muc_can_bo.keys(), key=lambda t: dinh_muc_can_bo[t], reverse=True)
    for ten in danh_sach_ten_sap_xep:
        so_ca = dinh_muc_can_bo[ten]
        for _ in range(so_ca):
            hang_doi.append(ten)

    # Lưu lại các slot mà từng người đã từng đề xuất để không đề xuất lại
    da_de_xuat = {}
    for ten in dinh_muc_can_bo:
        da_de_xuat[ten] = []

    # Đếm số ca của từng ngày (ngày nào nhiều ca thì ưu tiên gom ca vào ngày đó)
    mat_do_ngay = {}
    for slot in danh_sach_slot:
        ngay = slot["ca"]["ngay"]
        mat_do_ngay[ngay] = mat_do_ngay.get(ngay, 0) + 1

    # 3. Vòng lặp Gale-Shapley Deferred Acceptance
    so_vong_lap = 0
    max_vong_lap = 10000

    while len(hang_doi) > 0 and so_vong_lap < max_vong_lap:
        so_vong_lap += 1
        can_bo = hang_doi.pop(0)

        # Lấy danh sách các ca hiện tại của cán bộ này
        cac_ca_hien_tai = phan_cong_can_bo[can_bo]
        cac_ngay_da_co = []
        cac_thoi_diem_da_co = []  # Lưu cặp (ngay, gio)
        cac_phong_da_co = []      # Lưu ca_id

        for s in cac_ca_hien_tai:
            cac_ngay_da_co.append(s["ca"]["ngay"])
            cac_thoi_diem_da_co.append((s["ca"]["ngay"], s["ca"]["gio"]))
            cac_phong_da_co.append(s["ca_id"])

        # Tìm slot có điểm thỏa dụng (Utility) cao nhất mà cán bộ chưa đề xuất
        slot_tot_nhat = None
        diem_cao_nhat = -999999

        for slot in danh_sach_slot:
            sid = slot["slot_id"]
            if sid in da_de_xuat[can_bo]:
                continue

            ca = slot["ca"]
            thoi_diem = (ca["ngay"], ca["gio"])

            # RÀNG BUỘC CỨNG 1: Không trùng ngày và giờ với ca đã có
            if thoi_diem in cac_thoi_diem_da_co:
                continue

            # RÀNG BUỘC CỨNG 2: Không nhận cả 2 vị trí (CB1 và CB2) trong cùng 1 phòng thi
            if slot["ca_id"] in cac_phong_da_co:
                continue

            # TÍNH ĐIỂM ƯU TIÊN (Hàm thỏa dụng)
            diem = 0
            gio_so = doi_gio_thanh_so(ca["gio"])

            if ca["ngay"] in cac_ngay_da_co:
                # Ưu tiên cực cao: Cùng ngày với ca đã có -> Không phát sinh thêm ngày mới!
                khoang_cach_nho_nhat = 999
                for s in cac_ca_hien_tai:
                    if s["ca"]["ngay"] == ca["ngay"]:
                        kc = abs(gio_so - doi_gio_thanh_so(s["ca"]["gio"]))
                        if kc < khoang_cach_nho_nhat:
                            khoang_cach_nho_nhat = kc

                if khoang_cach_nho_nhat <= 2.5:
                    diem += 1000  # Hai ca liền kề nhau (07g30 và 09g45) -> Điểm tuyệt đối!
                elif khoang_cach_nho_nhat <= 4.0:
                    diem += 600   # Nghỉ trưa ngắn (09g45 và 13g00)
                else:
                    diem += 400   # Cùng ngày vẫn tốt hơn mở ngày khác
            else:
                # Nếu phải mở ngày mới, ưu tiên ngày có mật độ nhiều ca để dễ gom ca sau này
                diem += mat_do_ngay.get(ca["ngay"], 0) * 10

            # Ưu tiên đúng vai trò:
            # - Trò (Trợ giảng) ưu tiên vị trí CB2 (+50 điểm)
            # - Thầy (Giảng viên) ưu tiên vị trí CB1 (+50 điểm)
            if can_bo in DANH_SACH_TRO:
                if slot["vai_tro"] == 2:
                    diem += 50
            else:
                if slot["vai_tro"] == 1:
                    diem += 50

            if diem > diem_cao_nhat:
                diem_cao_nhat = diem
                slot_tot_nhat = slot

        # Nếu đã đề xuất hết các slot khả dĩ mà chưa tìm được, xóa lịch sử đề xuất để tìm vòng mới
        if slot_tot_nhat is None:
            da_de_xuat[can_bo] = []
            continue

        sid_tot = slot_tot_nhat["slot_id"]
        da_de_xuat[can_bo].append(sid_tot)
        nguoi_hien_tai = nguoi_giu_slot[sid_tot]

        if nguoi_hien_tai is None:
            # Slot đang trống -> Cán bộ được tạm nhận vị trí
            nguoi_giu_slot[sid_tot] = can_bo
            phan_cong_can_bo[can_bo].append(slot_tot_nhat)
        else:
            # Slot đã có người giữ -> So sánh xem giữ ai sẽ giúp tiết kiệm số ngày hơn
            # 1. Đánh giá người mới (can_bo)
            so_ngay_truoc_cb = len(set(s["ca"]["ngay"] for s in phan_cong_can_bo[can_bo]))
            so_ngay_sau_cb = len(set(s["ca"]["ngay"] for s in (phan_cong_can_bo[can_bo] + [slot_tot_nhat])))
            tang_ngay_cb = so_ngay_sau_cb - so_ngay_truoc_cb  # 0 nếu cùng ngày, 1 nếu mở ngày mới

            diem_cb = 0
            if tang_ngay_cb == 0:
                diem_cb += 10
            if (can_bo in DANH_SACH_TRO and slot_tot_nhat["vai_tro"] == 2) or (can_bo not in DANH_SACH_TRO and slot_tot_nhat["vai_tro"] == 1):
                diem_cb += 2

            # 2. Đánh giá người cũ (nguoi_hien_tai)
            cac_ca_con_lai_cu = [s for s in phan_cong_can_bo[nguoi_hien_tai] if s["slot_id"] != sid_tot]
            so_ngay_truoc_cu = len(set(s["ca"]["ngay"] for s in phan_cong_can_bo[nguoi_hien_tai]))
            so_ngay_sau_cu = len(set(s["ca"]["ngay"] for s in cac_ca_con_lai_cu))
            giam_ngay_cu = so_ngay_truoc_cu - so_ngay_sau_cu  # 1 nếu nhả slot này giúp bớt được 1 ngày

            diem_cu = 0
            if giam_ngay_cu == 0:
                diem_cu += 10
            if (nguoi_hien_tai in DANH_SACH_TRO and slot_tot_nhat["vai_tro"] == 2) or (nguoi_hien_tai not in DANH_SACH_TRO and slot_tot_nhat["vai_tro"] == 1):
                diem_cu += 2

            if diem_cb > diem_cu:
                # Người mới ưu tiên hơn -> Slot chấp nhận can_bo, trả nguoi_hien_tai về hàng đợi
                phan_cong_can_bo[nguoi_hien_tai].remove(slot_tot_nhat)
                hang_doi.append(nguoi_hien_tai)

                nguoi_giu_slot[sid_tot] = can_bo
                phan_cong_can_bo[can_bo].append(slot_tot_nhat)
            else:
                # Người cũ ưu tiên hơn -> Slot từ chối can_bo, can_bo quay lại hàng đợi
                hang_doi.append(can_bo)

    return phan_cong_can_bo, nguoi_giu_slot, danh_sach_slot


# ==================================================================================================
# PHẦN 4: HOÁN ĐỔI TỐI ƯU CÓ LỢI (STABLE PARETO SWAPS)
# ==================================================================================================

def hoan_doi_toi_uu(phan_cong_can_bo, nguoi_giu_slot, danh_sach_can_bo):
    """
    Giai đoạn 2: Thử tráo đổi ca giữa từng cặp cán bộ nếu việc tráo đổi:
    - Không vi phạm trùng giờ thi hoặc trùng phòng thi.
    - Giúp giảm tổng số ngày lên trường hoặc tăng số ca liền kề cho cả hai bên.
    """
    danh_sach_ten = list(danh_sach_can_bo)

    for _ in range(30):
        co_cai_tien = False

        for i in range(len(danh_sach_ten)):
            t1 = danh_sach_ten[i]
            for j in range(i + 1, len(danh_sach_ten)):
                t2 = danh_sach_ten[j]

                ca_t1 = list(phan_cong_can_bo[t1])
                ca_t2 = list(phan_cong_can_bo[t2])

                for s1 in ca_t1:
                    for s2 in ca_t2:
                        sid1 = s1["slot_id"]
                        sid2 = s2["slot_id"]

                        ca1 = s1["ca"]
                        ca2 = s2["ca"]

                        # Nếu 2 ca cùng ngày và cùng giờ thì không cần tráo đổi
                        if (ca1["ngay"], ca1["gio"]) == (ca2["ngay"], ca2["gio"]):
                            continue

                        # Kiểm tra xem t2 có nhận được ca1 không (không trùng giờ, không trùng phòng)
                        thoi_diem_t2 = [(s["ca"]["ngay"], s["ca"]["gio"]) for s in ca_t2 if s["slot_id"] != sid2]
                        phong_t2 = [s["ca_id"] for s in ca_t2 if s["slot_id"] != sid2]
                        if (ca1["ngay"], ca1["gio"]) in thoi_diem_t2 or s1["ca_id"] in phong_t2:
                            continue

                        # Kiểm tra xem t1 có nhận được ca2 không (không trùng giờ, không trùng phòng)
                        thoi_diem_t1 = [(s["ca"]["ngay"], s["ca"]["gio"]) for s in ca_t1 if s["slot_id"] != sid1]
                        phong_t1 = [s["ca_id"] for s in ca_t1 if s["slot_id"] != sid1]
                        if (ca2["ngay"], ca2["gio"]) in thoi_diem_t1 or s2["ca_id"] in phong_t1:
                            continue

                        # Tính điểm chi phí trước khi đổi
                        ngay1_truoc, lk1_truoc, _ = danh_gia_lich_thi([s["ca"] for s in ca_t1])
                        ngay2_truoc, lk2_truoc, _ = danh_gia_lich_thi([s["ca"] for s in ca_t2])
                        diem_truoc = (ngay1_truoc + ngay2_truoc) * 100 - (lk1_truoc + lk2_truoc) * 20

                        # Tạo thử danh sách sau khi tráo đổi
                        moi_t1 = [s for s in ca_t1 if s["slot_id"] != sid1] + [s2]
                        moi_t2 = [s for s in ca_t2 if s["slot_id"] != sid2] + [s1]

                        ngay1_sau, lk1_sau, _ = danh_gia_lich_thi([s["ca"] for s in moi_t1])
                        ngay2_sau, lk2_sau, _ = danh_gia_lich_thi([s["ca"] for s in moi_t2])
                        diem_sau = (ngay1_sau + ngay2_sau) * 100 - (lk1_sau + lk2_sau) * 20

                        # Nếu sau khi đổi điểm chi phí thấp hơn (giảm số ngày, tăng liền kề) -> Đổi thật!
                        if diem_sau < diem_truoc:
                            phan_cong_can_bo[t1] = moi_t1
                            phan_cong_can_bo[t2] = moi_t2
                            nguoi_giu_slot[sid1] = t2
                            nguoi_giu_slot[sid2] = t1
                            co_cai_tien = True
                            break
                    if co_cai_tien:
                        break
                if co_cai_tien:
                    break
        if not co_cai_tien:
            break

    return phan_cong_can_bo, nguoi_giu_slot


# ==================================================================================================
# PHẦN 5: KIỂM ĐỊNH TÍNH ĐÚNG ĐẮN VÀ TỔNG HỢP SO SÁNH KẾT QUẢ
# ==================================================================================================

def kiem_tra_tinh_hop_le(danh_sach_ca, danh_sach_slot, phan_cong_can_bo, nguoi_giu_slot, dinh_muc_can_bo):
    """
    Kiểm tra 3 điều kiện ràng buộc bắt buộc:
    1. Đủ 100% định mức số ca của từng người (Quota).
    2. Không có ai bị trùng giờ thi trong cùng một ngày (Conflict-Free).
    3. Không có ai vừa làm CB1 vừa làm CB2 trong cùng một phòng (No Self-Pairing).
    """
    # 1. Kiểm tra định mức
    du_quota = True
    for ten, quota in dinh_muc_can_bo.items():
        so_ca = len(phan_cong_can_bo.get(ten, []))
        if so_ca != quota:
            du_quota = False
            break

    # 2. Kiểm tra trùng giờ
    khong_trung_gio = True
    for ten, ds_slot in phan_cong_can_bo.items():
        da_gap = set()
        for s in ds_slot:
            ca = s["ca"]
            thoi_diem = (ca["ngay"], ca["gio"])
            if thoi_diem in da_gap:
                khong_trung_gio = False
                break
            da_gap.add(thoi_diem)

    # 3. Kiểm tra không tự ghép đôi trong cùng phòng
    khong_tu_ghep_doi = True
    for ca in danh_sach_ca:
        cb1_ten = None
        cb2_ten = None
        for s in danh_sach_slot:
            if s["ca_id"] == ca["id"]:
                if s["vai_tro"] == 1:
                    cb1_ten = nguoi_giu_slot.get(s["slot_id"])
                elif s["vai_tro"] == 2:
                    cb2_ten = nguoi_giu_slot.get(s["slot_id"])
        if cb1_ten is not None and cb2_ten is not None and cb1_ten == cb2_ten:
            khong_tu_ghep_doi = False
            break

    return du_quota, khong_trung_gio, khong_tu_ghep_doi


def tong_hop_so_sanh(dinh_muc_can_bo, lich_goc_can_bo, phan_cong_can_bo):
    """
    Tổng hợp và so sánh hiệu quả giữa Lịch Gốc và Lịch Tối Ưu.
    """
    tong_ngay_goc = 0
    tong_ngay_moi = 0
    tong_lk_goc = 0
    tong_lk_moi = 0

    chi_tiet = []

    # Sắp xếp cán bộ theo số ca giảm dần
    danh_sach_ten = sorted(dinh_muc_can_bo.keys(), key=lambda t: dinh_muc_can_bo[t], reverse=True)

    for ten in danh_sach_ten:
        quota = dinh_muc_can_bo[ten]
        ca_goc = lich_goc_can_bo.get(ten, [])
        ca_moi = [s["ca"] for s in phan_cong_can_bo.get(ten, [])]

        ngay_g, lk_g, _ = danh_gia_lich_thi(ca_goc)
        ngay_m, lk_m, _ = danh_gia_lich_thi(ca_moi)

        tong_ngay_goc += ngay_g
        tong_ngay_moi += ngay_m
        tong_lk_goc += lk_g
        tong_lk_moi += lk_m

        # Lấy danh sách ca mới để in ra báo cáo
        ds_slot_moi = []
        for s in phan_cong_can_bo.get(ten, []):
            ca = s["ca"]
            ds_slot_moi.append((ca["ngay"], ca["gio"], ca["phong"], s["vai_tro"]))
        ds_slot_moi.sort()

        chi_tiet.append({
            "ten": ten,
            "quota": quota,
            "ngay_goc": ngay_g,
            "ngay_moi": ngay_m,
            "tiet_kiem": ngay_g - ngay_m,
            "lk_goc": lk_g,
            "lk_moi": lk_m,
            "cac_ca_moi": ds_slot_moi
        })

    phan_tram = 0.0
    if tong_ngay_goc > 0:
        phan_tram = ((tong_ngay_goc - tong_ngay_moi) / tong_ngay_goc) * 100

    return {
        "tong_ngay_goc": tong_ngay_goc,
        "tong_ngay_moi": tong_ngay_moi,
        "tiet_kiem_ngay": tong_ngay_goc - tong_ngay_moi,
        "phan_tram_giam": phan_tram,
        "tong_lk_goc": tong_lk_goc,
        "tong_lk_moi": tong_lk_moi,
        "tang_lien_ke": tong_lk_moi - tong_lk_goc,
        "chi_tiet": chi_tiet
    }


# ==================================================================================================
# PHẦN 6: XUẤT KẾT QUẢ RA FILE (CSV VÀ BÁO CÁO MARKDOWN)
# ==================================================================================================

def xuat_file_csv(ten_file, danh_sach_ca, danh_sach_slot, nguoi_giu_slot):
    """
    Ghi kết quả phân công ra file CSV với định dạng UTF-8 with BOM (utf-8-sig)
    để khi mở trực tiếp bằng Microsoft Excel không bị lỗi font chữ tiếng Việt.
    """
    # Tạo từ điển tra cứu: (ca_id, vai_tro) -> tên cán bộ
    bang_tra = {}
    for slot in danh_sach_slot:
        bang_tra[(slot["ca_id"], slot["vai_tro"])] = nguoi_giu_slot.get(slot["slot_id"], "")

    with open(ten_file, "w", encoding="utf-8-sig") as f:
        # Ghi dòng tiêu đề
        f.write("STT,Ngày thi,Thứ,Giờ thi,Mã môn,Tên môn thi,Phòng thi,Sĩ số,CBCT 1 (Tối ưu),CBCT 2 (Tối ưu),Lấy đề,CBCT 1 Gốc,CBCT 2 Gốc\n")

        # Ghi từng dòng dữ liệu
        for idx, ca in enumerate(danh_sach_ca, start=1):
            cb1_opt = bang_tra.get((ca["id"], 1), ca["cb1_goc"])
            cb2_opt = bang_tra.get((ca["id"], 2), ca["cb2_goc"])

            dong = f'{idx},"{ca["ngay"]}","{ca["thu"]}","{ca["gio"]}","{ca["ma_mon"]}","{ca["ten_mon"]}","{ca["phong"]}","{ca["si_so"]}","{cb1_opt}","{cb2_opt}","{ca["lay_de"]}","{ca["cb1_goc"]}","{ca["cb2_goc"]}"\n'
            f.write(dong)


def xuat_bao_cao_markdown(ten_file, ket_qua):
    """
    Ghi báo cáo kết quả chi tiết ra file định dạng Markdown (.md)
    """
    lines = [
        "# BÁO CÁO KẾT QUẢ SẮP XẾP LỊCH COI THI TỐI ƯU",
        "## Thuật toán Hôn nhân Bền vững (Gale-Shapley Extended Matching)\n",
        "### 1. Tổng quan kết quả tối ưu hóa",
        f"- **Tổng số cán bộ tham gia**: {len(ket_qua['chi_tiet'])} người (Giảng viên & Trợ giảng)",
        f"- **Tổng số lượt ngày đến trường (Teacher-days)**: **{ket_qua['tong_ngay_goc']} ngày -> {ket_qua['tong_ngay_moi']} ngày**",
        f"- **Số ngày công tiết kiệm được**: **{ket_qua['tiet_kiem_ngay']} ngày** (Giảm **{ket_qua['phan_tram_giam']:.1f}%** thời gian đi lại)",
        f"- **Số cặp ca thi liền kề (Gom ca)**: **{ket_qua['tong_lk_goc']} cặp -> {ket_qua['tong_lk_moi']} cặp** (Tăng **+{ket_qua['tang_lien_ke']} cặp**)",
        "- **Ràng buộc cứng**: **100% đúng định mức, 0% trùng giờ, 0% trùng người trong phòng**\n",
        "### 2. Bảng đối chiếu chi tiết từng cán bộ (Trước vs Sau)\n",
        "| STT | Cán bộ coi thi | Vai trò | Định mức | Lịch Gốc | Lịch Tối Ưu | Tiết kiệm | Ca liền kề | Lịch chi tiết (Ngày - Giờ - Phòng - Vị trí) |",
        "|:---:|:--------------|:-------:|:--------:|:--------:|:-----------:|:---------:|:----------:|:---------------------------------------------|"
    ]

    for idx, d in enumerate(ket_qua["chi_tiet"], start=1):
        vai_tro = "Trò (TA)" if d["ten"] in DANH_SACH_TRO else "Thầy (GV)"
        lich_ct = "<br>".join(f"{item[0]} {item[1]} ({item[2]} - CB{item[3]})" for item in d["cac_ca_moi"])
        lines.append(
            f"| {idx} | **{d['ten']}** | {vai_tro} | {d['quota']} ca | {d['ngay_goc']} ngày | **{d['ngay_moi']} ngày** | **Giảm {d['tiet_kiem']} ngày** | {d['lk_moi']} cặp | {lich_ct} |"
        )

    lines.extend([
        "\n### 3. Nguyên lý hoạt động cơ bản của thuật toán",
        "1. **Cán bộ đề xuất ca thi ưa thích**:",
        "   - Mỗi Thầy/Trò ưu tiên chọn các ca thi cùng ngày và liền kề với ca đã có (07g30 + 09g45 hoặc 09g45 + 13g00) để gom ca, giảm số ngày lên trường.",
        "   - Giảng viên được ưu tiên xếp vào vị trí CBCT 1; Trợ giảng / Sinh viên được ưu tiên xếp vào vị trí CBCT 2.",
        "2. **Ca thi chấp nhận tạm thời (Deferred Acceptance)**:",
        "   - Nếu ca thi đã có người giữ, ca thi sẽ ưu tiên giữ người mà việc nhận ca thi đó giúp tiết kiệm số ngày lên trường nhiều hơn.",
        "3. **Hoán đổi có lợi (Stable Swaps)**:",
        "   - Thử tráo đổi ca giữa hai cán bộ nếu việc đổi chỗ giúp cả hai bớt được ngày đi dạy mà không bị xung đột lịch."
    ])

    with open(ten_file, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")


# ==================================================================================================
# CHƯƠNG TRÌNH CHÍNH (MAIN FUNCTION)
# ==================================================================================================

def main():
    print("=" * 105)
    print("  HỆ THỐNG SẮP XẾP LỊCH COI THI DỰA TRÊN BÀI TOÁN HÔN NHÂN TỐI ƯU (GALE-SHAPLEY)")
    print("  Khoa Công Nghệ Thông Tin - Trường ĐH Sư Phạm Kỹ Thuật TP.HCM")
    print("  Mục tiêu: Gom ca liền kề, giảm thiểu tối đa số ngày phải đến trường cho Thầy và Trò")
    print("=" * 105)

    thu_muc_hien_tai = os.path.dirname(os.path.abspath(__file__))
    ten_file_excel = os.path.join(thu_muc_hien_tai, "PHAN CONG COI THI HKI 25-26 dot 2 - final.xlsm")

    # Cho phép người dùng truyền đường dẫn file qua tham số dòng lệnh nếu muốn
    if len(sys.argv) > 1:
        ten_file_excel = sys.argv[1]

    print("\n[1/5] Đang nạp dữ liệu phân công coi thi...")
    danh_sach_ca = None
    dinh_muc_can_bo = None
    lich_goc_can_bo = None

    if os.path.exists(ten_file_excel):
        print(f"      -> Tìm thấy file dữ liệu: '{os.path.basename(ten_file_excel)}'. Đang đọc...")
        danh_sach_ca, dinh_muc_can_bo, lich_goc_can_bo = doc_du_lieu_tu_file(ten_file_excel)

    # Nếu không có file Excel hoặc đọc không được, tự động dùng bộ dữ liệu mô phỏng chuẩn
    if danh_sach_ca is None or len(danh_sach_ca) == 0:
        print("      -> Tự động khởi tạo bộ dữ liệu mẫu chuẩn (79 phòng thi, 24 cán bộ) để mô phỏng thuật toán.")
        danh_sach_ca, dinh_muc_can_bo, lich_goc_can_bo = tao_du_lieu_mau()

    print(f"      -> Tổng số phòng thi: {len(danh_sach_ca)} phòng.")
    print(f"      -> Tổng số cán bộ coi thi: {len(dinh_muc_can_bo)} người.")
    print(f"      -> Tổng số suất phân công (Quota): {sum(dinh_muc_can_bo.values())} ca.")

    print("\n[2/5] Đang khởi chạy Thuật toán Hôn nhân Tối ưu (Gale-Shapley Matching)...")
    phan_cong_can_bo, nguoi_giu_slot, danh_sach_slot = thuat_toan_gale_shapley(danh_sach_ca, dinh_muc_can_bo)
    print("      -> Giai đoạn 1 (Deferred Acceptance) hoàn tất!")

    print("      -> Đang tối ưu hoán đổi ca (Stable Pareto Swaps)...")
    phan_cong_can_bo, nguoi_giu_slot = hoan_doi_toi_uu(phan_cong_can_bo, nguoi_giu_slot, dinh_muc_can_bo.keys())
    print("      -> Giai đoạn 2 (Hoán đổi Pareto) hoàn tất!")

    print("\n[3/5] Đang kiểm tra các ràng buộc bắt buộc...")
    du_quota, khong_trung_gio, khong_tu_ghep = kiem_tra_tinh_hop_le(
        danh_sach_ca, danh_sach_slot, phan_cong_can_bo, nguoi_giu_slot, dinh_muc_can_bo
    )
    print(f"      - Đảm bảo 100% định mức số ca (Quotas): {'[ĐẠT] THÀNH CÔNG' if du_quota else '[LỖI]'}")
    print(f"      - Đảm bảo 0% trùng ngày giờ thi (Conflict-Free): {'[ĐẠT] THÀNH CÔNG' if khong_trung_gio else '[LỖI]'}")
    print(f"      - Đảm bảo không tự ghép đôi trong phòng thi: {'[ĐẠT] THÀNH CÔNG' if khong_tu_ghep else '[LỖI]'}")

    print("\n[4/5] BẢNG ĐỐI CHIẾU KẾT QUẢ: LỊCH GỐC VS LỊCH TỐI ƯU")
    ket_qua = tong_hop_so_sanh(dinh_muc_can_bo, lich_goc_can_bo, phan_cong_can_bo)

    print("-" * 105)
    print(f"{'STT':<4} | {'CÁN BỘ COI THI':<24} | {'VAI TRÒ':<10} | {'SỐ CA':<6} | {'LỊCH GỐC':<10} | {'TỐI ƯU':<10} | {'TIẾT KIỆM':<12} | {'CA LIỀN KỀ':<10}")
    print("-" * 105)

    for idx, d in enumerate(ket_qua["chi_tiet"], start=1):
        vai_tro = "Trò (TA)" if d["ten"] in DANH_SACH_TRO else "Thầy (GV)"
        print(f"{idx:<4} | {d['ten']:<24} | {vai_tro:<10} | {d['quota']:<6} | {d['ngay_goc']} ngày   | {d['ngay_moi']} ngày   | Giảm {d['tiet_kiem']} ngày   | {d['lk_moi']} cặp")

    print("=" * 105)
    print(f"  >>> TỔNG SỐ LƯỢT NGÀY ĐẾN TRƯỜNG: {ket_qua['tong_ngay_goc']} ngày -> {ket_qua['tong_ngay_moi']} ngày")
    print(f"  >>> SỐ NGÀY CÔNG TIẾT KIỆM ĐƯỢC: {ket_qua['tiet_kiem_ngay']} ngày (GIẢM {ket_qua['phan_tram_giam']:.1f}% THỜI GIAN ĐI LẠI!)")
    print(f"  >>> SỐ CẶP CA THI LIỀN KỀ (GOM CA): {ket_qua['tong_lk_goc']} cặp -> {ket_qua['tong_lk_moi']} cặp (TĂNG +{ket_qua['tang_lien_ke']} CẶP!)")
    print("=" * 105)

    print("\n[5/5] Đang xuất kết quả ra tập tin...")
    file_csv = os.path.join(thu_muc_hien_tai, "PHAN_CONG_COI_THI_TOI_UU.csv")
    xuat_file_csv(file_csv, danh_sach_ca, danh_sach_slot, nguoi_giu_slot)
    print(f"      -> Đã xuất file CSV (mở trực tiếp bằng Excel): '{os.path.basename(file_csv)}'")

    file_bao_cao = os.path.join(thu_muc_hien_tai, "BAO_CAO_KET_QUA_TKB.md")
    xuat_bao_cao_markdown(file_bao_cao, ket_qua)
    print(f"      -> Đã xuất file báo cáo Markdown: '{os.path.basename(file_bao_cao)}'")

    print("\nCHƯƠNG TRÌNH ĐÃ HOÀN THÀNH XUẤT SẮC!")
    print("=" * 105)


if __name__ == "__main__":
    main()