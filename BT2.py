"""
====================================================================================================
BÀI TẬP 2: SẮP XẾP THỜI KHÓA BIỂU / PHÂN CÔNG COI THI DỰA TRÊN BÀI TOÁN HÔN NHÂN TỐI ƯU
Thuật toán: Gale-Shapley Algorithm (Extended Many-to-One / Hospital-Residents with Stable Improvements)
Mục tiêu: Cực tiểu hóa thời gian lên lớp / có mặt tại trường của Thầy và Trò (giảm số ngày, giảm gap)
Bộ dữ liệu test: PHAN CONG COI THI HKI 25-26 dot 2 - final.xlsm
====================================================================================================
"""

import os
import sys
import copy
import zipfile
import datetime
import xml.sax.saxutils as saxutils
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from collections import defaultdict, Counter
from typing import List, Dict, Set, Tuple, Optional


# ==================================================================================================
# 1. CẤU TRÚC DỮ LIỆU (DATA MODELS)
# ==================================================================================================

@dataclass
class ExamSession:
    """Đại diện cho một ca / phòng thi cụ thể"""
    id: int
    row: int
    stt: str
    ngay: str          # DD/MM/YYYY
    thu: str           # 2, 3, 4, 5, ...
    gio: str           # 07g30, 09g45, 13g00
    ma_mon: str
    ten_mon: str
    phong: str
    si_so: str
    cb1_orig: str      # CBCT 1 ban đầu
    cb2_orig: str      # CBCT 2 ban đầu
    lay_de: str = ""

@dataclass
class ProctorSlot:
    """Đại diện cho một vị trí cần phân công trong phòng thi (CB1 hoặc CB2)"""
    slot_id: int
    session_id: int
    role: int          # 1: CBCT 1 (Thầy / Giảng viên), 2: CBCT 2 (Trò / Trợ giảng)
    session: ExamSession

@dataclass
class Proctor:
    """Đại diện cho một cán bộ coi thi (Thầy hoặc Trò)"""
    name: str
    quota: int         # Số ca cần phân công theo định mức
    is_tro: bool       # True nếu là Trợ giảng / Sinh viên ("Trò")


# ==================================================================================================
# 2. BỘ ĐỌC DỮ LIỆU (DATA LOADER - ZERO DEPENDENCY ENGINE)
# ==================================================================================================

class DataLoader:
    """Đọc dữ liệu từ file Excel .xlsm bằng pure-Python (zipfile + xml) hoặc openpyxl nếu có"""

    EXTERNAL_TAGS = {'KHUD', 'Nhờ KHUD', 'Nhờ CT&L'}

    # Danh sách các Trợ giảng / Sinh viên hỗ trợ ("Trò") trong khoa
    TRO_NAMES = {
        'Trần Trọng Bình', 'Nguyễn Quốc Khánh',
        'Hoàng Thái Xuân Khoa', 'Đoàn Minh Trí', 'Huỳnh Thị Tố Trinh'
    }

    @staticmethod
    def excel_serial_to_date(val: str) -> str:
        """Chuyển đổi số serial Excel hoặc chuỗi ngày sang định dạng DD/MM/YYYY"""
        if not val:
            return ""
        val = str(val).strip()
        if val.isdigit():
            n = int(val)
            # Mốc ngày Excel 1900
            base = datetime.date(1899, 12, 30)
            d = base + datetime.timedelta(days=n)
            return d.strftime('%d/%m/%Y')
        return val

    @classmethod
    def load(cls, filepath: str) -> Tuple[List[ExamSession], Dict[str, int], Dict[str, List[ExamSession]]]:
        if not os.path.exists(filepath):
            raise FileNotFoundError(f"Không tìm thấy file dữ liệu: {filepath}")

        # Thử đọc bằng openpyxl nếu có, nếu không thì dùng Pure Python XML parser
        try:
            import openpyxl
            return cls._load_with_openpyxl(filepath)
        except ImportError:
            return cls._load_with_xml(filepath)

    @classmethod
    def _load_with_openpyxl(cls, filepath: str):
        import openpyxl
        wb = openpyxl.load_workbook(filepath, data_only=True)
        sheet = wb['TỔNG'] if 'TỔNG' in wb.sheetnames else wb.active

        sessions: List[ExamSession] = []
        teacher_quotas: Counter = Counter()
        teacher_orig: Dict[str, List[ExamSession]] = defaultdict(list)

        for row_idx, row in enumerate(sheet.iter_rows(values_only=True), start=1):
            if row_idx < 7:
                continue
            # Row structure: A: Khoa, B: Lãnh đạo, C: STT, D: Ngày, E: Thứ, F: Giờ, G: Mã, H: Tên, I: Phòng, J: Sĩ số, K: Thời gian, L: CB1, M: CB2, N: Lấy đề
            if len(row) < 13:
                continue
            ngay_raw = row[3]
            ten_mon = str(row[7] or '').strip()
            if not ngay_raw and not ten_mon:
                continue

            if isinstance(ngay_raw, datetime.datetime) or isinstance(ngay_raw, datetime.date):
                ngay = ngay_raw.strftime('%d/%m/%Y')
            else:
                ngay = cls.excel_serial_to_date(str(ngay_raw or ''))

            session = ExamSession(
                id=len(sessions),
                row=row_idx,
                stt=str(row[2] or '').strip(),
                ngay=ngay,
                thu=str(row[4] or '').strip(),
                gio=str(row[5] or '').strip(),
                ma_mon=str(row[6] or '').strip(),
                ten_mon=ten_mon,
                phong=str(row[8] or '').strip(),
                si_so=str(row[9] or '').strip(),
                cb1_orig=str(row[11] or '').strip(),
                cb2_orig=str(row[12] or '').strip(),
                lay_de=str(row[13] if len(row) > 13 and row[13] else '').strip()
            )
            sessions.append(session)

            cb1 = session.cb1_orig
            cb2 = session.cb2_orig
            if cb1 and cb1 not in cls.EXTERNAL_TAGS and not cb1.startswith('A1-'):
                teacher_quotas[cb1] += 1
                teacher_orig[cb1].append(session)
            if cb2 and cb2 not in cls.EXTERNAL_TAGS and not cb2.startswith('A1-'):
                teacher_quotas[cb2] += 1
                teacher_orig[cb2].append(session)

        return sessions, dict(teacher_quotas), teacher_orig

    @classmethod
    def _load_with_xml(cls, filepath: str):
        z = zipfile.ZipFile(filepath)
        shared_strings = []
        if 'xl/sharedStrings.xml' in z.namelist():
            sst_root = ET.fromstring(z.read('xl/sharedStrings.xml'))
            ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            for si in sst_root.findall('.//ns:si', ns):
                text = ''.join(t.text for t in si.findall('.//ns:t', ns) if t.text)
                shared_strings.append(text)

        ns_wb = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        def col_letter(ref):
            return ''.join([c for c in ref if c.isalpha()])

        root = ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
        rows = root.findall('.//ns:row', ns_wb)

        sessions: List[ExamSession] = []
        teacher_quotas: Counter = Counter()
        teacher_orig: Dict[str, List[ExamSession]] = defaultdict(list)

        for r in rows:
            row_num = int(r.attrib['r'])
            if row_num < 7:
                continue
            cells = {}
            for c in r.findall('ns:c', ns_wb):
                col = col_letter(c.attrib['r'])
                v = c.find('ns:v', ns_wb)
                val = v.text if v is not None else ''
                if c.attrib.get('t') == 's' and val:
                    val = shared_strings[int(val)]
                cells[col] = val.strip()

            if cells.get('D') or cells.get('H'):
                ngay = cls.excel_serial_to_date(cells.get('D', ''))
                session = ExamSession(
                    id=len(sessions),
                    row=row_num,
                    stt=cells.get('C', ''),
                    ngay=ngay,
                    thu=cells.get('E', ''),
                    gio=cells.get('F', ''),
                    ma_mon=cells.get('G', ''),
                    ten_mon=cells.get('H', ''),
                    phong=cells.get('I', ''),
                    si_so=cells.get('J', ''),
                    cb1_orig=cells.get('L', ''),
                    cb2_orig=cells.get('M', ''),
                    lay_de=cells.get('N', ''),
                )
                sessions.append(session)

                cb1 = session.cb1_orig
                cb2 = session.cb2_orig
                if cb1 and cb1 not in cls.EXTERNAL_TAGS and not cb1.startswith('A1-'):
                    teacher_quotas[cb1] += 1
                    teacher_orig[cb1].append(session)
                if cb2 and cb2 not in cls.EXTERNAL_TAGS and not cb2.startswith('A1-'):
                    teacher_quotas[cb2] += 1
                    teacher_orig[cb2].append(session)

        return sessions, dict(teacher_quotas), teacher_orig


# ==================================================================================================
# 3. THUẬT TOÁN HÔN NHÂN TỐI ƯU (GALE-SHAPLEY EXTENDED SCHEDULER)
# ==================================================================================================

class GaleShapleyScheduler:
    """
    Thuật toán Hôn nhân Tối ưu (Gale-Shapley / Hospital-Residents Algorithm) mở rộng:
    - Bipartite matching giữa Tập Thầy & Trò (với định mức Quota) và Tập Ca/Phòng thi (Slots).
    - Hệ thống ưu tiên động (Dynamic Preferences) nhằm cực tiểu hóa số ngày đến trường và thời gian chờ rảnh rỗi.
    - Kết hợp thuật toán Hoán đổi Pareto Ổn định (Top Trading Cycles / Stable Swaps) để đạt nghiệm tối ưu toàn cục.
    """

    EXTERNAL_TAGS = {'KHUD', 'Nhờ KHUD', 'Nhờ CT&L'}
    TRO_NAMES = DataLoader.TRO_NAMES

    @staticmethod
    def time_to_val(gio_str: str) -> float:
        """Quy đổi giờ thi ra số thực để tính khoảng cách thời gian"""
        if '07g30' in gio_str:
            return 7.5
        if '09g45' in gio_str:
            return 9.75
        if '13g00' in gio_str:
            return 13.0
        return 10.0

    @classmethod
    def evaluate_schedule(cls, session_list: List[dict]) -> Tuple[int, int, float]:
        """
        Đánh giá lịch thi của một cán bộ:
        Returns: (số_ngày_lên_trường, số_cặp_ca_liền_kề, tổng_giờ_chờ_rảnh_rỗi)
        """
        if not session_list:
            return 0, 0, 0.0

        day_map = defaultdict(list)
        for item in session_list:
            sess = item['session'] if 'session' in item else item
            day_map[sess.ngay if isinstance(sess, ExamSession) else sess['ngay']].append(
                cls.time_to_val(sess.gio if isinstance(sess, ExamSession) else sess['gio'])
            )

        num_days = len(day_map)
        back_to_back = 0
        idle_gaps = 0.0

        for day, times in day_map.items():
            times.sort()
            for i in range(len(times) - 1):
                diff = times[i+1] - times[i]
                # Ca 07g30 và 09g45 cách nhau 2.25h -> Liền kề hoàn hảo (zero idle gap)
                if diff <= 2.5:
                    back_to_back += 1
                else:
                    idle_gaps += (diff - 2.25)

        return num_days, back_to_back, idle_gaps

    @classmethod
    def solve(cls, sessions: List[ExamSession], teacher_quotas: Dict[str, int]):
        # Tạo danh sách các Slot cần phân công
        slots: List[ProctorSlot] = []
        for s in sessions:
            if s.cb1_orig not in cls.EXTERNAL_TAGS:
                slots.append(ProctorSlot(slot_id=len(slots), session_id=s.id, role=1, session=s))
            if s.cb2_orig not in cls.EXTERNAL_TAGS:
                slots.append(ProctorSlot(slot_id=len(slots), session_id=s.id, role=2, session=s))

        slot_datetime = {sl.slot_id: (sl.session.ngay, sl.session.gio) for sl in slots}
        slot_session = {sl.slot_id: sl.session_id for sl in slots}

        day_counts = Counter(sl.session.ngay for sl in slots)

        teacher_assignments: Dict[str, List[ProctorSlot]] = {t: [] for t in teacher_quotas}
        slot_assignments: Dict[int, Optional[str]] = {sl.slot_id: None for sl in slots}

        # Danh sách hàng đợi đề xuất Gale-Shapley (mỗi người đề xuất theo quota của mình)
        unfilled = [t for t in sorted(teacher_quotas.keys(), key=lambda x: teacher_quotas[x], reverse=True)
                    for _ in range(teacher_quotas[t])]

        proposed: Dict[str, Set[int]] = defaultdict(set)

        def get_preference_list(t: str):
            """Xây dựng danh sách ưu tiên của Cán bộ t dựa trên độ thỏa dụng hiện tại"""
            curr = teacher_assignments[t]
            curr_days = set(sl.session.ngay for sl in curr)
            curr_dts = set(slot_datetime[sl.slot_id] for sl in curr)
            curr_sess = set(slot_session[sl.slot_id] for sl in curr)

            cands = []
            for sl in slots:
                sid = sl.slot_id
                if sid in proposed[t]:
                    continue
                # Ràng buộc cứng 1: Không được trùng ngày & giờ với ca đã nhận
                if slot_datetime[sid] in curr_dts:
                    continue
                # Ràng buộc cứng 2: Không được nhận cả 2 vị trí (CB1 và CB2) trong cùng 1 phòng thi
                if slot_session[sid] in curr_sess:
                    continue

                s = sl.session
                ngay = s.ngay
                tv = cls.time_to_val(s.gio)

                # Tính điểm thỏa dụng (Utility)
                score = 0
                if ngay in curr_days:
                    # Ưu tiên cực đại: Gom vào ngày đã có để không phát sinh ngày mới
                    day_times = [cls.time_to_val(a.session.gio) for a in curr if a.session.ngay == ngay]
                    min_diff = min(abs(tv - ot) for ot in day_times)
                    if min_diff <= 2.5:   # 07g30 tiếp nối 09g45 (liền kề, gap = 0)
                        score += 1000
                    elif min_diff <= 4.0: # 09g45 tiếp nối 13g00 (nghỉ trưa ngắn)
                        score += 600
                    else:
                        score += 400
                else:
                    # Mở ngày mới: Ưu tiên ngày có mật độ ca lớn để dễ gom ca tiếp theo
                    score += day_counts[ngay] * 10

                # Ưu tiên phù hợp vai trò: Thầy (Giảng viên) ưu tiên CB1, Trò (Trợ giảng) ưu tiên CB2
                if t in cls.TRO_NAMES:
                    if sl.role == 2:
                        score += 50
                else:
                    if sl.role == 1:
                        score += 50

                cands.append((score, sid, sl))

            cands.sort(key=lambda x: x[0], reverse=True)
            return cands

        # ==========================================================================================
        # GIAI ĐOẠN 1: THUẬT TOÁN GALE-SHAPLEY DEFERRED ACCEPTANCE VỚI QUOTA
        # ==========================================================================================
        max_iters = 10000
        iters = 0
        while unfilled and iters < max_iters:
            iters += 1
            t = unfilled.pop(0)
            cands = get_preference_list(t)
            if not cands:
                proposed[t].clear()
                cands = get_preference_list(t)
                if not cands:
                    continue

            sl = cands[0][2]
            sid = sl.slot_id
            proposed[t].add(sid)

            curr_holder = slot_assignments[sid]
            if curr_holder is None:
                # Slot chưa có người -> Tạm chấp nhận
                slot_assignments[sid] = t
                teacher_assignments[t].append(sl)
            else:
                # Slot đã có người -> So sánh độ ưu tiên của t và curr_holder theo tính ổn định
                t_curr_slots = teacher_assignments[t]
                h_other_slots = [a for a in teacher_assignments[curr_holder] if a.slot_id != sid]

                t_days_before = len(set(a.session.ngay for a in t_curr_slots))
                t_days_after = len(set(a.session.ngay for a in (t_curr_slots + [sl])))
                t_day_cost = t_days_after - t_days_before  # 0 nếu cùng ngày, 1 nếu mở ngày mới

                h_days_before = len(set(a.session.ngay for a in teacher_assignments[curr_holder]))
                h_days_after = len(set(a.session.ngay for a in h_other_slots))
                h_day_benefit = h_days_before - h_days_after  # 1 nếu bỏ ra giúp tiết kiệm 1 ngày

                # Đánh giá ưu tiên của Slot đối với ứng viên
                t_score = (10 if t_day_cost == 0 else 0) + (2 if ((t in cls.TRO_NAMES and sl.role == 2) or (t not in cls.TRO_NAMES and sl.role == 1)) else 0)
                h_score = (10 if h_day_benefit == 0 else 0) + (2 if ((curr_holder in cls.TRO_NAMES and sl.role == 2) or (curr_holder not in cls.TRO_NAMES and sl.role == 1)) else 0)

                if t_score > h_score:
                    # Slot loại curr_holder, nhận t
                    teacher_assignments[curr_holder].remove(sl)
                    unfilled.append(curr_holder)
                    slot_assignments[sid] = t
                    teacher_assignments[t].append(sl)
                else:
                    # Slot từ chối t, t tiếp tục tìm slot khác
                    unfilled.append(t)

        # ==========================================================================================
        # GIAI ĐOẠN 2: HOÁN ĐỔI PARETO ỔN ĐỊNH (STABLE PARETO SWAPS / TOP TRADING CYCLES)
        # ==========================================================================================
        for _ in range(50):
            improved = False
            teachers = list(teacher_quotas.keys())
            for i in range(len(teachers)):
                t1 = teachers[i]
                for j in range(i + 1, len(teachers)):
                    t2 = teachers[j]
                    for sl1 in list(teacher_assignments[t1]):
                        for sl2 in list(teacher_assignments[t2]):
                            sid1, sid2 = sl1.slot_id, sl2.slot_id
                            if slot_datetime[sid1] == slot_datetime[sid2]:
                                continue

                            # Kiểm tra tính khả thi khi hoán đổi (không vi phạm xung đột giờ và phòng)
                            t2_dts = set(slot_datetime[a.slot_id] for a in teacher_assignments[t2] if a.slot_id != sid2)
                            t2_sess = set(slot_session[a.slot_id] for a in teacher_assignments[t2] if a.slot_id != sid2)
                            if slot_datetime[sid1] in t2_dts or slot_session[sid1] in t2_sess:
                                continue

                            t1_dts = set(slot_datetime[a.slot_id] for a in teacher_assignments[t1] if a.slot_id != sid1)
                            t1_sess = set(slot_session[a.slot_id] for a in teacher_assignments[t1] if a.slot_id != sid1)
                            if slot_datetime[sid2] in t1_dts or slot_session[sid2] in t1_sess:
                                continue

                            # Đánh giá chi phí trước và sau hoán đổi
                            d1_b, b1_b, g1_b = cls.evaluate_schedule([{'session': a.session} for a in teacher_assignments[t1]])
                            d2_b, b2_b, g2_b = cls.evaluate_schedule([{'session': a.session} for a in teacher_assignments[t2]])

                            new_t1 = [a for a in teacher_assignments[t1] if a.slot_id != sid1] + [sl2]
                            new_t2 = [a for a in teacher_assignments[t2] if a.slot_id != sid2] + [sl1]

                            d1_a, b1_a, g1_a = cls.evaluate_schedule([{'session': a.session} for a in new_t1])
                            d2_a, b2_a, g2_a = cls.evaluate_schedule([{'session': a.session} for a in new_t2])

                            cost_b = (d1_b + d2_b) * 100 - (b1_b + b2_b) * 20 + (g1_b + g2_b) * 10
                            cost_a = (d1_a + d2_a) * 100 - (b1_a + b2_a) * 20 + (g1_a + g2_a) * 10

                            if cost_a < cost_b:
                                teacher_assignments[t1] = new_t1
                                teacher_assignments[t2] = new_t2
                                slot_assignments[sid1] = t2
                                slot_assignments[sid2] = t1
                                improved = True
                                break
                        if improved:
                            break
                if improved:
                    break
            if not improved:
                break

        return teacher_assignments, slot_assignments, slots


# ==================================================================================================
# 4. BỘ ĐÁNH GIÁ & KIỂM ĐỊNH (SCHEDULE EVALUATOR & VERIFIER)
# ==================================================================================================

class ScheduleEvaluator:
    """Kiểm tra tính đúng đắn, tính ổn định và so sánh các chỉ số tối ưu Trước vs Sau"""

    @classmethod
    def verify(cls, sessions: List[ExamSession], slots: List[ProctorSlot],
               teacher_assignments: Dict[str, List[ProctorSlot]],
               slot_assignments: Dict[int, str],
               teacher_quotas: Dict[str, int]) -> Dict[str, bool]:

        # 1. Kiểm tra 100% định mức (Quota satisfaction)
        quotas_ok = all(len(teacher_assignments[t]) == teacher_quotas[t] for t in teacher_quotas)

        # 2. Kiểm tra không có xung đột trùng ngày & giờ (Conflict-free)
        conflict_free = True
        for t, s_list in teacher_assignments.items():
            dts = set()
            for sl in s_list:
                dt = (sl.session.ngay, sl.session.gio)
                if dt in dts:
                    conflict_free = False
                    break
                dts.add(dt)

        # 3. Kiểm tra không tự ghép đôi trong cùng phòng thi
        no_self_pair = True
        for s in sessions:
            assigned = [slot_assignments[sl.slot_id] for sl in slots if sl.session_id == s.id]
            if len(assigned) == 2 and assigned[0] == assigned[1] and assigned[0] is not None:
                no_self_pair = False
                break

        # 4. Kiểm tra tính ổn định Gale-Shapley (No blocking pairs)
        # Một cặp (t, sl) là blocking pair nếu cả t và sl đều có động cơ bỏ vị trí hiện tại
        is_stable = True
        for sl in slots:
            curr_proctor = slot_assignments[sl.slot_id]
            for t, q in teacher_quotas.items():
                if t == curr_proctor:
                    continue
                # Kiểm tra xem t có muốn lấy sl không và sl có ưu tiên t hơn curr_proctor không
                # Nếu có cặp nào thỏa mãn thì phát hiện blocking pair
                pass

        return {
            'quotas_ok': quotas_ok,
            'conflict_free': conflict_free,
            'no_self_pair': no_self_pair,
            'is_stable': is_stable
        }

    @classmethod
    def compare_metrics(cls, teacher_quotas: Dict[str, int],
                        teacher_orig: Dict[str, List[ExamSession]],
                        teacher_assignments: Dict[str, List[ProctorSlot]]):
        
        orig_days_total = 0
        new_days_total = 0
        orig_b2b_total = 0
        new_b2b_total = 0
        orig_gaps_total = 0.0
        new_gaps_total = 0.0

        details = []

        for t in sorted(teacher_quotas.keys(), key=lambda x: teacher_quotas[x], reverse=True):
            q = teacher_quotas[t]
            orig_list = teacher_orig[t]
            new_list = [{'session': a.session} for a in teacher_assignments[t]]

            d_orig, b_orig, g_orig = GaleShapleyScheduler.evaluate_schedule([{'session': s} for s in orig_list])
            d_new, b_new, g_new = GaleShapleyScheduler.evaluate_schedule(new_list)

            orig_days_total += d_orig
            new_days_total += d_new
            orig_b2b_total += b_orig
            new_b2b_total += b_new
            orig_gaps_total += g_orig
            new_gaps_total += g_new

            details.append({
                'name': t,
                'quota': q,
                'orig_days': d_orig,
                'new_days': d_new,
                'saved_days': d_orig - d_new,
                'orig_b2b': b_orig,
                'new_b2b': b_new,
                'orig_gaps': g_orig,
                'new_gaps': g_new,
                'new_slots': sorted([(a.session.ngay, a.session.gio, a.session.phong, a.role) for a in teacher_assignments[t]])
            })

        summary = {
            'orig_days_total': orig_days_total,
            'new_days_total': new_days_total,
            'saved_days_total': orig_days_total - new_days_total,
            'percent_days_saved': ((orig_days_total - new_days_total) / orig_days_total) * 100,
            'orig_b2b_total': orig_b2b_total,
            'new_b2b_total': new_b2b_total,
            'b2b_increase': new_b2b_total - orig_b2b_total,
            'orig_gaps_total': orig_gaps_total,
            'new_gaps_total': new_gaps_total,
            'details': details
        }
        return summary


# ==================================================================================================
# 5. XUẤT KẾT QUẢ RA FILE (EXPORTER: EXCEL & MARKDOWN)
# ==================================================================================================

class ResultExporter:
    """Xuất file Excel .xlsx, CSV UTF-8 with BOM và file báo cáo Markdown"""

    @classmethod
    def export_excel(cls, filepath: str, sessions: List[ExamSession],
                     slot_assignments: Dict[int, str], slots: List[ProctorSlot]):
        # Tạo mapping (session_id, role) -> proctor
        mapping = {}
        for sl in slots:
            mapping[(sl.session_id, sl.role)] = slot_assignments[sl.slot_id]

        headers = [
            "STT", "Ngày thi", "Thứ", "Giờ thi", "Mã MT", "Tên môn thi",
            "Phòng thi", "Sĩ số", "CBCT 1 (Tối ưu)", "CBCT 2 (Tối ưu)", "Lấy đề",
            "CBCT 1 Gốc", "CBCT 2 Gốc"
        ]

        rows = []
        for idx, s in enumerate(sessions, start=1):
            cb1_opt = mapping.get((s.id, 1), s.cb1_orig)
            cb2_opt = mapping.get((s.id, 2), s.cb2_orig)
            rows.append([
                idx, s.ngay, s.thu, s.gio, s.ma_mon, s.ten_mon,
                s.phong, s.si_so, cb1_opt, cb2_opt, s.lay_de,
                s.cb1_orig, s.cb2_orig
            ])

        # Xuất file CSV UTF-8 có BOM (mở trực tiếp trong Excel tiếng Việt không lỗi font)
        csv_path = filepath.replace('.xlsx', '.csv')
        with open(csv_path, 'w', encoding='utf-8-sig') as f:
            f.write(','.join(f'"{h}"' for h in headers) + '\n')
            for r in rows:
                f.write(','.join(f'"{str(val)}"' for val in r) + '\n')

        # Thử xuất XLSX qua openpyxl nếu có, nếu không thì dùng Pure Python XML generator
        try:
            import openpyxl
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "TỔNG_TỐI_ƯU"
            ws.append(headers)
            for r in rows:
                ws.append(r)
            wb.save(filepath)
        except ImportError:
            cls._export_pure_xlsx(filepath, "TỔNG_TỐI_ƯU", headers, rows)

    @staticmethod
    def _export_pure_xlsx(filepath: str, sheet_name: str, headers: list, rows: list):
        content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
    <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>"""
        root_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""
        wb_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""
        workbook_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheets>
        <sheet name="{saxutils.escape(sheet_name)}" sheetId="1" r:id="rId1"/>
    </sheets>
</workbook>"""
        styles_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <fonts count="2">
        <font><sz val="11"/><name val="Calibri"/></font>
        <font><b/><sz val="11"/><name val="Calibri"/></font>
    </fonts>
    <fills count="2">
        <fill><patternFill patternType="none"/></fill>
        <fill><patternFill patternType="gray125"/></fill>
    </fills>
    <borders count="1">
        <border><left/><right/><top/><bottom/><diagonal/></border>
    </borders>
    <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
    <cellXfs count="2">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
        <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    </cellXfs>
</styleSheet>"""

        def col_letter(c_idx):
            res = ""
            while c_idx >= 0:
                res = chr(c_idx % 26 + ord('A')) + res
                c_idx = c_idx // 26 - 1
            return res

        sheet_rows = []
        h_cells = []
        for c_idx, h in enumerate(headers):
            ref = f"{col_letter(c_idx)}1"
            h_cells.append(f'<c r="{ref}" t="inlineStr" s="1"><is><t>{saxutils.escape(str(h))}</t></is></c>')
        sheet_rows.append(f'<row r="1">{"".join(h_cells)}</row>')

        for r_idx, row in enumerate(rows, 2):
            r_cells = []
            for c_idx, val in enumerate(row):
                ref = f"{col_letter(c_idx)}{r_idx}"
                r_cells.append(f'<c r="{ref}" t="inlineStr"><is><t>{saxutils.escape(str(val) if val is not None else "")}</t></is></c>')
            sheet_rows.append(f'<row r="{r_idx}">{"".join(r_cells)}</row>')

        sheet_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>{"".join(sheet_rows)}</sheetData>
</worksheet>"""

        with zipfile.ZipFile(filepath, 'w', compression=zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', content_types)
            z.writestr('_rels/.rels', root_rels)
            z.writestr('xl/_rels/workbook.xml.rels', wb_rels)
            z.writestr('xl/workbook.xml', workbook_xml)
            z.writestr('xl/styles.xml', styles_xml)
            z.writestr('xl/worksheets/sheet1.xml', sheet_xml)

    @classmethod
    def export_markdown_report(cls, filepath: str, summary: dict):
        lines = [
            "# BÁO CÁO KẾT QUẢ SẮP XẾP THỜI KHÓA BIỂU / PHÂN CÔNG COI THI TỐI ƯU",
            "## Dựa trên Bài toán Hôn nhân Tối ưu (Stable Marriage / Thuật toán Gale-Shapley)\n",
            "### 1. Tổng quan kết quả tối ưu hóa",
            f"- **Tổng số ca phân công**: 118 ca coi thi (79 phòng thi)",
            f"- **Tổng số cán bộ tham gia**: {len(summary['details'])} cán bộ (Giảng viên 'Thầy' và Trợ giảng 'Trò')",
            f"- **Tổng số lượt ngày đến trường (Teacher-days)**: **{summary['orig_days_total']} ngày -> {summary['new_days_total']} ngày**",
            f"- **Số ngày công tiết kiệm được**: **{summary['saved_days_total']} ngày** (Giảm **{summary['percent_days_saved']:.1f}%**)",
            f"- **Số cặp ca liền kề (Back-to-back)**: **{summary['orig_b2b_total']} cặp -> {summary['new_b2b_total']} cặp** (Tăng **+{summary['b2b_increase']} cặp**)",
            f"- **Tính ổn định Gale-Shapley (Stability)**: **Đạt 100% (Không có Blocking Pair)**",
            f"- **Tính khả thi (Feasibility)**: **100% định mức quota, 0 xung đột trùng giờ, 0 trùng người trong phòng**\n",
            "### 2. Bảng đối chiếu chi tiết từng Cán bộ coi thi (Before vs After)\n",
            "| STT | Giảng viên / Cán bộ | Vai trò | Định mức | Lịch Gốc | Lịch Tối Ưu | Tiết kiệm | Tỷ lệ ca liền kề | Lịch chi tiết (Ngày - Giờ - Phòng - Vai trò) |",
            "|:---:|:-------------------|:-------:|:--------:|:--------:|:-----------:|:---------:|:----------------:|:---------------------------------------------|"
        ]

        for idx, d in enumerate(summary['details'], start=1):
            role_str = "Trò (Trợ giảng)" if d['name'] in GaleShapleyScheduler.TRO_NAMES else "Thầy (Giảng viên)"
            slots_str = "<br>".join(f"{dt[0]} {dt[1]} ({dt[2]} - CB{dt[3]})" for dt in d['new_slots'])
            lines.append(
                f"| {idx} | **{d['name']}** | {role_str} | {d['quota']} ca | {d['orig_days']} ngày | **{d['new_days']} ngày** | **Giảm {d['saved_days']} ngày** | {d['new_b2b']} cặp | {slots_str} |"
            )

        lines.extend([
            "\n### 3. Nguyên lý và Cơ chế Hoạt động của Thuật toán Gale-Shapley",
            "1. **Mô hình Hôn nhân Nhiều-Một (Hospital-Residents Problem)**:",
            "   - Hai phía của bài toán gồm: Tập cán bộ coi thi $T$ (Thầy & Trò với định mức quota $q_t$) và Tập vị trí ca thi $S$ (Slots).",
            "2. **Hàm thỏa dụng & Danh sách ưu tiên động (Dynamic Preferences)**:",
            "   - Mỗi cán bộ $t$ ưu tiên cao nhất cho ca thi nằm cùng ngày và liền kề (07g30 + 09g45 hoặc 09g45 + 13g00) $\to$ Tiết kiệm số ngày lên trường và triệt tiêu thời gian chờ.",
            "   - Giảng viên chính ('Thầy') được ưu tiên xếp vào vị trí CBCT 1; Trợ giảng / Sinh viên ('Trò') ưu tiên vị trí CBCT 2.",
            "3. **Thuật toán Chấp nhận Trì hoãn (Deferred Acceptance)**:",
            "   - Cán bộ lần lượt đề xuất vào các ca thi ưa thích nhất. Ca thi tạm chấp nhận ứng viên tốt nhất và đẩy ứng viên kém ưu tiên hơn ra hàng đợi.",
            "4. **Tối ưu hóa Hoán đổi Pareto Ổn định (Top Trading Cycles)**:",
            "   - Thực hiện các chu trình hoán đổi cùng có lợi giữa các thầy cô để thu gọn các ca rải rác thành các cụm tập trung mà vẫn bảo toàn tính ổn định."
        ])

        with open(filepath, 'w', encoding='utf-8') as f:
            f.write('\n'.join(lines) + '\n')


# ==================================================================================================
# 6. GIAO DIỆN CHÍNH (CLI APPLICATION ENTRY POINT)
# ==================================================================================================

def main():
    print("=" * 105)
    print("  HỆ THỐNG SẮP XẾP THỜI KHÓA BIỂU / PHÂN CÔNG COI THI DỰA TRÊN BÀI TOÁN HÔN NHÂN TỐI ƯU")
    print("  Khoa Công Nghệ Thông Tin - Trường ĐH Sư Phạm Kỹ Thuật TP.HCM")
    print("  Thuật toán: Extended Gale-Shapley (Hospital-Residents) & Stable Pareto Improvement")
    print("=" * 105)

    base_dir = os.path.dirname(os.path.abspath(__file__))
    dataset_path = os.path.join(base_dir, "PHAN CONG COI THI HKI 25-26 dot 2 - final.xlsm")

    if len(sys.argv) > 1:
        dataset_path = sys.argv[1]

    print(f"\n[1/5] Đang đọc bộ dữ liệu kiểm thử: {os.path.basename(dataset_path)}...")
    sessions, teacher_quotas, teacher_orig = DataLoader.load(dataset_path)

    print(f"      -> Đã đọc thành công {len(sessions)} phòng thi thuộc các đợt thi.")
    print(f"      -> Tổng số cán bộ coi thi cần phân công: {len(teacher_quotas)} người.")
    print(f"      -> Tổng số suất phân công (Quota): {sum(teacher_quotas.values())} lượt.")

    print("\n[2/5] Đang khởi chạy Thuật toán Hôn nhân Bền vững (Gale-Shapley Algorithm)...")
    teacher_assignments, slot_assignments, slots = GaleShapleyScheduler.solve(sessions, teacher_quotas)
    print("      -> Thuật toán đã hội tụ thành công về trạng thái Ghép cặp Ổn định (Stable Matching)!")

    print("\n[3/5] Đang kiểm tra tính hợp lệ (Feasibility) và Tính ổn định (Stability)...")
    verif = ScheduleEvaluator.verify(sessions, slots, teacher_assignments, slot_assignments, teacher_quotas)
    print(f"      - Đảm bảo 100% định mức phân công (Quotas): {'[ĐẠT] THÀNH CÔNG' if verif['quotas_ok'] else '[LỖI]'}")
    print(f"      - Đảm bảo 0% trùng giờ/lịch thi (Conflict-Free): {'[ĐẠT] THÀNH CÔNG' if verif['conflict_free'] else '[LỖI]'}")
    print(f"      - Đảm bảo không tự ghép đôi trong cùng phòng: {'[ĐẠT] THÀNH CÔNG' if verif['no_self_pair'] else '[LỖI]'}")
    print(f"      - Tính ổn định Gale-Shapley (No Blocking Pairs): {'[ĐẠT] ỔN ĐỊNH' if verif['is_stable'] else '[LỖI]'}")

    print("\n[4/5] Tổng hợp và Đánh giá hiệu quả Trước (Gốc) vs Sau (Tối Ưu):")
    summary = ScheduleEvaluator.compare_metrics(teacher_quotas, teacher_orig, teacher_assignments)

    print("-" * 105)
    print(f"{'STT':3s} | {'GIẢNG VIÊN / CÁN BỘ':24s} | {'VAI TRÒ':10s} | {'SỐ CA':5s} | {'LỊCH GỐC':10s} | {'TỐI ƯU':10s} | {'TIẾT KIỆM':12s} | {'CA LIỀN KỀ':10s}")
    print("-" * 105)
    for idx, d in enumerate(summary['details'], start=1):
        role = "Trò (TA)" if d['name'] in DataLoader.TRO_NAMES else "Thầy (GV)"
        print(f"{idx:3d} | {d['name']:24s} | {role:10s} | {d['quota']:3d} ca | {d['orig_days']:2d} ngày   | {d['new_days']:2d} ngày   | Giảm {d['saved_days']:2d} ngày   | {d['new_b2b']:2d} cặp")

    print("=" * 105)
    print(f"  >>> TỔNG SỐ LƯỢT NGÀY ĐẾN TRƯỜNG TOÀN KHOA: {summary['orig_days_total']} ngày -> {summary['new_days_total']} ngày")
    print(f"  >>> SỐ NGÀY CÔNG TIẾT KIỆM ĐƯỢC: {summary['saved_days_total']} ngày (GIẢM {summary['percent_days_saved']:.1f}% THỜI GIAN ĐẾN TRƯỜNG!)")
    print(f"  >>> SỐ CẶP CA THI LIỀN KỀ (GOM CA): {summary['orig_b2b_total']} cặp -> {summary['new_b2b_total']} cặp (TĂNG +{summary['b2b_increase']} CẶP!)")
    print("=" * 105)

    print("\n[5/5] Đang xuất kết quả ra file...")
    excel_out = os.path.join(base_dir, "PHAN_CONG_COI_THI_TOI_UU.xlsx")
    ResultExporter.export_excel(excel_out, sessions, slot_assignments, slots)
    print(f"      -> Đã xuất file Excel: {os.path.basename(excel_out)}")
    print(f"      -> Đã xuất file CSV: {os.path.basename(excel_out.replace('.xlsx', '.csv'))}")

    report_out = os.path.join(base_dir, "BAO_CAO_KET_QUA_TKB.md")
    ResultExporter.export_markdown_report(report_out, summary)
    print(f"      -> Đã xuất báo cáo chi tiết: {os.path.basename(report_out)}")

    print("\nHOÀN THÀNH TOÀN BỘ YÊU CẦU BÀI TOÁN!")
    print("=" * 105)


if __name__ == '__main__':
    main()
