#!/usr/bin/env python3
"""
Fetch bond information from VSD (Vietnamese Securities Depository)
URL: https://www.vsd.vn/vi/tin-thi-truong-co-so

Crawl trang tin tức thị trường cơ sở:
1. Lấy danh sách tin có mã CK từ ngày gần nhất
2. Mở từng tin tức để extract chi tiết thông tin
3. Return danh sách mã + chi tiết thông tin quyền
"""

import requests
from bs4 import BeautifulSoup
import json
import sys
import time
import os
from datetime import datetime, timedelta, timezone
VN_TZ = timezone(timedelta(hours=7))
import re
import logging
import hashlib
from concurrent.futures import ThreadPoolExecutor, as_completed

try:
    import pandas as pd
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False

# ============================================================================
# CONFIGURATION - Chọn chế độ lấy dữ liệu:
# ============================================================================
# RUN_MODE có thể nhận các giá trị:
# - "KEEP_DAYS": Lấy các bản tin trong N ngày gần nhất (cấu hình KEEP_DAYS ở dưới)
# - "DATE_RANGE": Lấy các bản tin trong khoảng ngày cụ thể (cấu hình DATE_FROM/DATE_TO ở dưới)
# - "EXCEL_URLS": Chỉ lấy các bản tin có URL trong file Excel và lọc theo khoảng ngày DATE_FROM/DATE_TO
RUN_MODE = os.environ.get("RUN_MODE", "EXCEL_URLS")

# --- Chế độ 1: KEEP_DAYS ---
# Lấy các bản tin trong N ngày GẦN NHẤT tính từ hôm nay.
KEEP_DAYS = int(os.environ.get("KEEP_DAYS", "7"))  # Ví dụ: 90 = lấy 90 ngày gần nhất

# --- Chế độ 2 & 3: DATE_RANGE & EXCEL_URLS ---
# Lấy các bản tin được đăng trong KHOẢNG NGÀY CỤ THỂ (dd/mm/yyyy).
DATE_FROM = os.environ.get("DATE_FROM", "23/12/2025")
DATE_TO   = os.environ.get("DATE_TO", "23/04/2026")

# File Excel chứa danh sách URL cho chế độ EXCEL_URLS
EXCEL_URLS_FILE = "url_2fetch_round2.xlsx"

# ============================================================================

logging.basicConfig(level=logging.INFO, format='%(message)s', stream=sys.stderr)
logger = logging.getLogger(__name__)

# --- BUSINESS RULES HELPER FUNCTIONS ---

def remove_accents_and_lower(text):
    """
    Tiền xử lý chuỗi: chuyển sang lowercase, loại bỏ dấu tiếng Việt,
    loại bỏ khoảng trắng thừa và ký tự đặc biệt nếu có.
    """
    if not text:
        return ""
    # Chuyển về chữ thường
    text = str(text).lower().strip()
    
    # Loại bỏ khoảng trắng thừa ngang, giữ nguyên ký tự xuống dòng
    lines = [re.sub(r'[ \t\r\f\v]+', ' ', line).strip() for line in text.split('\n')]
    text = '\n'.join(lines)
    
    # Loại bỏ dấu tiếng Việt
    accents_map = {
        'a': 'áàảãạăắằẳẵặâấầẩẫậ',
        'e': 'éèẻẽẹêếềểễệ',
        'i': 'íìỉĩị',
        'o': 'óòỏõọôốồổỗộơớờởỡợ',
        'u': 'úùủũụưứừửữự',
        'y': 'ýỳỷỹỵ',
        'd': 'đ',
    }
    
    for char, accented_chars in accents_map.items():
        for ac in accented_chars:
            text = text.replace(ac, char)
            text = text.replace(ac.upper(), char.upper())
            
    return text


def get_last_day_of_month(month, year):
    """
    Trả về ngày cuối cùng của tháng trong năm tương ứng.
    """
    import calendar
    try:
        return calendar.monthrange(year, month)[1]
    except Exception:
        return 28  # Safe default fallback

def extract_earliest_date(text_segment):
    """
    Tìm chuỗi ngày sớm nhất trong text_segment (khoảng index nhỏ nhất).
    Hỗ trợ nhiều định dạng và chuyển đổi thang mm nam YYYY sang ngày cuối tháng.
    """
    if not text_segment:
        return None
        
    matches = []
    
    # 1. ngay dd thang mm nam YYYY
    pattern1 = re.compile(r'ngay\s+(\d{1,2})\s+thang\s+(\d{1,2})\s+nam\s+(\d{4})', re.IGNORECASE)
    for m in pattern1.finditer(text_segment):
        d, m_val, y = m.group(1), m.group(2), m.group(3)
        try:
            matches.append((m.start(), f"{int(d):02d}/{int(m_val):02d}/{y}"))
        except:
            pass
        
    # 2. thang mm nam YYYY
    pattern2 = re.compile(r'thang\s+(\d{1,2})\s+nam\s+(\d{4})', re.IGNORECASE)
    for m in pattern2.finditer(text_segment):
        m_val, y = m.group(1), m.group(2)
        try:
            last_day = get_last_day_of_month(int(m_val), int(y))
            matches.append((m.start(), f"{last_day:02d}/{int(m_val):02d}/{y}"))
        except:
            pass
        
    # 3. dd/mm/YYYY
    pattern3 = re.compile(r'(\d{1,2})/(\d{1,2})/(\d{4})')
    for m in pattern3.finditer(text_segment):
        d, m_val, y = m.group(1), m.group(2), m.group(3)
        try:
            matches.append((m.start(), f"{int(d):02d}/{int(m_val):02d}/{y}"))
        except:
            pass
        
    # 4. mm/YYYY
    pattern4 = re.compile(r'(\d{1,2})/(\d{4})')
    for m in pattern4.finditer(text_segment):
        m_val, y = m.group(1), m.group(2)
        try:
            start = m.start()
            if start > 0 and text_segment[start-1] in '/0123456789':
                continue
            end = m.end()
            if end < len(text_segment) and text_segment[end] in '/0123456789':
                continue
            last_day = get_last_day_of_month(int(m_val), int(y))
            matches.append((m.start(), f"{last_day:02d}/{int(m_val):02d}/{y}"))
        except:
            pass

    if not matches:
        return None
        
    # Sắp xếp theo vị trí xuất hiện (start index) để lấy chuỗi sớm nhất
    matches.sort(key=lambda x: x[0])
    return matches[0][1]

def extract_all_dates_in_segment(text_segment):
    """
    Trích xuất toàn bộ ngày có trong đoạn văn bản, sắp xếp theo thứ tự xuất hiện.
    """
    if not text_segment:
        return []
        
    matches = []
    
    # 1. ngay dd thang mm nam YYYY
    pattern1 = re.compile(r'ngay\s+(\d{1,2})\s+thang\s+(\d{1,2})\s+nam\s+(\d{4})', re.IGNORECASE)
    for m in pattern1.finditer(text_segment):
        d, m_val, y = m.group(1), m.group(2), m.group(3)
        try:
            matches.append((m.start(), f"{int(d):02d}/{int(m_val):02d}/{y}"))
        except:
            pass
        
    # 2. thang mm nam YYYY
    pattern2 = re.compile(r'thang\s+(\d{1,2})\s+nam\s+(\d{4})', re.IGNORECASE)
    for m in pattern2.finditer(text_segment):
        m_val, y = m.group(1), m.group(2)
        try:
            last_day = get_last_day_of_month(int(m_val), int(y))
            matches.append((m.start(), f"{last_day:02d}/{int(m_val):02d}/{y}"))
        except:
            pass
        
    # 3. dd/mm/YYYY
    pattern3 = re.compile(r'(\d{1,2})/(\d{1,2})/(\d{4})')
    for m in pattern3.finditer(text_segment):
        d, m_val, y = m.group(1), m.group(2), m.group(3)
        try:
            matches.append((m.start(), f"{int(d):02d}/{int(m_val):02d}/{y}"))
        except:
            pass
        
    # 4. mm/YYYY
    pattern4 = re.compile(r'(\d{1,2})/(\d{4})')
    for m in pattern4.finditer(text_segment):
        m_val, y = m.group(1), m.group(2)
        try:
            start = m.start()
            if start > 0 and text_segment[start-1] in '/0123456789':
                continue
            end = m.end()
            if end < len(text_segment) and text_segment[end] in '/0123456789':
                continue
            last_day = get_last_day_of_month(int(m_val), int(y))
            matches.append((m.start(), f"{last_day:02d}/{int(m_val):02d}/{y}"))
        except:
            pass

    if not matches:
        return []
        
    matches.sort(key=lambda x: x[0])
    return [x[1] for x in matches]

def extract_original_line_by_keyword(original_text, keywords):
    """
    Tìm một dòng trong original_text chứa một trong các keywords không dấu.
    Trả về toàn bộ dòng gốc (giữ nguyên dấu tiếng Việt).
    """
    if not original_text:
        return None
    lines = original_text.split('\n')
    for line in lines:
        pre_line = remove_accents_and_lower(line)
        for kw in keywords:
            if kw in pre_line:
                return line.strip()
    return None

def check_keyword_with_negative_prefixes(text, keyword, negative_prefixes):
    """
    Kiểm tra xem keyword có xuất hiện trong text hay không,
    với điều kiện nó không bị bắt đầu bởi bất kỳ negative_prefix nào.
    """
    if not text or not keyword:
        return False
    start = 0
    while True:
        pos = text.find(keyword, start)
        if pos == -1:
            return False
        # Kiểm tra xem có tiền tố phủ định nào đứng ngay trước keyword không
        is_negative = False
        for prefix in negative_prefixes:
            prefix_len = len(prefix)
            if pos >= prefix_len:
                preceding = text[pos - prefix_len:pos]
                if preceding == prefix:
                    is_negative = True
                    break
        if not is_negative:
            return True
        start = pos + 1

def find_earliest_keyword(text, keywords):
    """
    Tìm vị trí xuất hiện sớm nhất của một trong các keywords trong text.
    Trả về (vị trí, keyword) hoặc (-1, None) nếu không tìm thấy.
    """
    if not text:
        return -1, None
    earliest_pos = -1
    earliest_kw = None
    for kw in keywords:
        pos = text.find(kw)
        if pos != -1:
            if earliest_pos == -1 or pos < earliest_pos:
                earliest_pos = pos
                earliest_kw = kw
    return earliest_pos, earliest_kw

def parse_vietnamese_float(num_str):
    """
    Parse chuỗi số theo định dạng tiếng Việt (dùng dấu phẩy làm phần thập phân
    hoặc dấu chấm làm phần ngăn cách hàng nghìn) thành kiểu float.
    """
    if not num_str:
        return None
    num_str = num_str.strip()
    if '.' in num_str and ',' in num_str:
        num_str = num_str.replace('.', '').replace(',', '.')
    elif ',' in num_str:
        num_str = num_str.replace(',', '.')
    elif '.' in num_str:
        # Nhận dạng nếu là ngăn cách hàng nghìn (ví dụ "10.000") hay thập phân (ví dụ "2.5")
        parts = num_str.split('.')
        if len(parts[-1]) == 3 and len(parts) > 1:
            num_str = num_str.replace('.', '')
    try:
        return float(num_str)
    except ValueError:
        return None

class VSDFetcher:
    def __init__(self):
        """
        Khởi tạo VSDFetcher

        Số ngày cần lấy được điều chỉnh bằng hằng số KEEP_DAYS ở đầu file
        """
        self.base_url = "https://www.vsd.vn"
        self.news_url = "https://www.vsd.vn/vi/tin-thi-truong-co-so"
        self.session = requests.Session()
        self.vptoken = None  # Token để AJAX POST phân trang
        # --- Xác định chế độ chạy ---
        self.date_from = None
        self.date_to   = None
        
        # Đọc RUN_MODE từ global scope (mặc định AUTO để tương thích ngược)
        run_mode = globals().get('RUN_MODE', 'AUTO')
        
        if run_mode == 'EXCEL_URLS':
            try:
                if not DATE_FROM or not DATE_TO:
                    raise ValueError("Chế độ EXCEL_URLS yêu cầu phải cấu hình cả DATE_FROM và DATE_TO")
                self.date_from = datetime.strptime(DATE_FROM.strip(), '%d/%m/%Y').date()
                self.date_to   = datetime.strptime(DATE_TO.strip(),   '%d/%m/%Y').date()
                if self.date_from > self.date_to:
                    raise ValueError(f"DATE_FROM ({DATE_FROM}) phải <= DATE_TO ({DATE_TO})")
                self.mode = 'excel_urls'
                logger.info(f"⚙ Chế độ: EXCEL_URLS (Lọc khoảng ngày {DATE_FROM} → {DATE_TO} theo danh sách Excel)")
            except ValueError as e:
                logger.error(f"✗ Cấu hình EXCEL_URLS không hợp lệ: {e}")
                raise
        elif run_mode == 'DATE_RANGE':
            try:
                if not DATE_FROM or not DATE_TO:
                    raise ValueError("Chế độ DATE_RANGE yêu cầu phải cấu hình cả DATE_FROM và DATE_TO")
                self.date_from = datetime.strptime(DATE_FROM.strip(), '%d/%m/%Y').date()
                self.date_to   = datetime.strptime(DATE_TO.strip(),   '%d/%m/%Y').date()
                if self.date_from > self.date_to:
                    raise ValueError(f"DATE_FROM ({DATE_FROM}) phải <= DATE_TO ({DATE_TO})")
                self.mode = 'date_range'
                logger.info(f"⚙ Chế độ: DATE_RANGE ({DATE_FROM} → {DATE_TO})")
            except ValueError as e:
                logger.error(f"✗ Cấu hình DATE_RANGE không hợp lệ: {e}")
                raise
        elif run_mode == 'KEEP_DAYS':
            if not KEEP_DAYS:
                raise ValueError("Chế độ KEEP_DAYS yêu cầu phải cấu hình KEEP_DAYS")
            self.keep_days = int(KEEP_DAYS)
            self.mode = 'keep_days'
            logger.info(f"⚙ Chế độ: KEEP_DAYS={self.keep_days}")
        else:
            # Tự động nhận diện (Tương thích ngược)
            if DATE_FROM and DATE_TO:
                try:
                    self.date_from = datetime.strptime(DATE_FROM.strip(), '%d/%m/%Y').date()
                    self.date_to   = datetime.strptime(DATE_TO.strip(),   '%d/%m/%Y').date()
                    if self.date_from > self.date_to:
                        raise ValueError(f"DATE_FROM ({DATE_FROM}) phải <= DATE_TO ({DATE_TO})")
                    self.mode = 'date_range'
                    logger.info(f"⚙ Chế độ tự động: DATE_RANGE ({DATE_FROM} → {DATE_TO})")
                except ValueError as e:
                    logger.error(f"✗ Cấu hình DATE_RANGE không hợp lệ: {e}")
                    raise
            elif KEEP_DAYS:
                self.keep_days = int(KEEP_DAYS)
                self.mode = 'keep_days'
                logger.info(f"⚙ Chế độ tự động: KEEP_DAYS={self.keep_days}")
            else:
                raise ValueError("Phải cấu hình ít nhất một trong RUN_MODE, KEEP_DAYS hoặc DATE_FROM+DATE_TO")
        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            'Accept-Language': 'vi-VN,vi;q=0.9',
            'Connection': 'keep-alive',
        }

    def apply_business_rules(self, record):
        """
        Áp dụng các quy tắc nghiệp vụ từ rules_2finalize.md lên một bản ghi.
        Trả về bản ghi mới chứa đầy đủ 25 trường nghiệp vụ cùng các thông tin bổ sung.
        """
        # 1. Trích xuất thông tin thô
        code = record.get('code') or ''
        title = record.get('title') or ''
        ly_do = record.get('lý_do_mục_đích') or ''
        noi_gd = record.get('nơi_giao_dịch') or ''
        loai_ck = record.get('loại_chứng_khoán') or ''
        text_content = record.get('text_content') or ''
        
        # 2. Tiền xử lý (bỏ dấu tiếng Việt, lower key, bỏ khoảng trắng thừa)
        pre_title = remove_accents_and_lower(title)
        pre_ly_do = remove_accents_and_lower(ly_do)
        pre_noi_gd = remove_accents_and_lower(noi_gd)
        pre_loai_ck = remove_accents_and_lower(loai_ck)
        pre_text = remove_accents_and_lower(text_content)
        
        # 3. Tính toán các trường
        # MaChungKhoan
        ma_ck = code
        
        # NhomQuyen
        nhom_quyen = None
        # a. Tin huỷ
        # "huy" ngay sau ": " hoặc ":" đầu tiên, và không chứa "chung quyen"
        col_idx = pre_title.find(':')
        if col_idx != -1:
            after_col = pre_title[col_idx+1:].strip()
            if after_col.startswith('huy') and 'chung quyen' not in pre_title:
                nhom_quyen = "Tin huỷ"
        else:
            # Dự phòng: nếu không có dấu hai chấm nhưng bắt đầu bằng hủy
            if pre_title.startswith('huy') and 'chung quyen' not in pre_title:
                nhom_quyen = "Tin huỷ"
                
        # b. Thay đổi
        if nhom_quyen is None:
            tc_keywords = ["thay doi", "chuyen du lieu", "chuyen san", "dinh chinh", "dieu chinh"]
            if any(kw in pre_title for kw in tc_keywords) and 'chung quyen' not in pre_title:
                nhom_quyen = "Thay đổi"
                
        # c. Đăng ký Lưu ký
        if nhom_quyen is None:
            dk_keywords = ["dang ky chung chi", "dang ky chung khoan", "dang ky co phieu", "dang ky trai phieu", "luu ky chung chi", "luu ky chung khoan", "luu ky co phieu", "luu ky trai phieu"]
            neg_prefixes = ["to chuc ", "to chuc dang ky ", "to chuc luu ky ", "to chuc dang ky luu ky ", "to chuc luu ky dang ky "]
            has_dk = False
            for kw in dk_keywords:
                if check_keyword_with_negative_prefixes(pre_title, kw, neg_prefixes) or check_keyword_with_negative_prefixes(pre_ly_do, kw, neg_prefixes):
                    has_dk = True
                    break
            if has_dk:
                nhom_quyen = "Đăng ký Lưu ký"
                
        # d. Cổ tức cổ phiếu / Cổ phiếu thưởng
        if nhom_quyen is None:
            ctcp_keywords = ["phat hanh co phieu", "nhan co phieu", "nhan them co phieu", "co tuc bang co phieu", "co tuc co phieu", "chuyen doi thanh co phan", "co phieu de tra co tuc"]
            match_ctcp = any(kw in pre_title or kw in pre_ly_do for kw in ctcp_keywords)
            if not match_ctcp:
                cond1 = ("co phieu" in pre_title and "de tang von" in pre_title) or ("co tuc" in pre_title and "co phieu" in pre_title)
                cond2 = ("co phieu" in pre_ly_do and "de tang von" in pre_ly_do) or ("co tuc" in pre_ly_do and "co phieu" in pre_ly_do)
                match_ctcp = cond1 or cond2
            if match_ctcp:
                nhom_quyen = "Cổ tức cổ phiếu / Cổ phiếu thưởng"
                
        # e. Cổ tức tiền
        if nhom_quyen is None:
            ctt_keywords = ["co tuc bang tien", "co tuc tien", "thanh toan co tuc", "tam ung co tuc", "lai trai phieu", "chi co tuc", "mua lai trai phieu", "tra goc", "thanh toan goc", "tra lai", "thanh toan lai"]
            match_ctt = any(kw in pre_title or kw in pre_ly_do for kw in ctt_keywords)
            if not match_ctt:
                cond1 = ("co tuc" in pre_title and "bang tien" in pre_title) or ("mua lai" in pre_title and "trai phieu" in pre_title) or ("mua lai" in pre_title and "truoc han" in pre_title)
                cond2 = ("co tuc" in pre_ly_do and "bang tien" in pre_ly_do) or ("mua lai" in pre_ly_do and "trai phieu" in pre_ly_do) or ("mua lai" in pre_ly_do and "truoc han" in pre_ly_do)
                match_ctt = cond1 or cond2
            if match_ctt:
                nhom_quyen = "Cổ tức tiền"
                
        # f. Quyền biểu quyết
        if nhom_quyen is None:
            qbq_keywords = ["y kien co dong", "dai hoi co dong", "dai hoi dong co dong", "dai hoi nha dau tu", "bang van ban", "lay y kien", "thong qua phuong an", "dai hoi", "dong co dong", "quyen de cu"]
            if any(kw in pre_title or kw in pre_ly_do for kw in qbq_keywords):
                nhom_quyen = "Quyền biểu quyết"
                
        # g. Quyền mua
        if nhom_quyen is None:
            if "quyen mua" in pre_title or "quyen mua" in pre_ly_do:
                nhom_quyen = "Quyền mua"
                
        # h. Hoán đổi chuyển đổi
        if nhom_quyen is None:
            hd_keywords = ["hoan doi co phieu", "hoan doi trai phieu", "chuyen doi trai phieu", "chuyen doi co phieu", "chuyen quyen"]
            if any(kw in pre_title or kw in pre_ly_do for kw in hd_keywords):
                nhom_quyen = "Hoán đổi chuyển đổi"
                
        # i. Chứng quyền
        if nhom_quyen is None:
            cq_keywords = ["dang ky", "giay chung nhan", "khai bao", "dieu chinh", "thuc hien", "do dao han", "huy"]
            if "chung quyen" in pre_title and any(kw in pre_title for kw in cq_keywords):
                nhom_quyen = "Chứng quyền"
            elif "chung quyen" in pre_ly_do and any(kw in pre_ly_do for kw in cq_keywords):
                nhom_quyen = "Chứng quyền"

        # LoaiQuyen
        loai_quyen = None
        if nhom_quyen == "Cổ tức cổ phiếu / Cổ phiếu thưởng":
            lh_keywords = ["phat hanh co phieu", "nhan co phieu", "nhan them co phieu", "co phieu thuong", "thuong co phieu"]
            is_cpt = any(kw in pre_title or kw in pre_ly_do for kw in lh_keywords)
            if not is_cpt:
                is_cpt = ("co phieu" in pre_title and "tang von" in pre_title) or ("co phieu" in pre_ly_do and "tang von" in pre_ly_do)
            
            if is_cpt:
                loai_quyen = "Cổ phiếu thưởng"
            else:
                ctcp_lh_keywords = ["co tuc bang co phieu", "co tuc co phieu", "chuyen doi thanh co phan", "co phieu de tra co tuc"]
                is_ctcp = any(kw in pre_title or kw in pre_ly_do for kw in ctcp_lh_keywords)
                if not is_ctcp:
                    is_ctcp = ("co tuc" in pre_title and "co phieu" in pre_title) or ("co tuc" in pre_ly_do and "co phieu" in pre_ly_do)
                if is_ctcp:
                    loai_quyen = "Cổ tức cổ phiếu"
        elif nhom_quyen == "Cổ tức tiền":
            if "trai phieu" in pre_title or "trai phieu" in pre_ly_do:
                loai_quyen = "Trái phiếu"
            else:
                loai_quyen = "Cổ phiếu"
        elif nhom_quyen == "Đăng ký Lưu ký":
            if "dang ky" in pre_title:
                loai_quyen = "Đăng ký"
            elif "luu ky" in pre_title:
                loai_quyen = "Lưu ký"
        elif nhom_quyen == "Tin huỷ":
            if "huy dang ky chung khoan" in pre_title or "huy dang ky chung khoan" in pre_text:
                loai_quyen = "Hủy đăng ký chứng khoán"
            elif "huy dang ky chung quyen" in pre_title or "huy dang ky chung quyen" in pre_text:
                loai_quyen = "Hủy đăng ký chứng quyền"
            elif "huy dang ky trai phieu" in pre_title or "huy dang ky trai phieu" in pre_text:
                loai_quyen = "Hủy đăng ký trái phiếu"
            elif "huy dot chot danh sach thuc hien chung quyen" in pre_title or "huy dot chot danh sach thuc hien chung quyen" in pre_text:
                loai_quyen = "Hủy đợt chốt danh sách thực hiện chứng quyền"
            elif "huy danh sach nguoi so huu chung khoan" in pre_title or "huy danh sach nguoi so huu chung khoan" in pre_text:
                loai_quyen = "Hủy danh sách người sở hữu chứng khoán"
            elif "huy thong bao ngay dang ky cuoi cung" in pre_title or "huy thong bao ngay dang ky cuoi cung" in pre_text:
                loai_quyen = "Hủy thông báo ngày đăng ký cuối cùng"

        # MaISIN
        isin_val = record.get('mã_isin')
        if not isin_val:
            isin_pos = pre_text.find('isin')
            if isin_pos != -1:
                sub_text = pre_text[isin_pos:]
                dash_idx = sub_text.find('-')
                nl_idx = sub_text.find('\n')
                indices = [idx for idx in [dash_idx, nl_idx] if idx != -1]
                end_pos = min(indices) if indices else len(sub_text)
                
                target_sub = sub_text[:end_pos]
                isin_match = re.search(r'\b([a-z][a-z0-9]{10}\d)\b', target_sub)
                if isin_match:
                    isin_val = isin_match.group(1).upper()

        # MaTrongNuoc
        ma_trong_nuoc = None
        mn_kw_pos, mn_kw = find_earliest_keyword(pre_text, ["ma quyen mua", "ma trong nuoc"])
        if mn_kw_pos != -1:
            sub_text = pre_text[mn_kw_pos:]
            dash_idx = sub_text.find('-')
            nl_idx = sub_text.find('\n')
            indices = [idx for idx in [dash_idx, nl_idx] if idx != -1]
            end_pos = min(indices) if indices else len(sub_text)
            target_sub = sub_text[:end_pos]
            mn_match = re.search(r'\b([a-z][a-z0-9]{7}\d)\b', target_sub)
            if mn_match:
                ma_trong_nuoc = mn_match.group(1).upper()

        # NgayChot
        ngay_chot_val = record.get('ngày_đăng_KY_cuối') or record.get('ngày_đăng_ký_cuối') or record.get('ngày_đăng_ky_cuối')
        if not ngay_chot_val:
            nc_kw_pos, nc_kw = find_earliest_keyword(pre_text, ["ngay dang ky cuoi", "thoi gian dang ky cuoi", "ngay chot", "thoi gian chot"])
            if nc_kw_pos != -1:
                segment = pre_text[nc_kw_pos:nc_kw_pos + 200]
                date_found = extract_earliest_date(segment)
                if date_found:
                    ngay_chot_val = date_found

        # NgayGDKHQ
        ngay_gdkhq_val = None
        if ngay_chot_val:
            try:
                chot_date = datetime.strptime(ngay_chot_val, '%d/%m/%Y').date()
                gdkhq_date = chot_date - timedelta(days=1)
                ngay_gdkhq_val = gdkhq_date.strftime('%d/%m/%Y')
            except Exception as e:
                logger.error(f"Error calculating NgayGDKHQ: {e}")

        # NgayThucHien & UocTinhNgayThucHien
        ngay_thuc_hien = None
        uoc_tinh_ngay_thuc_hien = None
        if nhom_quyen == "Quyền biểu quyết":
            found = False
            nth_kw_pos, nth_kw = find_earliest_keyword(pre_text, ["thoi gian thuc hien", "ngay thuc hien", "thoi gian hien", "ngay hien", "thoi gian thuc", "ngay thuc"])
            if nth_kw_pos != -1:
                sub_text = pre_text[nth_kw_pos:]
                dash_idx = sub_text.find('-')
                plus_idx = sub_text.find('+')
                nl_idx = sub_text.find('\n')
                indices = [idx for idx in [dash_idx, plus_idx, nl_idx] if idx != -1]
                end_pos = min(indices) if indices else len(sub_text)
                
                target_sub = sub_text[:end_pos]
                dates_list = extract_all_dates_in_segment(target_sub)
                if dates_list:
                    ngay_thuc_hien = dates_list[-1]
                    found = True
            
            if not found:
                if ngay_chot_val:
                    try:
                        chot_date = datetime.strptime(ngay_chot_val, '%d/%m/%Y').date()
                        import calendar
                        y = chot_date.year
                        m = chot_date.month + 1
                        if m > 12:
                            m = 1
                            y += 1
                        last_day = calendar.monthrange(y, m)[1]
                        ngay_thuc_hien = f"{last_day:02d}/{m:02d}/{y}"
                        uoc_tinh_ngay_thuc_hien = 1
                    except Exception as e:
                        logger.error(f"Error calculating fallback NgayThucHien: {e}")

        # NgayThanhToan
        ngay_thanh_toan = None
        if nhom_quyen not in ["Quyền biểu quyết", "Cổ tức cổ phiếu / Cổ phiếu thưởng"] and nhom_quyen is not None:
            ntt_kw_pos, ntt_kw = find_earliest_keyword(pre_text, ["ngay thuc hien", "ngay thanh toan", "thoi gian thuc hien", "thoi gian thanh toan"])
            if ntt_kw_pos != -1:
                sub_text = pre_text[ntt_kw_pos:]
                dash_idx = sub_text.find('-')
                nl_idx = sub_text.find('\n')
                indices = [idx for idx in [dash_idx, nl_idx] if idx != -1]
                end_pos = min(indices) if indices else len(sub_text)
                
                target_sub = sub_text[:end_pos]
                dates_list = extract_all_dates_in_segment(target_sub)
                if dates_list:
                    ngay_thanh_toan = dates_list[0]

        # CNQuyenMuaTuNgay & CNQuyenMuaDenNgay
        cn_quyen_mua_tu_ngay = None
        cn_quyen_mua_den_ngay = None
        if nhom_quyen == "Quyền mua":
            cn_kw_pos, cn_kw = find_earliest_keyword(pre_text, ["thoi gian chuyen nhuong", "ngay chuyen nhuong", "han chuyen nhuong"])
            if cn_kw_pos != -1:
                sub_text = pre_text[cn_kw_pos:]
                dash_idx = sub_text.find('-')
                plus_idx = sub_text.find('+')
                nl_idx = sub_text.find('\n')
                indices = [idx for idx in [dash_idx, plus_idx, nl_idx] if idx != -1]
                end_pos = min(indices) if indices else len(sub_text)
                
                target_sub = sub_text[:end_pos]
                dates_list = extract_all_dates_in_segment(target_sub)
                if dates_list:
                    cn_quyen_mua_tu_ngay = dates_list[0]
                    cn_quyen_mua_den_ngay = dates_list[-1]

        # DKQuyenMuaTuNgay & DKQuyenMuaDenNgay
        dk_quyen_mua_tu_ngay = None
        dk_quyen_mua_den_ngay = None
        if nhom_quyen == "Quyền mua":
            dk_kw_pos, dk_kw = find_earliest_keyword(pre_text, ["thoi gian dang ky", "ngay dang ky", "han dang ky", "thoi gian dat", "ngay dat", "han dat", "thoi gian nop tien", "ngay nop tien", "han nop tien"])
            if dk_kw_pos != -1:
                sub_text = pre_text[dk_kw_pos:]
                dash_idx = sub_text.find('-')
                plus_idx = sub_text.find('+')
                nl_idx = sub_text.find('\n')
                indices = [idx for idx in [dash_idx, plus_idx, nl_idx] if idx != -1]
                end_pos = min(indices) if indices else len(sub_text)
                
                target_sub = sub_text[:end_pos]
                dates_list = extract_all_dates_in_segment(target_sub)
                if dates_list:
                    dk_quyen_mua_tu_ngay = dates_list[0]
                    dk_quyen_mua_den_ngay = dates_list[-1]

        # DonViHuongQuyen & GiaTriHuongQuyen
        don_vi_huong_quyen = None
        gia_tri_huong_quyen = None
        if nhom_quyen == "Quyền biểu quyết":
            don_vi_huong_quyen = 1
            gia_tri_huong_quyen = 1
        elif nhom_quyen is not None:
            scope_text = ""
            for line in text_content.split('\n'):
                pre_line = remove_accents_and_lower(line)
                pos, kw = find_earliest_keyword(pre_line, ["ty le thuc hien", "ty le thanh toan", "ti le thuc hien", "ti le thanh toan"])
                if pos != -1:
                    sub_line = pre_line[pos + len(kw):]
                    dash_idx = sub_line.find('-')
                    if dash_idx != -1:
                        scope_text = sub_line[:dash_idx]
                    else:
                        scope_text = sub_line
                    break
                    
            if scope_text:
                scope_text = re.sub(r'\s+', ' ', scope_text).strip()
                num_pattern = r'(\d+(?:\.\d+)*(?:\,\d+)?)'
                
                verb_pattern = r'(?:\s*(?:se|duoc|nhan|them))*\s*'
                
                pattern_a = re.compile(
                    num_pattern + r'(?:\s*\([^)]+\))?\s*(?:co phieu|trai phieu|chung chi quy)' + verb_pattern + num_pattern,
                    re.IGNORECASE
                )
                
                pattern_b = re.compile(
                    num_pattern + r'(?:\s*\([^)]+\))?\s*(?:co phieu|trai phieu|chung chi quy)\s*[-–—]\s*' + num_pattern + r'\s*quyen bieu quyet',
                    re.IGNORECASE
                )
                
                pattern_c = re.compile(
                    num_pattern + r'(?:\s*\([^)]+\))?\s*[:/]\s*' + num_pattern,
                    re.IGNORECASE
                )
                
                match = pattern_a.search(scope_text) or pattern_b.search(scope_text) or pattern_c.search(scope_text)
                
                if match:
                    x_str = match.group(1)
                    y_str = match.group(2)
                    
                    try:
                        X = int(x_str.replace('.', ''))
                        if ',' in y_str:
                            y_clean = y_str.replace('.', '')
                            parts = y_clean.split(',')
                            integer_part = parts[0]
                            decimal_part = parts[1]
                            
                            factor = 10 ** len(decimal_part)
                            Y_val = int(integer_part) * factor + int(decimal_part)
                            X_val = X * factor
                            
                            don_vi_huong_quyen = X_val
                            gia_tri_huong_quyen = Y_val
                        else:
                            y_clean = y_str.replace('.', '')
                            don_vi_huong_quyen = X
                            gia_tri_huong_quyen = int(y_clean)
                    except Exception as e:
                        logger.error(f"Error parsing X and Y: {e}")

        # TyLeMenhGia
        ty_le_menh_gia = None
        if nhom_quyen == "Cổ tức tiền":
            if loai_quyen == "Cổ phiếu":
                scope_text = ""
                tl_kw_pos, tl_kw = find_earliest_keyword(pre_text, ["ty le thuc hien", "ty le thanh toan", "ti le thuc hien", "ti le thanh toan"])
                if tl_kw_pos != -1:
                    sub_text = pre_text[tl_kw_pos + len(tl_kw):]
                    dash_idx = sub_text.find('-')
                    plus_idx = sub_text.find('+')
                    nl_idx = sub_text.find('\n')
                    indices = [idx for idx in [dash_idx, plus_idx, nl_idx] if idx != -1]
                    end_pos = min(indices) if indices else len(sub_text)
                    scope_text = sub_text[:end_pos]
                        
                if scope_text:
                    percent_match = re.search(r'(\d+(?:\.\d+)*(?:\,\d+)?)\s*%', scope_text)
                    if percent_match:
                        num_str = percent_match.group(1)
                        ty_le_menh_gia = parse_vietnamese_float(num_str)
                        
                if ty_le_menh_gia is None and gia_tri_huong_quyen is not None and don_vi_huong_quyen is not None:
                    import math
                    try:
                        k = math.log10(don_vi_huong_quyen)
                        ty_le_menh_gia = float(gia_tri_huong_quyen) / (10 ** (2 + k))
                    except Exception as e:
                        logger.error(f"Error calculating TyLeMenhGia for stock: {e}")
            elif loai_quyen == "Trái phiếu":
                menh_gia_val = None
                menh_gia_str = record.get('mệnh_giá')
                if menh_gia_str:
                    cleaned_mg = re.sub(r'[^\d]', '', str(menh_gia_str))
                    if cleaned_mg:
                        try:
                            menh_gia_val = float(cleaned_mg)
                        except ValueError:
                            pass
                if gia_tri_huong_quyen is not None and menh_gia_val:
                    dv_huong_quyen_val = float(don_vi_huong_quyen) if don_vi_huong_quyen else 1.0
                    ty_le_menh_gia = ((float(gia_tri_huong_quyen) / dv_huong_quyen_val) / menh_gia_val) * 100

        # GiaPhatHanh
        gia_phat_hanh = None
        gph_pos = pre_text.find("gia phat hanh")
        if gph_pos != -1:
            sub_text = pre_text[gph_pos:]
            dash_idx = sub_text.find('-')
            plus_idx = sub_text.find('+')
            nl_idx = sub_text.find('\n')
            indices = [idx for idx in [dash_idx, plus_idx, nl_idx] if idx != -1]
            end_pos = min(indices) if indices else len(sub_text)
            
            target_sub = sub_text[:end_pos]
            num_match = re.search(r'(\d+(?:[.,]\d+)*)', target_sub)
            if num_match:
                val_clean = num_match.group(1).replace('.', '').split(',')[0]
                try:
                    gia_phat_hanh = int(val_clean)
                except ValueError:
                    pass

        # TieuDe
        tieu_de = title
        if code and tieu_de.startswith(code):
            tieu_de = tieu_de[len(code):].lstrip(':').strip()
        else:
            col_idx = tieu_de.find(':')
            if col_idx != -1:
                tieu_de = tieu_de[col_idx+1:].strip()

        # NoiDung
        noi_dung = None
        title_content = tieu_de

        if nhom_quyen in ["Hoán đổi chuyển đổi", "Khai báo chứng quyền", "Chứng quyền", "Đăng ký Lưu ký", "Thay đổi"]:
            noi_dung = None
        elif nhom_quyen == "Tin huỷ":
            if "huy dang ky chung khoan" in pre_title or "huy dang ky chung khoan" in pre_text:
                noi_dung = "Hủy đăng ký chứng khoán"
            elif "huy dang ky chung quyen" in pre_title or "huy dang ky chung quyen" in pre_text:
                noi_dung = "Hủy đăng ký chứng quyền"
            elif "huy dang ky trai phieu" in pre_title or "huy dang ky trai phieu" in pre_text:
                noi_dung = "Hủy đăng ký trái phiếu"
            elif "huy dot chot danh sach thuc hien chung quyen" in pre_title or "huy dot chot danh sach thuc hien chung quyen" in pre_text:
                noi_dung = "Hủy đợt chốt danh sách thực hiện chứng quyền"
            elif "huy danh sach nguoi so huu chung khoan" in pre_title or "huy danh sach nguoi so huu chung khoan" in pre_text:
                noi_dung = "Hủy danh sách người sở hữu chứng khoán"
            elif "huy thong bao ngay dang ky cuoi cung" in pre_title or "huy thong bao ngay dang ky cuoi cung" in pre_text:
                noi_dung = "Hủy thông báo ngày đăng ký cuối cùng"
        elif nhom_quyen == "Quyền biểu quyết":
            noi_dung = title_content
        else:
            orig_text = text_content
            parts_list = [title_content]
            
            ty_le_info = extract_original_line_by_keyword(orig_text, ["ty le thuc hien", "ty le thanh toan"])
            if ty_le_info:
                ty_le_info = re.sub(r'^[-\–\—+•*]+\s*', '', ty_le_info)
                parts_list.append(ty_le_info)
                
            ngay_tt_info = extract_original_line_by_keyword(orig_text, ["ngay thanh toan", "ngay thuc hien", "thoi gian thuc hien", "thoi gian thanh toan"])
            if ngay_tt_info:
                ngay_tt_info = re.sub(r'^[-\–\—+•*]+\s*', '', ngay_tt_info)
                parts_list.append(ngay_tt_info)
                
            gia_ph_info = extract_original_line_by_keyword(orig_text, ["gia phat hanh"])
            if gia_ph_info:
                gia_ph_info = re.sub(r'^[-\–\—+•*]+\s*', '', gia_ph_info)
                parts_list.append(gia_ph_info)
                
            noi_dung = " - ".join(parts_list)
            if noi_dung:
                noi_dung = re.sub(r'-\s*-', '-', noi_dung)

        # is_completed
        is_completed = 0
        if nhom_quyen is not None and code and ngay_chot_val is not None:
            criteria_met = True
            if nhom_quyen == "Quyền mua":
                quyen_mua_fields = [
                    cn_quyen_mua_tu_ngay, cn_quyen_mua_den_ngay,
                    dk_quyen_mua_tu_ngay, dk_quyen_mua_den_ngay,
                    don_vi_huong_quyen, gia_tri_huong_quyen,
                    gia_phat_hanh, isin_val, ma_trong_nuoc
                ]
                if any(f is None for f in quyen_mua_fields):
                    criteria_met = False
            elif nhom_quyen == "Cổ tức tiền":
                if ty_le_menh_gia is None:
                    criteria_met = False
            elif nhom_quyen == "Cổ tức cổ phiếu / Cổ phiếu thưởng":
                if don_vi_huong_quyen is None or gia_tri_huong_quyen is None:
                    criteria_met = False
            elif nhom_quyen == "Quyền biểu quyết":
                if ngay_thuc_hien is None:
                    criteria_met = False
                    
            if criteria_met:
                is_completed = 1

        # is_special
        is_special = 0
        if nhom_quyen in ["Hoán đổi chuyển đổi", "Khai báo chứng quyền", "Chứng quyền", "Đăng ký Lưu ký", "Tin huỷ", "Thay đổi"] or ';' in pre_title:
            is_special = 1

        # Trả về bản ghi mới theo đúng 25 cột của rules_2finalize.md + các metadata trường dự phòng
        res = dict(record)
        
        # Xử lý nâng cấp ngày giờ (published_at, collected_at) từ các định dạng cũ/mới
        pub_at = res.pop('published_at', None) or res.pop('published_date', None) or res.pop('date', None)
        if pub_at and ' ' not in str(pub_at).strip():
            pub_at = f"{str(pub_at).strip()} 00:00:00"
            
        coll_at = res.pop('collected_at', None) or res.pop('collected_date', None)
        if coll_at and ' ' not in str(coll_at).strip():
            coll_at = f"{str(coll_at).strip()} 00:00:00"
            
        # Loại bỏ triệt để các trường ngày tháng kiểu cũ khác
        res.pop('published_date', None)
        res.pop('collected_date', None)
        res.pop('date', None)

        res.update({
            'published_at': pub_at,
            'collected_at': coll_at,
            'url': record.get('url'),
            'text_content': record.get('text_content'),
            'MaChungKhoan': ma_ck,
            'TieuDe': tieu_de,
            'NhomQuyen': nhom_quyen,
            'LoaiQuyen': loai_quyen,
            'MaISIN': isin_val,
            'MaTrongNuoc': ma_trong_nuoc,
            'NgayChot': ngay_chot_val,
            'NgayGDKHQ': ngay_gdkhq_val,
            'NgayThucHien': ngay_thuc_hien,
            'UocTinhNgayThucHien': uoc_tinh_ngay_thuc_hien,
            'NgayThanhToan': ngay_thanh_toan,
            'CNQuyenMuaTuNgay': cn_quyen_mua_tu_ngay,
            'CNQuyenMuaDenNgay': cn_quyen_mua_den_ngay,
            'DKQuyenMuaTuNgay': dk_quyen_mua_tu_ngay,
            'DKQuyenMuaDenNgay': dk_quyen_mua_den_ngay,
            'DonViHuongQuyen': don_vi_huong_quyen,
            'GiaTriHuongQuyen': gia_tri_huong_quyen,
            'TyLeMenhGia': ty_le_menh_gia,
            'GiaPhatHanh': gia_phat_hanh,
            'NoiDung': noi_dung,
            'is_completed': is_completed,
            'is_special': is_special
        })
        return res
    def parse_date(self, date_string):
        """Parse ngày từ string 'dd/mm/yyyy' thành datetime.date"""
        try:
            return datetime.strptime(date_string, '%d/%m/%Y').date()
        except:
            return None

    def parse_datetime(self, datetime_string):
        """Parse ngày/giờ từ string 'dd/mm/yyyy HH:MM:SS' hoặc 'dd/mm/yyyy' thành datetime"""
        if not datetime_string:
            return None
        datetime_string = str(datetime_string).strip()
        try:
            return datetime.strptime(datetime_string, '%d/%m/%Y %H:%M:%S')
        except ValueError:
            try:
                return datetime.strptime(datetime_string, '%d/%m/%Y')
            except ValueError:
                return None

    def generate_record_id(self, record, split_idx=None):
        """
        Generate stable, unique _record_id for a record.
        Same record will have same ID across runs (safe for daily updates).

        Strategy:
        - Combine ticker code and a hash suffix of (url + title) to ensure uniqueness
          across events and split records, even after Excel serialization.

        Args:
            record: The record dict with 'code', 'url', 'title', 'MaChungKhoan', etc.
            split_idx: Optional index, kept for backward compatibility

        Returns:
            Stable record ID string (e.g., "DHC_a1b2c3d4")
        """
        code = str(record.get('code') or record.get('MaChungKhoan') or '').strip()
        url = str(record.get('url') or '').strip()
        title = str(record.get('title') or record.get('TieuDe') or record.get('NoiDung') or '').strip()

        # Build stable content to hash
        hash_content = f"{url}_{title}"
        hash_suffix = hashlib.md5(hash_content.encode('utf-8')).hexdigest()[:8]

        if code:
            record_id = f"{code}_{hash_suffix}"
        else:
            record_id = f"rec_{hash_suffix}"

        # If split_idx is passed (kept for compatibility), append it
        if split_idx is not None:
            record_id = f"{record_id}_{split_idx}"

        return record_id

    def get_vptoken(self):
        """Extract VPToken từ <meta name='__VPToken'> trên trang list"""
        try:
            response = self.session.get(self.news_url, headers=self.headers, timeout=10)
            response.encoding = 'utf-8'
            soup = BeautifulSoup(response.content, 'html.parser')
            meta = soup.find('meta', {'name': '__VPToken'})
            if meta and meta.get('content'):
                self.vptoken = meta.get('content')
                logger.info(f"✓ Got VPToken: {self.vptoken[:20]}...")
                return self.vptoken
            else:
                logger.error("✗ VPToken not found in meta tag")
                return None
        except Exception as e:
            logger.error(f"✗ Error getting VPToken: {str(e)}")
            return None

    def extract_field_from_text(self, text, field_label, max_length=500):
        """
        Extract field value từ text content dựa trên label
        Hỗ trợ multi-line content và bullet points

        Ví dụ:
        "Địa điểm thực hiện: ..." => lấy text sau "Địa điểm thực hiện:"
        "Địa điểm thực hiện:\n+ Đối với..." => lấy tất cả bullet points
        """
        pattern = f"{field_label}[:\\s]+([^\\n]+(?:\\n\\s*[+\\-•]\\s*[^\\n]+)*)"
        match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)

        if match:
            extracted = match.group(1).strip()
            # Nếu quá dài, chỉ lấy phần đầu
            if len(extracted) > max_length:
                extracted = extracted[:max_length] + "..."
            return extracted if extracted else None
        return None

    def extract_field_bullets(self, text, field_label):
        """
        Extract field value và split thành list nếu có bullet points

        Ví dụ:
        "Tỷ lệ thực hiện:\n+ Quyền 1\n+ Quyền 2" => ['Quyền 1', 'Quyền 2']
        "Tỷ lệ thực hiện: Quyền duy nhất" => ['Quyền duy nhất']

        Returns:
            List of extracted values, or None if nothing found
        """
        pattern = f"{field_label}[:\\s]+([^\\n]+(?:\\n\\s*[+\\-•]\\s*[^\\n]+)*)"
        match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)

        if not match:
            return None

        extracted = match.group(1).strip()

        # Tìm tất cả bullet points
        bullet_pattern = r'[+\-•]\s*([^\n]+)'
        bullets = re.findall(bullet_pattern, extracted)

        if bullets:
            # Nếu tìm được bullet points, trả về list
            return [b.strip() for b in bullets if b.strip()]
        else:
            # Nếu không có bullet points, trả về single item
            return [extracted] if extracted else None

    def contains_keyword(self, text, keywords):
        """
        Check if text contains any of the keywords (case-insensitive)

        Args:
            text: Text content to search in
            keywords: List of keywords to search for

        Returns:
            True if any keyword is found, False otherwise
        """
        text_lower = text.lower()
        for keyword in keywords:
            if keyword.lower() in text_lower:
                return True
        return False

    def extract_quyền_values(self, text, value_keywords_map):
        """
        Extract specific quyền values from text using keyword mapping

        Args:
            text: Text content to search in
            value_keywords_map: Dict with {value: [keywords]} structure
                e.g., {'Quyền đại hội cổ đông thường niên': ['đại hội thường niên', 'ĐHĐCĐ thường']}

        Returns:
            Comma-separated string of found values, or None if none found
        """
        text_lower = text.lower()
        found_values = []

        for value, keywords in value_keywords_map.items():
            for keyword in keywords:
                if keyword.lower() in text_lower:
                    if value not in found_values:
                        found_values.append(value)
                    break

        return ', '.join(found_values) if found_values else None

    def extract_detail_from_article(self, url):
        """
        Mở URL tin tức và extract chi tiết thông tin từ HTML structure:
        <div class="col-md-4 item-info">Label:</div>
        <div class="col-md-8 item-info item-info-main">Value</div>

        Trả về tuple (info_dict, extracted_code, actual_update_date) nếu tìm được mã từ chi tiết
        actual_update_date là ngày "Cập nhật ngày" từ bài viết (chính xác hơn ngày listing)
        """
        try:
            # Retry logic để đảm bảo page load đầy đủ
            max_retries = 3
            response = None
            for attempt in range(max_retries):
                response = self.session.get(url, headers=self.headers, timeout=10)
                response.encoding = 'utf-8'

                if response.status_code == 200:
                    break

                if attempt < max_retries - 1:
                    time.sleep(0.2)  # Wait before retry

            if response is None or response.status_code != 200:
                return None, None, None

            soup = BeautifulSoup(response.content, 'html.parser')

            # Extract text content - thử nhiều selector
            main = soup.find('main') or soup.find('article') or soup.find('div', class_='main-content') or soup.find('div', class_='content')

            if not main:
                # Fallback: lấy toàn bộ body content nếu không tìm được main/article
                body = soup.find('body')
                if not body:
                    return None, None, None
                text_content = body.get_text()
            else:
                # Lấy chỉ nội dung chính của article, không lấy phần "Tin cùng tổ chức" hoặc bảng thống kê phía dưới
                text_content = main.get_text()

            # Tìm điểm kết thúc của nội dung chính (phần "Tin cùng tổ chức" hoặc các phần khác)
            cutoff_markers = [
                'tin cùng tổ chức',
                'mã ck hủy đăng ký',
                'mã ck chuyển sàn',
                'thành viên đã thu hồi'
            ]

            # Cắt text tại marker đầu tiên tìm thấy
            min_cutoff = len(text_content)
            for marker in cutoff_markers:
                pos = text_content.lower().find(marker)
                if pos > 0:
                    min_cutoff = min(min_cutoff, pos)

            # Nếu tìm thấy marker, chỉ lấy text trước đó
            if min_cutoff < len(text_content):
                text_content = text_content[:min_cutoff]

            # Initialize info với tất cả fields cần thiết
            info = {
                'title': None,
                'tên_tổ_chức_đăng_ký': None,
                'tên_chứng_khoán': None,
                'mã_isin': None,
                'nơi_giao_dịch': None,
                'loại_chứng_khoán': None,
                'mệnh_giá': None,
                'ngày_đăng_ký_cuối': None,
                'lý_do_mục_đích': None,
                'tỷ_lệ_thực_hiện': None,
                'thời_gian_thực_hiện': None,
                'địa_điểm_thực_hiện': None,
                # 9 cột quyền mới
                'quyền_họp_đại_hội_cổ_đông': None,
                'quyền_cổ_tức_tiền': None,
                'quyền_cổ_tức_cổ_phiếu': None,
                'quyền_mua': None,
                'quyền_hoán_đổi_chuyển_đổi': None,
                'chứng_quyền': None,
                'chấp_thuận_đăng_ký': None,
                'tin_húy': None,
                'thay_đổi': None,
                # Nội dung chính của bài viết
                'text_content': None
            }

            extracted_code = None

            # Find all label-value pairs trong HTML structure chuẩn
            label_divs = soup.find_all('div', class_='col-md-4')

            # Nếu không tìm được col-md-4, thử extract từ text_content
            if not label_divs:
                # Extract code từ text content (pattern: CODE: ...)
                code_match = re.search(r'^([A-Z0-9]{6,})\s*:', text_content.strip(), re.MULTILINE)
                if code_match:
                    extracted_code = code_match.group(1)

                # Extract fields từ pattern "Label: Value"
                # Tên chứng khoán
                name_match = re.search(r'Tên chứng khoán[:\s]+([^\n]+)', text_content, re.IGNORECASE)
                if name_match:
                    info['tên_chứng_khoán'] = name_match.group(1).strip()

                # Mã chứng khoán
                code_match2 = re.search(r'Mã chứng khoán[:\s]+([A-Z0-9]+)', text_content, re.IGNORECASE)
                if code_match2:
                    info['mã_chứng_khoán'] = code_match2.group(1).strip()
                    extracted_code = code_match2.group(1)

                # Mã ISIN
                isin_match = re.search(r'Mã ISIN[:\s]+([A-Z0-9]+)', text_content, re.IGNORECASE)
                if isin_match:
                    info['mã_isin'] = isin_match.group(1).strip()

                # Tên tổ chức đăng ký - thường ở sau "Tổng Công ty" hoặc "Công ty cổ phần"
                org_match = re.search(r'(?:Tổng Công ty|Công ty cổ phần|CTCP|Ngân hàng)[^\n]+(?:thông báo|khai báo)', text_content)
                if org_match:
                    org_text = org_match.group(0)
                    # Extract công ty name
                    org_name_match = re.search(r'(?:Tổng Công ty|Công ty cổ phần|CTCP|Ngân hàng)([^-]+)', org_text)
                    if org_name_match:
                        info['tên_tổ_chức_đăng_ký'] = ('Tổng Công ty' if 'Tổng' in org_text else 'Công ty cổ phần') + org_name_match.group(1).strip()

            for label_div in label_divs:
                # Get label text
                label = label_div.get_text(strip=True).lower()

                # Get value (next sibling with col-md-8)
                value_div = label_div.find_next('div', class_='col-md-8')
                if not value_div:
                    continue

                value = value_div.get_text(strip=True)

                # Map labels to info keys
                if 'tên tổ chức đăng ký' in label or 'tên tcđkck' in label or 'tcđkck' in label:
                    info['tên_tổ_chức_đăng_ký'] = value
                elif 'tên chứng khoán' in label:
                    info['tên_chứng_khoán'] = value
                elif 'mã chứng khoán' in label or 'mã ck' in label:
                    # Nếu có trường "Mã chứng khoán", lấy mã từ đây
                    extracted_code = value
                elif 'mã isin' in label:
                    info['mã_isin'] = value
                elif 'nơi giao dịch' in label:
                    info['nơi_giao_dịch'] = value
                elif 'loại chứng khoán' in label:
                    info['loại_chứng_khoán'] = value
                elif 'mệnh giá' in label:
                    info['mệnh_giá'] = value
                elif 'ngày đăng ký' in label and 'cuối' in label:
                    info['ngày_đăng_ký_cuối'] = value
                elif 'lý do' in label or 'mục đích' in label:
                    info['lý_do_mục_đích'] = value
                elif 'tỷ lệ' in label and 'thực hiện' in label:
                    info['tỷ_lệ_thực_hiện'] = value
                elif 'thời gian' in label and 'thực hiện' in label:
                    info['thời_gian_thực_hiện'] = value
                elif 'địa điểm' in label and 'thực hiện' in label:
                    info['địa_điểm_thực_hiện'] = value

            # Nếu không tìm được từ HTML structure, thử lấy từ text content
            # Vì thông tin này có thể ở dạng bullet points hoặc multi-line
            if not info['tỷ_lệ_thực_hiện']:
                info['tỷ_lệ_thực_hiện'] = self.extract_field_from_text(
                    text_content,
                    'Tỷ lệ thực hiện',
                    max_length=1000
                )

            if not info['thời_gian_thực_hiện']:
                info['thời_gian_thực_hiện'] = self.extract_field_from_text(
                    text_content,
                    'Thời gian thực hiện',
                    max_length=300
                )

            if not info['địa_điểm_thực_hiện']:
                info['địa_điểm_thực_hiện'] = self.extract_field_from_text(
                    text_content,
                    'Địa điểm thực hiện',
                    max_length=500
                )

            if not info['lý_do_mục_đích']:
                info['lý_do_mục_đích'] = self.extract_field_from_text(
                    text_content,
                    'Lý do|Mục đích',
                    max_length=300
                )

            if not info['mệnh_giá']:
                info['mệnh_giá'] = self.extract_field_from_text(
                    text_content,
                    'Mệnh giá',
                    max_length=100
                )

            # Nếu chưa có tên tổ chức đăng ký, thử extract từ text content
            # Tìm "Tên tổ chức đăng ký chứng khoán" hoặc "Tên TCĐKCK"
            if not info['tên_tổ_chức_đăng_ký']:
                # Pattern: "Tên tổ chức đăng ký chứng khoán:" hoặc "Tên TCĐKCK:" + value
                org_pattern = r'(?:Tên tổ chức đăng ký chứng khoán|Tên TCĐKCK)[:\s]+([^\n]+)'
                org_match = re.search(org_pattern, text_content, re.IGNORECASE)
                if org_match:
                    extracted_org = org_match.group(1).strip()
                    if extracted_org and extracted_org != '--':
                        info['tên_tổ_chức_đăng_ký'] = extracted_org
                        logger.debug(f"  ✓ Found org name from text: {extracted_org[:50]}")

            # Extract 9 new quyền fields từ text content + tiêu đề với các giá trị cụ thể
            # Sử dụng keyword mapping để tìm các giá trị cụ thể trong text
            # Include title trong search để bắt được những trang dạng danh sách
            title_tag = soup.find('title')
            search_text = text_content + (" " + title_tag.get_text() if title_tag else "")

            # 1. Quyền họp đại hội cổ đông
            dhdc_map = {
                'Quyền đại hội cổ đông thường niên': [
                    'đại hội đồng cổ đông thường niên',
                    'đại hội cổ đông thường niên',
                    'đại hội thường niên',
                    'đhđcđ thường niên',
                    'agm',
                    'annual general meeting'
                ],
                'Quyền lấy ý kiến cổ đông bằng văn bản': [
                    'lấy ý kiến cổ đông bằng văn bản',
                    'ý kiến bằng văn bản',
                    'written opinion'
                ],
                'Quyền đại hội cổ đông bất thường': [
                    'đại hội đồng cổ đông bất thường',
                    'đại hội cổ đông bất thường',
                    'đại hội bất thường',
                    'egm',
                    'extraordinary general meeting'
                ]
            }
            info['quyền_họp_đại_hội_cổ_đông'] = self.extract_quyền_values(search_text, dhdc_map)

            # 2. Quyền cổ tức tiền
            dividend_cash_map = {
                'Chi trả cổ tức bằng tiền': [
                    'chi trả cổ tức bằng tiền',
                    'cổ tức tiền',
                    'dividend cash'
                ],
                'Thanh toán lãi trái phiếu': [
                    'thanh toán lãi',
                    'lãi trái phiếu',
                    'bond interest',
                    'interest payment'
                ],
                'Thanh toán gốc, lãi': [
                    'thanh toán gốc',
                    'trả gốc',
                    'principal payment',
                    'maturity payment'
                ],
                'Mua lại trái phiếu trước hạn': [
                    'mua lại trái phiếu',
                    'early redemption',
                    'buyback'
                ]
            }
            info['quyền_cổ_tức_tiền'] = self.extract_quyền_values(search_text, dividend_cash_map)

            # 3. Quyền cổ_tức cổ phiếu
            dividend_share_map = {
                'Trả cổ tức bằng cổ phiếu': [
                    'trả cổ tức bằng cổ phiếu',
                    'cổ tức cổ phiếu',
                    'stock dividend'
                ],
                'Phát hành cổ phiếu': [
                    'phát hành cổ phiếu',
                    'share issuance',
                    'cổ phiếu thưởng',
                    'bonus shares'
                ]
            }
            info['quyền_cổ_tức_cổ_phiếu'] = self.extract_quyền_values(search_text, dividend_share_map)

            # 4. Quyền mua
            purchase_map = {
                'Thực hiện quyền mua Trái phiếu chuyển đổi': [
                    'quyền mua trái phiếu chuyển đổi',
                    'conversion bond purchase',
                    'convertible bond exercise'
                ],
                'Thực hiện quyền mua cổ phiếu': [
                    'quyền mua cổ phiếu',
                    'quyền mua',
                    'right issue',
                    'subscription right'
                ]
            }
            info['quyền_mua'] = self.extract_quyền_values(search_text, purchase_map)

            # 5. Quyền hoán đổi, chuyển đổi
            swap_map = {
                'Hoán đổi cổ phiếu': [
                    'hoán đổi cổ phiếu',
                    'swap shares',
                    'cổ phiếu hoán đổi'
                ],
                'Chuyển đổi trái phiếu': [
                    'chuyển đổi trái phiếu',
                    'convertible bond',
                    'bond conversion'
                ]
            }
            info['quyền_hoán_đổi_chuyển_đổi'] = self.extract_quyền_values(search_text, swap_map)

            # 6. Chứng quyền
            warrant_map = {
                'Có': [
                    'chứng quyền',
                    'warrant',
                    'call warrant',
                    'put warrant'
                ]
            }
            info['chứng_quyền'] = self.extract_quyền_values(search_text, warrant_map)

            # 7. Chấp thuận đăng ký
            approval_map = {
                'Đăng ký cổ phiếu, trái phiếu': [
                    'đăng ký cổ phiếu',
                    'đăng ký trái phiếu',
                    'registration approval',
                    'chấp thuận đăng ký'
                ]
            }
            info['chấp_thuận_đăng_ký'] = self.extract_quyền_values(search_text, approval_map)

            # 8. Tin hủy
            cancellation_map = {
                'Hủy ngày đăng ký cuối cùng': [
                    'hủy ngày đăng ký',
                    'cancel registration date'
                ],
                'Hủy danh sách người sở hữu chứng khoán': [
                    'hủy danh sách người sở hữu',
                    'hủy danh sách người sử hữu',
                    'hủy danh sách',
                    'cancel ownership list',
                    'cancel list'
                ],
                'Hủy đăng ký chứng khoán, trái phiếu': [
                    'hủy đăng ký',
                    'huỷ',
                    'delisting',
                    'deregistration'
                ]
            }
            info['tin_húy'] = self.extract_quyền_values(search_text, cancellation_map)

            # 9. Thay đổi
            change_map = {
                'Thay đổi thời gian thanh toán': [
                    'thay đổi thời gian thanh toán',
                    'thay đổi ngày thanh toán',
                    'payment date change'
                ],
                'Chuyển dữ liệu đăng ký (chuyển sàn)': [
                    'chuyển dữ liệu',
                    'chuyển sàn',
                    'data transfer',
                    'transfer between exchanges'
                ]
            }
            info['thay_đổi'] = self.extract_quyền_values(search_text, change_map)

            # Extract "Cập nhật ngày" từ bài viết (thay vì lấy từ listing page)
            actual_update_datetime = None
            # Pattern: "Cập nhật ngày DD/MM/YYYY" hoặc "Cập nhật ngày DD/MM/YYYY - HH:MM:SS"
            update_match = re.search(r'Cập nhật ngày\s+(\d{1,2}/\d{1,2}/\d{4})(?:\s*-\s*(\d{1,2}:\d{1,2}:\d{1,2}))?', text_content)
            if update_match:
                date_str = update_match.group(1)
                time_str = update_match.group(2) if update_match.group(2) else "00:00:00"
                datetime_str = f"{date_str} {time_str}"
                actual_update_datetime = self.parse_datetime(datetime_str)
                logger.debug(f"  ✓ Found actual update datetime: {datetime_str}")

            # Lấy title từ thẻ h3 class title-category hoặc addthis_inline_share_toolbox hoặc h2
            h3_title = soup.find('h3', class_='title-category')
            if h3_title:
                info['title'] = h3_title.get_text(strip=True)
            else:
                share_toolbox = soup.find('div', class_='addthis_inline_share_toolbox')
                if share_toolbox and share_toolbox.has_attr('data-title'):
                    info['title'] = share_toolbox['data-title'].strip()
                else:
                    h2_title = soup.find('h2')
                    if h2_title:
                        info['title'] = h2_title.get_text(strip=True)
                    else:
                        title_tag = soup.find('title')
                        if title_tag:
                            info['title'] = title_tag.get_text(strip=True)

            # Lưu nội dung chính của bài viết (dùng cho hiển thị full text)
            # Clean up: remove extra whitespace
            text_content = '\n'.join(line.strip() for line in text_content.split('\n') if line.strip())
            info['text_content'] = text_content if text_content else None

            return info, extracted_code, actual_update_datetime

        except Exception as e:
            logger.debug(f"  ! Error extracting detail: {str(e)[:50]}")
            return None, None, None

    def merge_records(self, new_records, old_records):
        """
        Merge new records and old records by URL with ticker-prefix deduplication rules.
        """
        # Group all records by URL
        url_to_new = {}
        for r in new_records:
            url = r.get('url')
            if url:
                url_to_new.setdefault(url, []).append(r)
                
        url_to_old = {}
        for r in old_records:
            url = r.get('url')
            if url:
                url_to_old.setdefault(url, []).append(r)
                
        merged = []
        
        # Helper to parse dates for sorting
        def parse_date_str(date_str):
            if not date_str:
                return datetime.min
            for fmt in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
                try:
                    return datetime.strptime(str(date_str).strip(), fmt)
                except ValueError:
                    continue
            return datetime.min
            
        # 1. Process URLs that are in new_records (replace old records with new records)
        for url, new_list in url_to_new.items():
            old_list = url_to_old.get(url, [])
            if old_list:
                for r_old in old_list:
                    # Find matching new record
                    target = None
                    if len(new_list) == 1:
                        target = new_list[0]
                    else:
                        # Try to match by _record_id
                        old_id = r_old.get('_record_id')
                        for r_new in new_list:
                            if r_new.get('_record_id') == old_id:
                                target = r_new
                                break
                        if not target:
                            # Try to match by title similarity
                            p_title = str(r_old.get('title') or '').strip().lower()
                            for r_new in new_list:
                                new_title = str(r_new.get('title') or '').strip().lower()
                                if new_title == p_title or new_title.split(':', 1)[-1].strip() == p_title.split(':', 1)[-1].strip():
                                    target = r_new
                                    break
                        if not target:
                            target = new_list[0]
                            
                    # Merge flags
                    for flag in ['status', 'confirmation_status', 'is_completed', 'is_special']:
                        if flag in r_old and r_old[flag] is not None and not pd.isna(r_old[flag]):
                            if flag == 'status' and r_old[flag] == 'pending':
                                continue
                            if flag == 'confirmation_status' and r_old[flag] == 'awaiting_review':
                                continue
                            target[flag] = r_old[flag]
                            
            merged.extend(new_list)
            
        # 2. Process URLs that are NOT in new_records (keep old records but deduplicate them)
        for url, old_list in url_to_old.items():
            if url in url_to_new:
                continue
                
            if len(old_list) == 1:
                merged.append(old_list[0])
                continue
                
            # We have multiple old records for this URL. Deduplicate them:
            ticker_prefixed = []
            non_prefixed = []
            for r in old_list:
                code = r.get('MaChungKhoan') or r.get('code') or ''
                title = r.get('title') or ''
                prefix = f"{code}:"
                if title.strip().startswith(prefix):
                    ticker_prefixed.append(r)
                else:
                    non_prefixed.append(r)
                    
            if ticker_prefixed:
                # Merge flags from non-prefixed to prefixed
                for r_non in non_prefixed:
                    # Find matching prefixed record
                    if len(ticker_prefixed) == 1:
                        target = ticker_prefixed[0]
                    else:
                        target = None
                        p_title = str(r_non.get('title') or '').strip().lower()
                        for r_pref in ticker_prefixed:
                            pref_title = str(r_pref.get('title') or '').split(':', 1)[-1].strip().lower()
                            if pref_title in p_title or p_title in pref_title:
                                target = r_pref
                                break
                        if not target:
                            target = ticker_prefixed[0]
                            
                    # Merge
                    for flag in ['status', 'confirmation_status', 'is_completed', 'is_special']:
                        if flag in r_non and r_non[flag] is not None and not pd.isna(r_non[flag]):
                            if flag == 'status' and r_non[flag] == 'pending':
                                continue
                            if flag == 'confirmation_status' and r_non[flag] == 'awaiting_review':
                                continue
                            target[flag] = r_non[flag]
                            
                merged.extend(ticker_prefixed)
            else:
                # Fallback: keep the one with the latest collected_at
                old_list_sorted = sorted(old_list, key=lambda x: parse_date_str(x.get('collected_at') or x.get('published_at')), reverse=True)
                merged.append(old_list_sorted[0])
                
        return merged

    def fetch_latest_news(self):
        """
        Crawl tất cả trang tin tức VSD từ ngày gần nhất:
        1. Lặp qua các trang (page=1, 2, 3, ...)
        2. Extract danh sách tin từ mỗi trang
        3. Dừng khi ngày tin giảm xuống (đã hết tin từ ngày gần nhất)
        4. Mở từng tin để lấy chi tiết
        5. Return danh sách mã + chi tiết
        """
        try:
            logger.info(f"🔍 VSD: Crawling tin tức thị trường cơ sở (multiple pages)...")
            filtered_news = []
            latest_date_found = datetime.now(VN_TZ).date()
            
            if self.mode == 'excel_urls':
                # Đọc file excel_urls
                excel_file = globals().get('EXCEL_URLS_FILE', 'url_2fetch_round2.xlsx')
                if not os.path.exists(excel_file):
                    # Thử tìm ở thư mục cha
                    parent_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), excel_file)
                    if os.path.exists(parent_path):
                        excel_file = parent_path
                    else:
                        raise FileNotFoundError(f"Không tìm thấy file Excel danh sách URL: {excel_file}")
                
                logger.info(f"  📊 Đang đọc danh sách URL từ file: {excel_file}")
                df_urls = pd.read_excel(excel_file)
                
                # Tìm cột chứa URL
                url_col = None
                for col in df_urls.columns:
                    if 'url' in str(col).lower():
                        url_col = col
                        break
                
                if url_col is None:
                    # Fallback dùng cột đầu tiên
                    url_col = df_urls.columns[0]
                
                raw_urls = df_urls[url_col].dropna().astype(str).tolist()
                
                # Làm sạch và loại bỏ trùng lặp
                urls_to_crawl = []
                seen_urls = set()
                for u in raw_urls:
                    u_clean = u.strip()
                    if u_clean.startswith('http') and u_clean not in seen_urls:
                        seen_urls.add(u_clean)
                        urls_to_crawl.append(u_clean)
                
                logger.info(f"  Found {len(urls_to_crawl)} unique URLs to process from Excel")
                
                # Tạo mock filtered_news
                for u in urls_to_crawl:
                    filtered_news.append({
                        'code': 'N/A',
                        'title': 'Excel Link',
                        'url': u,
                        'date': None,
                        'date_obj': None,
                        'source': 'VSD'
                    })
                
                page = 1
                
            else:
                # --- CHẾ ĐỘ CÀO PHÂN TRANG (Tương thích ngược) ---
                all_news = []
                page = 1
                latest_date_found = None
                max_pages = 25  # Tối đa 25 trang
                
                # Calculate cutoff date
                today = datetime.now(VN_TZ).date()
                if self.mode == 'date_range':
                    cutoff_date = self.date_from
                else:
                    cutoff_date = today - timedelta(days=self.keep_days)
                
                logger.info(f"  📅 Cutoff date (oldest date to keep): {cutoff_date}")
                
                while page <= max_pages:
                    logger.info(f"  📄 Crawling page {page}...")
                    try:
                        # Get VPToken từ trang đầu tiên nếu chưa có
                        if page == 1:
                            vptoken = self.get_vptoken()
                            if not vptoken:
                                logger.error("  ✗ Cannot get VPToken, stopping")
                                break

                        # Use AJAX POST with VPToken
                        ajax_headers = {
                            'User-Agent': 'Mozilla/5.0',
                            'Content-Type': 'application/json;charset=utf-8',
                            'X-Requested-With': 'XMLHttpRequest',
                            'Referer': self.news_url,
                            'Origin': self.base_url,
                            '__VPToken': vptoken
                        }
                        payload = {'SearchKey': 'TCPH', 'CurrentPage': page}

                        response = self.session.post(self.news_url, headers=ajax_headers, json=payload, timeout=10)
                        response.encoding = 'utf-8'

                        if response.status_code != 200:
                            logger.info(f"  ⚠ Page {page} failed (HTTP {response.status_code})")
                            break

                        soup = BeautifulSoup(response.content, 'html.parser')
                        news_items = soup.find_all('li')
                        page_news = []

                        logger.info(f"    📰 Total items on page: {len(news_items)}")
                        for item in news_items:
                            h3 = item.find('h3')
                            if not h3:
                                continue

                            link = h3.find('a')
                            if not link:
                                continue

                            title = link.get_text(strip=True)
                            url = link.get('href', '')

                            if not title or not url:
                                continue

                            # Chỉ lấy tin có mã CK - pattern: CODE: (where CODE is 2-10 chars)
                            if not re.search(r'[A-Z0-9]{2,10}:', title):
                                continue

                            # Extract mã CK from title
                            match = re.search(r'([A-Z0-9]{2,10}):', title)
                            if not match:
                                continue

                            code = match.group(1)

                            if not url.startswith('http'):
                                url = self.base_url + url

                            # Extract ngày
                            time_div = item.find('div', class_='time-news')
                            date_text = None
                            date_obj = None

                            if time_div:
                                time_text = time_div.get_text(strip=True)
                                date_match = re.search(r'(\d{1,2}/\d{1,2}/\d{4})', time_text)
                                if date_match:
                                    date_text = date_match.group(1)
                                    date_obj = self.parse_date(date_text)

                            page_news.append({
                                'code': code,
                                'title': title,
                                'url': url,
                                'date': date_text,
                                'date_obj': date_obj,
                                'source': 'VSD'
                            })

                        if not page_news:
                            logger.info(f"  ⚠ Page {page} không có tin nào")
                            break

                        page_dates = [n['date_obj'] for n in page_news if n['date_obj']]
                        if not page_dates:
                            page += 1
                            continue

                        page_latest_date = max(page_dates)
                        page_oldest_date = min(page_dates)

                        if latest_date_found is None:
                            latest_date_found = page_latest_date
                            logger.info(f"  ✓ Ngày gần nhất tìm thấy: {latest_date_found}")

                        all_news.extend(page_news)
                        logger.info(f"    ✓ Thêm {len(page_news)} tin từ {page_oldest_date} đến {page_latest_date}")

                        if page_oldest_date <= cutoff_date:
                            logger.info(f"  ⏹ Trang {page} có tin từ {page_oldest_date} <= {cutoff_date}, DỪNG crawl")
                            break

                        page += 1

                    except requests.exceptions.Timeout:
                        logger.error(f"  ✗ Page {page}: Request timeout")
                        break
                    except Exception as e:
                        logger.error(f"  ✗ Page {page}: {str(e)[:50]}")
                        break

                if not all_news:
                    logger.info(f"  ⚠ Không tìm thấy tin nào trên VSD")
                    return {
                        'status': 'not_found',
                        'message': 'Không tìm thấy tin trên VSD'
                    }

                # Lọc tin theo keep_days hoặc date_range
                if self.mode == 'date_range':
                    filtered_news = [n for n in all_news if n['date_obj'] and self.date_from <= n['date_obj'] <= self.date_to]
                    logger.info(f"  ✓ Lọc khoảng ngày: {len(filtered_news)} tin từ {self.date_from} đến {self.date_to} (crawled {page-1} pages)")
                else:
                    min_keep_date = latest_date_found - timedelta(days=self.keep_days - 1)
                    filtered_news = [n for n in all_news if n['date_obj'] and n['date_obj'] >= min_keep_date]
                    logger.info(f"  ✓ Tìm thấy {len(filtered_news)} tin từ {min_keep_date} đến {latest_date_found} (crawled {page-1} pages)")

            logger.info(f"  🔗 Extracting details từ tất cả {len(filtered_news)} records (concurrent, with retry)...")

            # Extract chi tiết từ tin tức - concurrent với retry để ensure page load
            result_data = []

            def extract_with_retry(news):
                """Extract chi tiết với retry logic"""
                max_retries = 2
                for attempt in range(max_retries):
                    try:
                        detail, extracted_code, actual_update_datetime = self.extract_detail_from_article(news['url'])
                        final_code = extracted_code if extracted_code else news['code']

                        # 1. Thu thập ngày/giờ đăng tin (published_at)
                        # Ưu tiên dùng ngày/giờ từ bài viết (actual_update_datetime), nếu không có thì dùng ngày listing (news['date_obj'])
                        if actual_update_datetime:
                            published_datetime = actual_update_datetime
                        elif news['date_obj']:
                            # news['date_obj'] là date object, kết hợp với time mặc định 00:00:00
                            published_datetime = datetime.combine(news['date_obj'], datetime.min.time())
                        elif news['date']:
                            # Dự phòng nếu chỉ có chuỗi news['date']
                            published_datetime = self.parse_datetime(news['date']) or self.parse_datetime(f"{news['date']} 00:00:00")
                        else:
                            published_datetime = None

                        if published_datetime:
                            published_at = published_datetime.strftime('%d/%m/%Y %H:%M:%S')
                        else:
                            published_at = f"{news['date']} 00:00:00" if news['date'] else None

                        # 2. Thu thập ngày/giờ lấy dữ liệu (collected_at)
                        now_datetime = datetime.now(VN_TZ)
                        collected_at = now_datetime.strftime('%d/%m/%Y %H:%M:%S')

                        result_item = {
                            'code': final_code,
                            'title': news['title'],
                            'url': news['url'],
                            'published_at': published_at,
                            'collected_at': collected_at,
                            'source': 'VSD',
                            'status': 'pending'
                        }

                        if detail:
                            result_item.update(detail)

                        return result_item
                    except Exception as e:
                        if attempt < max_retries - 1:
                            time.sleep(0.3)
                        else:
                            logger.error(f"Failed {news['code']}: {str(e)[:30]}")
                            # Tính toán các biến ngày giờ dự phòng
                            if news['date_obj']:
                                published_datetime = datetime.combine(news['date_obj'], datetime.min.time())
                            elif news['date']:
                                published_datetime = self.parse_datetime(news['date']) or self.parse_datetime(f"{news['date']} 00:00:00")
                            else:
                                published_datetime = None
                                
                            if published_datetime:
                                published_at = published_datetime.strftime('%d/%m/%Y %H:%M:%S')
                            else:
                                published_at = f"{news['date']} 00:00:00" if news['date'] else None
                                
                            now_datetime = datetime.now(VN_TZ)
                            collected_at = now_datetime.strftime('%d/%m/%Y %H:%M:%S')
                            
                            # Return basic item on final failure
                            return {
                                'code': news['code'],
                                'title': news['title'],
                                'url': news['url'],
                                'published_at': published_at,
                                'collected_at': collected_at,
                                'source': 'VSD',
                                'status': 'pending'
                            }

            # Extract từ tất cả records (concurrent)
            with ThreadPoolExecutor(max_workers=5) as executor:
                futures = []
                for idx, news in enumerate(filtered_news):
                    future = executor.submit(extract_with_retry, news)
                    futures.append((future, news['code']))
                    if idx % 10 == 0:  # Minimal delay every 10 items
                        time.sleep(0.05)

                for future, code in futures:
                    try:
                        result_item = future.result()

                        # Handle multiple purposes: if lý_do_mục_đích contains semicolon (;), split into multiple records
                        lý_do = result_item.get('lý_do_mục_đích')
                        if lý_do and isinstance(lý_do, str) and ';' in lý_do:
                            # Split by semicolon
                            purposes = [p.strip() for p in lý_do.split(';') if p.strip()]

                            if len(purposes) > 1:
                                # Try to split text_content and tỷ_lệ_thực_hiện by numbered sections
                                text_content = result_item.get('text_content')
                                tỷ_lệ = result_item.get('tỷ_lệ_thực_hiện')

                                # Split text_content by numbered sections (1., 2., 3., ...)
                                sections = {}
                                if text_content:
                                    # Find all numbered sections
                                    pattern = r'^\d+\.\s+(.+?)(?=\n\d+\.\s+|\Z)'
                                    matches = re.findall(pattern, text_content, re.MULTILINE | re.DOTALL)
                                    if len(matches) == len(purposes):
                                        sections = {i+1: match for i, match in enumerate(matches)}

                                # Create a record for each purpose
                                for idx, purpose in enumerate(purposes, 1):
                                    purpose_item = dict(result_item)
                                    purpose_item['lý_do_mục_đích'] = purpose

                                    # Extract tỷ_lệ_thực_hiện từ section tương ứng
                                    if idx in sections:
                                        section_text = sections[idx]
                                        # Extract tỷ_lệ từ section
                                        tỷ_lệ_match = re.search(r'Tỷ lệ thực hiện[:\s]+([^\n]+(?:\n-\s+[^\n]+)*)', section_text, re.IGNORECASE)
                                        if tỷ_lệ_match:
                                            purpose_item['tỷ_lệ_thực_hiện'] = tỷ_lệ_match.group(1).strip()[:1000]

                                    # Update title to show only the specific purpose for this split record
                                    code = result_item.get('code', 'N/A')
                                    purpose_item['title'] = f"{code}: {purpose}"

                                    # Add unique stable ID for split records (used for modal lookup)
                                    # split_idx makes it unique across runs
                                    purpose_item['_record_id'] = self.generate_record_id(purpose_item, split_idx=idx)

                                    result_data.append(purpose_item)
                                    logger.error(f"    ✓ Split {code} by purpose: [{idx}] {purpose[:60]}")
                            else:
                                result_data.append(result_item)
                        else:
                            result_data.append(result_item)

                        if len(result_data) % 100 == 0:
                            logger.info(f"    Extracted {len(result_data)}/{len(filtered_news)}")
                    except Exception as e:
                        logger.error(f"Future error {code}: {str(e)[:30]}")

            logger.info(f"  ✓ Hoàn thành extract chi tiết từ {len(result_data)} tin")

            # Add unique stable _record_id to all records (for modal lookup)
            # Split records already have _record_id (e.g., "GEX_1", "GEX_2")
            # Non-split records get _record_id based on code or hashed content
            for record in result_data:
                if '_record_id' not in record:
                    record['_record_id'] = self.generate_record_id(record)

            # Merge với records cũ để tránh duplicate
            merged_data = result_data  # Mặc định chỉ có data mới
            total_count = len(result_data)

            # Nếu file vsd_records.json tồn tại, load và merge
            # Try multiple paths for both local development and n8n container
            json_file_paths = [
                '/app/vps-automation-vhck/data/vsd_records.json',
                '/Users/hieudt/vps-automation-vhck/data/vsd_records.json',
                os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data', 'vsd_records.json')
            ]
            json_file_path = None
            for path in json_file_paths:
                if os.path.exists(path):
                    json_file_path = path
                    break

            if self.mode != 'excel_urls' and json_file_path:
                try:
                    with open(json_file_path, 'r', encoding='utf-8') as f:
                        existing_data = json.load(f)

                    # Get all records from JSON (handle both old format and new format)
                    if isinstance(existing_data, dict) and 'records' in existing_data:
                        existing_records = existing_data.get('records', [])
                    else:
                        # If JSON structure is different, try to use existing_data as list
                        existing_records = existing_data if isinstance(existing_data, list) else []

                    logger.info(f"  📚 Found {len(existing_records)} existing records, merging...")

                    merged_data = self.merge_records(result_data, existing_records)
                    logger.info(f"  ✓ Merged by URL and ticker-prefix rule: {len(result_data)} new + {len(merged_data) - len(result_data)} old = {len(merged_data)} total")
                    total_count = len(merged_data)

                except Exception as e:
                    logger.error(f"  ✗ Error merging records: {str(e)}")
                    logger.error(f"  ✗ Exception type: {type(e).__name__}")
                    import traceback
                    logger.error(f"  ✗ Traceback: {traceback.format_exc()[:100]}")
                    # Fallback: use only new data nếu merge failed
                    merged_data = result_data

            # --- FINAL FILTER BY DATE FOR EXCEL_URLS MODE ---
            if self.mode == 'excel_urls':
                final_filtered = []
                for r in merged_data:
                    # Lấy published_at của bản ghi, fallback sang published_date hoặc date
                    pub_at_str = r.get('published_at') or r.get('published_date') or r.get('date')
                    if pub_at_str:
                        try:
                            # Tách lấy phần ngày (bỏ phần giờ nếu có)
                            date_part = pub_at_str.split(' ')[0]
                            item_date = datetime.strptime(date_part, '%d/%m/%Y').date()
                            if self.date_from <= item_date <= self.date_to:
                                final_filtered.append(r)
                            else:
                                logger.info(f"    ⚠️ Lọc bỏ {r.get('code') or r.get('url')} vì ngày đăng {pub_at_str} nằm ngoài khoảng lọc {DATE_FROM} - {DATE_TO}")
                        except ValueError:
                            final_filtered.append(r)
                    else:
                        # Giữ lại các bản ghi không có ngày hoặc lỗi
                        final_filtered.append(r)
                
                merged_data = final_filtered
                total_count = len(merged_data)

            # --- APPLY BUSINESS RULES ---
            logger.info("  💼 Applying business rules from rules_2finalize.md to all records...")
            final_processed_data = []
            for r in merged_data:
                try:
                    final_processed_data.append(self.apply_business_rules(r))
                except Exception as e:
                    logger.error(f"  ✗ Error applying business rules to {r.get('code') or r.get('url')}: {e}")
                    final_processed_data.append(r)
            merged_data = final_processed_data
            total_count = len(merged_data)

            return {
                'status': 'success',
                'date': str(latest_date_found),
                'data': merged_data,
                'count': total_count,
                'url': self.news_url,
                'pages_crawled': page - 1 if self.mode != 'excel_urls' else 0,
                'fetched_at': datetime.now(VN_TZ).isoformat(),
                'merge_info': f'{len(result_data)} new records merged with existing' if self.mode != 'excel_urls' else 'EXCEL_URLS mode (no merge)'
            }

        except requests.exceptions.Timeout:
            logger.error("  ✗ Request timeout")
            return {
                'status': 'error',
                'message': 'Request timeout'
            }
        except Exception as e:
            logger.error(f"  ✗ VSD Error: {str(e)[:100]}")
            return {
                'status': 'error',
                'message': str(e)
            }

    def save_to_excel(self, data, output_path):
        """
        Tạo hoặc update file Excel từ dữ liệu records
        - Nếu file đã tồn tại: chỉ thêm records mới (code chưa có trong file cũ)
        - Nếu file chưa tồn tại: tạo file mới với tất cả records

        Args:
            data: Result dict từ fetch_latest_news() chứa 'data' key với danh sách records
            output_path: Đường dẫn output file Excel

        Returns:
            Dict với status, message, và file info
        """
        if not EXCEL_AVAILABLE:
            return {
                'status': 'error',
                'message': 'pandas hoặc openpyxl chưa được cài đặt'
            }

        try:
            new_records = data.get('data', [])

            if not new_records:
                return {
                    'status': 'error',
                    'message': 'Không có dữ liệu để export'
                }

            # Ensure output directory exists
            output_dir = os.path.dirname(output_path)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir, exist_ok=True)

            # Kiểm tra file cũ có tồn tại không
            final_records = list(new_records)
            new_count = len(new_records)

            if os.path.exists(output_path):
                try:
                    # Đọc file cũ
                    df_old = pd.read_excel(output_path, sheet_name='Tin chứng khoán')
                    df_old = df_old.where(pd.notnull(df_old), None)
                    
                    old_records = []
                    for _, row in df_old.iterrows():
                        old_record = row.to_dict()
                        if 'title' not in old_record or not old_record.get('title'):
                            old_record['title'] = old_record.get('NoiDung') or f"{old_record.get('MaChungKhoan')}: Tin chứng khoán"
                        
                        old_rid = old_record.get('_record_id') or self.generate_record_id(old_record)
                        old_record['_record_id'] = old_rid
                        
                        # Standardize dates
                        if 'published_at' not in old_record or not old_record.get('published_at'):
                            pub_at = old_record.get('published_date') or old_record.get('date')
                            if pub_at:
                                pub_at_str = str(pub_at).strip()
                                if ' ' not in pub_at_str:
                                    pub_at_str = f"{pub_at_str} 00:00:00"
                                old_record['published_at'] = pub_at_str
                            else:
                                old_record['published_at'] = None
                                
                        if 'collected_at' not in old_record or not old_record.get('collected_at'):
                            coll_at = old_record.get('collected_date')
                            if coll_at:
                                coll_at_str = str(coll_at).strip()
                                if ' ' not in coll_at_str:
                                    coll_at_str = f"{coll_at_str} 00:00:00"
                                old_record['collected_at'] = coll_at_str
                            else:
                                old_record['collected_at'] = None
                                
                        old_record.pop('date', None)
                        old_record.pop('published_date', None)
                        old_record.pop('collected_date', None)
                        old_records.append(old_record)
                        
                    final_records = self.merge_records(new_records, old_records)
                    logger.info(f"  ✓ Merged by URL and ticker-prefix rule: {new_count} new + {len(final_records) - new_count} kept = {len(final_records)} total")
                except Exception as e:
                    logger.warning(f"  ⚠ Could not read existing file: {str(e)[:50]}, will create new file")
                    # Fallback: just use new records

            # Create DataFrame từ final records
            df = pd.DataFrame(final_records)

            # Đảm bảo chỉ lấy các cột chuẩn theo đúng thứ tự trong rules_2finalize.md + cột ID và title để theo dõi chính xác
            STANDARD_COLUMNS = [
                'published_at', 'collected_at', 'url', 'text_content', 'MaChungKhoan',
                'TieuDe',
                'NhomQuyen', 'LoaiQuyen', 'MaISIN', 'MaTrongNuoc', 'NgayChot',
                'NgayGDKHQ', 'NgayThucHien', 'UocTinhNgayThucHien', 'NgayThanhToan',
                'CNQuyenMuaTuNgay', 'CNQuyenMuaDenNgay', 'DKQuyenMuaTuNgay', 'DKQuyenMuaDenNgay',
                'DonViHuongQuyen', 'GiaTriHuongQuyen', 'TyLeMenhGia', 'GiaPhatHanh',
                'NoiDung', 'is_completed', 'is_special',
                '_record_id', 'title'
            ]

            # Điền các cột còn thiếu là None và chỉ giữ lại đúng 25 cột
            for col in STANDARD_COLUMNS:
                if col not in df.columns:
                    df[col] = None
            df = df[STANDARD_COLUMNS]

            # Viết Excel file
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                df.to_excel(
                    writer,
                    sheet_name='Tin chứng khoán',
                    index=False,
                    startrow=0
                )

                # Format Excel
                workbook = writer.book
                worksheet = writer.sheets['Tin chứng khoán']

                # Auto-adjust column widths
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter

                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass

                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width

            # Check file was created
            if os.path.exists(output_path):
                file_size = os.path.getsize(output_path)
                merge_msg = f' (merged: {new_count} new + {len(final_records) - new_count} kept)' if len(final_records) > new_count else ''
                logger.info(f"✓ Excel file updated: {output_path} ({file_size} bytes)")
                return {
                    'status': 'success',
                    'message': f'Excel file saved with {len(final_records)} total records{merge_msg}',
                    'file': output_path,
                    'file_size': file_size,
                    'new_records': new_count,
                    'total_records': len(final_records),
                    'timestamp': datetime.now(VN_TZ).isoformat()
                }
            else:
                return {
                    'status': 'error',
                    'message': 'Excel file was not created'
                }

        except Exception as e:
            logger.error(f"✗ Error creating Excel: {str(e)}")
            return {
                'status': 'error',
                'message': f'Error creating Excel: {str(e)}'
            }

def main():
    fetcher = VSDFetcher()
    logger.info(f"Starting VSD fetch with KEEP_DAYS={KEEP_DAYS}")
    result = fetcher.fetch_latest_news()

    # Tạo Excel file nếu fetch thành công
    if result.get('status') == 'success' and result.get('data'):
        # Tìm path output (ưu tiên tương đối theo script để tương thích mọi môi trường, sau đó đến cứng /app/ và /Users/)
        possible_paths = [
            os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data', 'vsd_records.xlsx'),
            '/app/vps-automation-vhck/data/vsd_records.xlsx',
            '/Users/hieudt/vps-automation-vhck/data/vsd_records.xlsx'
        ]
        excel_output_path = None
        for path in possible_paths:
            try:
                os.makedirs(os.path.dirname(path), exist_ok=True)
                excel_output_path = path
                break
            except:
                continue

        if excel_output_path:
            excel_result = fetcher.save_to_excel(result, excel_output_path)
            # Thêm info về Excel vào result
            result['excel_info'] = excel_result

        # Lưu kết quả vào JSON file để đồng bộ với HTML report
        json_output_paths = [
            os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data', 'vsd_records.json'),
            '/app/vps-automation-vhck/data/vsd_records.json',
            '/Users/hieudt/vps-automation-vhck/data/vsd_records.json'
        ]
        json_output_path = None
        for path in json_output_paths:
            try:
                os.makedirs(os.path.dirname(path), exist_ok=True)
                json_output_path = path
                break
            except:
                continue

        if json_output_path:
            try:
                # Chuẩn bị dữ liệu cho HTML: gắn status field từ merged data (toàn bộ records)
                records_for_html = []

                # Lấy toàn bộ records từ Excel merge (để match Excel data)
                # Excel đã merge từ Excel file, vậy use final_records từ Excel logic
                # Nếu Excel merge thành công, Excel data sẽ đầy đủ, dùng đó
                # Otherwise fall back to result['data']
                if 'excel_info' in result and result['excel_info'].get('status') == 'success':
                    # Excel merge thành công, hãy read lại Excel để get đầy đủ merged data
                    try:
                        excel_file = result['excel_info'].get('file')
                        if excel_file and os.path.exists(excel_file):
                            df_excel = pd.read_excel(excel_file, sheet_name='Tin chứng khoán')
                            all_records = df_excel.to_dict('records')

                            # Convert NaN to None for JSON serialization
                            import math
                            for record in all_records:
                                for key, value in record.items():
                                    if isinstance(value, float) and math.isnan(value):
                                        record[key] = None

                            logger.info(f"  📊 Using {len(all_records)} records from Excel merge")
                        else:
                            all_records = result.get('data', [])
                    except:
                        all_records = result.get('data', [])
                else:
                    # Fallback: use result['data']
                    all_records = result.get('data', [])

                for record in all_records:
                    html_record = dict(record)
                    # Khôi phục các trường cần thiết cho dashboard từ cột nghiệp vụ
                    if 'code' not in html_record and html_record.get('MaChungKhoan'):
                        html_record['code'] = html_record.get('MaChungKhoan')
                    if 'title' not in html_record:
                        html_record['title'] = html_record.get('NoiDung') or f"{html_record.get('MaChungKhoan')}: Tin chứng khoán"
                    if '_record_id' not in html_record:
                        html_record['_record_id'] = fetcher.generate_record_id(html_record)

                    if 'status' not in html_record:
                        html_record['status'] = 'pending'
                    if 'confirmation_status' not in html_record:
                        html_record['confirmation_status'] = 'awaiting_review'
                    records_for_html.append(html_record)

                # Tạo JSON output cho HTML (với toàn bộ merged records)
                json_output = {
                    'status': result.get('status'),
                    'date': result.get('date'),
                    'records': records_for_html,  # Toàn bộ merged records, không chỉ new records
                    'total_records': len(records_for_html),
                    'count': result.get('count'),
                    'url': result.get('url'),
                    'pages_crawled': result.get('pages_crawled'),
                    'fetched_at': result.get('fetched_at'),
                    'merge_info': result.get('merge_info')
                }

                with open(json_output_path, 'w', encoding='utf-8') as f:
                    json.dump(json_output, f, ensure_ascii=False, indent=2)

                logger.info(f"✓ JSON file saved: {json_output_path}")
                result['json_info'] = {
                    'status': 'success',
                    'file': json_output_path,
                    'records_count': len(records_for_html)
                }
            except Exception as e:
                logger.error(f"✗ Error saving JSON: {str(e)}")
                result['json_info'] = {
                    'status': 'error',
                    'message': f'Error saving JSON: {str(e)}'
                }

    # Output JSON (remove 'data' to avoid printing massive target data to console)
    console_result = dict(result)
    console_result.pop('data', None)
    try:
        print(json.dumps(console_result, ensure_ascii=False, indent=2))
    except UnicodeEncodeError:
        print(json.dumps(console_result, ensure_ascii=True, indent=2))

if __name__ == '__main__':
    main()