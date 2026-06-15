#!/usr/bin/env python3
"""
Fetch bond/stock information from VSD (Vietnamese Securities Depository)
URL: https://www.vsd.vn/vi/tin-thi-truong-co-so

Crawl trang tin tức thị trường cơ sở:
1. Lấy danh sách tin có mã CK từ ngày gần nhất
2. Mở từng tin tức để extract chi tiết thông tin
3. Rẽ nhánh (Hybrid):
   - Nếu bản tin có nhiều quyền (nhận diện qua tiêu đề chứa ';' hoặc nội dung chứa các mẫu liệt kê):
     Gọi LLM local (hoặc Gemini API) để trích xuất danh sách các quyền có cấu trúc chính xác (is_4llm = 1).
   - Nếu bản tin đơn giản chỉ có 1 quyền:
     Dùng Regex tĩnh truyền thống để xử lý nhanh (is_4llm = 0).
4. Áp dụng các quy tắc nghiệp vụ chuẩn hóa bằng Python.
5. Ghi kết quả vào Excel và JSON.
"""

import requests
from bs4 import BeautifulSoup
import json
import sys
import time
import os
from datetime import datetime, timedelta, timezone
import calendar
import re
import logging
import hashlib
from concurrent.futures import ThreadPoolExecutor, as_completed

# Thư viện phục vụ LLM Hybrid
import openai
from openai import OpenAI
from pydantic import BaseModel, Field
from typing import Optional, List

VN_TZ = timezone(timedelta(hours=7))

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
RUN_MODE = os.environ.get("RUN_MODE", "EXCEL_URLS")
KEEP_DAYS = int(os.environ.get("KEEP_DAYS", "7"))
DATE_FROM = os.environ.get("DATE_FROM", "23/12/2025")
DATE_TO   = os.environ.get("DATE_TO", "23/04/2026")
EXCEL_URLS_FILE = "url_2fetch_round2.xlsx"

# --- CONFIGURATION LLM LOCAL / GEMINI API ---
# Trên máy cá nhân: thiết lập biến môi trường để chạy test qua Gemini API
# Trên server công ty: dùng mặc định LLM local Qwen
LLM_URL = os.environ.get("LLM_BASE_URL", "http://10.32.5.38:5000/v1")
LLM_KEY = os.environ.get("LLM_API_KEY", "sk-no-ley-required")
LLM_MODEL = os.environ.get("LLM_MODEL", "qwen3-30b-a3b-instruct")
LLM_TIMEOUT = int(os.environ.get("LLM_TIMEOUT", "180"))

# Khởi tạo client OpenAI đồng bộ (phù hợp với ThreadPoolExecutor hiện tại)
llm_client = OpenAI(
    base_url=LLM_URL,
    api_key=LLM_KEY
)

# ============================================================================

logging.basicConfig(level=logging.INFO, format='%(message)s', stream=sys.stderr)
logger = logging.getLogger(__name__)

# --- PYDANTIC SCHEMAS CHO LLM STRUCTURED OUTPUTS ---

class VsdSingleExtraction(BaseModel):
    trich_dan_nguon: str = Field(
        description="Đoạn văn bản nguyên văn trích từ CONTEXT trực tiếp mô tả quyền này. Trường này bắt buộc phải điền đầu tiên để định vị chính xác thông tin, tuyệt đối không được lẫn lộn với các phần mô tả của quyền khác."
    )
    ly_do_muc_dich: str = Field(
        description="Lý do mục đích cụ thể của quyền này (ví dụ: 'Chi trả cổ tức bằng tiền năm 2025' hoặc 'Thực hiện quyền mua cổ phiếu', 'Tổ chức Đại hội đồng cổ đông thường niên năm 2026')"
    )
    nhom_quyen: str = Field(
        description="Phân loại nhóm quyền. Bắt buộc phải chọn 1 trong: 'Cổ tức tiền', 'Cổ tức cổ phiếu / Cổ phiếu thưởng', 'Quyền biểu quyết', 'Quyền mua', 'Hoán đổi chuyển đổi', 'Chứng quyền', 'Đăng ký Lưu ký', 'Tin huỷ', 'Thay đổi'"
    )
    loai_quyen: Optional[str] = Field(
        description="""Loại quyền chi tiết, phụ thuộc chặt chẽ vào nhom_quyen:
1. Nếu nhom_quyen là 'Cổ tức cổ phiếu / Cổ phiếu thưởng':
   - Điền 'Cổ phiếu thưởng' (nếu phát hành cổ phiếu thưởng, thưởng cổ phiếu, phát hành cổ phiếu để tăng vốn).
   - Điền 'Cổ tức cổ phiếu' (nếu trả cổ tức bằng cổ phiếu, cổ tức cổ phiếu, cổ phiếu để trả cổ tức).
2. Nếu nhom_quyen là 'Cổ tức tiền':
   - Điền 'Trái phiếu' (nếu là thanh toán gốc/lãi trái phiếu, mua lại trái phiếu trước hạn).
   - Điền 'Cổ phiếu' (cho tất cả các trường hợp cổ tức tiền mặt còn lại).
3. Nếu nhom_quyen là 'Đăng ký Lưu ký':
   - Điền 'Đăng ký' (nếu là đăng ký chứng khoán/cổ phiếu/trái phiếu).
   - Điền 'Lưu ký' (nếu là lưu ký chứng khoán/cổ phiếu/trái phiếu).
4. Nếu nhom_quyen là 'Tin huỷ':
   - Phải chọn chính xác 1 trong các giá trị sau:
     * 'Hủy đăng ký chứng khoán'
     * 'Hủy đăng ký chứng quyền'
     * 'Hủy đăng ký trái phiếu'
     * 'Hủy đợt chốt danh sách thực hiện chứng quyền'
     * 'Hủy danh sách người sở hữu chứng khoán'
     * 'Hủy thông báo ngày đăng ký cuối cùng'
5. Các trường hợp nhom_quyen khác: Để trống (null)."""
    )
    ngay_chot_raw: Optional[str] = Field(
        description="Chuỗi ngày đăng ký cuối cùng (ngày chốt danh sách) tìm thấy trong văn bản. Giữ nguyên định dạng gốc như '22/12/2025', 'tháng 12 năm 2025', hoặc 'ngày 15 tháng 06 năm 2026'."
    )
    ngay_thuc_hien_raw: Optional[str] = Field(
        description="Chuỗi ngày thực hiện (thường dùng cho Quyền biểu quyết/đại hội). Giữ nguyên định dạng gốc."
    )
    ngay_thanh_toan_raw: Optional[str] = Field(
        description="Chuỗi ngày thanh toán (thường dùng cho Cổ tức tiền). Giữ nguyên định dạng gốc."
    )
    ty_le_thuc_hien_raw: Optional[str] = Field(
        description="Đoạn văn bản mô tả tỷ lệ thực hiện quyền (ví dụ: '100 cổ phiếu được nhận 73.500,7476 cổ phiếu' hoặc '10:1' hoặc '100:15')."
    )
    ty_le_menh_gia_percent: Optional[str] = Field(
        description="Tỷ lệ phần trăm (%) nhận cổ tức tiền mặt tìm thấy trực tiếp trong văn bản (ví dụ: '12%', '10%')."
    )
    gia_phat_hanh_raw: Optional[str] = Field(
        description="Giá phát hành quyền mua tìm thấy trong văn bản (ví dụ: '10.000 đồng')."
    )
    cn_quyen_mua_tu_ngay: Optional[str] = Field(description="Ngày bắt đầu chuyển nhượng quyền mua (DD/MM/YYYY hoặc chuỗi gốc)")
    cn_quyen_mua_den_ngay: Optional[str] = Field(description="Ngày kết thúc chuyển nhượng quyền mua (DD/MM/YYYY hoặc chuỗi gốc)")
    dk_quyen_mua_tu_ngay: Optional[str] = Field(description="Ngày bắt đầu đăng ký đặt mua (DD/MM/YYYY hoặc chuỗi gốc)")
    dk_quyen_mua_den_ngay: Optional[str] = Field(description="Ngày kết thúc đăng ký đặt mua (DD/MM/YYYY hoặc chuỗi gốc)")
    ma_isin: Optional[str] = Field(description="Mã ISIN của chứng khoán. Thường là chuỗi 12 ký tự bắt đầu bằng chữ và kết thúc bằng số.")
    ma_trong_nuoc: Optional[str] = Field(description="Mã trong nước của chứng khoán hoặc mã quyền mua (chuỗi 9 ký tự).")
    noi_dung: Optional[str] = Field(
        description="Mô tả tóm tắt nội dung chính của quyền này (Ví dụ: 'Trả cổ tức năm 2025 bằng tiền tỷ lệ 10% - Thanh toán ngày 15/06/2026')"
    )

class VsdMultipleExtractions(BaseModel):
    extractions: List[VsdSingleExtraction] = Field(
        description="Danh sách tất cả các quyền độc lập trích xuất được từ bản tin VSD này"
    )


# --- HELPERS CHO NHẬN DIỆN BẢN TIN NHIỀU QUYỀN (RẼ NHÁNH HYBRID) ---

def is_multi_rights_article(title: str, text_content: str) -> bool:
    """
    Xác định xem bản tin có chứa nhiều quyền hay không dựa trên title và text_content.
    Quy tắc:
    1. Nếu title chứa dấu chấm phẩy ';' -> Chắc chắn có nhiều quyền (Dùng LLM).
    2. Nếu title không chứa ';': Xét tiếp text_content xem có chứa các mẫu liệt kê danh sách
       như 1., 2. hoặc a., b. hoặc A., B. hoặc a), b) hoặc 1/, 2/ (đứng đầu dòng) hay không.
    """
    if not title:
        return False
        
    # 1. Kiểm tra tiêu đề chứa ';'
    if ';' in title:
        return True

    if not text_content:
        return False

    # 2. Định nghĩa các mẫu Regex liệt kê đầu dòng
    # Dạng số: 1. hoặc 1) hoặc 1/ (theo sau bởi khoảng trắng và ký tự để tránh ngày tháng)
    num_pattern = r'(?:^|\n)\s*\d+[\.\)\/]\s+[A-Za-zĐđ]'
    
    # Dạng chữ cái thường: a. hoặc a) hoặc a/
    alpha_lower_pattern = r'(?:^|\n)\s*[a-g][\.\)\/]\s+[A-Za-zĐđ]'
    
    # Dạng chữ cái hoa: A. hoặc A) hoặc A/
    alpha_upper_pattern = r'(?:^|\n)\s*[A-G][\.\)\/]\s+[A-Za-zĐđ]'
    
    # Dạng số La Mã: I. hoặc II. hoặc I) hoặc II)
    roman_pattern = r'(?:^|\n)\s*[IVXLCDMivxlcdm]+[\.\)\/]\s+[A-Za-zĐđ]'

    # Kiểm tra xem có khớp bất kỳ mẫu liệt kê nào không
    if re.search(num_pattern, text_content) or \
       re.search(alpha_lower_pattern, text_content) or \
       re.search(alpha_upper_pattern, text_content) or \
       re.search(roman_pattern, text_content):
        return True

    return False


# --- BUSINESS RULES HELPER FUNCTIONS ---

def remove_accents_and_lower(text):
    """
    Tiền xử lý chuỗi: chuyển sang lowercase, loại bỏ dấu tiếng Việt,
    loại bỏ khoảng trắng thừa và ký tự đặc biệt nếu có.
    """
    if not text:
        return ""
    text = str(text).lower().strip()
    
    # Loại bỏ khoảng trắng thừa ngang, giữ nguyên ký tự xuống dòng
    lines = [re.sub(r'[ \t\r\f\v]+', ' ', line).strip() for line in text.split('\n')]
    text = '\n'.join(lines)
    
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
    """Trả về ngày cuối cùng của tháng trong năm tương ứng."""
    try:
        return calendar.monthrange(year, month)[1]
    except Exception:
        return 28  # Fallback

def parse_raw_date_string(date_str: str) -> Optional[str]:
    """
    Hậu xử lý chuỗi ngày thô nhận được từ LLM sang định dạng DD/MM/YYYY.
    Hỗ trợ xử lý 'tháng mm năm YYYY' -> ngày cuối tháng.
    """
    if not date_str:
        return None
        
    date_str = date_str.strip()
    
    # 1. Nếu có định dạng chuẩn DD/MM/YYYY hoặc D/M/YYYY
    std_match = re.search(r'(\d{1,2})/(\d{1,2})/(\d{4})', date_str)
    if std_match:
        d, m, y = int(std_match.group(1)), int(std_match.group(2)), int(std_match.group(3))
        return f"{d:02d}/{m:02d}/{y}"

    # 2. Định dạng ngày dd tháng mm năm YYYY
    text_date_match = re.search(r'(?:ngay\s+)?(\d{1,2})\s+thang\s+(\d{1,2})\s+nam\s+(\d{4})', remove_accents_and_lower(date_str))
    if text_date_match:
        d, m, y = int(text_date_match.group(1)), int(text_date_match.group(2)), int(text_date_match.group(3))
        return f"{d:02d}/{m:02d}/{y}"

    # 3. Định dạng tháng mm năm YYYY -> lấy ngày cuối tháng
    month_year_text_match = re.search(r'thang\s+(\d{1,2})\s+nam\s+(\d{4})', remove_accents_and_lower(date_str))
    if month_year_text_match:
        m, y = int(month_year_text_match.group(1)), int(month_year_text_match.group(2))
        last_day = get_last_day_of_month(m, y)
        return f"{last_day:02d}/{m:02d}/{y}"

    # 4. Định dạng mm/YYYY -> lấy ngày cuối tháng
    month_year_slash_match = re.search(r'(?:^|[^\d])(\d{1,2})/(\d{4})(?:$|[^\d])', date_str)
    if month_year_slash_match:
        m, y = int(month_year_slash_match.group(1)), int(month_year_slash_match.group(2))
        last_day = get_last_day_of_month(m, y)
        return f"{last_day:02d}/{m:02d}/{y}"

    # Fallback: dùng regex cũ quét chuỗi
    return extract_earliest_date(date_str)

def extract_earliest_date(text_segment):
    """Tìm chuỗi ngày sớm nhất trong text_segment."""
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
        
    matches.sort(key=lambda x: x[0])
    return matches[0][1]

def extract_all_dates_in_segment(text_segment):
    """Trích xuất toàn bộ ngày có trong đoạn văn bản."""
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
    """Tìm một dòng trong original_text chứa một trong các keywords."""
    if not original_text:
        return None
    lines = original_text.split('\n')
    for line in lines:
        pre_line = remove_accents_and_lower(line)
        for kw in keywords:
            if kw in pre_line:
                return line.strip()
    return None

def parse_retry_delay(error_obj) -> Optional[float]:
    """
    Trích xuất số giây cần chờ (retryDelay) từ thông báo lỗi của Gemini/OpenAI API.
    Hỗ trợ cả định dạng text lẫn dict/JSON.
    """
    if not error_obj:
        return None
    err_str = str(error_obj)
    
    # 1. Định dạng text của Gemini API: "Please retry in 11.105552559s."
    match_a = re.search(r'please retry in (\d+(?:\.\d+)?)\s*s', err_str, re.IGNORECASE)
    if match_a:
        return float(match_a.group(1))
        
    # 2. Định dạng dict trong response: 'retryDelay': '11s' hoặc "retryDelay": "11s"
    match_b = re.search(r'[\'"]retryDelay[\'"]\s*:\s*[\'"](\d+)\s*s[\'"]', err_str, re.IGNORECASE)
    if match_b:
        return float(match_b.group(1))
        
    return None

def normalize_ratio(ratio_str: str) -> tuple:
    """
    Hậu xử lý tỷ lệ thực hiện (DonViHuongQuyen X và GiaTriHuongQuyen Y) bằng Python.
    Áp dụng quy tắc nhân 10^n nếu Y chứa phần thập phân.
    Ví dụ: '100 cổ phiếu được nhận 73.500,7476 cổ phiếu' -> X=1000000, Y=735007476.
    """
    if not ratio_str:
        return None, None

    # CHỈ loại bỏ ngoặc đơn nếu bên trong KHÔNG chứa chữ số (ví dụ '(một)', '(hai)')
    # Điều này giúp giữ lại '(01 cổ phiếu được nhận 180 đồng)' nguyên vẹn
    ratio_str_clean = re.sub(r'\(\s*[^)\d]+\s*\)', '', ratio_str)
    
    clean_lower = remove_accents_and_lower(ratio_str_clean)
    
    # 1. Tìm theo mẫu có cấu trúc: X cổ phiếu/trái phiếu... được nhận Y
    pattern_a = re.search(
        r'(\d+)\s*(?:co phieu|trai phieu|chung chi quy)[^0-9\n]*(\d+(?:\.\d+)*(?:\,\d+)?)',
        clean_lower
    )
    if pattern_a:
        x_raw = pattern_a.group(1)
        y_raw = pattern_a.group(2)
    else:
        # 2. Tìm theo định dạng X:Y hoặc X/Y
        pattern_c = re.search(r'(\d+)\s*[:/]\s*(\d+(?:\.\d+)*(?:\,\d+)?)', clean_lower)
        if pattern_c:
            x_raw = pattern_c.group(1)
            y_raw = pattern_c.group(2)
        else:
            # 3. Phương án dự phòng cuối: tìm tất cả các số và lấy 2 số đầu tiên
            numbers = re.findall(r'(\d+(?:\.\d+)*(?:\,\d+)?)', ratio_str_clean)
            if len(numbers) >= 2:
                x_raw = numbers[0]
                y_raw = numbers[1]
            else:
                return None, None

    try:
        # Chuẩn hóa x_raw thành số nguyên (Đơn vị thường là số nguyên như 1, 100, 1000)
        x_val = int(x_raw.replace('.', '').replace(',', ''))
        
        # Kiểm tra y_raw có dấu phẩy thập phân (tiếng Việt) hay không
        if ',' in y_raw:
            # Ví dụ: '73.500,7476'
            parts = y_raw.split(',')
            integer_part = parts[0].replace('.', '') # '73500'
            decimal_part = parts[1] # '7476'
            
            n = len(decimal_part) # Số chữ số thập phân
            factor = 10 ** n
            
            # Nhân cả hai số với 10^n
            y_val = int(integer_part) * factor + int(decimal_part)
            x_val = x_val * factor
        elif '.' in y_raw and len(y_raw.split('.')[-1]) <= 4 and not re.search(r'\d{3}\.\d{3}', y_raw):
            # Phòng ngừa trường hợp dấu chấm được dùng làm dấu thập phân thay vì dấu phẩy
            parts = y_raw.split('.')
            integer_part = parts[0]
            decimal_part = parts[1]
            n = len(decimal_part)
            factor = 10 ** n
            y_val = int(integer_part) * factor + int(decimal_part)
            x_val = x_val * factor
        else:
            # Y là số nguyên thường
            y_val = int(y_raw.replace('.', '').replace(',', ''))
            
        return x_val, y_val
    except Exception as e:
        logger.error(f"Error normalising ratio: {e}")
        return None, None


# --- CLASS CHÍNH: VSD FETCHER TÍCH HỢP LLM HYBRID ---

class VSDFetcher:
    def __init__(self):
        self.base_url = "https://www.vsd.vn"
        self.news_url = "https://www.vsd.vn/vi/tin-thi-truong-co-so"
        self.session = requests.Session()
        self.vptoken = None
        self.date_from = None
        self.date_to   = None
        
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

    # --- CALL LLM TO EXTRACT MULTIPLE RIGHTS ---
    
    def extract_multiple_purposes_via_llm(self, text_content: str) -> List[dict]:
        """
        Gọi LLM để trích xuất danh sách các quyền từ nội dung bản tin nhiều quyền.
        Sử dụng Structured Outputs của OpenAI SDK kết hợp Pydantic Model.
        """
        prompt = f"""You are an expert financial data analyst specializing in Vietnamese securities documents.
Your task is to analyze the VSD (Vietnam Securities Depository) announcement provided in <CONTEXT> and extract all distinct corporate actions/rights into the 'extractions' list.

CRITICAL RULES FOR EXTRACTION:
1. INFORMATION ISOLATION (Ngăn ngừa lẫn lộn thông tin):
   - A document may contain multiple corporate actions/rights (e.g., Cash Dividend and Stock Dividend described in different sections).
   - For each corporate action, identify its specific section/paragraph. You MUST extract attributes (ratios, dates, payment dates) ONLY from that specific section. Do not mix or swap attributes between different sections.
   - If a date (e.g., ngày đăng ký cuối cùng / ngày chốt) is mentioned as a general date for the entire announcement (usually at the beginning or end of the document), apply it to all relevant corporate actions. If a date is mentioned within a specific section, only apply it to that specific action.

2. TRICH DAN NGUON (Source quoting):
   - For each extraction, you MUST first fill in the 'trich_dan_nguon' field with the exact text segment describing the right from the CONTEXT. This anchors your attention and prevents mixing data.

3. RIGHTS CLASSIFICATION (nhom_quyen & loai_quyen):
   - 'Đăng ký Lưu ký': Only classify as this if it refers to registering or depositing securities for trading. Ignore references to the organization name "Tổng công ty Lưu ký và Bù trừ chứng khoán Việt Nam (VSDC)" or "tổ chức đăng ký...".
   - 'Tin huỷ': Check title or text for cancellations (e.g. 'hủy đăng ký', 'hủy danh sách').
   - For 'loai_quyen', you MUST strictly follow the Pydantic field description rules.

4. RAW DATE EXTRACTION:
   - For all date fields (ngay_chot_raw, ngay_thuc_hien_raw, ngay_thanh_toan_raw), extract the exact date string as written in the text. E.g., "tháng 12/2025", "ngày 22/12/2025", "ngày 15 tháng 06 năm 2026". Do not calculate or change them.

5. RATIO EXTRACTION (ty_le_thuc_hien_raw):
   - Capture the complete phrase detailing the execution ratio. E.g., "01 cổ phiếu nhận được 1.000 đồng", "100 cổ phiếu nhận 73.500,7476 cổ phiếu", "10:1", "100:15". Ensure you capture any decimal points.

6. NO HALLUCINATION:
   - If a field is not found in the context, leave it as null.

<CONTEXT>
{text_content}
</CONTEXT>"""

        max_retries = 3
        result_extractions = []
        next_wait_time = 0
        
        for attempt in range(max_retries):
            try:
                if attempt > 0:
                    # Nếu có thời gian chờ cụ thể từ API ở vòng lặp trước, dùng nó, ngược lại dùng lũy tiến 20s
                    wait_time = next_wait_time if next_wait_time > 0 else (20 * attempt)
                    logger.info(f"  ⚠ Bị giới hạn tần suất (429). Đang chờ {wait_time:.2f} giây trước khi thử lại (Lần {attempt + 1}/{max_retries})...")
                    time.sleep(wait_time)
                
                # Reset next_wait_time cho lần thử này
                next_wait_time = 0
                
                response = llm_client.beta.chat.completions.parse(
                    model=LLM_MODEL,
                    messages=[
                        {"role": "user", "content": prompt}
                    ],
                    response_format=VsdMultipleExtractions,
                    temperature=0,
                    timeout=LLM_TIMEOUT
                )
                parsed_result = response.choices[0].message.parsed
                if parsed_result and parsed_result.extractions:
                    result_extractions = [ext.model_dump() for ext in parsed_result.extractions]
                break  # Thành công, thoát vòng lặp retry
            except openai.RateLimitError as e:
                # Trích xuất retry delay từ lỗi
                delay = parse_retry_delay(e)
                if delay:
                    next_wait_time = delay + 2.0  # Cộng thêm 2 giây buffer cho an toàn
                else:
                    next_wait_time = 20 * (attempt + 1)
                    
                if attempt == max_retries - 1:
                    logger.error(f"✗ Vượt quá giới hạn tần suất sau {max_retries} lần thử: {e}")
                continue
            except Exception as e:
                # Bắt lỗi 429 trả về dưới dạng Exception khác hoặc thông báo chuỗi
                err_msg = str(e).lower()
                if "429" in err_msg or "rate limit" in err_msg or "quota" in err_msg or "too many requests" in err_msg:
                    delay = parse_retry_delay(e)
                    if delay:
                        next_wait_time = delay + 2.0
                    else:
                        next_wait_time = 20 * (attempt + 1)
                        
                    if attempt == max_retries - 1:
                        logger.error(f"✗ Vượt quá giới hạn tần suất/quota sau {max_retries} lần thử: {e}")
                    continue
                logger.error(f"✗ Lỗi khi gọi LLM để trích xuất: {e}")
                break
        
        # Áp dụng Throttling (ngủ 5 giây sau mỗi lần gọi LLM để khống chế dưới 15 RPM)
        time.sleep(5)
        return result_extractions

    # --- POST-PROCESSING LLM DATA WITH PYTHON ---
    
    def process_llm_extracted_record(self, ext_raw: dict, base_record: dict) -> dict:
        """
        Nhận kết quả trích xuất thô của 1 quyền từ LLM, áp dụng các quy tắc nghiệp vụ
        của Python để tính toán các trường và chuẩn hóa định dạng.
        """
        record = dict(base_record)
        ticker_code = record.get('code', 'N/A')
        
        # 1. Các trường cơ bản
        nhom_quyen = ext_raw.get('nhom_quyen')
        loai_quyen = ext_raw.get('loai_quyen')
        ly_do_muc_dich = ext_raw.get('ly_do_muc_dich')
        
        # 2. Xử lý ngày tháng bằng hàm Python chuẩn
        ngay_chot = parse_raw_date_string(ext_raw.get('ngay_chot_raw'))
        ngay_thuc_hien = parse_raw_date_string(ext_raw.get('ngay_thuc_hien_raw'))
        ngay_thanh_toan = parse_raw_date_string(ext_raw.get('ngay_thanh_toan_raw'))
        
        # Ràng buộc nghiệp vụ: Quyền biểu quyết/Cổ phiếu thưởng/Cổ tức bằng cổ phiếu -> NgayThanhToan = null
        if nhom_quyen in ["Quyền biểu quyết", "Cổ phiếu thưởng", "Cổ tức bằng cổ phiếu"]:
            ngay_thanh_toan = None
            
        # Ràng buộc nghiệp vụ: Chỉ Quyền biểu quyết mới có NgayThucHien
        if nhom_quyen != "Quyền biểu quyết":
            ngay_thuc_hien = None

        # Tính NgayGDKHQ = NgayChot - 1 ngày
        ngay_gdkhq = None
        if ngay_chot:
            try:
                chot_date = datetime.strptime(ngay_chot, '%d/%m/%Y').date()
                gdkhq_date = chot_date - timedelta(days=1)
                ngay_gdkhq = gdkhq_date.strftime('%d/%m/%Y')
            except Exception as e:
                logger.error(f"Error calculating NgayGDKHQ: {e}")

        # Tính UocTinhNgayThucHien cho Quyền biểu quyết
        uoc_tinh_ngay_thuc_hien = None
        if nhom_quyen == "Quyền biểu quyết" and not ngay_thuc_hien:
            if ngay_chot:
                try:
                    chot_date = datetime.strptime(ngay_chot, '%d/%m/%Y').date()
                    y = chot_date.year
                    m = chot_date.month + 1
                    if m > 12:
                        m = 1
                        y += 1
                    last_day = get_last_day_of_month(m, y)
                    ngay_thuc_hien = f"{last_day:02d}/{m:02d}/{y}"
                    uoc_tinh_ngay_thuc_hien = 1
                except Exception as e:
                    logger.error(f"Error calculating fallback NgayThucHien: {e}")

        # 3. Xử lý tỷ lệ thực hiện (DonViHuongQuyen và GiaTriHuongQuyen)
        don_vi_huong_quyen = None
        gia_tri_huong_quyen = None
        
        if nhom_quyen == "Quyền biểu quyết":
            don_vi_huong_quyen = 1
            gia_tri_huong_quyen = 1
        else:
            # Gọi hàm hậu xử lý tỷ lệ bằng Python
            don_vi_huong_quyen, gia_tri_huong_quyen = normalize_ratio(ext_raw.get('ty_le_thuc_hien_raw'))

        # 4. Xử lý TyLeMenhGia cho Cổ tức tiền
        ty_le_menh_gia = None
        if nhom_quyen == "Cổ tức tiền":
            if loai_quyen == "Cổ phiếu":
                percent_str = ext_raw.get('ty_le_menh_gia_percent')
                if percent_str:
                    try:
                        ty_le_menh_gia = float(percent_str.replace('%', '').strip())
                    except:
                        pass
                # Fallback tính toán bằng công thức: GiaTriHuongQuyen / (10 ** (2 + k))
                if ty_le_menh_gia is None and gia_tri_huong_quyen is not None and don_vi_huong_quyen is not None:
                    import math
                    try:
                        k = math.log10(don_vi_huong_quyen)
                        ty_le_menh_gia = float(gia_tri_huong_quyen) / (10 ** (2 + k))
                    except Exception as e:
                        logger.error(f"Error calculating TyLeMenhGia for stock: {e}")
            elif loai_quyen == "Trái phiếu":
                menh_gia_val = None
                menh_gia_str = base_record.get('mệnh_giá')
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

        # 5. Xử lý GiaPhatHanh
        gia_phat_hanh = None
        gph_raw = ext_raw.get('gia_phat_hanh_raw')
        if gph_raw:
            try:
                # Trích xuất số nguyên từ chuỗi giá phát hành
                num_match = re.search(r'(\d+(?:\.\d+)+|\d+)', gph_raw)
                if num_match:
                    gia_phat_hanh = num_match.group(1).replace('.', '')
            except:
                pass

        # 6. Mã ISIN & Trong nước
        ma_isin = ext_raw.get('ma_isin')
        if ma_isin:
            ma_isin = ma_isin.strip().upper()
            
        ma_trong_nuoc = ext_raw.get('ma_trong_nuoc')
        if ma_trong_nuoc:
            ma_trong_nuoc = ma_trong_nuoc.strip().upper()

        # 7. Xây dựng Nội dung và Tiêu đề chuẩn format
        tieu_de = f"{ticker_code}: {ly_do_muc_dich}"
        
        # Lấy nguyên văn tiếng Việt có dấu từ text_content cho NoiDung
        noi_dung = tieu_de
        if nhom_quyen in ["Hoán đổi chuyển đổi", "Khai báo chứng quyền", "Đăng ký Lưu ký", "Thay đổi"]:
            noi_dung = None
        elif nhom_quyen == "Tin huỷ":
            noi_dung = loai_quyen
        elif nhom_quyen == "Quyền biểu quyết":
            # Tiêu đề không kèm code
            noi_dung = ly_do_muc_dich
        else:
            # Ghép chuỗi chuẩn format tiếng Việt
            # Cô lập phạm vi tìm kiếm thông tin bằng cách ưu tiên sử dụng đoạn văn bản nguồn được LLM trích xuất cho quyền này
            orig_text = ext_raw.get('trich_dan_nguon') or base_record.get('text_content', '')
            parts_list = [ly_do_muc_dich]
            
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

        # 8. Xác định is_completed cho bản ghi LLM
        is_completed = 0
        if nhom_quyen and ticker_code != 'N/A' and ngay_chot:
            criteria_met = True
            if nhom_quyen == "Quyền mua":
                quyen_mua_fields = [
                    ext_raw.get('cn_quyen_mua_tu_ngay'), ext_raw.get('cn_quyen_mua_den_ngay'),
                    ext_raw.get('dk_quyen_mua_tu_ngay'), ext_raw.get('dk_quyen_mua_den_ngay'),
                    don_vi_huong_quyen, gia_tri_huong_quyen,
                    gia_phat_hanh, ma_isin, ma_trong_nuoc
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

        is_special = 0
        if nhom_quyen in ["Hoán đổi chuyển đổi", "Khai báo chứng quyền", "Đăng ký Lưu ký", "Tin huỷ", "Thay đổi"] or ';' in base_record.get('title', ''):
            is_special = 1

        # Trả về bản ghi cập nhật
        record.update({
            'TieuDe': ly_do_muc_dich,
            'title': tieu_de,
            'NhomQuyen': nhom_quyen,
            'LoaiQuyen': loai_quyen,
            'MaISIN': ma_isin,
            'MaTrongNuoc': ma_trong_nuoc,
            'NgayChot': ngay_chot,
            'NgayGDKHQ': ngay_gdkhq,
            'NgayThucHien': ngay_thuc_hien,
            'UocTinhNgayThucHien': uoc_tinh_ngay_thuc_hien,
            'NgayThanhToan': ngay_thanh_toan,
            'CNQuyenMuaTuNgay': ext_raw.get('cn_quyen_mua_tu_ngay'),
            'CNQuyenMuaDenNgay': ext_raw.get('cn_quyen_mua_den_ngay'),
            'DKQuyenMuaTuNgay': ext_raw.get('dk_quyen_mua_tu_ngay'),
            'DKQuyenMuaDenNgay': ext_raw.get('dk_quyen_mua_den_ngay'),
            'DonViHuongQuyen': don_vi_huong_quyen,
            'GiaTriHuongQuyen': gia_tri_huong_quyen,
            'TyLeMenhGia': ty_le_menh_gia,
            'GiaPhatHanh': gia_phat_hanh,
            'NoiDung': noi_dung,
            'is_completed': is_completed,
            'is_special': is_special,
            'is_4llm': 1 # Đánh dấu được xử lý bởi LLM
        })
        
        return record

    def apply_business_rules(self, record):
        """
        Áp dụng các quy tắc nghiệp vụ từ rules_2finalize.md lên một bản ghi.
        Phương thức này chỉ được gọi đối với các bản ghi đơn giản (is_4llm = 0).
        """
        code = record.get('code') or ''
        title = record.get('title') or ''
        ly_do = record.get('lý_do_mục_đích') or ''
        noi_gd = record.get('nơi_giao_dịch') or ''
        loai_ck = record.get('loại_chứng_khoán') or ''
        text_content = record.get('text_content') or ''
        
        pre_title = remove_accents_and_lower(title)
        pre_ly_do = remove_accents_and_lower(ly_do)
        pre_noi_gd = remove_accents_and_lower(noi_gd)
        pre_loai_ck = remove_accents_and_lower(loai_ck)
        pre_text = remove_accents_and_lower(text_content)
        
        ma_ck = code
        nhom_quyen = None
        
        # a. Tin huỷ
        col_idx = pre_title.find(':')
        if col_idx != -1:
            after_col = pre_title[col_idx+1:].strip()
            if after_col.startswith('huy') and 'chung quyen' not in pre_title:
                nhom_quyen = "Tin huỷ"
        else:
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
            # Ràng buộc loại trừ "to chuc" ngay trước keyword
            # Ví dụ: "to chuc dang ky" không được coi là quyền Đăng ký
            has_valid_dk = False
            for kw in dk_keywords:
                if kw in pre_title or kw in pre_ly_do:
                    # Kiểm tra xem có chứa "to chuc " ngay trước kw không
                    invalid_kw = f"to chuc {kw}"
                    if invalid_kw not in pre_title and invalid_kw not in pre_ly_do:
                        has_valid_dk = True
                        break
            if has_valid_dk:
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
                isin_match = re.search(r'([a-z][a-z0-9]{10}\d)', target_sub)
                if isin_match:
                    isin_val = isin_match.group(1).upper()

        # MaTrongNuoc
        ma_trong_nuoc = None
        for kw in ["ma quyen mua", "ma trong nuoc"]:
            kw_pos = pre_text.find(kw)
            if kw_pos != -1:
                sub_text = pre_text[kw_pos:]
                dash_idx = sub_text.find('-')
                nl_idx = sub_text.find('\n')
                indices = [idx for idx in [dash_idx, nl_idx] if idx != -1]
                end_pos = min(indices) if indices else len(sub_text)
                target_sub = sub_text[:end_pos]
                mn_match = re.search(r'([a-z][a-z0-9]{7}\d)', target_sub)
                if mn_match:
                    ma_trong_nuoc = mn_match.group(1).upper()
                    break

        # NgayChot
        ngay_chot_val = record.get('ngày_đăng_KY_cuối') or record.get('ngày_đăng_ký_cuối') or record.get('ngày_đăng_ky_cuối')
        if not ngay_chot_val:
            for kw in ["ngay dang ky cuoi", "thoi gian dang ky cuoi", "ngay chot", "thoi gian chot"]:
                kw_pos = pre_text.find(kw)
                if kw_pos != -1:
                    segment = pre_text[kw_pos:kw_pos + 200]
                    date_found = extract_earliest_date(segment)
                    if date_found:
                        ngay_chot_val = date_found
                        break

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
            for kw in ["thoi gian thuc hien", "ngay thuc hien", "thoi gian hien", "ngay hien", "thoi gian thuc", "ngay thuc"]:
                kw_pos = pre_text.find(kw)
                if kw_pos != -1:
                    sub_text = pre_text[kw_pos:]
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
                        break
            
            if not found:
                if ngay_chot_val:
                    try:
                        chot_date = datetime.strptime(ngay_chot_val, '%d/%m/%Y').date()
                        y = chot_date.year
                        m = chot_date.month + 1
                        if m > 12:
                            m = 1
                            y += 1
                        last_day = get_last_day_of_month(m, y)
                        ngay_thuc_hien = f"{last_day:02d}/{m:02d}/{y}"
                        uoc_tinh_ngay_thuc_hien = 1
                    except Exception as e:
                        logger.error(f"Error calculating fallback NgayThucHien: {e}")

        # NgayThanhToan
        ngay_thanh_toan = None
        if nhom_quyen not in ["Quyền biểu quyết", "Cổ phiếu thưởng", "Cổ tức bằng cổ phiếu"] and nhom_quyen is not None:
            for kw in ["ngay thuc hien", "ngay thanh toan", "thoi gian thuc hien", "thoi gian thanh toan"]:
                kw_pos = pre_text.find(kw)
                if kw_pos != -1:
                    sub_text = pre_text[kw_pos:]
                    dash_idx = sub_text.find('-')
                    nl_idx = sub_text.find('\n')
                    indices = [idx for idx in [dash_idx, nl_idx] if idx != -1]
                    end_pos = min(indices) if indices else len(sub_text)
                    
                    target_sub = sub_text[:end_pos]
                    dates_list = extract_all_dates_in_segment(target_sub)
                    if dates_list:
                        ngay_thanh_toan = dates_list[-1]
                        break

        # CNQuyenMuaTuNgay & CNQuyenMuaDenNgay
        cn_quyen_mua_tu_ngay = None
        cn_quyen_mua_den_ngay = None
        if nhom_quyen == "Quyền mua":
            for kw in ["thoi gian chuyen nhuong", "ngay chuyen nhuong", "han chuyen nhuong"]:
                kw_pos = pre_text.find(kw)
                if kw_pos != -1:
                    sub_text = pre_text[kw_pos:]
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
                        break

        # DKQuyenMuaTuNgay & DKQuyenMuaDenNgay
        dk_quyen_mua_tu_ngay = None
        dk_quyen_mua_den_ngay = None
        if nhom_quyen == "Quyền mua":
            for kw in ["thoi gian dang ky", "ngay dang ky", "han dang ky", "thoi gian dat", "ngay dat", "han dat", "thoi gian nop tien", "ngay nop tien", "han nop tien"]:
                kw_pos = pre_text.find(kw)
                if kw_pos != -1:
                    sub_text = pre_text[kw_pos:]
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
                        break

        # DonViHuongQuyen & GiaTriHuongQuyen
        don_vi_huong_quyen = None
        gia_tri_huong_quyen = None
        if nhom_quyen == "Quyền biểu quyết":
            don_vi_huong_quyen = 1
            gia_tri_huong_quyen = 1
        elif nhom_quyen is not None:
            scope_text = ""
            for kw in ["ty le thuc hien", "ty le thanh toan", "ti le thuc hien", "ti le thanh toan"]:
                kw_pos = pre_text.find(kw)
                if kw_pos != -1:
                    sub_text = pre_text[kw_pos + len(kw):]
                    dash_idx = sub_text.find('-')
                    nl_idx = sub_text.find('\n')
                    indices = [idx for idx in [dash_idx, nl_idx] if idx != -1]
                    end_pos = min(indices) if indices else len(sub_text)
                    scope_text = sub_text[:end_pos]
                    break
                    
            if scope_text:
                scope_text = re.sub(r'\s+', ' ', scope_text).strip()
                num_pattern = r'(\d+(?:\.\d+)*(?:\,\d+)?)'
                
                verb_pattern = r'(?:\s*(?:se|duoc|nhan|them))*\s*'
                pattern_a = re.compile(
                    r'(\d+)(?:\s*\([^)]+\))?\s*(?:co phieu|trai phieu|chung chi quy)' + verb_pattern + num_pattern,
                    re.IGNORECASE
                )
                pattern_b = re.compile(
                    r'(\d+)(?:\s*\([^)]+\))?\s*(?:co phieu|trai phieu|chung chi quy)\s*[-–—]\s*' + num_pattern + r'\s*quyen bieu quyet',
                    re.IGNORECASE
                )
                pattern_c = re.compile(
                    r'(\d+)(?:\s*\([^)]+\))?\s*[:/]\s*' + num_pattern,
                    re.IGNORECASE
                )
                
                match = pattern_a.search(scope_text) or pattern_b.search(scope_text) or pattern_c.search(scope_text)
                if match:
                    # Sử dụng helper chuẩn của chúng ta để quy đổi thập phân nếu có
                    don_vi_huong_quyen, gia_tri_huong_quyen = normalize_ratio(match.group(0))

        # TyLeMenhGia
        ty_le_menh_gia = None
        if nhom_quyen == "Cổ tức tiền":
            if loai_quyen == "Cổ phiếu":
                scope_text = ""
                for kw in ["ty le thuc hien", "ty le thanh toan", "ti le thuc hien", "ti le thanh toan"]:
                    kw_pos = pre_text.find(kw)
                    if kw_pos != -1:
                        sub_text = pre_text[kw_pos + len(kw):]
                        dash_idx = sub_text.find('-')
                        plus_idx = sub_text.find('+')
                        nl_idx = sub_text.find('\n')
                        indices = [idx for idx in [dash_idx, plus_idx, nl_idx] if idx != -1]
                        end_pos = min(indices) if indices else len(sub_text)
                        scope_text = sub_text[:end_pos]
                        break
                        
                if scope_text:
                    percent_match = re.search(r'(\d+(?:\.\d+)*(?:\,\d+)?)\s*%', scope_text)
                    if percent_match:
                        num_str = percent_match.group(1)
                        num_str = num_str.replace('.', '').replace(',', '.')
                        ty_le_menh_gia = float(num_str)
                        
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
            num_match = re.search(r'(\d+(?:\.\d+)+)', target_sub)
            if num_match:
                gia_phat_hanh = num_match.group(1).replace('.', '')

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

        if nhom_quyen in ["Hoán đổi chuyển đổi", "Khai báo chứng quyền", "Đăng ký Lưu ký", "Thay đổi"]:
            noi_dung = None
        elif nhom_quyen == "Tin huỷ":
            noi_dung = loai_quyen
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

        is_special = 0
        if nhom_quyen in ["Hoán đổi chuyển đổi", "Khai báo chứng quyền", "Đăng ký Lưu ký", "Tin huỷ", "Thay đổi"] or ';' in title:
            is_special = 1

        res = dict(record)
        pub_at = res.pop('published_at', None) or res.pop('published_date', None) or res.pop('date', None)
        if pub_at and ' ' not in str(pub_at).strip():
            pub_at = f"{str(pub_at).strip()} 00:00:00"
            
        coll_at = res.pop('collected_at', None) or res.pop('collected_date', None)
        if coll_at and ' ' not in str(coll_at).strip():
            coll_at = f"{str(coll_at).strip()} 00:00:00"
            
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
            'is_special': is_special,
            'is_4llm': 0 # Không xử lý bằng LLM
        })
        return res

    def parse_date(self, date_string):
        try:
            return datetime.strptime(date_string, '%d/%m/%Y').date()
        except:
            return None

    def parse_datetime(self, datetime_string):
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
        code = str(record.get('code') or record.get('MaChungKhoan') or '').strip()
        url = str(record.get('url') or '').strip()
        title = str(record.get('title') or record.get('TieuDe') or record.get('NoiDung') or '').strip()

        hash_content = f"{url}_{title}"
        hash_suffix = hashlib.md5(hash_content.encode('utf-8')).hexdigest()[:8]

        if code:
            record_id = f"{code}_{hash_suffix}"
        else:
            record_id = f"rec_{hash_suffix}"

        if split_idx is not None:
            record_id = f"{record_id}_{split_idx}"

        return record_id

    def get_vptoken(self):
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
        pattern = f"{field_label}[:\\s]+([^\\n]+(?:\\n\\s*[+\\-•]\\s*[^\\n]+)*)"
        match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
        if match:
            extracted = match.group(1).strip()
            if len(extracted) > max_length:
                extracted = extracted[:max_length] + "..."
            return extracted if extracted else None
        return None

    def extract_field_bullets(self, text, field_label):
        pattern = f"{field_label}[:\\s]+([^\\n]+(?:\\n\\s*[+\\-•]\\s*[^\\n]+)*)"
        match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
        if not match:
            return None
        extracted = match.group(1).strip()
        bullet_pattern = r'[+\-•]\s*([^\n]+)'
        bullets = re.findall(bullet_pattern, extracted)
        if bullets:
            return [b.strip() for b in bullets if b.strip()]
        else:
            return [extracted] if extracted else None

    def contains_keyword(self, text, keywords):
        text_lower = text.lower()
        for keyword in keywords:
            if keyword.lower() in text_lower:
                return True
        return False

    def extract_quyền_values(self, text, value_keywords_map):
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
        Mở URL bài viết, crawl nội dung thô và parse bằng Regex truyền thống.
        Phương thức này chỉ đóng vai trò lấy text_content và là fallback khi LLM không được kích hoạt.
        """
        try:
            max_retries = 3
            response = None
            for attempt in range(max_retries):
                response = self.session.get(url, headers=self.headers, timeout=10)
                response.encoding = 'utf-8'
                if response.status_code == 200:
                    break
                if attempt < max_retries - 1:
                    time.sleep(0.2)

            if response is None or response.status_code != 200:
                return None, None, None

            soup = BeautifulSoup(response.content, 'html.parser')
            main = soup.find('main') or soup.find('article') or soup.find('div', class_='main-content') or soup.find('div', class_='content')
            if not main:
                body = soup.find('body')
                if not body:
                    return None, None, None
                text_content = body.get_text()
            else:
                text_content = main.get_text()

            cutoff_markers = [
                'tin cùng tổ chức',
                'mã ck hủy đăng ký',
                'mã ck chuyển sàn',
                'thành viên đã thu hồi'
            ]
            min_cutoff = len(text_content)
            for marker in cutoff_markers:
                pos = text_content.lower().find(marker)
                if pos > 0:
                    min_cutoff = min(min_cutoff, pos)
            if min_cutoff < len(text_content):
                text_content = text_content[:min_cutoff]

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
                'quyền_họp_đại_hội_cổ_đông': None,
                'quyền_cổ_tức_tiền': None,
                'quyền_cổ_tức_cổ_phiếu': None,
                'quyền_mua': None,
                'quyền_hoán_đổi_chuyển_đổi': None,
                'chứng_quyền': None,
                'chấp_thuận_đăng_ký': None,
                'tin_húy': None,
                'thay_đổi': None,
                'text_content': None
            }

            extracted_code = None
            label_divs = soup.find_all('div', class_='col-md-4')

            if not label_divs:
                code_match = re.search(r'^([A-Z0-9]{6,})\s*:', text_content.strip(), re.MULTILINE)
                if code_match:
                    extracted_code = code_match.group(1)
                name_match = re.search(r'Tên chứng khoán[:\s]+([^\n]+)', text_content, re.IGNORECASE)
                if name_match:
                    info['tên_chứng_khoán'] = name_match.group(1).strip()
                code_match2 = re.search(r'Mã chứng khoán[:\s]+([A-Z0-9]+)', text_content, re.IGNORECASE)
                if code_match2:
                    info['mã_chứng_khoán'] = code_match2.group(1).strip()
                    extracted_code = code_match2.group(1)
                isin_match = re.search(r'Mã ISIN[:\s]+([A-Z0-9]+)', text_content, re.IGNORECASE)
                if isin_match:
                    info['mã_isin'] = isin_match.group(1).strip()
                org_match = re.search(r'(?:Tổng Công ty|Công ty cổ phần|CTCP|Ngân hàng)[^\n]+(?:thông báo|khai báo)', text_content)
                if org_match:
                    org_text = org_match.group(0)
                    org_name_match = re.search(r'(?:Tổng Công ty|Công ty cổ phần|CTCP|Ngân hàng)([^\-]+)', org_text)
                    if org_name_match:
                        info['tên_tổ_chức_đăng_ký'] = ('Tổng Công ty' if 'Tổng' in org_text else 'Công ty cổ phần') + org_name_match.group(1).strip()

            for label_div in label_divs:
                label = label_div.get_text(strip=True).lower()
                value_div = label_div.find_next('div', class_='col-md-8')
                if not value_div:
                    continue
                value = value_div.get_text(strip=True)

                if 'tên tổ chức đăng ký' in label or 'tên tcđkck' in label or 'tcđkck' in label:
                    info['tên_tổ_chức_đăng_ký'] = value
                elif 'tên chứng khoán' in label:
                    info['tên_chứng_khoán'] = value
                elif 'mã chứng khoán' in label or 'mã ck' in label:
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

            if not info['tỷ_lệ_thực_hiện']:
                info['tỷ_lệ_thực_hiện'] = self.extract_field_from_text(text_content, 'Tỷ lệ thực hiện', max_length=1000)
            if not info['thời_gian_thực_hiện']:
                info['thời_gian_thực_hiện'] = self.extract_field_from_text(text_content, 'Thời gian thực hiện', max_length=300)
            if not info['địa_điểm_thực_hiện']:
                info['địa_điểm_thực_hiện'] = self.extract_field_from_text(text_content, 'Địa điểm thực hiện', max_length=500)
            if not info['lý_do_mục_đích']:
                info['lý_do_mục_đích'] = self.extract_field_from_text(text_content, 'Lý do|Mục đích', max_length=300)
            if not info['mệnh_giá']:
                info['mệnh_giá'] = self.extract_field_from_text(text_content, 'Mệnh giá', max_length=100)
            if not info['tên_tổ_chức_đăng_ký']:
                org_pattern = r'(?:Tên tổ chức đăng ký chứng khoán|Tên TCĐKCK)[:\s]+([^\n]+)'
                org_match = re.search(org_pattern, text_content, re.IGNORECASE)
                if org_match:
                    extracted_org = org_match.group(1).strip()
                    if extracted_org and extracted_org != '--':
                        info['tên_tổ_chức_đăng_ký'] = extracted_org

            # Trích xuất 9 trường phân loại phụ cho mục đích hiển thị/dashboard cũ
            title_tag = soup.find('title')
            search_text = text_content + (" " + title_tag.get_text() if title_tag else "")
            
            dhdc_map = {
                'Quyền đại hội cổ đông thường niên': ['đại hội đồng cổ đông thường niên', 'đại hội cổ đông thường niên', 'đại hội thường niên', 'đhđcđ thường niên', 'agm', 'annual general meeting'],
                'Quyền lấy ý kiến cổ đông bằng văn bản': ['lấy ý kiến cổ đông bằng văn bản', 'ý kiến bằng văn bản', 'written opinion'],
                'Quyền đại hội cổ đông bất thường': ['đại hội đồng cổ đông bất thường', 'đại hội cổ đông bất thường', 'đại hội bất thường', 'egm', 'extraordinary general meeting']
            }
            info['quyền_họp_đại_hội_cổ_đông'] = self.extract_quyền_values(search_text, dhdc_map)

            dividend_cash_map = {
                'Chi trả cổ tức bằng tiền': ['chi trả cổ tức bằng tiền', 'cổ tức tiền', 'dividend cash'],
                'Thanh toán lãi trái phiếu': ['thanh toán lãi', 'lãi trái phiếu', 'bond interest', 'interest payment'],
                'Thanh toán gốc, lãi': ['thanh toán gốc', 'trả gốc', 'principal payment', 'maturity payment'],
                'Mua lại trái phiếu trước hạn': ['mua lại trái phiếu', 'early redemption', 'buyback']
            }
            info['quyền_cổ_tức_tiền'] = self.extract_quyền_values(search_text, dividend_cash_map)

            dividend_share_map = {
                'Trả cổ tức bằng cổ phiếu': ['trả cổ tức bằng cổ phiếu', 'cổ tức cổ phiếu', 'stock dividend'],
                'Phát hành cổ phiếu': ['phát hành cổ phiếu', 'share issuance', 'cổ phiếu thưởng', 'bonus shares']
            }
            info['quyền_cổ_tức_cổ_phiếu'] = self.extract_quyền_values(search_text, dividend_share_map)

            purchase_map = {
                'Thực hiện quyền mua Trái phiếu chuyển đổi': ['quyền mua trái phiếu chuyển đổi', 'conversion bond purchase', 'convertible bond exercise'],
                'Thực hiện quyền mua cổ phiếu': ['quyền mua cổ phiếu', 'quyền mua', 'right issue', 'subscription right']
            }
            info['quyền_mua'] = self.extract_quyền_values(search_text, purchase_map)

            swap_map = {
                'Hoán đổi cổ phiếu': ['hoán đổi cổ phiếu', 'swap shares', 'cổ phiếu hoán đổi'],
                'Chuyển đổi trái phiếu': ['chuyển đổi trái phiếu', 'convertible bond', 'bond conversion']
            }
            info['quyền_hoán_đổi_chuyển_đổi'] = self.extract_quyền_values(search_text, swap_map)

            warrant_map = {'Có': ['chứng quyền', 'warrant', 'call warrant', 'put warrant']}
            info['chứng_quyền'] = self.extract_quyền_values(search_text, warrant_map)

            approval_map = {'Đăng ký cổ phiếu, trái phiếu': ['đăng ký cổ phiếu', 'đăng ký trái phiếu', 'registration approval', 'chấp thuận đăng ký']}
            info['chấp_thuận_đăng_ký'] = self.extract_quyền_values(search_text, approval_map)

            cancellation_map = {
                'Hủy ngày đăng ký cuối cùng': ['hủy ngày đăng ký', 'cancel registration date'],
                'Hủy danh sách người sở hữu chứng khoán': ['hủy danh sách người sở hữu', 'hủy danh sách người sử hữu', 'hủy danh sách', 'cancel ownership list', 'cancel list'],
                'Hủy đăng ký chứng khoán, trái phiếu': ['hủy đăng ký', 'huỷ', 'delisting', 'deregistration']
            }
            info['tin_húy'] = self.extract_quyền_values(search_text, cancellation_map)

            change_map = {
                'Thay đổi thời gian thanh toán': ['thay đổi thời gian thanh toán', 'thay đổi ngày thanh toán', 'payment date change'],
                'Chuyển dữ liệu đăng ký (chuyển sàn)': ['chuyển dữ liệu', 'chuyển sàn', 'data transfer', 'transfer between exchanges']
            }
            info['thay_đổi'] = self.extract_quyền_values(search_text, change_map)

            actual_update_datetime = None
            update_match = re.search(r'Cập nhật ngày\s+(\d{1,2}/\d{1,2}/\d{4})(?:\s*-\s*(\d{1,2}:\d{1,2}:\d{1,2}))?', text_content)
            if update_match:
                date_str = update_match.group(1)
                time_str = update_match.group(2) if update_match.group(2) else "00:00:00"
                datetime_str = f"{date_str} {time_str}"
                actual_update_datetime = self.parse_datetime(datetime_str)

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

            text_content = '\n'.join(line.strip() for line in text_content.split('\n') if line.strip())
            info['text_content'] = text_content if text_content else None

            return info, extracted_code, actual_update_datetime
        except Exception as e:
            logger.debug(f"  ! Error extracting detail: {str(e)[:50]}")
            return None, None, None

    def merge_records(self, new_records, old_records):
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
        
        def parse_date_str(date_str):
            if not date_str:
                return datetime.min
            for fmt in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
                try:
                    return datetime.strptime(str(date_str).strip(), fmt)
                except ValueError:
                    continue
            return datetime.min
            
        for url, new_list in url_to_new.items():
            old_list = url_to_old.get(url, [])
            if old_list:
                for r_old in old_list:
                    target = None
                    if len(new_list) == 1:
                        target = new_list[0]
                    else:
                        old_id = r_old.get('_record_id')
                        for r_new in new_list:
                            if r_new.get('_record_id') == old_id:
                                target = r_new
                                break
                        if not target:
                            p_title = str(r_old.get('title') or '').strip().lower()
                            for r_new in new_list:
                                new_title = str(r_new.get('title') or '').strip().lower()
                                if new_title == p_title or new_title.split(':', 1)[-1].strip() == p_title.split(':', 1)[-1].strip():
                                    target = r_new
                                    break
                        if not target:
                            target = new_list[0]
                            
                    for flag in ['status', 'confirmation_status', 'is_completed', 'is_special', 'is_4llm']:
                        if flag in r_old and r_old[flag] is not None:
                            if flag == 'status' and r_old[flag] == 'pending':
                                continue
                            if flag == 'confirmation_status' and r_old[flag] == 'awaiting_review':
                                continue
                            target[flag] = r_old[flag]
                            
            merged.extend(new_list)
            
        for url, old_list in url_to_old.items():
            if url in url_to_new:
                continue
            if len(old_list) == 1:
                merged.append(old_list[0])
                continue
                
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
                for r_non in non_prefixed:
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
                            
                    for flag in ['status', 'confirmation_status', 'is_completed', 'is_special', 'is_4llm']:
                        if flag in r_non and r_non[flag] is not None:
                            if flag == 'status' and r_non[flag] == 'pending':
                                continue
                            if flag == 'confirmation_status' and r_non[flag] == 'awaiting_review':
                                continue
                            target[flag] = r_non[flag]
                merged.extend(ticker_prefixed)
            else:
                old_list_sorted = sorted(old_list, key=lambda x: parse_date_str(x.get('collected_at') or x.get('published_at')), reverse=True)
                merged.append(old_list_sorted[0])
                
        return merged

    def fetch_latest_news(self):
        """
        Crawl các trang tin tức VSD.
        Tích hợp Hybrid Rẽ nhánh:
        - Nếu là tin nhiều quyền -> Gọi LLM + Post-process bằng Python.
        - Nếu là tin đơn giản -> Chạy Regex tĩnh truyền thống.
        """
        try:
            logger.info(f"🔍 VSD: Crawling tin tức thị trường cơ sở (multiple pages)...")
            filtered_news = []
            latest_date_found = datetime.now(VN_TZ).date()
            
            if self.mode == 'excel_urls':
                excel_file = globals().get('EXCEL_URLS_FILE', 'url_2fetch_round2.xlsx')
                if not os.path.exists(excel_file):
                    parent_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), excel_file)
                    if os.path.exists(parent_path):
                        excel_file = parent_path
                    else:
                        raise FileNotFoundError(f"Không tìm thấy file Excel danh sách URL: {excel_file}")
                
                logger.info(f"  📊 Đang đọc danh sách URL từ file: {excel_file}")
                df_urls = pd.read_excel(excel_file)
                
                url_col = None
                for col in df_urls.columns:
                    if 'url' in str(col).lower():
                        url_col = col
                        break
                if url_col is None:
                    url_col = df_urls.columns[0]
                
                raw_urls = df_urls[url_col].dropna().astype(str).tolist()
                urls_to_crawl = []
                seen_urls = set()
                for u in raw_urls:
                    u_clean = u.strip()
                    if u_clean.startswith('http') and u_clean not in seen_urls:
                        seen_urls.add(u_clean)
                        urls_to_crawl.append(u_clean)
                
                logger.info(f"  Found {len(urls_to_crawl)} unique URLs to process from Excel")
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
                all_news = []
                page = 1
                latest_date_found = None
                max_pages = 25
                
                today = datetime.now(VN_TZ).date()
                if self.mode == 'date_range':
                    cutoff_date = self.date_from
                else:
                    cutoff_date = today - timedelta(days=self.keep_days)
                
                logger.info(f"  📅 Cutoff date (oldest date to keep): {cutoff_date}")
                
                while page <= max_pages:
                    logger.info(f"  📄 Crawling page {page}...")
                    try:
                        if page == 1:
                            vptoken = self.get_vptoken()
                            if not vptoken:
                                logger.error("  ✗ Cannot get VPToken, stopping")
                                break

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

                            if not re.search(r'[A-Z0-9]{2,10}:', title):
                                continue

                            match = re.search(r'([A-Z0-9]{2,10}):', title)
                            if not match:
                                continue
                            code = match.group(1)

                            if not url.startswith('http'):
                                url = self.base_url + url

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
                    return {'status': 'not_found', 'message': 'Không tìm thấy tin trên VSD'}

                if self.mode == 'date_range':
                    filtered_news = [n for n in all_news if n['date_obj'] and self.date_from <= n['date_obj'] <= self.date_to]
                    logger.info(f"  ✓ Lọc khoảng ngày: {len(filtered_news)} tin từ {self.date_from} đến {self.date_to} (crawled {page-1} pages)")
                else:
                    min_keep_date = latest_date_found - timedelta(days=self.keep_days - 1)
                    filtered_news = [n for n in all_news if n['date_obj'] and n['date_obj'] >= min_keep_date]
                    logger.info(f"  ✓ Tìm thấy {len(filtered_news)} tin từ {min_keep_date} đến {latest_date_found} (crawled {page-1} pages)")

            logger.info(f"  🔗 Extracting details từ tất cả {len(filtered_news)} records (concurrent)...")
            result_data = []

            def extract_with_retry(news):
                max_retries = 2
                for attempt in range(max_retries):
                    try:
                        detail, extracted_code, actual_update_datetime = self.extract_detail_from_article(news['url'])
                        final_code = extracted_code if extracted_code else news['code']

                        if actual_update_datetime:
                            published_datetime = actual_update_datetime
                        elif news['date_obj']:
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

                        # Build basic item
                        result_item = {
                            'code': final_code,
                            'title': detail.get('title') if detail and detail.get('title') else news['title'],
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
                            
                            return {
                                'code': news['code'],
                                'title': news['title'],
                                'url': news['url'],
                                'published_at': published_at,
                                'collected_at': collected_at,
                                'source': 'VSD',
                                'status': 'pending'
                            }

            # Lấy chi tiết thô của tất cả tin tức (Đa luồng)
            raw_details = []
            with ThreadPoolExecutor(max_workers=5) as executor:
                futures = []
                for idx, news in enumerate(filtered_news):
                    future = executor.submit(extract_with_retry, news)
                    futures.append((future, news['code']))
                    if idx % 10 == 0:
                        time.sleep(0.05)

                for future, code in futures:
                    try:
                        raw_details.append(future.result())
                    except Exception as e:
                        logger.error(f"Future error {code}: {str(e)[:30]}")

            # --- RẼ NHÁNH HYBRID: DÙNG LLM CHO TIN NHIỀU QUYỀN ---
            logger.info("  ⚡ Hybrid check: Identifying multi-rights articles to invoke LLM...")
            
            for item in raw_details:
                title = item.get('title') or ''
                text_content = item.get('text_content') or ''
                ticker = item.get('code') or 'N/A'
                
                # Gọi hàm kiểm tra rẽ nhánh dựa trên title và text_content
                if is_multi_rights_article(title, text_content):
                    logger.info(f"    [LLM Mode] 🤖 Article for Ticker '{ticker}' has MULTIPLE rights. Invoking LLM...")
                    
                    # Gọi LLM trích xuất danh sách các quyền
                    llm_extractions = self.extract_multiple_purposes_via_llm(text_content)
                    
                    if llm_extractions:
                        logger.info(f"    ✓ LLM found {len(llm_extractions)} separate rights for Ticker '{ticker}'. Applying Python post-processing...")
                        # Map từng quyền thô từ LLM thành record nghiệp vụ đầy đủ
                        for idx, ext_raw in enumerate(llm_extractions, 1):
                            processed_record = self.process_llm_extracted_record(ext_raw, item)
                            # Tạo record ID duy nhất
                            processed_record['_record_id'] = self.generate_record_id(processed_record, split_idx=idx)
                            result_data.append(processed_record)
                    else:
                        logger.warning(f"    ⚠ LLM failed to extract rights for Ticker '{ticker}'. Falling back to Regex split.")
                        # Fallback về split thô như cũ nếu LLM lỗi
                        self._fallback_regex_split(item, result_data)
                else:
                    # Bản tin đơn giản chỉ có 1 quyền: Chạy Regex tĩnh để tối ưu
                    logger.info(f"    [Regex Mode] ⚡ Article for Ticker '{ticker}' has a SINGLE right. Applying regex rules.")
                    processed_record = self.apply_business_rules(item)
                    processed_record['_record_id'] = self.generate_record_id(processed_record)
                    processed_record['is_4llm'] = 0
                    result_data.append(processed_record)

            logger.info(f"  ✓ Hoàn thành extract chi tiết từ {len(result_data)} tin")

            # Merge với records cũ để tránh duplicate
            merged_data = result_data
            total_count = len(result_data)

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

                    if isinstance(existing_data, dict) and 'records' in existing_data:
                        existing_records = existing_data.get('records', [])
                    else:
                        existing_records = existing_data if isinstance(existing_data, list) else []

                    logger.info(f"  📚 Found {len(existing_records)} existing records, merging...")
                    merged_data = self.merge_records(result_data, existing_records)
                    logger.info(f"  ✓ Merged by URL and ticker-prefix rule: {len(result_data)} new + {len(merged_data) - len(result_data)} old = {len(merged_data)} total")
                    total_count = len(merged_data)
                except Exception as e:
                    logger.error(f"  ✗ Error merging records: {str(e)}")
                    merged_data = result_data

            if self.mode == 'excel_urls':
                final_filtered = []
                for r in merged_data:
                    pub_at_str = r.get('published_at') or r.get('published_date') or r.get('date')
                    if pub_at_str:
                        try:
                            date_part = pub_at_str.split(' ')[0]
                            item_date = datetime.strptime(date_part, '%d/%m/%Y').date()
                            if self.date_from <= item_date <= self.date_to:
                                final_filtered.append(r)
                            else:
                                logger.info(f"    ⚠️ Lọc bỏ {r.get('code') or r.get('url')} vì ngày đăng {pub_at_str} nằm ngoài khoảng lọc {DATE_FROM} - {DATE_TO}")
                        except ValueError:
                            final_filtered.append(r)
                    else:
                        final_filtered.append(r)
                merged_data = final_filtered
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
            return {'status': 'error', 'message': 'Request timeout'}
        except Exception as e:
            logger.error(f"  ✗ VSD Error: {str(e)[:100]}")
            import traceback
            logger.error(traceback.format_exc())
            return {'status': 'error', 'message': str(e)}

    def _fallback_regex_split(self, result_item, result_data):
        """Fallback split bằng regex tĩnh cũ khi LLM lỗi ở tin nhiều quyền"""
        lý_do = result_item.get('lý_do_mục_đích')
        if lý_do and isinstance(lý_do, str) and ';' in lý_do:
            purposes = [p.strip() for p in lý_do.split(';') if p.strip()]
            if len(purposes) > 1:
                text_content = result_item.get('text_content')
                sections = {}
                if text_content:
                    pattern = r'^\d+\.\s+(.+?)(?=\n\d+\.\s+|\Z)'
                    matches = re.findall(pattern, text_content, re.MULTILINE | re.DOTALL)
                    if len(matches) == len(purposes):
                        sections = {i+1: match for i, match in enumerate(matches)}

                for idx, purpose in enumerate(purposes, 1):
                    purpose_item = dict(result_item)
                    purpose_item['lý_do_mục_đích'] = purpose
                    if idx in sections:
                        section_text = sections[idx]
                        tỷ_lệ_match = re.search(r'Tỷ lệ thực hiện[:\s]+([^\n]+(?:\n-\s+[^\n]+)*)', section_text, re.IGNORECASE)
                        if tỷ_lệ_match:
                            purpose_item['tỷ_lệ_thực_hiện'] = tỷ_lệ_match.group(1).strip()[:1000]

                    code = result_item.get('code', 'N/A')
                    purpose_item['title'] = f"{code}: {purpose}"
                    
                    processed_record = self.apply_business_rules(purpose_item)
                    processed_record['_record_id'] = self.generate_record_id(processed_record, split_idx=idx)
                    processed_record['is_4llm'] = 0
                    result_data.append(processed_record)
                    logger.info(f"    ✓ [Fallback Regex Split] Split {code} by purpose: [{idx}] {purpose[:60]}")
                return
                
        # Nếu không split được
        processed_record = self.apply_business_rules(result_item)
        processed_record['_record_id'] = self.generate_record_id(processed_record)
        processed_record['is_4llm'] = 0
        result_data.append(processed_record)

    def save_to_excel(self, data, output_path):
        if not EXCEL_AVAILABLE:
            return {'status': 'error', 'message': 'pandas hoặc openpyxl chưa được cài đặt'}

        try:
            new_records = data.get('data', [])
            if not new_records:
                return {'status': 'error', 'message': 'Không có dữ liệu để export'}

            output_dir = os.path.dirname(output_path)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir, exist_ok=True)

            final_records = list(new_records)
            new_count = len(new_records)

            if os.path.exists(output_path):
                try:
                    df_old = pd.read_excel(output_path, sheet_name='Tin chứng khoán')
                    df_old = df_old.where(pd.notnull(df_old), None)
                    
                    old_records = []
                    for _, row in df_old.iterrows():
                        old_record = row.to_dict()
                        if 'title' not in old_record or not old_record.get('title'):
                            old_record['title'] = old_record.get('NoiDung') or f"{old_record.get('MaChungKhoan')}: Tin chứng khoán"
                        
                        old_rid = old_record.get('_record_id') or self.generate_record_id(old_record)
                        old_record['_record_id'] = old_rid
                        
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

            df = pd.DataFrame(final_records)

            # Cập nhật STANDARD_COLUMNS có chứa thêm cột is_4llm
            STANDARD_COLUMNS = [
                'published_at', 'collected_at', 'url', 'text_content', 'MaChungKhoan',
                'TieuDe',
                'NhomQuyen', 'LoaiQuyen', 'MaISIN', 'MaTrongNuoc', 'NgayChot',
                'NgayGDKHQ', 'NgayThucHien', 'UocTinhNgayThucHien', 'NgayThanhToan',
                'CNQuyenMuaTuNgay', 'CNQuyenMuaDenNgay', 'DKQuyenMuaTuNgay', 'DKQuyenMuaDenNgay',
                'DonViHuongQuyen', 'GiaTriHuongQuyen', 'TyLeMenhGia', 'GiaPhatHanh',
                'NoiDung', 'is_completed', 'is_special', 'is_4llm',
                '_record_id', 'title'
            ]

            for col in STANDARD_COLUMNS:
                if col not in df.columns:
                    df[col] = None
            df = df[STANDARD_COLUMNS]

            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Tin chứng khoán', index=False, startrow=0)
                workbook = writer.book
                worksheet = writer.sheets['Tin chứng khoán']

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
                return {'status': 'error', 'message': 'Excel file was not created'}
        except Exception as e:
            logger.error(f"✗ Error creating Excel: {str(e)}")
            return {'status': 'error', 'message': f'Error creating Excel: {str(e)}'}


def main():
    fetcher = VSDFetcher()
    logger.info(f"Starting VSD Hybrid LLM fetch with KEEP_DAYS={KEEP_DAYS}")
    result = fetcher.fetch_latest_news()

    if result.get('status') == 'success' and result.get('data'):
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
            result['excel_info'] = excel_result

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
                records_for_html = []
                if 'excel_info' in result and result['excel_info'].get('status') == 'success':
                    try:
                        excel_file = result['excel_info'].get('file')
                        if excel_file and os.path.exists(excel_file):
                            df_excel = pd.read_excel(excel_file, sheet_name='Tin chứng khoán')
                            all_records = df_excel.to_dict('records')
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
                    all_records = result.get('data', [])

                for record in all_records:
                    html_record = dict(record)
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

                json_output = {
                    'status': result.get('status'),
                    'date': result.get('date'),
                    'records': records_for_html,
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
