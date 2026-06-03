#!/usr/bin/env python3
"""
Recover corrupted VSD records containing "[OOPS!]" error page text.
This script scans data/vsd_records.json for corrupted records, re-crawls their URLs
with proper rate-limiting and retry logic, and merges the restored data back into
both JSON and Excel files.
"""

import json
import os
import sys
import time
from datetime import datetime

# Add parent directory to path to enable imports
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from fetch_vsd import VSDFetcher

def main():
    json_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data', 'vsd_records.json')
    excel_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data', 'vsd_records.xlsx')
    
    if not os.path.exists(json_path):
        print(f"✗ File JSON không tồn tại tại {json_path}")
        return

    # Load existing data
    with open(json_path, 'r', encoding='utf-8') as f:
        database = json.load(f)
    
    records = database.get('records', [])
    print(f"📚 Đã tải {len(records)} bản ghi từ {json_path}")

    # Find corrupted records
    corrupted_records = []
    for idx, r in enumerate(records):
        text = str(r.get('text_content') or '')
        if "[OOPS!]" in text or "Bài viết không có" in text:
            corrupted_records.append((idx, r))

    total_corrupted = len(corrupted_records)
    if total_corrupted == 0:
        print("✓ Không tìm thấy bản ghi nào bị lỗi [OOPS!]. Database hoàn toàn sạch!")
        return

    print(f"⚠️ Phát hiện {total_corrupted} bản ghi bị lỗi [OOPS!] (nội dung bị VSD chặn/trang lỗi).")
    print("🚀 Bắt đầu quá trình khôi phục tự động...")

    fetcher = VSDFetcher()
    recovered_count = 0
    fail_count = 0

    for idx, (original_index, record) in enumerate(corrupted_records, 1):
        url = record.get('url')
        code = record.get('code') or record.get('MaChungKhoan') or 'N/A'
        print(f"🔄 [{idx}/{total_corrupted}] Đang khôi phục mã {code} (URL: {url})...")

        try:
            # Re-fetch và parse detail
            detail, extracted_code, actual_update_date = fetcher.extract_detail_from_article(url)
            
            if detail and detail.get('text_content') and "[OOPS!]" not in detail.get('text_content'):
                # Cập nhật thông tin chi tiết vào record cũ
                record.update(detail)
                
                # Áp dụng lại business rules
                if extracted_code:
                    record['code'] = extracted_code
                record = fetcher.apply_business_rules(record)
                
                # Cập nhật ngược lại danh sách gốc
                records[original_index] = record
                recovered_count += 1
                print(f"  ✓ Khôi phục THÀNH CÔNG cho mã {record.get('MaChungKhoan')}!")
            else:
                fail_count += 1
                print(f"  ✗ Thất bại: VSD vẫn trả về trang lỗi hoặc nội dung trống.")
                
            # Sleep 1.0s giữa các request để đảm bảo lịch sự và tránh bị rate limit
            time.sleep(1.0)
            
        except Exception as e:
            fail_count += 1
            print(f"  ✗ Lỗi khi tải URL {url}: {e}")
            time.sleep(2.0)

    print(f"\n==================================================")
    print(f"🏁 Hoàn thành khôi phục:")
    print(f"  - Tổng số bản ghi bị lỗi: {total_corrupted}")
    print(f"  - Khôi phục THÀNH CÔNG: {recovered_count}")
    print(f"  - Thất bại (vẫn bị chặn): {fail_count}")
    print(f"==================================================")

    if recovered_count > 0:
        # Save updated JSON database
        database['records'] = records
        database['total_records'] = len(records)
        database['fetched_at'] = datetime.now().isoformat()
        
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(database, f, ensure_ascii=False, indent=2)
        print(f"✓ Đã lưu cơ sở dữ liệu JSON mới với các bản ghi được khôi phục tại: {json_path}")

        # Save to Excel
        print("📊 Đang cập nhật tệp Excel...")
        try:
            excel_result = fetcher.save_to_excel({'data': records}, excel_path)
            if excel_result.get('status') == 'success':
                print(f"✓ Đã lưu Excel thành công tại: {excel_path}")
            else:
                print(f"✗ Lỗi khi lưu Excel: {excel_result.get('message')}")
        except Exception as e:
            print(f"✗ Lỗi khi lưu Excel: {e}")

        # Update HTML Web Interface
        embed_script = os.path.join(os.path.dirname(__file__), 'embed_data.py')
        if os.path.exists(embed_script):
            print("🔄 Cập nhật giao diện web (HTML)...")
            try:
                import subprocess
                subprocess.run([sys.executable, embed_script], check=True)
                print("✓ Đã cập nhật giao diện web thành công!")
            except Exception as e:
                print(f"✗ Lỗi khi cập nhật giao diện web: {e}")
    else:
        print("ℹ Không có thay đổi nào được lưu vì không có bản ghi nào khôi phục thành công.")

if __name__ == '__main__':
    main()
