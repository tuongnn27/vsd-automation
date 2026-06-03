import json
import os
import re

def embed_data():
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    json_path = os.path.join(base_dir, 'data', 'vsd_records.json')
    html_paths = [
        os.path.join(base_dir, 'index.html'),
        os.path.join(base_dir, 'web', 'vps_automation_vhck.html'),
        os.path.join(base_dir, 'docs', 'index.html')
    ]

    if not os.path.exists(json_path):
        print(f"Error: {json_path} not found")
        return

    print(f"Reading data from {json_path}...")
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    json_str = json.dumps(data, ensure_ascii=False)
    new_line = f'      window.EMBEDDED_DATA = {json_str};'
    pattern = r'^\s*window\.EMBEDDED_DATA\s*=\s*\{.*?\};'
    pattern_flexible = r'window\.EMBEDDED_DATA\s*=\s*\{.*?\};'

    for html_path in html_paths:
        if not os.path.exists(html_path):
            print(f"Warning: {html_path} not found, skipping")
            continue

        print(f"Reading HTML from {html_path}...")
        with open(html_path, 'r', encoding='utf-8') as f:
            html_content = f.read()

        new_html = re.sub(pattern, lambda _: new_line, html_content, flags=re.MULTILINE)

        if new_html == html_content:
            print(f"Warning: No changes made to HTML {os.path.basename(html_path)}. Pattern might not have matched. Trying flexible pattern...")
            new_html = re.sub(pattern_flexible, lambda _: new_line, html_content)
            
        print(f"Writing updated HTML to {html_path}...")
        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(new_html)

        print(f"Successfully embedded data into {os.path.basename(html_path)}")

if __name__ == '__main__':
    embed_data()

