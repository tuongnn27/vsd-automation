with open('scripts/fetch_vsd.py', 'r', encoding='utf-8') as f:
    content = f.read()

import re
match = re.search(r'STANDARD_COLUMNS\s*=\s*\[(.*?)\]', content, re.DOTALL)
if match:
    with open('scratch/columns.txt', 'w', encoding='utf-8') as out:
        out.write(match.group(0))
    print("STANDARD_COLUMNS written to scratch/columns.txt!")
else:
    print("STANDARD_COLUMNS not found!")
