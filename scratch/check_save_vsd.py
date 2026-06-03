with open('scripts/fetch_vsd.py', 'r', encoding='utf-8') as f:
    content = f.read()

import re

# Find Excel and JSON export paths or references
for match in re.finditer(r'(xlsx|json|excel|to_excel|pandas|DataFrame)', content, re.IGNORECASE):
    start = max(0, match.start() - 50)
    end = min(len(content), match.end() + 50)
    snippet = content[start:end].replace('\n', ' ')
    print(f"Match at {match.start()}: ... {snippet} ...")
