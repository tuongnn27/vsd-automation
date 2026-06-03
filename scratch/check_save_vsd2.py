import re

with open('scripts/fetch_vsd.py', 'r', encoding='utf-8') as f:
    content = f.read()

findings = []

# Find all occurrences of Excel, XLSX, JSON saving or DataFrame creation
keywords = ['xlsx', 'json', 'excel', 'save', 'write', 'export', 'columns']
for kw in keywords:
    matches = list(re.finditer(re.escape(kw), content, re.IGNORECASE))
    findings.append(f"Keyword '{kw}' found {len(matches)} times.")
    for m in matches[:10]:
        start = max(0, m.start() - 60)
        end = min(len(content), m.end() + 60)
        snippet = content[start:end].replace('\n', ' ')
        findings.append(f"  - Position {m.start()}: ... {snippet} ...")

with open('scratch/findings.txt', 'w', encoding='utf-8') as f:
    f.write('\n'.join(findings))

print("Findings written to scratch/findings.txt successfully!")
