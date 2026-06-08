with open(r'c:\Users\MSI\SourceCode\vps-automation-vhck\index.html', 'r', encoding='utf-8') as f:
    for i, line in enumerate(f, 1):
        if 'normalizeRecord' in line:
            print(f"{i}: {line.strip()}")
