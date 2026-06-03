with open('index.html', 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "function normalizeRecord" in line:
        print(f"normalizeRecord: Line {i+1}")
    if "function initializeTable" in line:
        print(f"initializeTable: Line {i+1}")
    if "// Initialize table on load" in line:
        print(f"Initialize table load: Line {i+1}")
    if "function filterTable" in line:
        print(f"filterTable: Line {i+1}")
    if "function filterByStatus" in line:
        print(f"filterByStatus: Line {i+1}")
