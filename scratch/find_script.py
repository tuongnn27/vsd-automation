with open('index.html', 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "function normalizeRecord" in line:
        print(f"normalizeRecord: Line {i+1}")
    if "function loadDataFromFile" in line:
        print(f"loadDataFromFile: Line {i+1}")
    if "function initializeTable" in line:
        print(f"initializeTable: Line {i+1}")
    if "function updateStats" in line:
        print(f"updateStats: Line {i+1}")
    if "function copyAllRecordDetails" in line:
        print(f"copyAllRecordDetails: Line {i+1}")
    if "function renderModalField" in line:
        print(f"renderModalField: Line {i+1}")
    if "function enableInlineEdit" in line:
        print(f"enableInlineEdit: Line {i+1}")
    if "function applySmartHighlighting" in line:
        print(f"applySmartHighlighting: Line {i+1}")
    if "function showModal" in line:
        print(f"showModal: Line {i+1}")
    if "function filterTable" in line:
        print(f"filterTable: Line {i+1}")
    if "function filterByStatus" in line:
        print(f"filterByStatus: Line {i+1}")
