import json

categories = {}

with open('VN/danhmuc.txt', 'r', encoding='utf-8') as f:
    for line in f:
        line = line.strip()
        if not line or '|' not in line:
            continue
        parts = line.split('|')
        parent = parts[0].strip()
        child = parts[1].strip()
        if parent not in categories:
            categories[parent] = []
        categories[parent].append(child)

with open('scratch_danhmuc.json', 'w', encoding='utf-8') as out:
    json.dump(categories, out, ensure_ascii=False, indent=4)
print("Parsed and saved to scratch_danhmuc.json")
