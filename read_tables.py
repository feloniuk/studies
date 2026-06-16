import sys, os
sys.stdout.reconfigure(encoding='utf-8')
from docx import Document

base = 'D:\\feloniuk\\studies\\2_\u0422\u0430\u0439\u043c \u043c\u0435\u043d\u0435\u0434\u0436\u043c\u0435\u043d\u0442'
fpath = os.path.join(base, '\u041c\u043e\u0434\u0443\u043b\u044c 5_\u0406\u043d\u0434\u0438\u0432\u0456\u0434\u0443\u0430\u043b\u044c\u043d\u0435 \u0437\u0430\u0432\u0434\u0430\u043d\u043d\u044f', '\u041c\u043e\u0434\u0443\u043b\u044c 5_\u041f\u0440\u0430\u043a\u0442\u0438\u0447\u043d\u0435 \u0437\u0430\u0432\u0434\u0430\u043d\u043d\u044f 1_\u0421\u043a\u043b\u0430\u0434\u0430\u043d\u043d\u044f \u0445\u0440\u043e\u043d\u043e\u043a\u0430\u0440\u0442\u0438.docx')
doc = Document(fpath)
print("TABLES:")
for i, table in enumerate(doc.tables):
    print(f"\n--- TABLE {i+1} ({len(table.rows)} rows x {len(table.columns)} cols) ---")
    for j, row in enumerate(table.rows):
        cells = [cell.text.strip() for cell in row.cells]
        print(f"  Row {j}: {' | '.join(cells)}")
    if i > 3:
        print("... (more tables)")
        break
