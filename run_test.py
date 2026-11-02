import json, sys
sys.path.insert(0, "/home/claude/memoire-formatter")
from formatter import format_docx
data = open(sys.argv[1], "rb").read()
out, rep = format_docx(data, {}, with_pages=True)
open(sys.argv[2], "wb").write(out)
print(json.dumps(rep, ensure_ascii=False, indent=1)[:3500])
