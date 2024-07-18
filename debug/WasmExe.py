from pathlib import Path
import wasmtime.loader
import potato_sheets

import typst_env
# egg.excel_to_json(bb,b"Sheet1")

print(dir(egg))

ssname = b"Sheet1"
bb =  Path("Book.xlsx").read_bytes()


egg.excel_to_json(*typst_env.typstArgs(bb,ssname))