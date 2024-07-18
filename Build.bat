cargo build
python stuber.py "./target/wasm32-unknown-unknown/debug/tySheetSu.wasm" "tySheetSu.wasm"
move tySheetSu.wasm ./package
del tempfile.wat
pause