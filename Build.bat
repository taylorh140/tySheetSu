cargo build --release
python stuber.py "./target/wasm32-unknown-unknown/release/tySheetSu.wasm" "tySheetSu.wasm"
move tySheetSu.wasm ./package
del tempfile.wat
pause