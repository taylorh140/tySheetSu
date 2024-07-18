from pathlib import Path
import re
import sys
import subprocess


# Check if the correct number of arguments is provided
if len(sys.argv) != 3:
    print("Usage: ./Patch.sh Input_wasm Output_wasm")
    sys.exit(1)

# Retrieve the command-line arguments
arg1 = sys.argv[1]
arg2 = sys.argv[2]

# Your Python script logic here, using arg1 and arg2
print("Argument 1:", arg1)
print("Argument 2:", arg2)


result = subprocess.run(['wasm2wat', arg1], stdout=subprocess.PIPE)
wat = result.stdout.decode('utf-8') #Path('Pre-patched-dot.wat').read_text()

types={}
exports = {}
imports = {}
replacments = {}

for _ in re.findall(r"\s*\(type \(;(\d+);\) \(func ?(\(param ([^\)]+)\))? ?(\(result ([^\)]+)\))?\)\)",wat):
    types[_[0]]=(_[1],_[3],_[4])

for _ in re.findall(r"(\(import ([^\(]+)\(func \(;(\d+);\) \(type (\d+)\)\)\))",wat):
    fimport,fname,fnum,ftype = _
    fname=fname.strip()
    params,rtn_statement,rtn_type = types[ftype]
    imports[fname]=dict(fimport=fimport,fname=fname,fnum=fnum,ftype=ftype,params=params,rtn_statement=rtn_statement,rtn_type=rtn_type)
    print("check sane",imports[fname])

    rep=f"(func (;{fnum};) {params} {rtn_statement} {'' if rtn_type=='' else rtn_type+'.const 0 return '} )"

    if "\"typst_env\"" in fimport:
        continue

    if fimport in wat:
        replacments[fimport]=rep
    #wat=wat.replace(_[0],rep)

for _ in re.findall(r"(\(export ([^\(]+)\(func (\d+)\)\))",wat):
    fexport,fname,fnum = _
    fname=fname.strip()
    fmatch = fname.replace("___",'" "')
    print("EXPORT:",fname,fnum,fmatch)
    exports[fmatch]=dict(fexport=fexport,fname=fname,fnum=fnum)

    print(fmatch)
    for key in list(replacments.keys()):
        if fmatch in key:
            imp = imports[fmatch]
            #replacments[imp['fimport']]="" 
            replacments[fexport]=""
            replacments[f"call {imp['fnum']}\\b"]=f"call {fnum}"


for x,y in replacments.items():
    print("replacing:",x,"with",y)
    if "\\b" in x:
        print('SUBD')
        wat = re.sub(x,y,wat)
    else:
        wat=wat.replace(x,y)

Path('tempfile.wat').write_text(wat)

subprocess.run(['wat2wasm', 'tempfile.wat' , '-o',arg2], stdout=subprocess.PIPE)