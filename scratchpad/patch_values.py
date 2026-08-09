"""Evaluate the workbook's formulas in Python and bake cached <v> values into the
xlsx XML (LibreOffice can't run in this sandbox). Supports the small formula subset
used by build_xlsx.py: + - * /, cell/range refs, cross-sheet 'Sheet'!ref, SUM, SUMIF,
IF, ABS, comparisons, string literals."""
import re, sys, zipfile, shutil, os
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string

PATH = sys.argv[1]
wb = load_workbook(PATH, data_only=False)
memo = {}

TOKEN_RE = re.compile(r"""
    \s*(?:
      (?P<str>"(?:[^"]*)")
    | (?P<sheetref>'[^']*'!\$?[A-Za-z]+\$?[0-9]+)
    | (?P<range>\$?[A-Za-z]+\$?[0-9]+:\$?[A-Za-z]+\$?[0-9]+)
    | (?P<cell>\$?[A-Za-z]+\$?[0-9]+)
    | (?P<func>[A-Za-z]+)\(
    | (?P<num>[0-9]*\.?[0-9]+)
    | (?P<op><=|>=|<>|[+\-*/<>=(),])
    )
""", re.VERBOSE)

def tokenize(s):
    toks=[]; i=0
    while i < len(s):
        if s[i].isspace(): i+=1; continue
        m=TOKEN_RE.match(s,i)
        if not m: raise ValueError(f"bad token at {s[i:]!r} in {s!r}")
        i=m.end()
        for k,v in m.groupdict().items():
            if v is not None:
                toks.append((k,v)); break
    return toks

def split_ref(ref):
    """'Sheet'!H3 or H3 -> (sheet_or_None, coord_without_$)"""
    if ref.startswith("'"):
        j=ref.index("'",1); sheet=ref[1:j]; coord=ref[j+2:]
    else:
        sheet=None; coord=ref
    return sheet, coord.replace("$","")

def cell_value(sheet, coord):
    key=(sheet,coord)
    if key in memo: return memo[key]
    memo[key]=0  # cycle guard
    c=wb[sheet][coord]; v=c.value
    if v is None: r=0
    elif isinstance(v,(int,float)): r=v
    elif isinstance(v,str):
        if v.startswith("="): r=eval_expr(sheet, tokenize(v[1:]), [0])
        else: r=v
    else: r=v
    memo[key]=r
    return r

def expand_range(sheet, rng):
    a,b=rng.replace("$","").split(":")
    ca=re.match(r"([A-Za-z]+)([0-9]+)",a); cb=re.match(r"([A-Za-z]+)([0-9]+)",b)
    c1=column_index_from_string(ca.group(1)); r1=int(ca.group(2))
    c2=column_index_from_string(cb.group(1)); r2=int(cb.group(2))
    cells=[]
    for rr in range(min(r1,r2),max(r1,r2)+1):
        for cc in range(min(c1,c2),max(c1,c2)+1):
            cells.append(f"{get_column_letter(cc)}{rr}")
    return [cell_value(sheet, co) for co in cells]

# recursive-descent over token list; pos is [int] mutable
def eval_expr(sheet, toks, pos):
    return parse_cmp(sheet, toks, pos)

def parse_cmp(sheet, toks, pos):
    left=parse_add(sheet,toks,pos)
    while pos[0]<len(toks) and toks[pos[0]]==('op',) : break
    while pos[0]<len(toks) and toks[pos[0]][0]=='op' and toks[pos[0]][1] in ('<','>','=','<=','>=','<>'):
        op=toks[pos[0]][1]; pos[0]+=1
        right=parse_add(sheet,toks,pos)
        if op=='<': left = left<right
        elif op=='>': left = left>right
        elif op=='=': left = left==right
        elif op=='<=': left = left<=right
        elif op=='>=': left = left>=right
        elif op=='<>': left = left!=right
    return left

def parse_add(sheet,toks,pos):
    left=parse_mul(sheet,toks,pos)
    while pos[0]<len(toks) and toks[pos[0]][0]=='op' and toks[pos[0]][1] in ('+','-'):
        op=toks[pos[0]][1]; pos[0]+=1
        right=parse_mul(sheet,toks,pos)
        left = left+right if op=='+' else left-right
    return left

def parse_mul(sheet,toks,pos):
    left=parse_unary(sheet,toks,pos)
    while pos[0]<len(toks) and toks[pos[0]][0]=='op' and toks[pos[0]][1] in ('*','/'):
        op=toks[pos[0]][1]; pos[0]+=1
        right=parse_unary(sheet,toks,pos)
        left = left*right if op=='*' else (left/right if right!=0 else 0)
    return left

def parse_unary(sheet,toks,pos):
    if pos[0]<len(toks) and toks[pos[0]]==('op','-'):
        pos[0]+=1; return -parse_unary(sheet,toks,pos)
    return parse_primary(sheet,toks,pos)

def parse_args(sheet,toks,pos):
    args=[]
    if pos[0]<len(toks) and toks[pos[0]]==('op',')'):
        pos[0]+=1; return args
    while True:
        # a raw range token as an arg
        if toks[pos[0]][0]=='range':
            args.append(('range',toks[pos[0]][1])); pos[0]+=1
        else:
            args.append(('val', parse_cmp(sheet,toks,pos)))
        t=toks[pos[0]]
        if t==('op',','): pos[0]+=1; continue
        if t==('op',')'): pos[0]+=1; break
        raise ValueError(f"expected , or ) got {t}")
    return args

def parse_primary(sheet,toks,pos):
    k,v=toks[pos[0]]
    if k=='num': pos[0]+=1; return float(v)
    if k=='str': pos[0]+=1; return v[1:-1]
    if k=='sheetref':
        pos[0]+=1; sh,co=split_ref(v); return cell_value(sh,co)
    if k=='cell':
        pos[0]+=1; return cell_value(sheet, v.replace("$",""))
    if k=='range':
        pos[0]+=1; return sum(expand_range(sheet,v))
    if k=='func':
        name=v.upper(); pos[0]+=1
        args=parse_args(sheet,toks,pos)
        if name=='SUM':
            tot=0
            for a in args:
                if a[0]=='range': tot+=sum(expand_range(sheet,a[1]))
                else: tot+=a[1]
            return tot
        if name=='ABS': return abs(args[0][1])
        if name=='IF':
            cond=args[0][1]; return args[1][1] if cond else args[2][1]
        if name=='SUMIF':
            crit_cells=expand_range(sheet,args[0][1]) if args[0][0]=='range' else [args[0][1]]
            criterion=args[1][1]
            sum_cells_rng=args[2][1] if args[2][0]=='range' else None
            # need actual cell coords for sum range aligned with crit range
            a,b=args[2][1].replace("$","").split(":")
            ca=re.match(r"([A-Za-z]+)([0-9]+)",a)
            scol=column_index_from_string(ca.group(1)); srow=int(ca.group(2))
            # crit range coords
            ka,kb=args[0][1].replace("$","").split(":")
            kca=re.match(r"([A-Za-z]+)([0-9]+)",ka); kcol=column_index_from_string(kca.group(1)); krow=int(kca.group(2))
            kcb=re.match(r"([A-Za-z]+)([0-9]+)",kb); kend=int(kcb.group(2))
            tot=0
            for idx,rr in enumerate(range(krow,kend+1)):
                cv=cell_value(sheet, f"{get_column_letter(kcol)}{rr}")
                if cv==criterion:
                    tot+=cell_value(sheet, f"{get_column_letter(scol)}{srow+idx}")
            return tot
        raise ValueError(f"unknown func {name}")
    if k=='op' and v=='(':
        pos[0]+=1; r=parse_cmp(sheet,toks,pos)
        assert toks[pos[0]]==('op',')'); pos[0]+=1; return r
    raise ValueError(f"unexpected token {k}:{v}")

# ---- evaluate every formula cell ----
results={}  # sheetname -> {coord: value}
for sh in wb.sheetnames:
    ws=wb[sh]; results[sh]={}
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value,str) and c.value.startswith("="):
                results[sh][c.coordinate]=cell_value(sh, c.coordinate)

# ---- map sheet name -> sheetN.xml ----
import xml.etree.ElementTree as ET
z=zipfile.ZipFile(PATH)
wbxml=z.read("xl/workbook.xml").decode()
rels=z.read("xl/_rels/workbook.xml.rels").decode()
ns={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
rid_to_target={}
for m in re.finditer(r'Id="([^"]+)"\s+Type="[^"]*worksheet"\s+Target="([^"]+)"', rels):
    rid_to_target[m.group(1)]=m.group(2)
# also handle attr order variations
for m in re.finditer(r'<Relationship\b[^>]*>', rels):
    tag=m.group(0)
    if 'worksheet' in tag:
        rid=re.search(r'Id="([^"]+)"',tag).group(1)
        tgt=re.search(r'Target="([^"]+)"',tag).group(1)
        rid_to_target[rid]=tgt
name_to_file={}
for m in re.finditer(r'<sheet\b[^>]*/?>', wbxml):
    tag=m.group(0)
    nm=re.search(r'name="([^"]+)"',tag).group(1)
    rid=re.search(r'r:id="([^"]+)"',tag).group(1)
    tgt=rid_to_target[rid].lstrip("/")
    if not tgt.startswith("xl/"): tgt="xl/"+tgt
    name_to_file[nm]=tgt
z.close()

def xesc(s):
    return s.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")

def fmt_num(v):
    if isinstance(v,bool): return None
    if abs(v-round(v))<1e-9:
        iv=round(v)
        return str(int(iv))
    return repr(round(v,10))

# ---- patch each sheet xml ----
tmp=PATH+".tmp"
zin=zipfile.ZipFile(PATH,'r')
zout=zipfile.ZipFile(tmp,'w',zipfile.ZIP_DEFLATED)
for item in zin.infolist():
    data=zin.read(item.filename)
    # find which sheet this is
    sheetname=None
    for nm,fn in name_to_file.items():
        if fn==item.filename: sheetname=nm; break
    if sheetname is not None:
        xml=data.decode()
        vals=results[sheetname]
        for coord,val in vals.items():
            # locate the <c r="coord" ...><f>...</f><v /></c>
            pat=re.compile(r'(<c r="'+coord+r'"[^>]*?)(>)(<f[^>]*>.*?</f>)<v\s*/>', re.DOTALL)
            if isinstance(val,str):
                repl_open=r'\1 t="str"\2'
                newv='<v>'+xesc(val)+'</v>'
            elif isinstance(val,bool):
                repl_open=r'\1 t="b"\2'
                newv='<v>'+('1' if val else '0')+'</v>'
            else:
                repl_open=r'\1\2'
                num=fmt_num(val)
                newv='<v>'+num+'</v>'
            def _r(m):
                return re.sub(r'\1\\2', '', '')  # placeholder
            xml2=pat.sub(lambda m: (m.group(1)+(' t="str"' if isinstance(val,str) else (' t="b"' if isinstance(val,bool) else ''))+m.group(2)+m.group(3)+newv), xml)
            xml=xml2
        # guard: any leftover empty <v /> on formula cells -> set 0 to avoid blanks
        data=xml.encode()
    zout.writestr(item, data)
zin.close(); zout.close()
os.replace(tmp, PATH)
print("patched sheets:", list(name_to_file.items()))
print("formula cells:", {s:len(v) for s,v in results.items()})
