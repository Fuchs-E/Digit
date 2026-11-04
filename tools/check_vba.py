"""Static checks on the VBA sources, run before publishing a release.

VBA only reports these as compile errors when a module is first used, which the
self-tests may never do (ribbon callbacks, for example), so check them here:

  1. A call to a procedure that is not defined (or is Private in another module).
  2. A declaration (Const, Dim, Declare, Type, Enum) after the first procedure.
  3. A local variable declared twice in one procedure, or with a parameter's name.
  4. An assignment to a name that is declared nowhere (Option Explicit error).

Names that VBA, Excel or MSForms provide are listed in BUILTINS. If the check
flags a real built-in, add it there.

Usage: python tools/check_vba.py     (exit code 1 when something is wrong)
"""
import glob, os, re, sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FILES = (sorted(glob.glob(os.path.join(ROOT, 'src', 'vba', '*.bas'))) +
         sorted(glob.glob(os.path.join(ROOT, 'src', '**', '*.code.vba'), recursive=True)) +
         [os.path.join(ROOT, 'lib', 'JsonConverter.bas'),
          os.path.join(ROOT, 'build', 'DigitFormBuilder.bas')])

BUILTINS = set("""
Abs Array Asc AscW Atn CBool CByte CCur CDate CDbl CDec CInt CLng CLngLng CLngPtr CSng CStr CVar CVErr
Chr ChrW Choose Command CreateObject CurDir Date DateAdd DateDiff DatePart DateSerial DateValue Day
Dir DoEvents Environ Environ$ EOF Error Err Exp FileDateTime FileLen Filter Fix Format Format$ FreeFile
GetObject GetSetting SaveSetting DeleteSetting Hex Hex$ Hour IIf InputBox InStr InStrB InStrRev Int IsArray IsDate IsEmpty
IsError IsMissing IsNull IsNumeric IsObject Join Kill LBound LCase LCase$ Left Left$ LeftB Len LenB
LoadPicture Loc LOF Log LTrim LTrim$ Mid Mid$ MidB Minute MkDir Month MsgBox Now Oct ObjPtr Replace
RGB Right Right$ RmDir Rnd Round RTrim RTrim$ Second Seek Sgn Shell Sin Space Space$ Split Sqr
StrComp StrConv String String$ StrPtr StrReverse Switch Tab Tan Time Timer TimeSerial TimeValue Trim
Trim$ TypeName UBound UCase UCase$ Val VarPtr VarType Weekday Year Debug Application ActiveCell
ActiveSheet ActiveWorkbook Selection ThisWorkbook Workbooks Worksheets Sheets Range Cells Columns Rows
Intersect Union Me Nothing Unload Load Beep ChDir ChDrive FileCopy Name Open Close Put Get Print
Write Input Line Randomize Reset SetAttr GetAttr Erase ReDim Preserve Stop End Exit Call Set Let
Dim Static Const If Then Else ElseIf For Each Next To Step Do Loop While Wend Until Select Case With
GoTo GoSub Return On Resume Not And Or Xor Mod Is Like New Private Public Friend Sub Function
Property Optional ByVal ByRef ParamArray As Boolean Byte Integer Long LongLong LongPtr Single Double
Currency Decimal Variant Object Collection Dictionary True False Empty Null Attribute Option
Explicit Compare Base Text Binary Module Declare PtrSafe Lib Alias Type Enum Global Event RaiseEvent
WithEvents Implements Rem Access Read Shared Lock Append Output Random Len
CallByName CVDate AppActivate SendKeys Calendar MacScript Partition IMEStatus
""".split())

STRING_RE = re.compile(r'"(?:[^"]|"")*"')
PROC_RE = re.compile(r'^\s*(?:(Public|Private|Friend)\s+)?(?:Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_]\w*)\s*(\(.*)?', re.I)
DECL_RE = re.compile(r'^((Public|Private|Global)\s+)?(Const|Dim|Declare|Type|Enum)\s', re.I)

def code_lines(path):
    """Logical lines with strings blanked and comments removed (continuations joined)."""
    out, buf, start = [], '', None
    for n, raw in enumerate(open(path, encoding='utf-8', errors='replace').read().split('\n'), 1):
        line = raw.rstrip('\r')
        s = STRING_RE.sub('""', line)
        q = s.find("'")
        if q >= 0: s = s[:q]
        if re.match(r'^\s*Rem\b', s, re.I): s = ''
        if start is None: start = n
        if s.rstrip().endswith(' _'):
            buf += s.rstrip()[:-1] + ' '
            continue
        out.append((start, raw, buf + s))
        buf, start = '', None
    return out

def main():
    public, private, names_in = set(), {}, {}
    parsed = {}
    for f in FILES:
        lines = code_lines(f); parsed[f] = lines
        priv, local = set(), set()
        in_type = False
        for n, raw, s in lines:
            m = PROC_RE.match(s)
            if m:
                (priv if (m.group(1) or '').lower() == 'private' else public).add(m.group(3).lower())
                for p in re.findall(r'(?:ByVal|ByRef|Optional|ParamArray)?\s*([A-Za-z_]\w*)\s+As\b', m.group(4) or ''):
                    local.add(p.lower())
                for p in re.findall(r'[(,]\s*(?:Optional\s+)?(?:ByVal\s+|ByRef\s+)?([A-Za-z_]\w*)', m.group(4) or ''):
                    local.add(p.lower())
                continue
            for kw in ('Dim', 'Const', 'Static', 'Private', 'Public', 'Global', 'ReDim'):
                mm = re.match(r'^\s*(?:Public\s+|Private\s+|Global\s+)?' + kw + r'\s+(.*)', s, re.I)
                if mm and not re.match(r'^\s*(Public|Private)\s+(Sub|Function|Property|Declare|Type|Enum|Const)\b', s, re.I):
                    for v in re.findall(r'(?:^|,)\s*(?:Preserve\s+)?(?:WithEvents\s+)?([A-Za-z_]\w*)', mm.group(1)):
                        local.add(v.lower())
            md = re.match(r'^\s*(?:Public\s+|Private\s+)?Declare\s+PtrSafe\s+(?:Sub|Function)\s+(\w+)', s, re.I)
            if md: public.add(md.group(1).lower())
            mc = re.match(r'^\s*(?:Public\s+|Private\s+|Global\s+)?Const\s+(\w+)', s, re.I)
            if mc: public.add(mc.group(1).lower())
            mt = re.match(r'^\s*(?:Public\s+|Private\s+)?(Type|Enum)\s+(\w+)', s, re.I)
            if mt: public.add(mt.group(2).lower()); in_type = True; continue
            if in_type:
                if re.match(r'^\s*End\s+(Type|Enum)', s, re.I): in_type = False
                else:
                    mi = re.match(r'^\s*(\w+)', s)
                    if mi: public.add(mi.group(1).lower())
            for lbl in re.findall(r'^\s*([A-Za-z_]\w*):(?!=)', s):
                local.add(lbl.lower())
        private[f] = priv; names_in[f] = local
    # VB_Name of each module (Module.Proc calls) and form control names (DigitFormBuilder)
    modules = set()
    for f in FILES:
        m = re.search(r'Attribute VB_Name = "(\w+)"', open(f, encoding='utf-8').read())
        if m: modules.add(m.group(1).lower())
    modules |= {'frmdigitsettings', 'scripting', 'msforms', 'vbide', 'office'}
    builder = open(os.path.join(ROOT, 'build', 'DigitFormBuilder.bas'), encoding='utf-8').read()
    controls = {c.lower() for c in re.findall(r'"((?:lbl|lst|btn|cbo|txt|img|chk|fra|mp)\w*?)"', builder)}
    controls |= {c.lower() + str(i) for c in ('cboSlot', 'txtKey', 'lblSlot') for i in range(1, 6)}
    builtins = {b.lower() for b in BUILTINS}

    problems = []
    for f in FILES:
        rel = os.path.relpath(f, ROOT)
        known = public | private[f] | names_in[f] | modules | builtins | controls
        in_procs = False
        for n, raw, s in parsed[f]:
            if PROC_RE.match(s): in_procs = True; continue
            if in_procs and DECL_RE.match(s) and not s.startswith((' ', '\t')):
                problems.append(f'{rel}:{n}: declaration after the first procedure: {raw.strip()}')
            if not in_procs: continue
            for stmt in re.split(r':(?!=)|\bThen\b|\bElse\b', s):
                for m in re.finditer(r'(?<![\w.!])([A-Za-z_]\w*\$?)\s*\(', stmt):
                    name = m.group(1)
                    if name.lower() not in known and name.rstrip('$').lower() not in known:
                        problems.append(f'{rel}:{n}: "{name}" is not defined anywhere: {raw.strip()}')
                ma = re.match(r'^\s*(?:Set\s+|Let\s+)?([A-Za-z_]\w*)\s*=', stmt)
                if ma and ma.group(1).lower() not in known:
                    problems.append(f'{rel}:{n}: "{ma.group(1)}" is assigned but never declared: {raw.strip()}')
                m = re.match(r'^\s*(?:Call\s+)?([A-Za-z_]\w*)(\s+[^=(]|\s*$)', stmt)
                if m and m.group(1).lower() not in known and not re.match(r'^\s*\w+\s*=', stmt):
                    problems.append(f'{rel}:{n}: "{m.group(1)}" is not defined anywhere: {raw.strip()}')
        # Duplicate local declarations within a procedure
        cur_proc, seen, stack = None, {}, []
        for n, raw, s in parsed[f]:
            m = PROC_RE.match(s)
            if m:
                cur_proc, seen, stack = m.group(3), {}, []
                for p in re.findall(r'(?:Optional\s+)?(?:ByVal\s+|ByRef\s+)?(?:ParamArray\s+)?([A-Za-z_]\w*)(?:\(\))?\s+As\b', m.group(4) or ''):
                    seen[p.lower()] = n
                continue
            if re.match(r'^\s*End\s+(Sub|Function|Property)\b', s, re.I):
                cur_proc = None; continue
            if cur_proc is None: continue
            # #If / #Else branches may declare the same name: each branch starts
            # from what was declared before the #If.
            if re.match(r'^\s*#If\b', s, re.I): stack.append(dict(seen)); continue
            if re.match(r'^\s*#Else(If)?\b', s, re.I) and stack: seen = dict(stack[-1]); continue
            if re.match(r'^\s*#End\s*If\b', s, re.I) and stack: stack.pop(); continue
            mm = re.match(r'^\s*(?:Dim|Static)\s+(.*)', s, re.I)
            if mm:
                for v in re.findall(r'(?:^|,)\s*(?:WithEvents\s+)?([A-Za-z_]\w*)', mm.group(1)):
                    if v.lower() in seen:
                        problems.append(f'{rel}:{n}: "{v}" is already declared in {cur_proc} (line {seen[v.lower()]}): {raw.strip()}')
                    else:
                        seen[v.lower()] = n
        # Private procedures of other modules
        for g in FILES:
            if g == f: continue
            for p in private[g] - public - names_in[f] - private[f]:
                for n, raw, s in parsed[f]:
                    if re.search(r'(?<![\w.])' + re.escape(p) + r'\b', s, re.I) and not PROC_RE.match(s):
                        problems.append(f'{rel}:{n}: "{p}" is Private in {os.path.basename(g)}: {raw.strip()}')
    for p in sorted(set(problems)): print(p)
    print(f'check_vba: {len(set(problems))} problem(s) in {len(FILES)} files')
    return 1 if problems else 0

if __name__ == '__main__':
    sys.exit(main())
