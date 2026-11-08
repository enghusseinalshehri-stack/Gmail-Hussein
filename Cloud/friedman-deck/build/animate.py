import json, re, sys, zipfile, shutil
src, dst = sys.argv[1], sys.argv[2]
D = json.load(open("anims.json")); anims, trans = D["anims"], D["trans"]

TRANS = {
 "fade": '<p:transition spd="med"><p:fade/></p:transition>',
 "push": '<p:transition spd="med"><p:push dir="l"/></p:transition>',
 "wipe": '<p:transition spd="med"><p:wipe dir="l"/></p:transition>',
}
def tgt(s): return f'<p:tgtEl><p:spTgt spid="{s}"/></p:tgtEl>'
class Ids:
    def __init__(s, n): s.n = n
    def __call__(s): s.n += 1; return s.n

def vis(i, s, val="visible"):
    return (f'<p:set><p:cBhvr><p:cTn id="{i()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
            f'{tgt(s)}<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>'
            f'<p:to><p:strVal val="{val}"/></p:to></p:set>')
def fade(i, s, d): return f'<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="{i()}" dur="{d}"/>{tgt(s)}</p:cBhvr></p:animEffect>'
def wipe(i, s, d, dr): return f'<p:animEffect transition="in" filter="wipe({dr})"><p:cBhvr><p:cTn id="{i()}" dur="{d}"/>{tgt(s)}</p:cBhvr></p:animEffect>'
def anim(i, s, d, attr, f, t):
    return (f'<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base"><p:cTn id="{i()}" dur="{d}" fill="hold"/>{tgt(s)}'
            f'<p:attrNameLst><p:attrName>{attr}</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst>'
            f'<p:tav tm="0"><p:val><p:strVal val="{f}"/></p:val></p:tav><p:tav tm="100000"><p:val><p:strVal val="{t}"/></p:val></p:tav></p:tavLst></p:anim>')

def effect(i, s, fx, d, node):
    cls, pid, sub = "entr", 10, 0
    if fx == "wipeL": pid, sub = 22, 8
    elif fx == "wipeR": pid, sub = 22, 2
    elif fx == "float": pid, sub = 42, 0
    elif fx == "zoom": pid, sub = 53, 16
    elif fx == "appear": pid, sub = 1, 0
    elif fx == "disappear": cls, pid, sub = "exit", 1, 0
    cid = i()
    if fx == "fade": body = vis(i, s) + fade(i, s, d)
    elif fx == "wipeL": body = vis(i, s) + wipe(i, s, d, "left")
    elif fx == "wipeR": body = vis(i, s) + wipe(i, s, d, "right")
    elif fx == "float": body = vis(i, s) + fade(i, s, d) + anim(i, s, d, "ppt_y", "#ppt_y+.08", "#ppt_y")
    elif fx == "zoom": body = vis(i, s) + anim(i, s, d, "ppt_w", "0", "#ppt_w") + anim(i, s, d, "ppt_h", "0", "#ppt_h") + fade(i, s, d)
    elif fx == "appear": body = vis(i, s)
    elif fx == "disappear": body = vis(i, s, "hidden")
    return (f'<p:par><p:cTn id="{cid}" presetID="{pid}" presetClass="{cls}" presetSubtype="{sub}" fill="hold" nodeType="{node}">'
            f'<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>{body}</p:childTnLst></p:cTn></p:par>')

def timing(effs, spid):
    i = Ids(3)
    groups = {}
    for e in effs: groups.setdefault(e["at"], []).append(e)
    inner = ""
    for at in sorted(groups):
        gid = i(); pars = ""
        for k, e in enumerate(groups[at]):
            pars += effect(i, spid[e["name"]], e["fx"], e["dur"], "afterEffect" if k == 0 else "withEffect")
        inner += f'<p:par><p:cTn id="{gid}" fill="hold"><p:stCondLst><p:cond delay="{at}"/></p:stCondLst><p:childTnLst>{pars}</p:childTnLst></p:cTn></p:par>'
    return ('<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
            '<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>'
            '<p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond></p:stCondLst><p:childTnLst>'
            f'{inner}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn>'
            '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
            '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>'
            '</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>')

zin = zipfile.ZipFile(src); zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
missing = []
for it in zin.infolist():
    data = zin.read(it.filename)
    m = re.fullmatch(r"ppt/slides/slide(\d+)\.xml", it.filename)
    if m:
        n = m.group(1); xml = data.decode("utf8")
        spid = {}
        for mm in re.finditer(r'<p:cNvPr id="(\d+)" name="([^"]*)"', xml): spid[mm.group(2)] = mm.group(1)
        texts = [v for k,v in spid.items() if re.fullmatch(r'Text \d+', k)]
        effs = anims.get(n, [])
        for e in effs:
            if e['name'] not in spid and texts: spid[e['name']] = texts[0]
        for e in effs:
            if e["name"] not in spid: missing.append((n, e["name"]))
        effs = [e for e in effs if e["name"] in spid]
        add = TRANS.get(trans.get(n, "fade"), TRANS["fade"]) + (timing(effs, spid) if effs else "")
        xml = xml.replace("</p:sld>", add + "</p:sld>")
        data = xml.encode("utf8")
    zout.writestr(it, data)
zout.close()
print("missing:", missing)
