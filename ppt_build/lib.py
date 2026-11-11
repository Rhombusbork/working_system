from lxml import etree
import copy

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
ns = {'p': P, 'a': A, 'r': R}
E = 914400

WHITE = 'FFFFFF'
PANEL = 'F7F7F7'
BLACK = '000000'
GREY = '8C8C8C'
DIM = '595959'

FONT = 'Pretendard'
FONT_KW = 'Pretendard ExtraBold'
FONT_SB = 'Pretendard SemiBold'


def emu(v):
    return str(int(round(v * E)))


class Slide:
    def __init__(self, path):
        self.path = path
        self.tree = etree.parse(path)
        self.root = self.tree.getroot()
        self.sp = self.root.find('.//p:spTree', ns)
        self.next_id = max(int(e.get('id')) for e in self.root.iter('{%s}cNvPr' % P)) + 1

    def nid(self):
        i = self.next_id
        self.next_id += 1
        return i

    def get(self, name):
        for el in self.sp:
            c = el.find('.//p:cNvPr', ns)
            if c is not None and c.get('name') == name:
                return el
        raise KeyError(name)

    def sid(self, el):
        return int(el.find('.//p:cNvPr', ns).get('id'))

    def remove(self, *names):
        for n in names:
            el = self.get(n)
            el.getparent().remove(el)

    def move(self, name_or_el, x=None, y=None, w=None, h=None):
        el = self.get(name_or_el) if isinstance(name_or_el, str) else name_or_el
        xf = el.find('.//a:xfrm', ns)
        o = xf.find('a:off', ns)
        e = xf.find('a:ext', ns)
        if x is not None: o.set('x', emu(x))
        if y is not None: o.set('y', emu(y))
        if w is not None: e.set('cx', emu(w))
        if h is not None: e.set('cy', emu(h))
        return el

    # ---------- text ----------
    @staticmethod
    def _runs_xml(runs):
        """runs: list of paragraphs; each paragraph = (align, [ (text, size, color, font) ... ], lnspc)"""
        out = []
        for algn, rs, lnspc in runs:
            ppr = f'<a:pPr algn="{algn}"><a:lnSpc><a:spcPct val="{lnspc}"/></a:lnSpc></a:pPr>'
            rr = ''
            for (t, sz, col, font) in rs:
                t = t.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                rr += (f'<a:r><a:rPr lang="ko-KR" altLang="en-US" sz="{int(sz*100)}" b="0" i="0" dirty="0">'
                       f'<a:solidFill><a:srgbClr val="{col}"/></a:solidFill>'
                       f'<a:latin typeface="{font}"/><a:ea typeface="{font}"/><a:cs typeface="{font}"/>'
                       f'</a:rPr><a:t xml:space="preserve">{t}</a:t></a:r>')
            out.append(f'<a:p>{ppr}{rr}</a:p>')
        return ''.join(out)

    def text(self, x, y, w, h, paras, anchor='t', name=None, wrap=True):
        """paras: list of (align, runs, lnspc) or a shortcut string tuple."""
        i = self.nid()
        name = name or f'Text {i}'
        xml = (f'<p:sp xmlns:p="{P}" xmlns:a="{A}"><p:nvSpPr><p:cNvPr id="{i}" name="{name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
               f'<p:spPr><a:xfrm><a:off x="{emu(x)}" y="{emu(y)}"/><a:ext cx="{emu(w)}" cy="{emu(h)}"/></a:xfrm>'
               f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
               f'<p:txBody><a:bodyPr wrap="{"square" if wrap else "none"}" lIns="0" tIns="0" rIns="0" bIns="0" anchor="{anchor}" rtlCol="0"><a:noAutofit/></a:bodyPr><a:lstStyle/>'
               f'{self._runs_xml(paras)}</p:txBody></p:sp>')
        el = etree.fromstring(xml)
        self.sp.append(el)
        return el

    def t1(self, x, y, w, h, s, size, color=WHITE, font=FONT, align='l', anchor='t', lnspc=100000, name=None):
        lines = s.split('\n')
        paras = [(align, [(ln, size, color, font)], lnspc) for ln in lines]
        return self.text(x, y, w, h, paras, anchor=anchor, name=name)

    def retext(self, name, s, size, color=WHITE, font=FONT, align=None, anchor=None, lnspc=100000):
        """Replace text of an existing template text box keeping its geometry."""
        el = self.get(name)
        tb = el.find('p:txBody', ns)
        bp = tb.find('a:bodyPr', ns)
        if anchor: bp.set('anchor', anchor)
        for k in ('tIns', 'bIns', 'lIns', 'rIns'):
            bp.set(k, '0')
        oldp = tb.find('a:p', ns)
        al = align or (oldp.find('a:pPr', ns).get('algn') if oldp is not None and oldp.find('a:pPr', ns) is not None else 'l') or 'l'
        for p_ in tb.findall('a:p', ns):
            tb.remove(p_)
        frag = etree.fromstring(f'<x xmlns:a="{A}">' + self._runs_xml([(al, [(ln, size, color, font)], lnspc) for ln in s.split('\n')]) + '</x>')
        for p_ in frag:
            tb.append(p_)
        return el

    # ---------- shapes ----------
    def shape(self, prst, x, y, w, h, fill=None, line=None, lw=2.0, dash=None, rot=0, adj=None, name=None,
              text=None, size=40, tcolor=BLACK, font=FONT, anchor='ctr', flipH=False):
        i = self.nid()
        name = name or f'{prst} {i}'
        fillx = f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>' if fill else '<a:noFill/>'
        if line:
            d = f'<a:prstDash val="{dash}"/>' if dash else ''
            lnx = f'<a:ln w="{int(lw*12700)}"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill>{d}</a:ln>'
        else:
            lnx = '<a:ln><a:noFill/></a:ln>'
        av = ''.join(f'<a:gd name="{k}" fmla="val {v}"/>' for k, v in (adj or {}).items())
        rotx = f' rot="{int(rot*60000)}"' if rot else ''
        fl = ' flipH="1"' if flipH else ''
        if text is not None:
            body = (f'<p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="{anchor}" rtlCol="0"><a:noAutofit/></a:bodyPr><a:lstStyle/>'
                    + self._runs_xml([('ctr', [(ln, size, tcolor, font)], 100000) for ln in text.split('\n')]) + '</p:txBody>')
        else:
            body = ''
        xml = (f'<p:sp xmlns:p="{P}" xmlns:a="{A}"><p:nvSpPr><p:cNvPr id="{i}" name="{name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
               f'<p:spPr><a:xfrm{rotx}{fl}><a:off x="{emu(x)}" y="{emu(y)}"/><a:ext cx="{emu(w)}" cy="{emu(h)}"/></a:xfrm>'
               f'<a:prstGeom prst="{prst}"><a:avLst>{av}</a:avLst></a:prstGeom>{fillx}{lnx}</p:spPr>{body}</p:sp>')
        el = etree.fromstring(xml)
        self.sp.append(el)
        return el

    def line(self, x1, y1, x2, y2, color=WHITE, lw=2.0, dash=None, head=None, name=None):
        i = self.nid()
        name = name or f'Line {i}'
        fh = ' flipH="1"' if x2 < x1 else ''
        fv = ' flipV="1"' if y2 < y1 else ''
        x, y = min(x1, x2), min(y1, y2)
        w, h = abs(x2 - x1), abs(y2 - y1)
        d = f'<a:prstDash val="{dash}"/>' if dash else ''
        he = f'<a:tailEnd type="{head}" w="med" len="med"/>' if head else ''
        xml = (f'<p:cxnSp xmlns:p="{P}" xmlns:a="{A}"><p:nvCxnSpPr><p:cNvPr id="{i}" name="{name}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>'
               f'<p:spPr><a:xfrm{fh}{fv}><a:off x="{emu(x)}" y="{emu(y)}"/><a:ext cx="{emu(w)}" cy="{emu(h)}"/></a:xfrm>'
               f'<a:prstGeom prst="line"><a:avLst/></a:prstGeom>'
               f'<a:ln w="{int(lw*12700)}"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill>{d}{he}</a:ln></p:spPr></p:cxnSp>')
        el = etree.fromstring(xml)
        self.sp.append(el)
        return el

    def xmark(self, cx, cy, s, color, lw=2.0):
        h = s / 2
        return [self.line(cx - h, cy - h, cx + h, cy + h, color, lw), self.line(cx - h, cy + h, cx + h, cy - h, color, lw)]

    def plus(self, cx, cy, s, color, lw=2.0):
        h = s / 2
        return [self.line(cx - h, cy, cx + h, cy, color, lw), self.line(cx, cy - h, cx, cy + h, color, lw)]

    def to_front(self, el):
        self.sp.remove(el)
        self.sp.append(el)

    def group(self, els, name=None):
        """Wrap elements (already in spTree) into a p:grpSp so they animate together."""
        i = self.nid()
        name = name or f'Group {i}'
        xs, ys, xe, ye = [], [], [], []
        for el in els:
            xf = el.find('.//a:xfrm', ns)
            o, e = xf.find('a:off', ns), xf.find('a:ext', ns)
            x, y, w, h = int(o.get('x')), int(o.get('y')), int(e.get('cx')), int(e.get('cy'))
            xs.append(x); ys.append(y); xe.append(x + w); ye.append(y + h)
        X, Y, W, H = min(xs), min(ys), max(xe) - min(xs), max(ye) - min(ys)
        g = etree.fromstring(
            f'<p:grpSp xmlns:p="{P}" xmlns:a="{A}"><p:nvGrpSpPr><p:cNvPr id="{i}" name="{name}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            f'<p:grpSpPr><a:xfrm><a:off x="{X}" y="{Y}"/><a:ext cx="{W}" cy="{H}"/><a:chOff x="{X}" y="{Y}"/><a:chExt cx="{W}" cy="{H}"/></a:xfrm></p:grpSpPr></p:grpSp>')
        pos = min(list(self.sp).index(el) for el in els)
        for el in els:
            self.sp.remove(el)
            g.append(el)
        self.sp.insert(pos, g)
        return g

    # ---------- animation ----------
    def builds(self, steps):
        """steps: list of lists of elements; each list appears on one click (Appear effect)."""
        for t in self.root.findall('p:timing', ns):
            self.root.remove(t)
        cid = [2]

        def n():
            cid[0] += 1
            return cid[0]
        pars = ''
        for step in steps:
            outer, inner = n(), n()
            eff = ''
            for k, el in enumerate(step):
                spid = self.sid(el)
                a, b = n(), n()
                kind = 'clickEffect' if k == 0 else 'withEffect'
                eff += (f'<p:par><p:cTn id="{a}" presetID="1" presetClass="entr" presetSubtype="0" fill="hold" nodeType="{kind}">'
                        f'<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:set><p:cBhvr>'
                        f'<p:cTn id="{b}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
                        f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
                        f'</p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set></p:childTnLst></p:cTn></p:par>')
            pars += (f'<p:par><p:cTn id="{outer}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>'
                     f'<p:par><p:cTn id="{inner}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>{eff}</p:childTnLst></p:cTn></p:par>'
                     f'</p:childTnLst></p:cTn></p:par>')
        xml = (f'<p:timing xmlns:p="{P}"><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
               f'<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>{pars}</p:childTnLst></p:cTn>'
               f'<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
               f'<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>'
               f'</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>')
        t = etree.fromstring(xml)
        clr = self.root.find('p:clrMapOvr', ns)
        clr.addnext(t)

    def save(self):
        self.tree.write(self.path, xml_declaration=True, encoding='UTF-8', standalone=True)
