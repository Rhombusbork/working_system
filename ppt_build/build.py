import math, os, shutil, subprocess, sys
sys.path.insert(0, os.path.dirname(__file__))
from lib import *

W = os.path.dirname(os.path.abspath(__file__))
SRC, DST = f'{W}/deck', f'{W}/work'
shutil.rmtree(DST, ignore_errors=True)
shutil.copytree(SRC, DST)
S = lambda n: Slide(f'{DST}/ppt/slides/slide{n}.xml')

KW = 80          # 키워드 (20in 캔버스. 13.33in 기준 약 53pt → 아래 큰 글자와 함께 씀)
SRC_PT = 16      # 출처 (13.33in 기준 약 11pt)
SRC_ON_LIGHT = '6B6B6B'


def header(s, box, kw):
    el = s.retext(box, kw, KW, WHITE, FONT_KW, align='l', anchor='ctr')
    s.move(el, x=0.97, y=0.40, w=12.8, h=1.55)


def pill(s, box, tag):
    s.retext(box, tag, 26, WHITE, FONT_SB, align='ctr', anchor='ctr')


# ---------------- 1컷: 표지 (템플릿 1장) ----------------
s = S(1)
el = s.retext('TextBox 3', '린 스타트업', 150, WHITE, FONT_KW, align='l', anchor='t')
s.move(el, x=0.97, y=1.55, w=16.0, h=2.6)
s.remove('TextBox 7', 'TextBox 5')
s.retext('TextBox 4', '『린 스타트업』 10~14장', 20, BLACK, FONT, align='l', anchor='ctr')
s.retext('TextBox 6', '발표자 이름', 20, BLACK, FONT, align='r', anchor='ctr')
s.save()

# ---------------- 2컷: 질문 → 답 (템플릿 3장) ----------------
s = S(3)
s.remove('Picture 2', 'TextBox 3', 'TextBox 4', 'TextBox 5', 'Picture 8')
s.move('Picture 7', y=2.9)
s.shape('rect', 8.71, 2.9, 2.60, 2.60, text='?', size=110, tcolor=BLACK, font=FONT_KW, name='Question')
ans = s.retext('TextBox 6', '지속 가능한 성장', 130, WHITE, FONT_KW, align='ctr', anchor='ctr')
s.move(ans, x=0.5, y=6.1, w=19.0, h=2.4)
s.builds([[ans]])
s.save()

# ---------------- 3컷: 새 고객 vs 기존 고객 (템플릿 11장) ----------------
s = S(11)
s.remove('TextBox 3', 'TextBox 4', 'TextBox 7')
pill(s, 'TextBox 11', '성장 엔진')
el = s.retext('TextBox 6', '새 고객', 110, GREY, FONT_KW, align='l', anchor='t')
s.move(el, x=0.97, y=4.1, w=8.6, h=2.0)
el = s.retext('TextBox 5', '기존 고객', 130, BLACK, FONT_KW, align='l', anchor='t')
s.move(el, x=10.97, y=4.0, w=8.8, h=2.2)
s.retext('TextBox 14', 'VS', 26, WHITE, FONT_SB, align='ctr', anchor='ctr')
s.save()

# ---------------- 4컷: 성장 엔진 셋 (템플릿 9장) ----------------
s = S(9)
header(s, 'TextBox 3', '한 엔진에 집중')
pill(s, 'TextBox 7', '성장 엔진')
names = [('TextBox 20', '재방문'), ('TextBox 10', '바이럴'), ('TextBox 15', '유료')]
for box, nm in names:
    s.retext(box, nm, 56, BLACK, FONT_KW, align='ctr', anchor='ctr')
s.remove('TextBox 19')
s.retext('TextBox 9', '100만 명 / 6개월', 28, BLACK, FONT, align='ctr', anchor='t')
s.retext('TextBox 14', '획득 비용 < 생애 가치', 26, BLACK, FONT, align='ctr', anchor='t')
for b in ('TextBox 22', 'TextBox 12', 'TextBox 17'):
    s.get(b)  # 01 02 03 그대로
src = s.t1(6.96, 10.15, 6.10, 0.4, '핫메일 · 『린 스타트업』 220쪽', SRC_PT, SRC_ON_LIGHT, align='ctr')
g1 = s.group([s.get(n) for n in ('Picture 18', 'TextBox 20', 'Picture 21', 'TextBox 22')])
g2 = s.group([s.get(n) for n in ('Picture 8', 'TextBox 9', 'TextBox 10', 'Picture 11', 'TextBox 12')] + [src])
g3 = s.group([s.get(n) for n in ('Picture 13', 'TextBox 14', 'TextBox 15', 'Picture 16', 'TextBox 17')])
# 마지막 빌드: 원 하나만 검게 채운다(집중)
cx, cy, d = 10.01, 6.91, 6.0
fill = [s.shape('ellipse', cx - d / 2, cy - d / 2, d, d, fill=BLACK)]
fill.append(s.shape('ellipse', 9.42, 5.08, 1.15, 1.15, fill=PANEL, text='02', size=33, tcolor=BLACK, font=FONT_SB))
fill.append(s.t1(7.72, 6.43, 4.54, 0.96, '바이럴', 56, PANEL, FONT_KW, align='ctr', anchor='ctr'))
fill.append(s.t1(7.44, 7.42, 5.11, 0.6, '100만 명 / 6개월', 28, PANEL, FONT, align='ctr'))
g4 = s.group(fill)
s.builds([[g1], [g2], [g3], [g4]])
s.save()

# ---------------- 5컷: 속도 → 품질 (템플릿 18장) ----------------
s = S(18)
s.remove('TextBox 6', 'TextBox 11', 'TextBox 12', 'TextBox 13')
pill(s, 'TextBox 10', '품질')
el = s.retext('TextBox 5', '속도', 150, GREY, FONT_KW, align='l', anchor='t')
s.move(el, x=1.24, y=3.6, w=7.4, h=2.6)
el = s.retext('TextBox 4', '품질', 150, BLACK, FONT_KW, align='l', anchor='t')
s.move(el, x=10.61, y=3.6, w=7.4, h=2.6)
s.save()

# ---------------- 6컷: 다섯 번 왜 (템플릿 16장) ----------------
s = S(16)
for n in (2, 3, 4, 10, 11, 12, 13, 14):
    s.remove(f'TextBox {n}')
for n in range(17, 23):
    s.remove(f'Picture {n}')
header(s, 'TextBox 5', '다섯 번 왜')
pill(s, 'TextBox 9', '품질')
s.t1(0.97, 1.98, 9.0, 0.5, '새 버전 기능 고장', 26, GREY, FONT)
L, Rr, top, bot = 0.62, 19.40, 2.64, 10.61
cw = (Rr - L) / 5
mid = (top + bot) / 2
for k in range(1, 5):
    bx = L + cw * k
    s.line(bx - 0.3, top, bx + 0.3, mid, WHITE, 1.0)
    s.line(bx + 0.3, mid, bx - 0.3, bot, WHITE, 1.0)
answers = ['서버\n오류', '시스템\n오용', '사용법\n모름', '교육\n없음', '바쁜\n관리자']
steps = []
for k, a in enumerate(answers):
    x0 = L + cw * k + 0.55
    els = []
    last = (k == 4)
    if last:
        els.append(s.shape('rect', L + cw * k + 0.45, top + 0.02, Rr - (L + cw * k + 0.45), bot - top - 0.04, fill=PANEL))
    col_t = BLACK if last else WHITE
    col_l = SRC_ON_LIGHT if last else GREY
    els.append(s.t1(x0, 4.25, 3.0, 0.6, f'왜? {k + 1}', 28, col_l, FONT_SB))
    els.append(s.t1(x0, 5.05, 3.1, 2.4, a, 60, col_t, FONT_KW, lnspc=95000))
    steps.append([s.group(els)])
s.t1(0.62, 10.72, 8, 0.35, '『린 스타트업』 238~239쪽', SRC_PT, GREY)
s.builds(steps)
s.save()

# ---------------- 7컷: 혁신 샌드박스 (템플릿 7장) ----------------
s = S(7)
header(s, 'TextBox 2', '혁신 샌드박스')
pill(s, 'TextBox 6', '혁신')
s.remove('TextBox 9', 'TextBox 11', 'TextBox 12', 'Picture 7')
el = s.retext('TextBox 10', '모기업', 60, WHITE, FONT_KW, align='ctr', anchor='ctr')
s.move(el, x=2.67, y=5.68, w=6.0, h=1.2)
cx, cy, d = 14.34, 6.28, 6.26
s.shape('ellipse', cx - d / 2, cy - d / 2, d, d, line=GREY, lw=2.5, dash='dash', name='Sandbox')
sd = 3.1
s.shape('ellipse', cx - sd / 2, 4.15, sd, sd, fill=PANEL, text='혁신 팀', size=40, tcolor=BLACK, font=FONT_KW, name='Team')
s.t1(cx - 2.5, 7.55, 5.0, 0.6, '고객 5%', 32, WHITE, FONT_SB, align='ctr')
s.t1(cx - 2.5, 8.2, 5.0, 0.4, '예시', 20, GREY, FONT, align='ctr')
s.save()

# ---------------- 8컷: 인용 (템플릿 3장 복제) ----------------
s = S(21)
s.remove('TextBox 3', 'TextBox 4')
src = s.retext('TextBox 5', '피터 드러커 · 『린 스타트업』 283쪽', 22, GREY, FONT, align='ctr', anchor='ctr')
s.move(src, x=3.0, w=14.0)
q = s.get('TextBox 6')
q.getparent().remove(q)
G = GREY
s.text(1.0, 5.45, 18.0, 3.3, [
    ('ctr', [('전혀 해서는 안 될 일을', 64, G, FONT_KW)], 100000),
    ('ctr', [('매우 효율적으로 하는 것만큼', 64, G, FONT_KW)], 100000),
    ('ctr', [('무용한 짓', 64, WHITE, FONT_KW), ('은 확실히 없다', 64, G, FONT_KW)], 100000),
], anchor='t', name='Quote')
s.save()

# ---------------- 9컷: 주제 전환 (템플릿 20장) ----------------
s = S(20)
s.remove('TextBox 2', 'Picture 4', 'Picture 5')
el = s.retext('TextBox 3', '스타트업만의\n이야기일까?', 130, WHITE, FONT_KW, align='r', anchor='b', lnspc=95000)
s.move(el, x=5.0, y=4.3, w=14.15, h=6.5)
s.save()

# ---------------- 10컷: 작은 승리 (템플릿 5장) ----------------
s = S(5)
header(s, 'TextBox 5', '작은 승리')
pill(s, 'TextBox 9', '인생')
for n in (2, 3, 4, 10, 11, 12, 19, 20):
    s.remove(f'TextBox {n}')
s.t1(0.97, 1.98, 8.0, 0.4, 'Karl Weick, 1984', SRC_PT, GREY)
s.remove('Picture 15')
ctop, cbot = 3.35, 10.95
s.shape('rect', 0.45, ctop, 4.3, cbot - ctop, fill=PANEL, name='Win 1')
s.xmark(1.625, 9.625, 1.45, BLACK, 2.0)
icons = [('plus', 6.72, 9.44, 1.6), ('x', 11.625, 9.625, 1.45), ('plus', 16.73, 9.57, 1.6)]
steps = []
for k, (kind, ix, iy, sz) in enumerate(icons, start=1):
    els = [s.shape('rect', 5.0 * k + 0.25, ctop, 4.5, cbot - ctop, fill=PANEL, name=f'Win {k + 1}')]
    els += (s.plus(ix, iy, sz, BLACK, 2.0) if kind == 'plus' else s.xmark(ix, iy, sz, BLACK, 2.0))
    steps.append([s.group(els)])
s.builds(steps)
s.save()

# ---------------- 11컷: 피드백 사례 (템플릿 17장) ----------------
s = S(17)
header(s, 'TextBox 9', '피드백')
pill(s, 'TextBox 13', '인생')
s.remove('TextBox 4', 'TextBox 5', 'TextBox 6', 'TextBox 7', 'TextBox 8', 'Picture 3')
# 왼쪽: 선으로 그린 속도 표시 전광판
s.shape('roundRect', 2.85, 3.55, 5.35, 4.15, line=WHITE, lw=3, adj={'adj': 8000}, name='Sign')
s.t1(2.85, 3.95, 5.35, 0.8, 'YOUR SPEED', 44, WHITE, FONT_KW, align='ctr')
s.shape('rect', 3.65, 4.95, 3.75, 2.15, line=GREY, lw=2, name='Display')
s.t1(0.90, 8.15, 9.24, 0.7, '속도 표시 전광판', 34, WHITE, FONT_SB, align='ctr')
s.t1(0.90, 10.52, 9.24, 0.35, 'Wired, 2011', SRC_PT, GREY)
# 오른쪽: 만보기
r = [s.shape('roundRect', 10.86, 2.83, 8.24, 7.57, fill=PANEL, adj={'adj': 7000}, name='Card')]
r.append(s.t1(10.86, 4.55, 8.24, 2.0, '+2,000보', 120, BLACK, FONT_KW, align='ctr', anchor='ctr'))
r.append(s.t1(10.86, 8.15, 8.24, 0.7, '만보기 · 하루', 34, BLACK, FONT_SB, align='ctr'))
r.append(s.t1(10.86, 10.52, 8.24, 0.35, 'Bravata 외, 2007, JAMA', SRC_PT, GREY))
s.builds([[s.group(r)]])
s.save()

# ---------------- 12컷: 만들기-측정-학습 고리 (템플릿 8장) ----------------
s = S(8)
header(s, 'TextBox 2', '구조로 돌린다')
pill(s, 'TextBox 6', '인생')
for n in range(9, 22):
    try:
        s.remove(f'TextBox {n}')
    except KeyError:
        pass
for n in (12, 13, 14, 15):
    s.remove(f'Picture {n}')
CX, CY, RAD = 10.0, 6.35, 3.25
s.move('Picture 7', x=CX - RAD, y=CY - RAD, w=2 * RAD, h=2 * RAD)
el = s.retext('TextBox 8', 'IMVU', 40, GREY, FONT_SB, align='ctr', anchor='ctr')
s.move(el, x=CX - 2, y=CY - 0.4, w=4, h=0.8)


def loop_labels(s, CX, CY, RAD, labels, size, heads=True):
    angs = [-90, 30, 150]
    for (t, w), a in zip(labels, angs):
        x = CX + RAD * math.cos(math.radians(a))
        y = CY + RAD * math.sin(math.radians(a))
        s.shape('rect', x - w / 2, y - 0.55, w, 1.1, fill=BLACK, text=t, size=size, tcolor=WHITE, font=FONT_KW)
    if heads:
        for a in (-30, 90, 210):
            x = CX + RAD * math.cos(math.radians(a))
            y = CY + RAD * math.sin(math.radians(a))
            s.shape('triangle', x - 0.24, y - 0.22, 0.48, 0.44, fill=WHITE, rot=(a + 180) % 360)


loop_labels(s, CX, CY, RAD, [('만들기', 3.2), ('측정', 2.4), ('학습', 2.4)], 48)
s.t1(CX - 5, 10.45, 10, 0.35, '『린 스타트업』 194~195쪽', SRC_PT, GREY, align='ctr')
s.save()

# ---------------- 13컷: 한계 1/3 (템플릿 17장 복제) ----------------
s = S(22)
header(s, 'TextBox 9', '한계')
pill(s, 'TextBox 13', '인생')
s.remove('Picture 3', 'TextBox 5', 'TextBox 6', 'TextBox 7', 'TextBox 8')
cx0 = (20 - 9.24) / 2
s.move('Picture 2', x=cx0)
el = s.retext('TextBox 4', '1/3', 240, GREY, FONT_KW, align='ctr', anchor='ctr')
s.move(el, x=cx0, y=3.4, w=9.24, h=4.0)
s.t1(cx0, 7.75, 9.24, 0.7, '성과 하락', 36, WHITE, FONT_SB, align='ctr')
s.t1(cx0, 10.52, 9.24, 0.35, 'Kluger & DeNisi, 1996', SRC_PT, GREY, align='ctr')
s.save()

# ---------------- 14컷: AI 비서 (템플릿 8장 복제) ----------------
s = S(23)
header(s, 'TextBox 2', '목표-피드백 루프')
pill(s, 'TextBox 6', '실천')
for n in (9, 11, 17, 19, 21):
    s.remove(f'TextBox {n}')
el = s.retext('TextBox 8', 'AI 비서', 52, WHITE, FONT_KW, align='ctr', anchor='ctr')
s.move(el, x=2.25, y=5.73, w=5.38, h=1.0)
for box, t, y in (('TextBox 10', '작은 작업 단위', 3.62), ('TextBox 16', '작은 프로젝트', 5.25),
                  ('TextBox 18', '자동 피드백', 6.87), ('TextBox 20', '주의사항', 8.5)):
    el = s.retext(box, t, 38, WHITE, FONT_SB, align='l', anchor='ctr')
    s.move(el, x=12.45, y=y - 0.4, w=6.8, h=0.8)
# 마지막 빌드: 비서 전체를 감싸는 고리
cx, cy, rr = 4.94, 6.23, 3.45
ring = [s.shape('ellipse', cx - rr, cy - rr, 2 * rr, 2 * rr, line=WHITE, lw=2.5, dash='dash')]
for a in (-90, 90):
    x = cx + rr * math.cos(math.radians(a))
    y = cy + rr * math.sin(math.radians(a))
    ring.append(s.shape('triangle', x - 0.26, y - 0.24, 0.52, 0.48, fill=WHITE, rot=(a + 180) % 360))
ring.append(s.t1(cx - 3.5, 9.95, 7.0, 0.5, '만들기 → 측정 → 학습', 26, GREY, FONT_SB, align='ctr'))
s.builds([[s.group(ring)]])
s.save()

# ---------------- 15컷: 감사합니다 (템플릿 20장 복제) ----------------
s = S(24)
s.remove('TextBox 2', 'Picture 4', 'Picture 5')
el = s.retext('TextBox 3', '감사합니다', 170, WHITE, FONT_KW, align='r', anchor='b')
s.move(el, x=4.5, y=6.6, w=14.65, h=3.6)
s.save()

out = f'{W}/린스타트업_발표.pptx'
if os.path.exists(out):
    os.remove(out)
subprocess.run(['zip', '-XrDq', out, '.'], cwd=DST, check=True)
print('wrote', out)
