"""폰으로 보는 발표 대본 PDF (장마다: 슬라이드 그림 + 대본 + 그 장의 예상 질문 요약).

준비: 슬라이드를 PNG로 렌더링한 폴더(slide01.png ~ slide22.png)
실행: python make_phone_script.py <png 폴더>  → slides/script_phone.pdf
대본은 build_deck.py 의 SCRIPTS, 예상 질문은 qa_short.py 의 QA_SHORT(요약본)에서 읽는다.
한 장의 내용이 한 쪽을 넘으면 '(계속)' 쪽으로 이어진다.
"""
import os, re, sys
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))
from qa_short import QA_SHORT

src = (HERE / "build_deck.py").read_text(encoding="utf-8")
ns = {}
exec(src[src.index("SCRIPTS = {"):src.index("# ---------------- 예상 질문")], ns)
SCRIPTS = ns["SCRIPTS"]

PNG = Path(sys.argv[1])
FONT = Path(os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\Windows\Fonts"))
_cache = {}
def f(w, s):
    if (w, s) not in _cache:
        _cache[w, s] = ImageFont.truetype(str(FONT / f"Pretendard-{w}.ttf"), s)
    return _cache[w, s]

W, H, M, BOTTOM = 1080, 1920, 64, 1885   # 폰 세로 화면 비율
INK, SUB, MUTE, BLUE, VIOLET, BG, LINE, PALE = "#191F28", "#4E5968", "#8B95A1", "#3182F6", "#7C5CFF", "#FFFFFF", "#E5E8EB", "#F5F3FF"
measure = ImageDraw.Draw(Image.new("RGB", (1, 1)))


def wrap(text, font, width):
    """한국어는 어절 단위로, 너무 긴 어절은 글자 단위로 줄바꿈."""
    lines, cur = [], ""
    for word in text.split(" "):
        test = (cur + " " + word).strip()
        if measure.textlength(test, font=font) <= width:
            cur = test
            continue
        if cur:
            lines.append(cur)
        cur = word
        while measure.textlength(cur, font=font) > width:
            for i in range(len(cur), 0, -1):
                if measure.textlength(cur[:i], font=font) <= width:
                    lines.append(cur[:i]); cur = cur[i:]; break
    if cur:
        lines.append(cur)
    return lines


def paragraph(text, font, color, lh, after, indent=0):
    """줄 목록 (글, 글꼴, 색, 줄 높이, 들여쓰기) + 끝 간격."""
    rows = [(ln, font, color, lh, indent) for ln in wrap(text, font, W - 2 * M - indent)]
    return rows + [("", None, None, after, 0)]


pages, elapsed = [], 0
for n in sorted(SCRIPTS):
    name, secs, text = SCRIPTS[n]
    elapsed += secs
    rows = []
    for sent in re.split(r"(?<=[.?!])\s+", text.strip()):
        rows += paragraph(sent, f("Medium", 42), INK, 64, 14)
    if QA_SHORT.get(n):
        rows.append(("__qa__", None, None, 80, 0))
        for i, (q, ans) in enumerate(QA_SHORT[n], 1):
            rows += paragraph(f"Q{i}. {q}", f("Bold", 33), VIOLET, 46, 2)
            rows += paragraph(ans, f("Regular", 31), SUB, 43, 14, indent=26)

    def new_page(first):
        page = Image.new("RGB", (W, H), BG)
        d = ImageDraw.Draw(page)
        d.text((M, 56), f"{n} / {len(SCRIPTS)}", font=f("ExtraBold", 44), fill=BLUE)
        d.text((M + 170, 62), name + ("" if first else "  · 예상 질문 (계속)"), font=f("SemiBold", 38), fill=SUB)
        tag = f"{secs}초 · 누적 {elapsed // 60}:{elapsed % 60:02d}"
        d.text((W - M - d.textlength(tag, font=f("Medium", 32)), 66), tag, font=f("Medium", 32), fill=MUTE)
        y = 150
        if first:
            shot = Image.open(PNG / f"slide{n:02d}.png").convert("RGB")
            sw = int((W - 2 * M) * 0.62)
            shot = shot.resize((sw, int(shot.height * sw / shot.width)), Image.LANCZOS)
            page.paste(shot, (M, 140))
            d.rectangle([M, 140, M + sw, 140 + shot.height], outline=LINE, width=2)
            y = 140 + shot.height + 40
        return page, d, y

    page, d, y = new_page(True)
    for txt, font, color, lh, ind in rows:
        if txt and y + lh > BOTTOM:
            pages.append(page)
            page, d, y = new_page(False)
        if txt == "__qa__":
            d.rounded_rectangle([M, y + 12, M + 200, y + 64], radius=26, fill=PALE)
            d.text((M + 26, y + 19), "예상 질문", font=f("Bold", 31), fill=VIOLET)
        elif txt:
            d.text((M + ind, y), txt, font=font, fill=color)
        y += lh
    pages.append(page)

out = HERE / "script_phone.pdf"
pages[0].save(out, save_all=True, append_images=pages[1:], resolution=150)
print(f"saved {out} · {len(pages)}쪽 (발표 {len(SCRIPTS)}장)")
