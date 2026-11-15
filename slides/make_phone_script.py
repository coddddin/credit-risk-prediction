"""폰으로 보는 발표 대본 PDF (장마다 한 쪽: 슬라이드 그림 + 대본 + 목표 시간).

준비: 슬라이드를 PNG로 렌더링한 폴더(slide01.png ~ slide22.png)
실행: python make_phone_script.py <png 폴더>  → slides/script_phone.pdf
대본은 build_deck.py 의 SCRIPTS 에서 읽는다.
"""
import os, re, sys
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

HERE = Path(__file__).parent
src = (HERE / "build_deck.py").read_text(encoding="utf-8")
ns = {}
exec(src[src.index("SCRIPTS = {"):src.index("# ---------------- 예상 질문")], ns)
SCRIPTS = ns["SCRIPTS"]

PNG = Path(sys.argv[1])
FONT = Path(os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\Windows\Fonts"))
f = lambda w, s: ImageFont.truetype(str(FONT / f"Pretendard-{w}.ttf"), s)

W, H, M = 1080, 1920, 64            # 폰 세로 화면 비율
INK, SUB, MUTE, BLUE, BG, LINE = "#191F28", "#4E5968", "#8B95A1", "#3182F6", "#FFFFFF", "#E5E8EB"
body_font, LH = f("Medium", 50), 82


def wrap(draw, text, font, width):
    """한국어는 어절 단위로, 너무 긴 어절은 글자 단위로 줄바꿈."""
    lines, cur = [], ""
    for word in text.split(" "):
        test = (cur + " " + word).strip()
        if draw.textlength(test, font=font) <= width:
            cur = test
            continue
        if cur:
            lines.append(cur)
        cur = word
        while draw.textlength(cur, font=font) > width:
            for i in range(len(cur), 0, -1):
                if draw.textlength(cur[:i], font=font) <= width:
                    lines.append(cur[:i]); cur = cur[i:]; break
    if cur:
        lines.append(cur)
    return lines


pages, elapsed = [], 0
total = sum(v[1] for v in SCRIPTS.values())
for n in sorted(SCRIPTS):
    name, secs, text = SCRIPTS[n]
    page = Image.new("RGB", (W, H), BG)
    d = ImageDraw.Draw(page)
    # 머리: 장 번호 · 이름 · 목표 시간 · 누적 시간
    d.text((M, 56), f"{n} / {len(SCRIPTS)}", font=f("ExtraBold", 44), fill=BLUE)
    d.text((M + 170, 62), name, font=f("SemiBold", 38), fill=SUB)
    elapsed += secs
    tag = f"{secs}초 · 누적 {elapsed // 60}:{elapsed % 60:02d}"
    d.text((W - M - d.textlength(tag, font=f("Medium", 32)), 66), tag, font=f("Medium", 32), fill=MUTE)
    # 슬라이드 그림
    shot = Image.open(PNG / f"slide{n:02d}.png").convert("RGB")
    sw = W - 2 * M
    shot = shot.resize((sw, int(shot.height * sw / shot.width)), Image.LANCZOS)
    page.paste(shot, (M, 140))
    d.rectangle([M, 140, M + sw, 140 + shot.height], outline=LINE, width=2)
    # 대본 (문장마다 새 줄로 나눠 읽기 쉽게)
    y = 140 + shot.height + 56
    for sent in re.split(r"(?<=[.?!])\s+", text.strip()):
        for ln in wrap(d, sent, body_font, W - 2 * M):
            d.text((M, y), ln, font=body_font, fill=INK)
            y += LH
        y += 26
    if y > H - 40:
        raise SystemExit(f"{n}장 대본이 한 쪽을 넘침 (y={y})")
    pages.append(page)

out = HERE / "script_phone.pdf"
pages[0].save(out, save_all=True, append_images=pages[1:], resolution=150)
print(f"saved {out} · {len(pages)}쪽 · 목표 합계 {total // 60}분 {total % 60}초")
