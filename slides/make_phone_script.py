"""폰으로 보는 발표 대본 PDF (장마다: 슬라이드 그림 + 대본 + 그 장의 예상 질문 요약).

대본 원본
  - slides/script_edit.md 가 있으면 그 파일을 읽는다 (PPT 제출 뒤 대본을 고칠 때는 이 파일만 고친다).
  - 없으면 build_deck.py 의 SCRIPTS 와 qa_short.py 의 QA_SHORT 를 읽는다.

실행
  python make_phone_script.py --init            # script_edit.md 를 처음 한 번 만든다 (이미 있으면 덮어쓰지 않음)
  python make_phone_script.py [png 폴더]         # → slides/script_phone.pdf
  png 폴더 기본값은 slides/submitted/png (제출한 PPT를 render_png.ps1 로 내보낸 그림)
한 장의 내용이 한 쪽을 넘으면 '(계속)' 쪽으로 이어진다.
"""
import os, re, sys
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

HERE = Path(__file__).parent
EDIT = HERE / "script_edit.md"
SPEED = 5  # 공백 빼고 초당 글자 수 (build_deck.py 의 예상 시간과 같은 기준)


def from_build_deck():
    sys.path.insert(0, str(HERE))
    from qa_short import QA_SHORT
    src = (HERE / "build_deck.py").read_text(encoding="utf-8")
    ns = {}
    exec(src[src.index("SCRIPTS = {"):src.index("# ---------------- 예상 질문")], ns)
    return {n: (name, secs, text, QA_SHORT.get(n, [])) for n, (name, secs, text) in ns["SCRIPTS"].items()}


def write_edit_file(slides):
    out = ["# 발표 대본 (편집용)", "",
           "- 이 파일만 고치면 된다. 저장 뒤 `python make_phone_script.py` 를 실행하면 폰용 대본 PDF가 다시 만들어진다.",
           "- 형식: `## 장번호. 장 이름 · 목표 초` 아래에 대본, `### 예상 질문` 아래에 `- Q:` / `  A:` 두 줄씩.",
           "- 문장마다 줄을 바꾸면 PDF에서도 문장마다 줄이 나뉜다. PPT(제출본)는 이 파일과 상관없이 그대로다.", ""]
    for n, (name, secs, text, qa) in sorted(slides.items()):
        out += [f"## {n}. {name} · {secs}초", ""]
        out += [s for s in re.split(r"(?<=[.?!])\s+", text.strip())]
        if qa:
            out += ["", "### 예상 질문"]
            for q, a in qa:
                out += [f"- Q: {q}", f"  A: {a}"]
        out.append("")
    EDIT.write_text("\n".join(out), encoding="utf-8")


def read_edit_file():
    slides, cur, in_qa = {}, None, False
    for raw in EDIT.read_text(encoding="utf-8").splitlines():
        line = raw.rstrip()
        m = re.match(r"^##\s+(\d+)\.\s*(.+?)\s*·\s*(\d+)\s*초\s*$", line)
        if m:
            cur = int(m.group(1))
            slides[cur] = [m.group(2), int(m.group(3)), [], []]
            in_qa = False
            continue
        if cur is None or not line.strip():
            continue
        if line.startswith("### 예상 질문"):
            in_qa = True
        elif in_qa and line.lstrip().startswith("- Q:"):
            slides[cur][3].append([line.split("Q:", 1)[1].strip(), ""])
        elif in_qa and line.lstrip().startswith("A:") and slides[cur][3]:
            slides[cur][3][-1][1] = line.split("A:", 1)[1].strip()
        elif not in_qa:
            slides[cur][2].append(line.strip())
    return {n: (v[0], v[1], " ".join(v[2]), [tuple(q) for q in v[3]]) for n, v in slides.items()}


if "--init" in sys.argv:
    if EDIT.exists():
        sys.exit(f"{EDIT.name} 이 이미 있어 덮어쓰지 않았습니다.")
    write_edit_file(from_build_deck())
    sys.exit(f"saved {EDIT}")

SLIDES = read_edit_file() if EDIT.exists() else from_build_deck()
args = [a for a in sys.argv[1:] if not a.startswith("--")]
PNG = Path(args[0]) if args else HERE / "submitted" / "png"
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


pages, elapsed, spoken = [], 0, 0
for n in sorted(SLIDES):
    name, secs, text, qa = SLIDES[n]
    elapsed += secs
    est = round(len(re.sub(r"\s", "", text)) / SPEED)
    spoken += est
    rows = []
    for sent in re.split(r"(?<=[.?!])\s+", text.strip()):
        rows += paragraph(sent, f("Medium", 42), INK, 64, 14)
    if qa:
        rows.append(("__qa__", None, None, 80, 0))
        for i, (q, ans) in enumerate(qa, 1):
            rows += paragraph(f"Q{i}. {q}", f("Bold", 33), VIOLET, 46, 2)
            rows += paragraph(ans, f("Regular", 31), SUB, 43, 14, indent=26)

    def new_page(first):
        page = Image.new("RGB", (W, H), BG)
        d = ImageDraw.Draw(page)
        d.text((M, 56), f"{n} / {len(SLIDES)}", font=f("ExtraBold", 44), fill=BLUE)
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
src = EDIT.name if EDIT.exists() else "build_deck.py"
print(f"saved {out} · {len(pages)}쪽 (발표 {len(SLIDES)}장, 원본 {src}) · "
      f"목표 {elapsed // 60}분 {elapsed % 60}초 · 대본 예상 {spoken // 60}분 {spoken % 60}초")
