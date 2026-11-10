"""Lab PDF에서 코드 블록을 잘라 이미지로 저장한다 (장표의 'Lab 원본' 캡처용)."""
from pathlib import Path

import pymupdf

LAB_DIR = Path(r"C:\Users\heeju\Desktop\ALGORYITHM\수업 자료")
OUT_DIR = Path(__file__).resolve().parent.parent / "figures" / "lab_captures"


def capture(pdf_name, page_no, start_text, end_text, out_name, zoom=3.0, pad=8, mark=None):
    """page_no 쪽에서 start_text 줄부터 그 아래 처음 나오는 end_text 줄까지 잘라 저장.
    mark: 이 글자가 있는 줄에 빨간 테두리를 그린다 (우리가 고친 줄 표시)."""
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    doc = pymupdf.open(LAB_DIR / pdf_name)
    page = doc[page_no]
    start = page.search_for(start_text)[0]
    ends = [r for r in page.search_for(end_text) if r.y0 >= start.y0]
    end = ends[0]
    region = pymupdf.Rect(start.x0 - pad, start.y0 - pad, page.rect.width, end.y1 + pad)
    right = max(w[2] for w in page.get_text("words", clip=region))  # 글자가 끝나는 오른쪽 끝
    clip = pymupdf.Rect(region.x0, region.y0, right + pad * 2, region.y1)
    pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), clip=clip)
    path = OUT_DIR / out_name
    pix.save(path)
    if mark:
        from PIL import Image, ImageDraw
        hit = [r for r in page.search_for(mark) if clip.contains(r)][0]
        # 같은 줄의 끝까지 테두리를 늘린다
        line_words = [w for w in page.get_text("words", clip=clip) if abs(w[1] - hit.y0) < 2]
        x1 = max([hit.x1] + [w[2] for w in line_words])
        to_px = lambda vx, vy: ((vx - clip.x0) * zoom, (vy - clip.y0) * zoom)
        a = to_px(hit.x0 - 3, hit.y0 - 2.5)
        b = to_px(x1 + 3, hit.y1 + 2.5)
        im = Image.open(path).convert("RGB")
        ImageDraw.Draw(im).rounded_rectangle([a, b], radius=10, outline=(240, 68, 82), width=5)
        im.save(path)
    return path


if __name__ == "__main__":
    print(capture("Lab1_Loading ML Data.pdf", 3, "# Load CSV using Pandas", "print(data.shape)", "lab1_read_csv.png"))
    print(capture("Lab1_Loading ML Data.pdf", 7, "# Class Distribution", "print(class_counts)", "lab1_class_dist.png"))
