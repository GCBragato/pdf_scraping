import os.path

import fitz
from fitz import Document, Page, Rect


# For visualizing the rects that PyMuPDF uses compared to what you see in the PDF
VISUALIZE = True

input_path = "_SPTs_test.pdf"
doc: Document = fitz.open(input_path)

for i in range(len(doc)):
    page: Page = doc[i]

    # Hard-code the rect you need
    rect = Rect(506, 115, 549, 128)

    if VISUALIZE:
        # Draw a red box to visualize the rect's area (text)
        page.draw_rect(rect, width=1.5, color=(1, 0, 0))

    text = page.get_textbox(rect)

    print(text)


if VISUALIZE:
    head, tail = os.path.split(input_path)
    viz_name = os.path.join(head, "viz_" + tail)
    doc.save(viz_name)
    #print(page.rect.width, page.rect.height)