from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.enum.text import PP_ALIGN
from PIL import Image
import requests
import io
import os

def create_pptx(slides_data, output_path="output.pptx"):
    """
    slides_data: [
        {
            "title": ...,
            "content": ...,
            "img_url": ... # 可选
        }
    ]
    output_path: str
    """
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)

    for slide_data in slides_data:
        slide = prs.slides.add_slide(prs.slide_layouts[5])  # 空白布局

        # 标题
        left = Inches(1)
        top = Inches(0.5)
        width = Inches(11)
        height = Inches(1)
        title_box = slide.shapes.add_textbox(left, top, width, height)
        title_frame = title_box.text_frame
        title_frame.text = slide_data.get("title", "")
        p = title_frame.paragraphs[0]
        p.font.size = Pt(40)
        p.font.bold = True
        p.alignment = PP_ALIGN.LEFT

        # 主要内容
        left = Inches(1)
        top = Inches(2)
        width = Inches(11)
        height = Inches(3)
        content_box = slide.shapes.add_textbox(left, top, width, height)
        content_frame = content_box.text_frame
        content_frame.text = slide_data.get("content", "")
        p2 = content_frame.paragraphs[0]
        p2.font.size = Pt(28)
        p2.alignment = PP_ALIGN.LEFT

        # 图片（如果提供）
        img_url = slide_data.get("img_url")
        if img_url:
            try:
                response = requests.get(img_url)
                img_stream = io.BytesIO(response.content)
                slide.shapes.add_picture(img_stream, Inches(10), Inches(5), width=Inches(2), height=Inches(2))
            except Exception as e:
                print(f"图片下载失败: {e}")

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    prs.save(output_path)
    return output_path
