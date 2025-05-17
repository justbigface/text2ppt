from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import os

def add_card_item(slide, top, num, subtitle, desc):
    # 1. 圆角矩形
    left = Inches(1)
    width = Inches(10.5)
    height = Inches(1.5)
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = RGBColor(255, 255, 255)  # 白底
    card.line.color.rgb = RGBColor(180, 180, 180)
    # python-pptx does not support shadow directly, simulate with border or other shapes if needed
    # card.shadow.inherit = False

    # 2. 左侧大圆
    circle_left = left - Inches(0.65)
    circle_top = top + Inches(0.25)
    circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, circle_left, circle_top, Inches(1), Inches(1))
    circle.fill.solid()
    circle.fill.fore_color.rgb = RGBColor(242, 241, 247)  # 淡灰紫色
    circle.line.color.rgb = RGBColor(90, 70, 160)
    circle.line.width = Pt(3)

    # 圆内编号
    num_frame = circle.text_frame
    p = num_frame.paragraphs[0]
    p.text = str(num) # Ensure num is string
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = RGBColor(80, 70, 160)
    p.alignment = PP_ALIGN.CENTER

    # 3. 圆角矩形里的内容
    # 副标题
    sub_box = slide.shapes.add_textbox(left + Inches(1.1), top + Inches(0.2), Inches(8.5), Inches(0.5))
    sub_frame = sub_box.text_frame
    sub_frame.text = subtitle
    p = sub_frame.paragraphs[0]
    p.font.size = Pt(24)
    p.font.bold = True
    p.font.color.rgb = RGBColor(90, 70, 160)
    p.alignment = PP_ALIGN.LEFT

    # 描述内容
    desc_box = slide.shapes.add_textbox(left + Inches(1.1), top + Inches(0.7), Inches(8.2), Inches(0.7))
    desc_frame = desc_box.text_frame
    desc_frame.text = desc
    p = desc_frame.paragraphs[0]
    p.font.size = Pt(16)
    p.font.color.rgb = RGBColor(50, 50, 50)
    p.alignment = PP_ALIGN.LEFT

def create_job_summary_ppt(title, items, output_path="output.pptx"):
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[5])  # 空白

    # 标题
    title_box = slide.shapes.add_textbox(Inches(1), Inches(0.3), Inches(11), Inches(1))
    title_frame = title_box.text_frame
    title_frame.text = title
    p = title_frame.paragraphs[0]
    p.font.size = Pt(40)
    p.font.bold = True
    p.alignment = PP_ALIGN.LEFT

    # 循环绘制每个卡片条目
    for idx, item in enumerate(items):
        add_card_item(
            slide,
            top=Inches(1.5) + idx * Inches(2),  # 每行间距
            num=item["num"],
            subtitle=item["subtitle"],
            desc=item["desc"]
        )

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    prs.save(output_path)
    return output_path

# Example usage (optional, for testing)
# if __name__ == "__main__":
#     items = [
#         {"num": "01", "subtitle": "项目推进", "desc": "这里写推进描述……"},
#         {"num": "02", "subtitle": "项目成果", "desc": "这里写成果描述……"},
#         {"num": "03", "subtitle": "项目亮点", "desc": "这里写亮点描述……"},
#     ]
#     create_job_summary_ppt("工作内容概述", items, "output.pptx")
