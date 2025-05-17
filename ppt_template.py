from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
import os

def create_job_summary_slide(prs, title, items):
    """
    添加类似“工作内容概述”的PPT页面。
    items: 列表，每项是 {'num': '01', 'subtitle': '项目推进', 'desc': '这里是详细内容...'}
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # 空白布局

    # 添加大标题
    title_box = slide.shapes.add_textbox(Inches(0), Inches(0.4), prs.slide_width, Inches(1))
    tf = title_box.text_frame
    tf.text = title
    p = tf.paragraphs[0]
    p.font.size = Pt(38)
    p.font.bold = True
    p.alignment = PP_ALIGN.CENTER

    # 每一条内容块的起始高度
    start_top = 1.5
    block_height = 1.4
    gap = 0.2

    for i, item in enumerate(items):
        top = Inches(start_top + i * (block_height + gap))
        left = Inches(1)
        width = prs.slide_width - Inches(2)
        height = Inches(block_height)

        # 添加内容块（圆角矩形带边框）
        shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height
        )
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(255, 255, 255)
        shape.line.color.rgb = RGBColor(180, 180, 180)
        shape.line.width = Pt(2)
        shape.shadow.inherit = False

        # 编号圆形
        circle_size = Inches(0.7)
        circle_left = left + Inches(0.3)
        circle_top = top + (height - circle_size) / 2
        circ = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, circle_left, circle_top, circle_size, circle_size
        )
        circ.fill.solid()
        circ.fill.fore_color.rgb = RGBColor(0, 0, 0)
        circ.line.color.rgb = RGBColor(200, 200, 200)
        circ.line.width = Pt(1.2)
        circ.shadow.inherit = False

        # 圆形里的编号
        circ_tf = circ.text_frame
        circ_tf.text = item['num']
        circ_p = circ_tf.paragraphs[0]
        circ_p.font.size = Pt(22)
        circ_p.font.bold = True
        circ_p.font.color.rgb = RGBColor(255, 255, 255)
        circ_p.alignment = PP_ALIGN.CENTER
        circ_tf.vertical_anchor = MSO_ANCHOR.MIDDLE

        # 副标题+正文内容（加在内容块矩形内，适当缩进）
        txt_left = circle_left + circle_size + Inches(0.3)
        txt_top = top + Inches(0.25)
        txt_width = width - (txt_left - left) - Inches(0.4)
        txt_height = height - Inches(0.5)

        txtbox = slide.shapes.add_textbox(txt_left, txt_top, txt_width, txt_height)
        txt_frame = txtbox.text_frame

        # 副标题
        p1 = txt_frame.add_paragraph()
        p1.text = item['subtitle']
        p1.font.size = Pt(22)
        p1.font.bold = True
        p1.font.color.rgb = RGBColor(40, 40, 40)
        p1.alignment = PP_ALIGN.LEFT

        # 正文
        p2 = txt_frame.add_paragraph()
        p2.text = item['desc']
        p2.font.size = Pt(15)
        p2.font.color.rgb = RGBColor(80, 80, 80)
        p2.alignment = PP_ALIGN.LEFT
        txt_frame.margin_top = 0
        txt_frame.margin_bottom = 0
        txt_frame.margin_left = 0
        txt_frame.margin_right = 0

def create_job_summary_ppt(title, items, output_path="output.pptx"):
    prs = Presentation()
    # 16:9
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)

    create_job_summary_slide(prs, title, items)
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    prs.save(output_path)
    print(f"已生成：{output_path}")

# 示例数据
if __name__ == "__main__":
    items = [
        {"num": "01", "subtitle": "项目推进", "desc": "负责项目各环节推进，包括计划制定、进度跟踪和协调资源，确保项目顺利实施。"},
        {"num": "02", "subtitle": "项目成果", "desc": "项目顺利交付，达成预期目标，获得客户高度认可，提升了团队整体协作能力。"},
        {"num": "03", "subtitle": "项目亮点", "desc": "引入创新管理方法，优化流程，提高了效率，实现了项目价值最大化。"},
    ]
    create_job_summary_ppt("工作内容概述", items, "output.pptx")
