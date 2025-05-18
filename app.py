import os
from flask import Flask, request, jsonify, send_file
from ppt_templates.card_style import create_card_ppt
from ppt_templates.triple_column import create_triple_column_ppt
from ppt_templates.image_right import create_image_right_ppt
from ppt_templates.icons_grid import create_icons_grid_ppt
from ppt_templates.cover_big_image import create_cover_big_image_ppt

app = Flask(__name__)

@app.route('/')
def home():
    return "OK", 200

@app.route('/generate_ppt', methods=['POST'])
def generate_ppt():
    """
    请求体示例 (新样式):
    {
        "title": "演示文稿标题",
        "items": [
            {"num": "01", "subtitle": "项目推进", "desc": "这里写推进描述……"},
            {"num": "02", "subtitle": "项目成果", "desc": "这里写成果描述……"},
            ...
        ],
        "output_path": "custom_name.pptx"  # 可选
    }
    """
    try:
        data = request.get_json(force=True)
        title = data['title']
        items = data['items']
        output_path = data.get('output_path', 'output.pptx')

        # 保证文件保存到当前目录下
        output_path = os.path.basename(output_path)
        output_path = os.path.join(os.path.dirname(__file__), output_path)

        template_style = data.get('template', 'card_style') # 默认使用card_style模板

        if template_style == 'card_style':
            ppt_path = create_card_ppt(title, items, output_path=output_path)
        elif template_style == 'triple_column':
            ppt_path = create_triple_column_ppt(title, items, output_path=output_path)
        elif template_style == 'image_right':
            # 假设image_right模板需要img_path参数，从kwargs中获取或使用默认值
            img_path = data.get('img_path', 'your_default_image.jpg')
            ppt_path = create_image_right_ppt(title, items, output_path=output_path, img_path=img_path)
        elif template_style == 'icons_grid':
            ppt_path = create_icons_grid_ppt(title, items, output_path=output_path)
        elif template_style == 'cover_big_image':
            subtitle = data.get('subtitle', '')
            img = data.get('img')
            ppt_path = create_cover_big_image_ppt(title, subtitle=subtitle, img=img, output_path=output_path)
        else:
            return jsonify({'error': f'Unknown template style: {template_style}'}), 400

        return send_file(
            ppt_path,
            as_attachment=True,
            download_name=os.path.basename(ppt_path),
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 8080))
    app.run(host='0.0.0.0', port=port)