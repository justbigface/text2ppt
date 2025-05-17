import os
from flask import Flask, request, jsonify, send_file
from ppt_template import create_pptx

app = Flask(__name__)

@app.route('/')
def home():
    return "OK", 200

@app.route('/generate_ppt', methods=['POST'])
def generate_ppt():
    """
    请求体示例:
    {
        "slides": [
            {
                "title": "第一页标题",
                "content": "第一页内容",
                "img_url": " `https://example.com/image.jpg` "  # 可选
            },
            ...
        ],
        "output_path": "custom_name.pptx"  # 可选
    }
    """
    try:
        data = request.get_json(force=True)
        slides_data = data['slides']
        output_path = data.get('output_path', 'output.pptx')

        # 保证文件保存到当前目录下
        output_path = os.path.basename(output_path)
        output_path = os.path.join(os.path.dirname(__file__), output_path)

        ppt_path = create_pptx(slides_data, output_path)

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