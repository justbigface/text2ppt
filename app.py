import os
from flask import Flask, request, jsonify, send_file
from ppt_template import create_job_summary_ppt

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

        ppt_path = create_job_summary_ppt(title, items, output_path)

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