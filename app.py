from flask import Flask, request, jsonify, send_file
from ppt_template import create_pptx
import os

app = Flask(__name__)

@app.route('/')
def home():
    return "PPT生成服务已启动，请POST到 /generate_ppt 生成PPT"

@app.route('/generate_ppt', methods=['POST'])
def generate_ppt():
    """接收POST请求生成PPT
    
    请求体格式：
    {
        "slides": [
            {
                "title": "标题1",
                "content": "内容1",
                "img_url": "http://example.com/image1.jpg"  # 可选
            },
            ...
        ],
        "output_path": "custom_name.pptx"  # 可选
    }
    """
    try:
        data = request.get_json()
        if not data or 'slides' not in data:
            return jsonify({'error': '无效的请求数据'}), 400

        slides_data = data['slides']
        output_path = data.get('output_path', 'output.pptx')
        
        # 确保输出路径在当前目录下
        output_path = os.path.basename(output_path)
        output_path = os.path.join(os.path.dirname(__file__), output_path)
        
        # 生成PPT
        ppt_path = create_pptx(slides_data, output_path)
        
        # 返回生成的PPT文件
        return send_file(
            ppt_path,
            as_attachment=True,
            download_name=os.path.basename(ppt_path),
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )

    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)