import os
from flask import Flask, request, jsonify, send_file
from ppt_template import create_pptx

app = Flask(__name__)

@app.route('/')
def home():
    return "PPT服务运行中"

@app.route('/generate_ppt', methods=['POST'])
def generate_ppt():
    try:
        data = request.get_json()
        slides_data = data['slides']
        output_path = "output.pptx"
        ppt_path = create_pptx(slides_data, output_path)
        return send_file(ppt_path, as_attachment=True)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port, debug=False)