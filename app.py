import os
from flask import Flask, request, jsonify, send_file
from ppt_template import create_pptx

app = Flask(__name__)

@app.get("/")
def health():
    return "OK", 200

@app.post("/generate_ppt")
def generate_ppt():
    try:
        data = request.get_json(force=True)
        slides = data["slides"]
        ppt_path = create_pptx(slides, "output.pptx")
        return send_file(ppt_path, as_attachment=True)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    port = int(os.getenv("PORT", 8000))
    # 不开 debug，交给 Gunicorn 统一管理
    app.run(host="0.0.0.0", port=port)