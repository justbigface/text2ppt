## 文本转PPT接口部署和使用说明

### 一、接口说明

* `POST /generate_ppt`

  * 请求体示例（JSON）：

    ```json
    {
        "slides": [
            {"title": "第一页", "content": "这是第一页的内容", "img_url": " `https://...` "},
            {"title": "第二页", "content": "这是第二页内容"}
        ],
        "output_path": "demo.pptx"
    }
    ```
  * 返回：PPTX二进制文件（直接下载）
* `GET /`

  * 健康检查，返回 "OK"

### 二、部署到 Zeabur 或其他云平台

* 默认监听 8080 端口，平台服务端口需一致。
* 推荐直接用 Dockerfile 部署。

### 三、依赖

* Flask、gunicorn、python-pptx、Pillow、requests

### 四、本地启动

```bash
pip install -r requirements.txt
python app.py
```

* 默认端口 8080，可用 `curl` 或 Postman 测试。

---

## **你可以直接把这套代码打包上传到 GitHub 或 Zeabur，一键部署即用。**

如果需要多模板，可以用 `ppt_template.py` 里再扩展参数或加不同生成函数，支持自定义样式、主题等。

如有后续自动化、文件回传、对接n8n等需求，也可以直接基于这套API扩展！

---

如还需前端上传界面或n8n工作流范例，随时可以补！