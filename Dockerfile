FROM python:3.10-slim

# 更快的安装 & 字体支持（可选）
RUN apt-get update && apt-get install -y \
    fonts-dejavu-core && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# 默认用 $PORT，Zeabur/Render/Heroku 都兼容
ENV PORT=5000
EXPOSE 5000

# 2 workers, 1 thread each；可按需调整
CMD sh -c 'gunicorn -w 2 -k gthread --threads 1 -b 0.0.0.0:${PORT:-5000} app:app'