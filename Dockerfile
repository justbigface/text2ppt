FROM python:3.10-slim

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 8080
ENV PORT=8080

CMD sh -c 'gunicorn -w 2 -k gthread --threads 1 -b 0.0.0.0:${PORT:-8080} app:app'