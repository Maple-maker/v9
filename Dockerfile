FROM python:3.11-slim
WORKDIR /app
COPY requirements.txt /app/requirements.txt
RUN pip install --no-cache-dir -r /app/requirements.txt
COPY app.py /app/app.py
EXPOSE 8080
CMD ["bash","-lc","gunicorn -w 2 -k gthread -t 120 -b 0.0.0.0:${PORT:-8080} app:app"]
