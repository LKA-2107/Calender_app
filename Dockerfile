FROM python:3.11-slim

WORKDIR /app
COPY app/requirements.txt /app/requirements.txt
RUN pip install --no-cache-dir -r /app/requirements.txt

COPY app/main.py /app/main.py

# Data dir for token + sqlite
RUN mkdir -p /data
ENV DATA_DIR=/data

CMD ["python", "/app/main.py"]