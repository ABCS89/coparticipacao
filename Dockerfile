FROM python:3.13-slim-buster

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

CMD gunicorn teste:app --workers 3 --bind 0.0.0.0:8000