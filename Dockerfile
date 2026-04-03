FROM python:3.12-slim

WORKDIR /app

COPY requirements.txt .
COPY packages.txt .

RUN apt-get update && \
    xargs -a packages.txt apt-get install -y --no-install-recommends && \
    rm -rf /var/lib/apt/lists/*

RUN pip install --no-cache-dir -r requirements.txt

RUN mkdir -p .streamlit
COPY ./streamlit-config.toml .streamlit/config.toml

COPY . .
EXPOSE 3000

CMD ["streamlit", "run", "app.py"]
