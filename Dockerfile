FROM python:3.12-slim-bullseye

ENV DEBIAN_FRONTEND=noninteractive

WORKDIR /app

RUN apt-get update && \
    apt-get install -y --no-install-recommends \
        libreoffice-core \
        libreoffice-impress \
        libreoffice-writer \
        default-jre-headless \
        fonts-dejavu \
        curl \
        unzip \
        python3-dev \
        build-essential \
        && apt-get clean \
        && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY app ./app

EXPOSE 8000

CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8000"]
