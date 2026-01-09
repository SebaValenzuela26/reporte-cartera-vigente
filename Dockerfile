FROM python:3.12-slim-bullseye

# Evitar prompts de apt
ENV DEBIAN_FRONTEND=noninteractive

# Carpeta de trabajo
WORKDIR /app

# Instalar dependencias del sistema necesarias para LibreOffice y Python
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

# Copiar requirements e instalar Python packages
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copiar el código de la app
COPY app ./app

# Exponer puerto
EXPOSE 8000

# Comando por defecto
CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8000"]
