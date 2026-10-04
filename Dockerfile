FROM python:3.12-slim
WORKDIR /work

RUN apt-get update && apt-get install -y --no-install-recommends \
    gcc g++ libgdal-dev \
    && rm -rf /var/lib/apt-get/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY main.py .
CMD ["python", "main.py"]