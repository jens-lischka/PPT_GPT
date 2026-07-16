FROM python:3.12-slim

# lxml/Pillow build deps kept out by using binary wheels; libxml present in wheels.
WORKDIR /app
COPY requirements.txt runtime_requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt -r runtime_requirements.txt

COPY main.py agent.py jobs.py ow_default.pptx ./
COPY src ./src

# Cloud Run injects $PORT; default 8080 for local docker run.
ENV PORT=8080
CMD exec uvicorn main:app --host 0.0.0.0 --port ${PORT}
