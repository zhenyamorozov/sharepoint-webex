FROM python:3.11-slim
RUN apt-get update && apt-get install -y git && apt-get clean
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
EXPOSE 5000
CMD ["python", "web.py"]

