FROM python:3.11-slim
RUN apt-get update && apt-get install -y git && apt-get clean
WORKDIR /app
COPY requirements.txt .
# AppRunner deployment does not run pip directly, calling it as Python module
RUN python -m pip install --no-cache-dir -r requirements.txt
COPY . .
CMD ["python", "web.py"]
EXPOSE 5000
