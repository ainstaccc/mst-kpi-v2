# 1. 使用官方 Python 3.10
FROM python:3.10-slim

# 2. 建立工作目錄
WORKDIR /app

# 3. 複製需求與程式碼
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# 4. 使用 Streamlit 執行
EXPOSE 8080
CMD ["streamlit", "run", "app.py", "--server.port=8080", "--server.address=0.0.0.0"]
