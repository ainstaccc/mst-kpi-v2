# 1. 用Python官方映像檔
FROM python:3.11

# 2. 設定工作目錄
WORKDIR /app

# 3. 複製程式和需求檔案
COPY requirements.txt requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

# 4. 複製專案全部檔案
COPY . .

# 5. 告訴Cloud Run要跑Streamlit
CMD ["streamlit", "run", "app.py", "--server.port", "8080", "--server.address", "0.0.0.0"]
