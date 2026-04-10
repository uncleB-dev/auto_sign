# Python 3.11 슬림 이미지를 기반으로 사용
FROM python:3.11-slim

# 작업 디렉토리 설정
WORKDIR /app

# 시스템 의존성 설치 (OpenCV 및 PyMuPDF를 위해 필요)
RUN apt-get update && apt-get install -y \
    build-essential \
    libgl1-mesa-glx \
    libglib2.0-0 \
    && rm -rf /var/lib/apt/lists/*

# 종속성 파일 복사 및 설치
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 프로젝트 파일 복사 (폰트, 이미지 포함)
COPY . .

# Streamlit 기본 포트 노출
EXPOSE 8501

# 환경 변수 설정
ENV STREAMLIT_SERVER_PORT=8501
ENV STREAMLIT_SERVER_ADDRESS=0.0.0.0

# 앱 실행
CMD ["streamlit", "run", "app.py"]
