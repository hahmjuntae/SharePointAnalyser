# 1. 가상 환경 생성
python3 -m venv venv

# 2. 가상 환경 활성화
source venv/bin/activate

# 3. 패키지 설치
pip3 install pandas openpyxl

# 4. 애플리케이션 실행
python3 audit_log_analyser.py

# 가상 환경 비활성화
deactivate
