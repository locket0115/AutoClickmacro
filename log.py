import os
import logging

# 디렉토리들
cur_dir = os.getcwd()
log_dir = os.path.join(cur_dir, 'logs')

# 디렉토리 생성
if not os.path.exists(log_dir):
    os.makedirs(log_dir)


# 루트 로거 생성
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)  # 모든 로그 메시지를 처리하기 위해 DEBUG 레벨로 설정

# 로그 파일 (logger.log) 핸들러 설정
file_handler = logging.FileHandler(os.path.join(log_dir, 'logger.log'))
file_handler.setLevel(logging.INFO)  # 모든 로그를 기록

# 에러 로그 파일 (error.log) 핸들러 설정
error_handler = logging.FileHandler(os.path.join(log_dir, 'error.log'))
error_handler.setLevel(logging.ERROR)  # ERROR 레벨 이상만 기록

# 콘솔 핸들러 설정
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.DEBUG)  # 모든 로그를 콘솔에 출력

# 로그 포맷 설정
formatter = logging.Formatter('%(asctime)s  %(levelname)s: %(message)s')
file_handler.setFormatter(formatter)
error_handler.setFormatter(formatter)
console_handler.setFormatter(formatter)

# 핸들러 추가
logger.addHandler(file_handler)
logger.addHandler(error_handler)
logger.addHandler(console_handler)