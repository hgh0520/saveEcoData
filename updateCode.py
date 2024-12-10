import schedule
import subprocess
import time
import threading
import sys


# 종료 플래그
exit_flag = False

update_period = 15

# 스케줄러에서 실행
def run_script():
    # 'saveEcoData.py' 파일 실행
    print("Running saveEcoData.py...")
    subprocess.run(["python", "saveEcoData.py"])

# 키입력으로 즉시 실행
def run_script_now():
    # 'saveEcoData.py' 파일 실행
    print("Running saveEcoData.py...")
    subprocess.run(["python", "saveEcoData.py"])

# 55분마다 실행하도록 스케줄 설정
schedule.every(update_period).minutes.do(run_script)

# 키 입력 처리 함수
def key_input_listener():
    global exit_flag
    while True:
        key = input("Enter 'x' to execute immediately or 'q' to quit: ")
        if key.lower() == "x":
            run_script_now()
        elif key.lower() == "q":
            print("프로그램 종료.")
            exit_flag = True
            sys.exit()

# 스케줄러 실행 함수
def scheduler_runner():
    while not exit_flag:
        schedule.run_pending()
        time.sleep(1)  # CPU 점유율을 낮추기 위해 1초 대기

# 멀티스레드로 키 입력 감지 및 스케줄러 실행
if __name__ == "__main__":
    # 초기 스크립트 실행
    run_script()

    # 키 입력 감지를 별도 스레드로 실행
    input_thread = threading.Thread(target=key_input_listener, daemon=True)
    input_thread.start()

    print(f"{update_period}분마다 작업이 실행됩니다. 즉시 실행하려면 'x'을 입력하세요. 종료하려면 'q'를 입력하세요.")
    scheduler_runner()
