# 설명
계정 내에 생성(배포)되어 있는 AWS 리소스 정보를 확인하기 위해 개발 중인 Python 스크립트입니다.

# 사전 작업
0. Excel 설치
1. Python 설치(3.12.3 버전에서 테스트 완료)
2. Python 모듈 설치
    - pip install -r requirements.txt -U
3. AWS IAM User 생성, 액세스 키 생성, 권한 부여(ReadOnlyAccess - AWS 관리형)
4. conf 디렉터리 내, config.py 파일 값 수정
    -   AWS_ACCESS_KEY_ID 수정        ex) 'AKE .. 생략 .. VVCBND'

    -   AWS_SECRET_ACCESS_KEY 수정        ex) 's2312RSDWQDD .. 생략 .. 7kjEKSJ24sa2'

    -   AWS_SESSION_TOKEN 수정(생략 가능!!)        ex) 'Qskdtksd .. 생략 .. wYY/'
    
    -   region 수정        ex) 'ap-northeast-2'

    -   output_file_name 수정        ex) 'aws_services_info.xlsx'

    -   aws_services 수정        ex) 'ec2' : 'on' or 'ec2' : 'off' 등

5. main.py 스크립트 실행
    -   python main.py

# 참고
    - AttributeError: module 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9' has no attribute 'CLSIDToClassMap' 에러 발생 시
        해결 방법 => C:\Users\<your username>\AppData\Local\Temp\gen_py 디렉토리 제거
        참고 링크 => https://stackoverflow.com/questions/52889704/python-win32com-excel-com-model-started-generating-errors