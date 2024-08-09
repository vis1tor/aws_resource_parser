# AWS 액세스 키 및 시크릿 키 설정
AWS_ACCESS_KEY_ID=""
AWS_SECRET_ACCESS_KEY=""
AWS_SESSION_TOKEN="" # 생략 가능 

# AWS 리전 설정
region = 'ap-northeast-2'

# Output 파일명 설정(ex) ./output/aws_service.xlsx)
output_file_name = 'aws_service.xlsx' 

# AWS 서비스 수집 여부 설정(아래 목록만 현재 수집 가능함)
aws_services = {
    # # 주랭이
    # 'ec2' : 'on',
    # 'sg' : 'on', # Security Groups
    'elb' : 'on', # gatewaylb 작업 추가 필요
    # 'eks' : 'on',
    # 'ecr' : 'on',
    # # 초랭이
    # 's3' : 'on',
    # 'efs' : 'on',
    # # 파랭이
    # 'rds' : 'on', # 클러스터 조회 업데이트 추후 필요
    # 'redis': 'on',
    # 'codecommit' : 'on',
    # 'codebuild' : 'on', # 추후 업데이트 필요
    # # 보랭이
    # 'vpce' : 'on', # Gateway 방식 추후 업데이트 필요
    # 'route53' : 'on',
    # # 분홍이
    # 'cw_logs' : 'on', # CloudWatch Logs
    # 'sns' : 'on'
}