import os
import win32com.client as win32 # 엑셀 파일 셀 크기 자동 조절 https://abooda.tistory.com/30
import boto3
from openpyxl import Workbook
from conf.config import AWS_ACCESS_KEY_ID, AWS_SECRET_ACCESS_KEY, AWS_SESSION_TOKEN, region, output_file_name, aws_services
from aws_services import boto3_client
from aws_services import ec2, sg, elb, eks, ecr
from aws_services import s3, efs
from aws_services import rds, redis, codecommit, codebuild
from aws_services import vpce, route53
from aws_services import cw_logs, sns

def main():
    output_dir = 'output'
    output_file = output_dir+'/'+ output_file_name
    
    if not os.path.exists(os.getcwd()+'\\'+output_dir):
        os.makedirs(output_dir)

    workbook = Workbook()
    
    #common client
    ec2_client = boto3_client.get_aws_client('ec2', AWS_ACCESS_KEY_ID, AWS_SECRET_ACCESS_KEY, AWS_SESSION_TOKEN, region)
    
    ###################
    # 주랭이 서비스
    ###################
    # EC2 정보 조회 및 엑셀 파일 생성 - 위에서 client 정의하여 생략
    if aws_services.get('ec2') == 'on':
        ec2.export_ec2_info_to_excel(workbook, ec2_client)
    # Security Group 정보 조회 및 엑셀 파일 생성
    if aws_services.get('sg') == 'on':
        sg.export_sg_info_to_excel(workbook, ec2_client)
    # ELB 정보 조회 및 엑셀 파일 생성
    if aws_services.get('elb') == 'on':
        elbv2_client = boto3_client.get_aws_client('elbv2', AWS_ACCESS_KEY_ID, AWS_SECRET_ACCESS_KEY, AWS_SESSION_TOKEN, region)
        elb.export_elb_info_to_excel(workbook, elbv2_client, ec2_client)
    # EKS 정보 조회 및 엑셀 파일 생성
    if aws_services.get('eks') == 'on':
        eks_client = boto3_client.get_aws_client('eks', AWS_ACCESS_KEY_ID, AWS_SECRET_ACCESS_KEY, AWS_SESSION_TOKEN, region)
        eks.export_eks_info_to_excel(workbook, eks_client, ec2_client)
    # ECR 정보 조회 및 엑셀 파일 생성
    if aws_services.get('ecr') == 'on':
        private_ecr_client = boto3_client.get_aws_client('ecr', AWS_ACCESS_KEY_ID, AWS_SECRET_ACCESS_KEY, AWS_SESSION_TOKEN, region)
        public_ecr_client = boto3_client.get_aws_client('ecr-public', AWS_ACCESS_KEY_ID, AWS_SECRET_ACCESS_KEY, AWS_SESSION_TOKEN, 'us-east-1') # public ecr 리전은 반드시 us-east-1
        ecr.export_ecr_info_to_excel(workbook, private_ecr_client, public_ecr_client)
    ###################
    # 초랭이 서비스
    ###################
    # S3 정보 조회 및 엑셀 파일 생성
    if aws_services.get('s3') == 'on':
        s3_client = boto3_client.get_aws_client('s3', AWS_ACCESS_KEY_ID, AWS_SECRET_ACCESS_KEY, AWS_SESSION_TOKEN, region)
        s3.export_s3_info_to_excel(workbook, s3_client)
    # EFS 정보 조회 및 엑셀 파일 생성
    if aws_services.get('efs') == 'on':
        efs_client = boto3_client.get_aws_client('efs', AWS_ACCESS_KEY_ID, AWS_SECRET_ACCESS_KEY, AWS_SESSION_TOKEN, region)
        efs.export_efs_info_to_excel(workbook, efs_client, ec2_client)
    ###################
    # 파랭이 서비스
    ###################
    # RDS 정보 조회 및 엑셀 파일 생성
    if aws_services.get('rds') == 'on':
        rds_client = boto3_client.get_aws_client('rds', AWS_ACCESS_KEY_ID, AWS_SECRET_ACCESS_KEY, AWS_SESSION_TOKEN, region)
        rds.export_rds_info_to_excel(workbook, rds_client, ec2_client)
    # Redis 정보 조회 및 엑셀 파일 생성
    if aws_services.get('redis') == 'on':
        redis_client = boto3_client.get_aws_client('elasticache', AWS_ACCESS_KEY_ID, AWS_SECRET_ACCESS_KEY, AWS_SESSION_TOKEN, region)
        redis.export_redis_info_to_excel(workbook, redis_client, ec2_client)
    # CodeCommit 정보 조회 및 엑셀 파일 생성
    if aws_services.get('codecommit') == 'on':
        codecommit_client = boto3_client.get_aws_client('codecommit', AWS_ACCESS_KEY_ID, AWS_SECRET_ACCESS_KEY, AWS_SESSION_TOKEN, region)
        codecommit.export_codecommit_info_to_excel(workbook, codecommit_client)
    # CodeBuild 정보 조회 및 엑셀 파일 생성
    if aws_services.get('codebuild') == 'on':
        codebuild_client = boto3_client.get_aws_client('codebuild', AWS_ACCESS_KEY_ID, AWS_SECRET_ACCESS_KEY, AWS_SESSION_TOKEN, region)
        codebuild.export_codebuild_info_to_excel(workbook, codebuild_client)
    ###################
    # 보랭이 서비스
    ###################
    # VPC Endpoint 정보 조회 및 엑셀 파일 생성
    if aws_services.get('vpce') == 'on':
        vpce.export_vpce_info_to_excel(workbook, ec2_client)
    # Route53 정보 조회 및 엑셀 파일 생성
    if aws_services.get('route53') == 'on':
        route53_client = boto3_client.get_aws_client('route53', AWS_ACCESS_KEY_ID, AWS_SECRET_ACCESS_KEY, AWS_SESSION_TOKEN, region)
        route53.export_route53_info_to_excel(workbook, route53_client)
    ###################
    # 분홍이 서비스
    ###################
    # CloudWatch Log Group 정보 조회 및 엑셀 파일 생성
    if aws_services.get('cw_logs') == 'on':
        cw_logs_client = boto3_client.get_aws_client('logs', AWS_ACCESS_KEY_ID, AWS_SECRET_ACCESS_KEY, AWS_SESSION_TOKEN, region)
        cw_logs.export_cw_logs_info_to_excel(workbook, cw_logs_client)
    # SNS 정보 조회 및 엑셀 파일 생성
    if aws_services.get('sns') == 'on':
        sns_client = boto3_client.get_aws_client('sns', AWS_ACCESS_KEY_ID, AWS_SECRET_ACCESS_KEY, AWS_SESSION_TOKEN, region)
        sns.export_sns_info_to_excel(workbook, sns_client)
    
    workbook.save(output_file)

    # 엑셀 시트 셀 크기 자동 조절
    try:
        excel = win32.gencache.EnsureDispatch('Excel.Application') 
        load_wb = excel.Workbooks.Open(os.getcwd()+'\\'+output_file)

        sheet_list = []
        for sh in load_wb.Sheets:
            ws = load_wb.Worksheets(sh.Name)
            ws.Columns.AutoFit()
            ws.Rows.AutoFit()
        
        load_wb.Save()
        excel.Application.Quit()
    except TypeError as e:
        print(e)
        print('열려있는 Excel 파일을 모두 닫아주세요...')

if __name__ == '__main__':
    main()