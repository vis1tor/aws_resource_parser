import boto3

def get_aws_client(service_name, AWS_ACCESS_KEY_ID, AWS_SECRET_ACCESS_KEY, AWS_SESSION_TOKEN, region):
    return boto3.client(service_name, aws_access_key_id=AWS_ACCESS_KEY_ID, aws_secret_access_key=AWS_SECRET_ACCESS_KEY, aws_session_token=AWS_SESSION_TOKEN, region_name=region)

# 토큰 추가 전
# def get_aws_resource(service_name, aws_access_key, aws_secret_key, region):
#     return boto3.client(service_name, aws_access_key_id=aws_access_key, aws_secret_access_key=aws_secret_key, region_name=region)

