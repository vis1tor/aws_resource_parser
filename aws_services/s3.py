# 개선 : 리전별 버킷 웹 사이트 엔드포인트 형식 지정 및 추가 필요
# 참고 : https://docs.aws.amazon.com/ko_kr/general/latest/gr/s3.html#s3_website_region_endpoints
# 위치 : get_bucket_static_web_hosting() -> bucket_static_web_hosting = 'Enabled' 하위

# 2023-12 수명주기 여러 개일 때 내용 반영 수정 중...
from botocore.exceptions import ClientError
from conf.sheet_style import front_header_font,header_font,header_fill,header_alignment,content_alignment,multiple_content_alignment,content_border,header_border

def get_bucket_tags(s3_client, bucket_name):
    # S3는 태그 미존재 시, 에러 발생하여 예외 처리
    try:
        response = s3_client.get_bucket_tagging(Bucket=bucket_name)
        s3_tag_list = []
        
        for tag in response['TagSet']:
            s3_tag_list.append(f"{tag['Key']} : {tag['Value']}")

        # s3_tag_list = ', '.join(s3_tag_list) - 불필요
        s3_tag_list = '\n'.join(s3_tag_list)
        
    except ClientError as e:
        if e.response['Error']['Code'] == 'NoSuchTagSet':
            s3_tag_list = '-'
        else:
            print('Unknown Error...')

    return s3_tag_list


def get_bucket_enc(s3_client, bucket_name):
    response = s3_client.get_bucket_encryption(Bucket=bucket_name)
    bucket_enc_info = response['ServerSideEncryptionConfiguration']['Rules'][0]['ApplyServerSideEncryptionByDefault']

    # KMSMasterKeyID 키 사라져서 임시 변경
    bucket_enc_type = bucket_enc_info['SSEAlgorithm']
    bucket_enc_key = '-'
    
    return bucket_enc_type, bucket_enc_key
    # AES25 암호화 예외처리
    # bucket_enc_type = bucket_enc_info['SSEAlgorithm']
    # if bucket_enc_type == 'AES256':
    #     bucket_enc_key = '-'
    # else:
    #     bucket_enc_key = bucket_enc_info['KMSMasterKeyID']
    
    # return bucket_enc_type, bucket_enc_key

def get_bucket_versioning(s3_client, bucket_name):
    response = s3_client.get_bucket_versioning(Bucket=bucket_name)
    
    if response.get('Status'):
        bucket_versioning = response.get('Status')
    else:
        bucket_versioning = 'Disabled'
    
    return bucket_versioning


def get_bucket_logging(s3_client, bucket_name):

    response = s3_client.get_bucket_logging(Bucket=bucket_name)
    
    # print(response)
    # 액세스 로깅 비활성/활성 예외처리
    if response.get('LoggingEnabled'):
        bucket_logging = 'Enabled'
        bucket_logging_target = response['LoggingEnabled']['TargetBucket']

        if response['LoggingEnabled']['TargetPrefix'] == '':
            bucket_logging_prefix = '-'
        else:
            bucket_logging_prefix = response['LoggingEnabled']['TargetPrefix']

        if 'TargetObjectKeyFormat' in response['LoggingEnabled']:
            if 'PartitionedPrefix' in response['LoggingEnabled']['TargetObjectKeyFormat']:
                bucket_log_format = f"PartitionedPrefix({response['LoggingEnabled']['TargetObjectKeyFormat']['PartitionedPrefix']['PartitionDateSource']})"
            if 'SimplePrefix' in response['LoggingEnabled']['TargetObjectKeyFormat']:
                bucket_log_format = list(response['LoggingEnabled']['TargetObjectKeyFormat'].keys())[0]
        else: 
            bucket_log_format = '-'
    else:
        bucket_logging = 'Disabled'
        bucket_logging_target = '-'
        bucket_logging_prefix = '-'
        bucket_log_format = '-'
    
    return bucket_logging, bucket_logging_target, bucket_logging_prefix, bucket_log_format

# 임시 주석
# def get_bucket_logging(s3_client, bucket_name):
#     response = s3_client.get_bucket_logging(Bucket=bucket_name)
    
#     # 액세스 로깅 비활성/활성 예외처리
#     if response.get('LoggingEnabled'):
#         bucket_logging = 'Enabled'
#         bucket_logging_target = response['LoggingEnabled']['TargetBucket']
#         # PartitionedPrefix 로그 포맷 사용 시
#         if 'PartitionedPrefix' in response['LoggingEnabled']['TargetObjectKeyFormat']:
#             bucket_log_format = f"PartitionedPrefix({response['LoggingEnabled']['TargetObjectKeyFormat']['PartitionedPrefix']['PartitionDateSource']})"
#             # 접두사가 없으면 '-'
#             if response['LoggingEnabled']['TargetPrefix'] == '':
#                 bucket_logging_prefix = '-'
#             else:
#                 bucket_logging_prefix = response['LoggingEnabled']['TargetPrefix']
#         # SimplePrefix 로그(default) 포맷 사용 시
#         if 'SimplePrefix' in response['LoggingEnabled']['TargetObjectKeyFormat']:
#             bucket_log_format = list(response['LoggingEnabled']['TargetObjectKeyFormat'].keys())[0]
#             # 접두사가 없으면 '-'
#             if response['LoggingEnabled']['TargetPrefix'] == '':
#                 bucket_logging_prefix = '-'
#             else:
#                 bucket_logging_prefix = response['LoggingEnabled']['TargetPrefix']
#     # 로깅 비활성화 시
#     else:
#         bucket_logging = 'Disabled'
#         bucket_logging_target = '-'
#         bucket_logging_prefix = '-'
#         bucket_log_format = '-'
    
#     return bucket_logging, bucket_logging_target, bucket_logging_prefix, bucket_log_format


def get_bucket_static_web_hosting(s3_client, bucket_name):
    try:
        response = s3_client.get_bucket_website(Bucket=bucket_name)

        # 정적 호스팅 구분
        if response.get('RedirectAllRequestsTo') or response.get('IndexDocument'):
            bucket_static_web_hosting = 'Enabled'
            # 객체에 대한 요청 리디렉션일 때
            if response.get('RedirectAllRequestsTo'):
                bucket_hosting_type = 'Redirect Request'
                bucket_hosting_host_name = response['RedirectAllRequestsTo']['HostName']
                bucket_hosting_index_doc = '-'
                bucket_hosting_error_doc = '-'
                # 프로토콜 존재할 때
                if response['RedirectAllRequestsTo'].get('Protocol'):
                    bucket_hosting_protocol = response['RedirectAllRequestsTo']['Protocol']
                else:
                    bucket_hosting_protocol = '-'
            
            # 정적 웹 사이트 호스팅일 때 
            else:
                bucket_hosting_type = 'Bucket Hosting'
                bucket_hosting_host_name = '-'
                bucket_hosting_index_doc = response['IndexDocument']['Suffix']
                bucket_hosting_protocol = '-'
                # 에러 페이지 존재할 때
                if response.get('ErrorDocument'):
                    bucket_hosting_error_doc = response['ErrorDocument']['Key']
                else:
                    bucket_hosting_error_doc = '-'
        
        return bucket_static_web_hosting, bucket_hosting_type, bucket_hosting_host_name, bucket_hosting_protocol, bucket_hosting_index_doc, bucket_hosting_error_doc

    except ClientError as e:
        if e.response['Error']['Code'] == 'NoSuchWebsiteConfiguration':
            bucket_static_web_hosting = 'Disabled'
            bucket_hosting_type = '-'
            bucket_hosting_host_name = '-'
            bucket_hosting_protocol = '-'
            bucket_hosting_index_doc = '-'
            bucket_hosting_error_doc = '-'
        return bucket_static_web_hosting, bucket_hosting_type, bucket_hosting_host_name, bucket_hosting_protocol, bucket_hosting_index_doc, bucket_hosting_error_doc


def get_bucket_public_access_block(s3_client, bucket_name):
    response = s3_client.get_public_access_block(Bucket=bucket_name)

    public_access_block_conf = response.get('PublicAccessBlockConfiguration')
    
    # all() -> 모든 요소(BlockPublicAcls, IgnorePublicAcls, BlockPublicPolicy, RestrictPublicBuckets)가 True일 때 True 아니면 False 반환
    if all(value for value in public_access_block_conf.values()):
        bucket_public_access_block = 'Enabled'
        bucket_public_access_block_BlockPublicAcls = 'Enabled'
        bucket_public_access_block_IgnorePublicAcls = 'Enabled'
        bucket_public_access_block_BlockPublicPolicy = 'Enabled'
        bucket_public_access_block_RestrictPublicBuckets = 'Enabled'
    else:
        bucket_public_access_block = 'Diabled'
        bucket_public_access_block_BlockPublicAcls = public_access_block_conf.get('BlockPublicAcls')
        bucket_public_access_block_IgnorePublicAcls = public_access_block_conf.get('IgnorePublicAcls')
        bucket_public_access_block_BlockPublicPolicy = public_access_block_conf.get('BlockPublicPolicy')
        bucket_public_access_block_RestrictPublicBuckets = public_access_block_conf.get('RestrictPublicBuckets')
    
    return bucket_public_access_block, bucket_public_access_block_BlockPublicAcls, bucket_public_access_block_IgnorePublicAcls, bucket_public_access_block_BlockPublicPolicy, bucket_public_access_block_RestrictPublicBuckets


def get_bucket_lifecycle(s3_client, bucket_name):
    # Lifecycle 미존재 시, 예외 처리
    try:
        response = s3_client.get_bucket_lifecycle_configuration(Bucket=bucket_name)
        
        bucket_lifecycle = 'Exist'
        bucket_lifecycle_info = {
            'rule_name': [],
            'rule_status': [],
            # 'rule_scope': [],
            'rule_expiration_day': [],
            'rule_transitions_day': [],
            'rule_transitions_storageclass': []
        }

        for rule in response['Rules']:
            # print(rule)
            bucket_lifecycle_info['rule_name'].append(rule['ID'])
            bucket_lifecycle_info['rule_status'].append(rule['Status'])
            # bucket_lifecycle_info['rule_scope'].append(rule['Filter'])
            if 'Expiration' in rule:
                bucket_lifecycle_info['rule_expiration_day'].append(str(rule['Expiration']['Days'])) # Days는 Int형이라 문자열로 변경
                
            if 'Transitions' in rule:
                for transitions in rule['Transitions']:
                    bucket_lifecycle_info['rule_transitions_day'].append(str(transitions['Days']))
                    bucket_lifecycle_info['rule_transitions_storageclass'].append(transitions['StorageClass']) # Days는 Int형이라 문자열로 변경
        
        bucket_lifecycle_rule_name = '\n'.join(bucket_lifecycle_info['rule_name'])
        bucket_lifecycle_rule_status = '\n'.join(bucket_lifecycle_info['rule_status'])
        # bucket_lifecycle_rule_scope = '\n'.join(bucket_lifecycle_info['rule_scope'])
        bucket_lifecycle_rule_expiration_day = '\n'.join(bucket_lifecycle_info['rule_expiration_day'])
        bucket_lifecycle_rule_transitions_day = '\n'.join(bucket_lifecycle_info['rule_transitions_day'])
        bucket_lifecycle_rule_transitions_storageclass = '\n'.join(bucket_lifecycle_info['rule_transitions_storageclass'])

        # print(bucket_lifecycle_rule_name,bucket_lifecycle_rule_status, bucket_lifecycle_rule_expiration_day, bucket_lifecycle_rule_transitions_day, bucket_lifecycle_rule_transitions_storageclass)
        bucket_lifecycle_str = f"{bucket_lifecycle_rule_transitions_day}일 {bucket_lifecycle_rule_transitions_storageclass}로 객체 이동, {bucket_lifecycle_rule_expiration_day}일 객체 만료"

    except ClientError as e:
        if e.response['Error']['Code'] == 'NoSuchLifecycleConfiguration':
            bucket_lifecycle = '-'
            bucket_lifecycle_rule_name = '-'
            bucket_lifecycle_rule_status = '-'
            # bucket_lifecycle_rule_scope = '-'
            bucket_lifecycle_rule_expiration_day = '-'
            bucket_lifecycle_rule_transitions_day = '-'
            bucket_lifecycle_str = '-' 
        else:
            print('Unknown Error...')
            print(e.response)
    
    return bucket_lifecycle, bucket_lifecycle_rule_name, bucket_lifecycle_rule_status, bucket_lifecycle_str#, bucket_lifecycle_rule_scope, bucket_lifecycle_rule_expiration_day, bucket_lifecycle_rule_transitions_day, bucket_lifecycle_rule_transitions_storageclass
    
    
def export_s3_info_to_excel(workbook, s3_client):
    s3_info = s3_client.list_buckets()
    
    # S3 시트 생성
    worksheet = workbook.create_sheet('S3')
    
    # S3 열 정보 추가
    worksheet.append(['S3'])
    worksheet.cell(1, 1).font = front_header_font

    s3_headers = [
    # S3 기본 정보
    'Bucket', 'Encryption Type', 'Encryption Key', 'Tags', 'Bucket Versioning', 
    # S3 액세스 로깅
    'Server Access Logging', 'Target Bucket', 'Target Bucket Prefix', 'Target Bucket Log Format',
    # S3 웹 호스팅
    'Static Web Site Hosting', 'Hosting Type', 'Host', 'Hosting Protocol', 'Hosting Index Document', 'Hosting Error Document', 
    # S3 퍼블릭 접근 차단
    'Public Access Block', 'BlockPublicAcls', 'IgnorePublicAcls', 'BlockPublicPolicy', 'RestrictPublicBuckets',
    # S3 수명주기
    'Lifecycle', 'Lifecycle Rule', 'Lifecycle Rule Status', 'Transition and Expiration'#, 'Lifecycle Rule Scope', 'Lifecycle Rule Expiration Day', 'Lifecycle Rule Transitions Day', 'Lifecycle Rule Transitions StorageClass'
    ]

    for col_num, header in enumerate(s3_headers,1):
        worksheet.cell(2, col_num, value=header).font = header_font
        worksheet.cell(2, col_num, value=header).fill = header_fill
        worksheet.cell(2, col_num, value=header).alignment = header_alignment
        worksheet.cell(2, col_num, value=header).border = header_border
    
    # 현재 행 위치
    header_row = worksheet.max_row
    
    # auto_filter 적용
    worksheet.auto_filter.ref = f"A{header_row}:{chr(64 + len(s3_headers))}{header_row}"

    # S3 정보를 엑셀에 쓰기
    for bucket in s3_info['Buckets']:
        bucket_name = bucket['Name']
        bucket_enc_type, bucket_enc_key = get_bucket_enc(s3_client, bucket_name)
        bucket_tags = get_bucket_tags(s3_client, bucket_name)
        bucket_versioning = get_bucket_versioning(s3_client, bucket_name)
        bucket_logging, bucket_logging_target, bucket_logging_prefix, bucket_log_format = get_bucket_logging(s3_client, bucket_name)
        bucket_static_web_hosting, bucket_hosting_type, bucket_hosting_host_name, bucket_hosting_protocol, bucket_hosting_index_doc, bucket_hosting_error_doc = get_bucket_static_web_hosting(s3_client, bucket_name)
        bucket_public_access_block, bucket_public_access_block_BlockPublicAcls, bucket_public_access_block_IgnorePublicAcls, bucket_public_access_block_BlockPublicPolicy, bucket_public_access_block_RestrictPublicBuckets = get_bucket_public_access_block(s3_client, bucket_name)
        bucket_lifecycle, bucket_lifecycle_rule_name, bucket_lifecycle_rule_status, bucket_lifecycle_str = get_bucket_lifecycle(s3_client, bucket_name) #, bucket_lifecycle_rule_expiration_day, bucket_lifecycle_rule_transitions_day, bucket_lifecycle_rule_transitions_storageclass, bucket_lifecycle_rule_scope,
        # bucket_tags = bucket_tags.replace(', ', '\n') - 불필요

        variables = [
            bucket_name, bucket_enc_type, bucket_enc_key, bucket_tags, bucket_versioning, 
            bucket_logging, bucket_logging_target, bucket_logging_prefix, bucket_log_format,
            bucket_static_web_hosting, bucket_hosting_type, bucket_hosting_host_name, bucket_hosting_protocol, bucket_hosting_index_doc, bucket_hosting_error_doc,
            bucket_public_access_block, str(bucket_public_access_block_BlockPublicAcls), str(bucket_public_access_block_IgnorePublicAcls), str(bucket_public_access_block_BlockPublicPolicy), str(bucket_public_access_block_RestrictPublicBuckets),
            bucket_lifecycle, bucket_lifecycle_rule_name, bucket_lifecycle_rule_status, bucket_lifecycle_str#, bucket_lifecycle_rule_expiration_day, bucket_lifecycle_rule_transitions_day, bucket_lifecycle_rule_transitions_storageclass, bucket_lifecycle_rule_scope
        ]
        
        # 시트에 데이터 쓰기
        worksheet.append(variables)
        
        # 모든 셀 텍스트 높이 가운데 맞춤
        for index, value in enumerate(variables, start=1):
            # '\n' 즉, 개행이 포함되어 있으면 즉, 셀 값이 다중값이면 텍스트 자동 줄바꿈
            if '\n' in value:
                cell = worksheet.cell(row=worksheet.max_row, column=index)
                cell.alignment = multiple_content_alignment
                cell.border = content_border
            else:
                cell = worksheet.cell(row=worksheet.max_row, column=index)
                cell.alignment = content_alignment
                cell.border = content_border