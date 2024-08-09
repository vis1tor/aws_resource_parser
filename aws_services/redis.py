from botocore.exceptions import ClientError
from conf import convert
from conf.sheet_style import front_header_font,header_font,header_fill,header_alignment,content_alignment,multiple_content_alignment,content_border,header_border

def get_redis_tags(redis_client, redis_arn):
    response = redis_client.list_tags_for_resource(ResourceName=redis_arn)

    redis_tag_list = []

    # Redis 태그 미존재 시 예외 처리
    if len(response['TagList']) == 0:
        redis_tag_list = '-'
    else:
        for tag in response['TagList']:
            redis_tag_list.append(f"{tag['Key']} : {tag['Value']}")
        
        redis_tag_list = '\n'.join(redis_tag_list)
    
    return redis_tag_list

def get_redis_subnet_groups(ec2_client, redis_client, redis_sbng):
    response = redis_client.describe_cache_subnet_groups(CacheSubnetGroupName=redis_sbng)
    
    redis_vpc_id = response['CacheSubnetGroups'][0]['VpcId']
    redis_vpc = convert.vpc_info(ec2_client, redis_vpc_id) + '('+redis_vpc_id+')'
    
    redis_subnet_name = []

    for subnet in response['CacheSubnetGroups'][0]['Subnets']:
        redis_subnet_id = subnet['SubnetIdentifier']
        redis_subnet_name.append('- ' + convert.subnet_info(ec2_client, redis_subnet_id) + '('+redis_subnet_id+')')
    
    redis_subnet_name.sort()
    redis_subnet_name = '\n'.join(redis_subnet_name)

    return redis_vpc, redis_subnet_name

def export_redis_info_to_excel(workbook, redis_client, ec2_client):
    redis_info = redis_client.describe_cache_clusters()
    
    # EC2 시트 생성
    worksheet = workbook.create_sheet('Redis')
    
    # EC2 열 정보 추가
    worksheet.append(['Redis'])
    worksheet.cell(1, 1).font = front_header_font

    redis_headers = [
        # Redis 기본 정보
        'Name', 'Type', 'Engine', 'Engine Version', 'Tags',
        # Redis 네트워크 정보
        'VPC','Subnet Group', 'Prefer AZ', 'Security Group',
        # Redis 파라미터 정보
        'Parameter Group',
        # 기타 정보
        'Prefer Maintain', 'Node' 
    ]

    for col_num, header in enumerate(redis_headers,1):
        worksheet.cell(2, col_num, value=header).font = header_font
        worksheet.cell(2, col_num, value=header).fill = header_fill
        worksheet.cell(2, col_num, value=header).alignment = header_alignment
        worksheet.cell(2, col_num, value=header).border = header_border
    
    # 현재 행 위치
    header_row = worksheet.max_row
    
    # auto_filter 적용
    worksheet.auto_filter.ref = f"A{header_row}:{chr(64 + len(redis_headers))}{header_row}"
    
    for redis_list in redis_info['CacheClusters']:
        redis_arn = redis_list['ARN']
        redis_name = redis_list['CacheClusterId']
        redis_type = redis_list['CacheNodeType']
        redis_engine = redis_list['Engine']
        redis_engine_version = redis_list['EngineVersion']
        redis_node = redis_list['NumCacheNodes']
        redis_prefer_az = redis_list['PreferredAvailabilityZone']
        redis_prefer_maintain = redis_list['PreferredMaintenanceWindow']
        redis_parg = redis_list['CacheParameterGroup']['CacheParameterGroupName']
        redis_sbng = redis_list['CacheSubnetGroupName']
        redis_vpc, redis_subnet_name = get_redis_subnet_groups(ec2_client, redis_client, redis_sbng)
        redis_tags = get_redis_tags(redis_client, redis_arn)
        
        redis_sg_name = []
        for sg in redis_list['SecurityGroups']:
            redis_sg_id = sg['SecurityGroupId']
            redis_sg_name.append(convert.sg_info(ec2_client, redis_sg_id) + '('+redis_sg_id+')')
        
        redis_sg_name.sort()
        redis_sg_name = '\n'.join(redis_sg_name)

        variables = [
            # Redis 기본 정보
            redis_name, redis_type, redis_engine, redis_engine_version, redis_tags,
            # Redis 네트워크 정보
            redis_vpc, redis_sbng +'\n'+ redis_subnet_name, redis_prefer_az, redis_sg_name,
            # Redis 파라미터 정보
            redis_parg,
            # 기타 정보
            redis_prefer_maintain, str(redis_node)
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