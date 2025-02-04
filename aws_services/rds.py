# 시작 템플릿 추가 필요
from botocore.exceptions import ClientError
from conf import convert
from conf.sheet_style import front_header_font,header_font,header_fill,header_alignment,content_alignment,multiple_content_alignment,content_border,header_border

def get_rds_tags(rds_tags):
    rds_tag_list = []
    
    # RDS 태그 미존재 시 예외 처리
    if len(rds_tags) == 0:
        rds_tag_list = '-'
    else:
        for tag in rds_tags:
            rds_tag_list.append(f"{tag['Key']} : {tag['Value']}")

        rds_tag_list = '\n'.join(rds_tag_list)

    return rds_tag_list

def export_rds_info_to_excel(workbook, rds_client, ec2_client):
#====================================== RDS Instance Section ======================================
    # RDS 클러스터 목록 조회
    # rds_cluster_list = rds_client.describe_db_clusters()
    
    # RDS 인스턴스 목록 조회
    rds_instance_list = rds_client.describe_db_instances()
    
    # RDS 시트 생성
    worksheet = workbook.create_sheet('RDS')
    
    # RDS 열 정보 추가
    worksheet.append(['RDS'])
    worksheet.cell(1, 1).font = front_header_font
    
    rds_instance_headers = [
    # RDS 기본 정보
    'Instance', 'Class', 'Engine', 'Engine Version', 'Master User Name', 'Endpoint', 'Port', 'Tags',
    # RDS 네트워크 정보
    'VPC', 'Subnet Group & Subnet', 'Multi AZ', 'Availability Zone', 'Security Group',
    # 스토리지 정보
    'Storage Type', 'Storage Size', 'Storage Iops', 'Storage Throughput', 'Storage Encrypted', 'Storage KMS Key',
    # 파라미터 정보
    'Parameter Group'
    ]

    for col_num, header in enumerate(rds_instance_headers,1):
        worksheet.cell(2, col_num, value=header).font = header_font
        worksheet.cell(2, col_num, value=header).fill = header_fill
        worksheet.cell(2, col_num, value=header).alignment = header_alignment
        worksheet.cell(2, col_num, value=header).border = header_border

    # 현재 행 위치
    header_row = worksheet.max_row
    
    # auto_filter 적용
    worksheet.auto_filter.ref = f"A{header_row}:{chr(64 + len(rds_instance_headers))}{header_row}"

    for rds_instance in rds_instance_list['DBInstances']:
        # RDS 기본 정보
        rds_instance_name = rds_instance['DBInstanceIdentifier']
        rds_instance_class = rds_instance['DBInstanceClass']
        rds_instance_engine = rds_instance['Engine']
        rds_instance_engine_version = rds_instance['EngineVersion']
        rds_instance_master = rds_instance['MasterUsername']
        rds_instance_endpoint = rds_instance['Endpoint']['Address']
        rds_instance_port = str(rds_instance['Endpoint']['Port'])
        rds_instance_tags = get_rds_tags(rds_instance['TagList'])

        # 네트워크 정보
        rds_instance_vpc_id = rds_instance['DBSubnetGroup']['VpcId']
        rds_instance_vpc_name = convert.vpc_info(ec2_client, rds_instance_vpc_id) + '('+rds_instance_vpc_id+')'

        rds_instance_sbng = rds_instance['DBSubnetGroup']['DBSubnetGroupName']
        rds_instance_subnet_name = []
        
        for subnet in rds_instance['DBSubnetGroup']['Subnets']:
            rds_instance_subnet_id = subnet['SubnetIdentifier']
            rds_instance_subnet_name.append('- ' + convert.subnet_info(ec2_client, rds_instance_subnet_id) + '('+rds_instance_subnet_id+')')
        
        rds_instance_subnet_name.sort()
        rds_instance_subnet_name = '\n'.join(rds_instance_subnet_name)

        rds_instance_multi_az = str(rds_instance['MultiAZ'])
        rds_instance_az =  rds_instance['AvailabilityZone']

        rds_instance_sg_name = []
        for sg in rds_instance['VpcSecurityGroups']:
            rds_instance_sg_id = sg['VpcSecurityGroupId']
            rds_instance_sg_name.append(convert.sg_info(ec2_client, rds_instance_sg_id) + '('+rds_instance_sg_id+')')
        
        rds_instance_sg_name.sort()
        rds_instance_sg_name = '\n'.join(rds_instance_sg_name)

        # 파라미터 정보
        rds_instance_parg = rds_instance['DBParameterGroups'][0]['DBParameterGroupName']

        # 스토리지 정보
        rds_instance_storage_type = rds_instance['StorageType']
        rds_instance_storage_size = str(rds_instance['AllocatedStorage'])
        rds_instance_storage_iops = str(rds_instance['Iops'])
        rds_instance_storage_throughput = str(rds_instance['StorageThroughput'])
        rds_instance_storage_encrypted = str(rds_instance['StorageEncrypted'])
        try:
            rds_instance_storage_kms_Key = rds_instance['KmsKeyId']
        except:
            rds_instance_storage_kms_Key = '-'

        variables = [
            # RDS 기본 정보
            rds_instance_name, rds_instance_class, rds_instance_engine, rds_instance_engine_version, rds_instance_master, rds_instance_endpoint, rds_instance_port, rds_instance_tags,
            # 네트워크 정보
            rds_instance_vpc_name, rds_instance_sbng+'\n'+ rds_instance_subnet_name, rds_instance_multi_az, rds_instance_az, rds_instance_sg_name,
            # 스토리지 정보
            rds_instance_storage_type, str(rds_instance_storage_size), rds_instance_storage_iops, rds_instance_storage_throughput, rds_instance_storage_encrypted, rds_instance_storage_kms_Key,
            # 파라미터 정보
            rds_instance_parg
        ]

        worksheet.append(variables)
            
        # 모든 셀 텍스트 높이 가운데 맞춤
        for index, value in enumerate(variables, start=1):
            cell = worksheet.cell(row=worksheet.max_row, column=index)
            cell.alignment = content_alignment
            cell.border = content_border
            # '\n' 즉, 개행이 포함되어 있으면 즉, 셀 값이 다중값이면 텍스트 자동 줄바꿈
            if '\n' in value:
                cell = worksheet.cell(row=worksheet.max_row, column=index)
                cell.alignment = multiple_content_alignment
                cell.border = content_border