# 개선 : 수명주기 관리, DNS 이름
from botocore.exceptions import ClientError
from conf import convert
from conf.sheet_style import front_header_font,header_font,header_fill,header_alignment,content_alignment,multiple_content_alignment,content_border,header_border

def get_efs_tags(efs_tags):
    efs_tag_list = []
    
    # EFS 태그 미존재 시 예외 처리
    if len(efs_tags) == 0:
        efs_tag_list = '-'
    else:
        for tag in efs_tags:
            efs_tag_list.append(f"{tag['Key']} : {tag['Value']}")

        efs_tag_list = '\n'.join(efs_tag_list)

    return efs_tag_list


def get_efs_network_info(efs_client, efs_id):
    response = efs_client.describe_mount_targets(FileSystemId=efs_id)
    
    #EFS 네트워크 정보가 저장될 딕셔너리 선언
    efs_network_info = {
            'efs_ip': [],
            'efs_subnet_id': [],
            'efs_az': [],
            'efs_sg_id': []
    }

    # response['MountTargets'] 의 AvailabilityZoneName 값 기준으로 오름차순 정렬(sort)하여 출력
    for efs_mnt_tg_info in sorted(response['MountTargets'], key=lambda subnet_zone: (subnet_zone['AvailabilityZoneName'])):
        efs_mnt_tg_id = efs_mnt_tg_info['MountTargetId']
        efs_network_info['efs_ip'].append(efs_mnt_tg_info['IpAddress'])
        efs_network_info['efs_subnet_id'].append(efs_mnt_tg_info['SubnetId'])
        efs_network_info['efs_az'].append(efs_mnt_tg_info['AvailabilityZoneName'])
        efs_network_info['efs_sg_id'].append(get_efs_sg(efs_client, efs_mnt_tg_id))
    
    efs_ip = efs_network_info['efs_ip']
    efs_subnet_id  = efs_network_info['efs_subnet_id']
    efs_az = efs_network_info['efs_az']
    efs_sg_id = list(set(efs_network_info['efs_sg_id'])) #보안 그룹 ID는 중복 가능하므로 SET 설정 후, 다시 List 형식으로 변환

    return efs_ip, efs_subnet_id, efs_az, efs_sg_id


# EFS 마운트 타겟 ID 정보로 Security Group ID 정보 확인
def get_efs_sg(efs_client, efs_mnt_tg_id):
    response = efs_client.describe_mount_target_security_groups(MountTargetId=efs_mnt_tg_id)
    
    efs_sg_id =  response['SecurityGroups'][0]
    
    return efs_sg_id


def export_efs_info_to_excel(workbook, efs_client, ec2_client):
    efs_info = efs_client.describe_file_systems()
    
    # EFS 시트 생성
    worksheet = workbook.create_sheet('EFS')
    
    # EFS 열 정보 추가
    worksheet.append(['EFS'])
    worksheet.cell(1, 1).font = front_header_font

    efs_headers = [
    # EFS 기본 정보
    'FileSystem', 'FileSystem ID', 'Availability Zone', 'Subntet', 'FileSystem IP', 'Security Group', 'Encryption Type', 'Encryption Key', 'Tags',
    'Throughput Mode', 'Performance Mode', 'FileSystem Protection'
    ]

    for col_num, header in enumerate(efs_headers,1):
        worksheet.cell(2, col_num, value=header).font = header_font
        worksheet.cell(2, col_num, value=header).fill = header_fill
        worksheet.cell(2, col_num, value=header).alignment = header_alignment
        worksheet.cell(2, col_num, value=header).border = header_border

    # 현재 행 위치
    header_row = worksheet.max_row
    
    # auto_filter 적용
    worksheet.auto_filter.ref = f"A{header_row}:{chr(64 + len(efs_headers))}{header_row}"

    for efs in efs_info['FileSystems']:
        efs_name = efs['Name']
        efs_id = efs['FileSystemId']
        efs_throughputmode = efs['ThroughputMode']
        efs_performancemode = efs['PerformanceMode']
        efs_filesystem_protection = efs['FileSystemProtection']['ReplicationOverwriteProtection']
        efs_tag = get_efs_tags(efs['Tags'])

        if efs['Encrypted'] == False:
            efs_enc = 'Disabled'
            efs_enc_key = '-'
        else:
            efs_enc = 'Enabled'
            efs_enc_key = efs['KmsKeyId']

        # EFS 네트워크 정보 구한 후, Security Group/Subnet ID를 Name으로 변환
        efs_ip, efs_subnet_id, efs_az, efs_sg_id = get_efs_network_info(efs_client, efs_id)

        efs_sg_name = []
        efs_subnet_name = []
        
        for sg_id in efs_sg_id:
            result = convert.sg_info(ec2_client, sg_id)
            efs_sg_name.append(result)

        for subnet_id in efs_subnet_id:
            result = convert.subnet_info(ec2_client, subnet_id)
            efs_subnet_name.append(result)
        
        # 리스트를 문자열 변수로 변환 후 ,으로 연결
        efs_az = '\n'.join(efs_az)
        efs_subnet_name = '\n'.join(efs_subnet_name)
        efs_ip = '\n'.join(efs_ip)
        efs_sg_name = '\n'.join(efs_sg_name)
        
        # 시트에 데이터 쓰기
        variables = [
                efs_name, efs_id, efs_az, efs_subnet_name, efs_ip, efs_sg_name, efs_enc, efs_enc_key, efs_tag,
                efs_throughputmode, efs_performancemode, efs_filesystem_protection
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