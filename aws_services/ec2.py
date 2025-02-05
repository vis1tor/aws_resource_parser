from botocore.exceptions import ClientError
from conf import convert
from conf.sheet_style import front_header_font,header_font,header_fill,header_alignment,content_alignment,multiple_content_alignment,content_border,header_border

# 태그 정보 정리
def get_ec2_tags(ec2_tags):
    ec2_tag_list = []
    
    # EC2 태그 미존재 시 예외 처리
    if len(ec2_tags) == 0:
        ec2_tag_list = '-'
    else:
        for tag in ec2_tags:
            ec2_tag_list.append(f"{tag['Key']} : {tag['Value']}")
            
            if tag['Key'] == 'Name':
                ec2_name = tag['Value']

        ec2_tag_list = '\n'.join(ec2_tag_list)

    return ec2_name, ec2_tag_list

# 볼륨 정보 정리
def get_ec2_vol_info(ec2_client, ec2_vol_id, ec2_root_device_name):
    response = ec2_client.describe_volumes(VolumeIds=[ec2_vol_id])

    ec2_vol_name = '-'

    for ec2_vol in response['Volumes']:
        ec2_vol_size = str(ec2_vol['Size'])+' Gib'
        ec2_vol_type = str(ec2_vol['VolumeType'])
        if ec2_vol['Encrypted'] == True:
            ec2_vol_enc = str(ec2_vol['Encrypted'])
            ec2_vol_kms_id = ec2_vol['KmsKeyId']
        else:
            ec2_vol_enc = str(ec2_vol['Encrypted'])
            ec2_vol_kms_id = '-'

        for ec2_vol_attach in ec2_vol['Attachments']:
            if ec2_root_device_name == ec2_vol_attach['Device']:
                ec2_device_name = ec2_vol_attach['Device']+' - root'
                ec2_vol_id = ec2_vol_attach['VolumeId']+' - root'

                # EC2 태그 미존재 시 예외 처리
                try:
                    for tag in ec2_vol['Tags']:
                        if tag['Key'] == 'Name':
                            ec2_vol_name = tag['Value'] +' - root'
                except:
                    ec2_vol_name = '-'
            else:
                ec2_device_name = ec2_vol_attach['Device']
                ec2_vol_id = ec2_vol_attach['VolumeId']
                
                try:
                    for tag in ec2_vol['Tags']:
                        if tag['Key'] == 'Name':
                            ec2_vol_name = tag['Value']
                except:
                    ec2_vol_name = '-'
    
    return ec2_vol_name, ec2_vol_id, ec2_device_name, ec2_vol_enc, ec2_vol_kms_id, ec2_vol_size, ec2_vol_type

# AMI 정보 정리
def get_ec2_ami_info(ec2_client, ec2_image_id):
    response = ec2_client.describe_images(ImageIds=[ec2_image_id])
    
    try:
        ec2_image = response['Images'][0]['Name']
    except:
        ec2_image = "Not Exist"
    
    return ec2_image

# vCPU, 메모리 정보 정리
def get_ec2_mem_info(ec2_client, ec2_instance_type):
    response = ec2_client.describe_instance_types(InstanceTypes=[ec2_instance_type])
    
    ec2_cpu_size = str(response['InstanceTypes'][0]['VCpuInfo']['DefaultVCpus'])
    ec2_mem_size = str(int(response['InstanceTypes'][0]['MemoryInfo']['SizeInMiB'] / 1024)) + ' Gib'
    
    return ec2_cpu_size, ec2_mem_size

def export_ec2_info_to_excel(workbook, ec2_client):
    ec2_info = ec2_client.describe_instances()

    # EC2 시트 생성
    worksheet = workbook.create_sheet('EC2')
    
    # EC2 열 정보 추가
    worksheet.append(['EC2'])
    worksheet.cell(1, 1).font = front_header_font

    ec2_headers = [
    # EC2 기본 정보
    'Instance', 'Instance ID', 'AMI', 'AMI ID', 'Instance Type', 'Key Pair',
    # EC2 네트워크 정보
    'VPC', 'Subnet', 'Availability Zone', 'Private IP', 'Public IP', 'Security Group',
    # ETC...
    'IAM Role', 'Tags',
    # vCPU, 메모리 정보
    'vCPU', 'Memory Size',
    # Volume 정보
    'Volume Name', 'Volume ID', 'Volume Device', 'Volume Type', 'Volume Size', 'Volume Encrypted', 'KMS ID'
    ]

    for col_num, header in enumerate(ec2_headers,1):
        worksheet.cell(2, col_num, value=header).font = header_font
        worksheet.cell(2, col_num, value=header).fill = header_fill
        worksheet.cell(2, col_num, value=header).alignment = header_alignment
        worksheet.cell(2, col_num, value=header).border = header_border
    
    # 현재 행 위치
    header_row = worksheet.max_row
    
    # auto_filter 적용
    worksheet.auto_filter.ref = f"A{header_row}:{chr(64 + len(ec2_headers))}{header_row}"

    for ec2_list in ec2_info['Reservations']:
        for ec2 in ec2_list['Instances']:
            ec2_name, ec2_tags = get_ec2_tags(ec2['Tags'])
            ec2_instance_id = ec2['InstanceId']
            
            # AMI 정보
            ec2_image_id = ec2['ImageId']
            ec2_image = get_ec2_ami_info(ec2_client, ec2_image_id)

            ec2_instance_type = ec2['InstanceType']
            ec2_az = ec2['Placement']['AvailabilityZone']
            ec2_root_device_name = ec2['RootDeviceName']
            
            if 'KeyName' in ec2:
                ec2_key_name = ec2['KeyName']
            else:
                ec2_key_name = '-'

            if 'IamInstanceProfile' in ec2:
                ec2_iam_role = ec2['IamInstanceProfile']['Arn'].split('/')[1]
            else:
                ec2_iam_role = '-'
            
            ec2_vpc_list = []
            ec2_subnet_list = []
            ec2_sg_list = []
            ec2_pri_ip_list = []
            ec2_pub_ip_list = []
            ec2_vol_name_list = []
            ec2_vol_id_list = []
            ec2_device_name_list = []
            ec2_vol_enc_list = []
            ec2_vol_kms_id_list = []
            ec2_vol_size_list = []
            ec2_vol_type_list = []

            # 볼륨 정보
            for ec2_vol in ec2['BlockDeviceMappings']:
                ec2_vol_name, ec2_vol_id, ec2_device_name, ec2_vol_enc, ec2_vol_kms_id, ec2_vol_size, ec2_vol_type = get_ec2_vol_info(ec2_client,ec2_vol['Ebs']['VolumeId'], ec2_root_device_name)
                ec2_vol_name_list.append(ec2_vol_name)
                ec2_vol_id_list.append(ec2_vol_id)
                ec2_device_name_list.append(ec2_device_name)
                ec2_vol_enc_list.append(ec2_vol_enc)
                ec2_vol_kms_id_list.append(ec2_vol_kms_id)
                ec2_vol_size_list.append(ec2_vol_size)
                ec2_vol_type_list.append(ec2_vol_type)
            
            # vCPU, 메모리 정보
            ec2_cpu_size, ec2_mem_size = get_ec2_mem_info(ec2_client, ec2_instance_type)

            # 네트워크 정보
            for ec2_network in ec2['NetworkInterfaces']:
                ec2_vpc_list.append(convert.vpc_info(ec2_client,ec2_network['VpcId']))
                ec2_subnet_list.append(convert.subnet_info(ec2_client,ec2_network['SubnetId']))
                
                for ec2_sg in ec2_network['Groups']:
                    ec2_sg_list.append(ec2_sg['GroupName']+'('+ec2_sg['GroupId']+')')
                
                for ec2_private in ec2_network['PrivateIpAddresses']:
                    ec2_pri_ip_list.append(ec2_private['PrivateIpAddress'])
                    if 'Association' in ec2_private:
                        ec2_pub_ip_list.append(ec2_private['Association']['PublicIp'])
                    else:
                        ec2_pub_ip_list.append('-')

            # Value 2개 이상 시, 자동 개행
            ec2_vpc_list = '\n'.join(ec2_vpc_list)
            ec2_subnet_list = '\n'.join(ec2_subnet_list)            
            ec2_sg_list = '\n'.join(ec2_sg_list)
            ec2_pri_ip_list = '\n'.join(ec2_pri_ip_list)
            ec2_pub_ip_list = '\n'.join(ec2_pub_ip_list)
            ec2_vol_name_list = '\n'.join(ec2_vol_name_list)
            ec2_vol_id_list = '\n'.join(ec2_vol_id_list)
            ec2_device_name_list = '\n'.join(ec2_device_name_list)
            ec2_vol_enc_list = '\n'.join(ec2_vol_enc_list)
            ec2_vol_kms_id_list = '\n'.join(ec2_vol_kms_id_list)
            ec2_vol_size_list = '\n'.join(ec2_vol_size_list)
            ec2_vol_type_list = '\n'.join(ec2_vol_type_list)

            
            variables = [
                # EC2 기본 정보
                ec2_name, ec2_instance_id, ec2_image, ec2_image_id, ec2_instance_type, ec2_key_name,
                # EC2 네트워크 정보
                ec2_vpc_list, ec2_subnet_list, ec2_az, ec2_pri_ip_list, ec2_pub_ip_list, ec2_sg_list,
                # ETC...
                ec2_iam_role, ec2_tags,
                # vCPU, 메모리 정보
                ec2_cpu_size, ec2_mem_size,
                # Volume 정보
                ec2_vol_name_list, ec2_vol_id_list, ec2_device_name_list, ec2_vol_type_list, ec2_vol_size_list, ec2_vol_enc_list, ec2_vol_kms_id_list
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