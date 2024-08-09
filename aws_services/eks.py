# 시작 템플릿 추가 필요
from botocore.exceptions import ClientError
from conf import convert
from conf.sheet_style import front_header_font,header_font,header_fill,header_alignment,content_alignment,multiple_content_alignment,content_border,header_border

def get_launch_template_info(ec2_client, lt_id):
    response = ec2_client.describe_launch_templates(LaunchTemplateIds=[lt_id])
    
    if 'Tags' in response['LaunchTemplates'][0]:
        lt_tags = get_lt_tags(response['LaunchTemplates'][0]['Tags'])
    else:
        lt_tags = '-'
    
    return lt_tags

def get_launch_template_version_info(ec2_client, lt_id, lt_version):
    response = ec2_client.describe_launch_template_versions(LaunchTemplateId=lt_id, Versions=[lt_version])
    
    lt_instance_tag_list = []
    lt_volume_tag_list = []
    lt_eni_tag_list = []

    # LT 리소스 태그 정보
    for tag in response['LaunchTemplateVersions'][0]['LaunchTemplateData']['TagSpecifications']:
        if tag['ResourceType'] == 'instance':
            lt_instance_tag_list = get_lt_tags(tag['Tags'])
        elif tag['ResourceType'] == 'volume':
            lt_volume_tag_list = get_lt_tags(tag['Tags'])
        elif tag['ResourceType'] == 'network-interface':
            lt_eni_tag_list = get_lt_tags(tag['Tags'])
    
    # LT 보안그룹 정보
    lt_sg_list = []
    try:
        for eni in response['LaunchTemplateVersions'][0]['LaunchTemplateData']['NetworkInterfaces']: # NetworkInterfaces 키 존재 시
            for sg_id in eni['Groups']:
                lt_sg_list.append(convert.sg_info(ec2_client, sg_id)+'('+sg_id+')')
    except:
        pass
    
    try:
        for sg_id in response['LaunchTemplateVersions'][0]['LaunchTemplateData']['SecurityGroupIds']: # SecurityGroupIds 키 존재 시
            lt_sg_list.append(convert.sg_info(ec2_client, sg_id)+'('+sg_id+')')
    except:
        pass

    # LT 인스턴트 타입 정보
    lt_instance_type = response['LaunchTemplateVersions'][0]['LaunchTemplateData']['InstanceType']
    
    # LT 이미지 정보
    try:
        lt_image_id = response['LaunchTemplateVersions'][0]['LaunchTemplateData']['ImageId']
    except:
        lt_image_id = '-'
    
    try:
        lt_image = get_ec2_ami_info(ec2_client, lt_image_id) + '('+lt_image_id+')'
    except:
        lt_image = '-'

    # LT 스토리지 정보
    lt_vol_device_name_list = []
    lt_vol_enc_list = []
    lt_vol_kms_id_list = []
    lt_vol_size_list = []
    lt_vol_type_list = [] 

    if 'BlockDeviceMappings' in response['LaunchTemplateVersions'][0]['LaunchTemplateData']:
        for device in response['LaunchTemplateVersions'][0]['LaunchTemplateData']['BlockDeviceMappings']:
            lt_vol_device_name_list.append(device['DeviceName'])
            lt_vol_size_list.append(str(device['Ebs']['VolumeSize']))
            try:
                lt_vol_type_list.append(device['Ebs']['VolumeType'])
            except:
                lt_vol_type_list.append('-')
            try:
                lt_vol_enc_list.append(str(device['Ebs']['Encrypted']))
            except:
                lt_vol_enc_list.append('-')
            try:
                lt_vol_kms_id_list.append(device['Ebs']['KmsKeyId'])
            except:
                lt_vol_kms_id_list.append('-')

    lt_sg_list = '\n'.join(lt_sg_list)
    lt_vol_device_name_list = '\n'.join(lt_vol_device_name_list)
    lt_vol_size_list = '\n'.join(lt_vol_size_list)
    lt_vol_type_list = '\n'.join(lt_vol_type_list)
    lt_vol_enc_list = '\n'.join(lt_vol_enc_list)
    lt_vol_kms_id_list = '\n'.join(lt_vol_kms_id_list)

    return lt_image, lt_instance_type, lt_vol_device_name_list, lt_vol_size_list, lt_vol_type_list, lt_vol_enc_list, lt_vol_kms_id_list, lt_sg_list, lt_instance_tag_list, lt_volume_tag_list, lt_eni_tag_list

def get_ec2_ami_info(ec2_client, ec2_image_id):
    response = ec2_client.describe_images(ImageIds=[ec2_image_id])
    
    ec2_image = response['Images'][0]['Name']
    
    return ec2_image

# EKS 태그 정보 정리 - EKS는 태그가 딕셔너리 형태이므로 함수 내용 일부 수정
def get_eks_tags(eks_tags):
    eks_tag_list = []
    
    # EKS 태그 미존재 시 예외 처리
    if len(eks_tags) == 0:
        eks_tag_list = '-'
    else:
        for key,value in eks_tags.items():
            eks_tag_list.append(f"{key} : {value}")

        eks_tag_list = '\n'.join(eks_tag_list)

    return eks_tag_list

def get_lt_tags(lt_tags):
    lt_tag_list = []
    
    # LT 태그 미존재 시 예외 처리
    if len(lt_tags) == 0:
        lt_tag_list = '-'
    else:
        for tag in lt_tags:
            lt_tag_list.append(f"{tag['Key']} : {tag['Value']}")

        lt_tag_list = '\n'.join(lt_tag_list)

    return lt_tag_list

def export_eks_info_to_excel(workbook, eks_client, ec2_client):
#====================================== EKS Section ======================================
    # EKS 클러스터 목록 조회
    eks_list = eks_client.list_clusters()

    # EKS 시트 생성
    worksheet = workbook.create_sheet('EKS')
    
    # EKS 열 정보 추가
    worksheet.append(['EKS'])
    worksheet.cell(1, 1).font = front_header_font

    eks_headers = [
        # EKS 기본 정보
        'Cluster', 'K8S Version', #'EKS Platform Version',
        # EKS 네트워크 정보
        'VPC', 'Subnet', 'K8S Service Network', 'Tags', 'Logging Types' , 
        # 값이 길어서 뒤로 뺌
        'Access Auth Mode', 'API Server Endpoint Access ', 'API Server Endpoint Public Access CIDR', 'Cluster Security Group', 'Security Group', 'API Server Endpoint', 'OIDC URL', 'Cluster IAM Role'
    ]

    for col_num, header in enumerate(eks_headers,1):
        worksheet.cell(2, col_num, value=header).font = header_font
        worksheet.cell(2, col_num, value=header).fill = header_fill
        worksheet.cell(2, col_num, value=header).alignment = header_alignment
        worksheet.cell(2, col_num, value=header).border = header_border
    
    # 현재 행 위치
    header_row = worksheet.max_row
    
    # auto_filter 적용
    worksheet.auto_filter.ref = f"A{header_row}:{chr(64 + len(eks_headers))}{header_row}"

    for cluster in eks_list['clusters']:
        eks_info = eks_client.describe_cluster(name=cluster)

        cluster_name = eks_info['cluster']['name']
        k8s_version = eks_info['cluster']['version']
        k8s_svc_network = eks_info['cluster']['kubernetesNetworkConfig']['serviceIpv4Cidr']
        #cluster_platform_version = eks_info['cluster']['platformVersion']
        cluster_ep = eks_info['cluster']['endpoint']
        cluster_role = eks_info['cluster']['roleArn']
        cluster_oidc = eks_info['cluster']['identity']['oidc']['issuer']
        cluster_tag_list = get_eks_tags(eks_info['cluster']['tags'])
        cluster_access_auth_mode = eks_info['cluster']['accessConfig']['authenticationMode']
        cluster_pub_ep_access = str(eks_info['cluster']['resourcesVpcConfig']['endpointPublicAccess'])
        cluster_pri_ep_access = str(eks_info['cluster']['resourcesVpcConfig']['endpointPrivateAccess'])
        cluster_pub_ep_access_cidr = eks_info['cluster']['resourcesVpcConfig']['publicAccessCidrs']

        # 클러스터 로깅 확인
        cluster_logging_type = []
        for logging_status in eks_info['cluster']['logging']['clusterLogging']:
            if logging_status['enabled'] == True:
                for logging_type in logging_status['types']:
                    cluster_logging_type.append(logging_type.capitalize())
            elif logging_status['enabled'] == False and len(logging_status['types']) == 5:
                cluster_logging_type = '-'
        cluster_logging_type = '\n'.join(cluster_logging_type)
        
        # 클러스터 API 서버 앤드포인트 액세스 확인
        if cluster_pub_ep_access and cluster_pri_ep_access == True:
            cluster_ep_access = 'Public & Private'
            if len(cluster_pub_ep_access_cidr) == 0:
                cluster_pub_ep_access_cidr = '-'
        elif cluster_pub_ep_access == True:
            cluster_ep_access = 'Public'
            cluster_pub_ep_access_cidr = '-'
        else:
            cluster_ep_access = 'Private'
            cluster_pub_ep_access_cidr = '-'
        
        vpc_id = eks_info['cluster']['resourcesVpcConfig']['vpcId']
        vpc_name = convert.vpc_info(ec2_client, vpc_id)+'('+vpc_id+')'

        cluster_sg_id = eks_info['cluster']['resourcesVpcConfig']['clusterSecurityGroupId']
        cluster_sg_name = convert.sg_info(ec2_client, cluster_sg_id)+'('+cluster_sg_id+')'

        subnet_name_list = []
        sg_name_list = []
        
        for subnet_id in eks_info['cluster']['resourcesVpcConfig']['subnetIds']:
            subnet_name_list.append(convert.subnet_info(ec2_client, subnet_id)+'('+subnet_id+')')
        
        for sg_id in eks_info['cluster']['resourcesVpcConfig']['securityGroupIds']:
            sg_name_list.append(convert.sg_info(ec2_client, sg_id)+'('+sg_id+')')

        subnet_name_list = '\n'.join(subnet_name_list)
        sg_name_list = '\n'.join(sg_name_list)

        
        variables = [
                # EKS 기본 정보
                cluster_name, k8s_version, #cluster_platform_version,
                # EKS 네트워크 정보
                vpc_name, subnet_name_list, k8s_svc_network, cluster_tag_list, cluster_logging_type,
                # ETC - 값이 길어서 뒤로 뺌
                cluster_access_auth_mode, cluster_ep_access, cluster_pub_ep_access_cidr, cluster_ep, cluster_sg_name, sg_name_list, cluster_oidc, cluster_role,
            ]

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
            
            # 아랫 부분 문제 없으면 추후 제거 필요
            # worksheet.cell(row=worksheet.max_row, column=index).alignment = content_alignment
            # # '\n' 즉, 개행이 포함되어 있으면 즉, 셀 값이 다중값이면 텍스트 자동 줄바꿈
            # if '\n' in value:
            #     worksheet.cell(row=worksheet.max_row, column=index).alignment = multiple_content_alignment

#====================================== Node Group Section ======================================
    # Node Group 열 정보 추가
    worksheet.append([None])
    worksheet.append(['Node Group'])
    max_row = worksheet.max_row
    worksheet.cell(max_row, 1).font = front_header_font

    ng_headers = [
        'Cluster', 'Node Group', 'AMI', #'AMI Type', 
        'Subnet',
        'Capacity Type', 'Tags', 'Desired_size', 'Minimum size', 'Maximum size', 'Launch Template'
    ]

    for col_num, header in enumerate(ng_headers,1):
        worksheet.cell(max_row+1, col_num, value=header).font = header_font
        worksheet.cell(max_row+1, col_num, value=header).fill = header_fill
        worksheet.cell(max_row+1, col_num, value=header).alignment = header_alignment
        worksheet.cell(max_row+1, col_num, value=header).border = header_border

    # Node Group 정보를 엑셀에 쓰기
    # 노드 그룹 목록 조회를 위한 클러스터 반복
    for cluster in eks_list['clusters']:
        ng_list = eks_client.list_nodegroups(clusterName=cluster)
        # 노드 그룹 조회를 위한 노드 그룹 반복
        for ng in ng_list['nodegroups']:
            ng_info = eks_client.describe_nodegroup(clusterName=cluster,nodegroupName=ng)

            ng_name = ng_info['nodegroup']['nodegroupName']
            ng_cluster_name = ng_info['nodegroup']['clusterName']
            ng_image_id = ng_info['nodegroup']['releaseVersion']
            
            if ng_info['nodegroup']['releaseVersion'].startswith('ami-'):
                ng_image = get_ec2_ami_info(ec2_client, ng_info['nodegroup']['releaseVersion']) + '('+ng_image_id+')'
            else:
                ng_image = ng_info['nodegroup']['releaseVersion']

            #ng_ami_type = ng_info['nodegroup']['amiType']
            ng_capacitytype = ng_info['nodegroup']['capacityType']
            ng_desired_size = str(ng_info['nodegroup']['scalingConfig']['desiredSize'])
            ng_min_size = str(ng_info['nodegroup']['scalingConfig']['minSize'])
            ng_max_size = str(ng_info['nodegroup']['scalingConfig']['maxSize'])
            ng_tags = get_eks_tags(ng_info['nodegroup']['tags'])
            
            if 'name' in ng_info['nodegroup']['launchTemplate']: 
                ng_lt = ng_info['nodegroup']['launchTemplate']['name'] + '('+ng_info['nodegroup']['launchTemplate']['id']+')'


            ng_subnet_name_list = []
            
            for subnet_id in ng_info['nodegroup']['subnets']:
                ng_subnet_name_list.append(convert.subnet_info(ec2_client, subnet_id)+'('+subnet_id+')')
            
            ng_subnet_name_list = '\n'.join(ng_subnet_name_list)

            # 시트에 데이터 쓰기
            variables = [
                ng_cluster_name, ng_name, ng_image, #ng_ami_type,
                ng_subnet_name_list,
                ng_capacitytype, ng_tags, ng_desired_size, ng_min_size, ng_max_size, ng_lt
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
#====================================== Launch Template Section ======================================
    # Launch Template 열 정보 추가
    worksheet.append([None])
    worksheet.append(['Launch Template'])
    max_row = worksheet.max_row
    worksheet.cell(max_row, 1).font = front_header_font
    
    lt_headers = [
        'Cluster', 'Node Group', 'Launch Template', 'AMI', 'Instance Type', 'Tags', 
        'Volume Device', 'Volume Type', 'Volume Size', 'Volume Encrypted', 'KMS ID', "Security Group",
        'Instance Tag', 'Volume Tag', 'ENI Tag'
    ]
    
    for col_num, header in enumerate(lt_headers,1):
        worksheet.cell(max_row+1, col_num, value=header).font = header_font
        worksheet.cell(max_row+1, col_num, value=header).fill = header_fill
        worksheet.cell(max_row+1, col_num, value=header).alignment = header_alignment
        worksheet.cell(max_row+1, col_num, value=header).border = header_border

    # Node Group 정보를 엑셀에 쓰기
    # 노드 그룹 목록 조회를 위한 클러스터 반복
    for cluster in eks_list['clusters']:
        ng_list = eks_client.list_nodegroups(clusterName=cluster)
        # 노드 그룹 조회를 위한 노드 그룹 반복
        for ng in ng_list['nodegroups']:
            ng_info = eks_client.describe_nodegroup(clusterName=cluster,nodegroupName=ng)

            lt_ng_name = ng_info['nodegroup']['nodegroupName']
            lt_cluster_name = ng_info['nodegroup']['clusterName']

            if 'name' in ng_info['nodegroup']['launchTemplate']: 
                lt_id = ng_info['nodegroup']['launchTemplate']['id']
                lt_version = ng_info['nodegroup']['launchTemplate']['version']
                lt_name = ng_info['nodegroup']['launchTemplate']['name'] + '('+lt_id+')'
                lt_tags = get_launch_template_info(ec2_client, lt_id)
                lt_image, lt_instance_type, lt_vol_device_name_list, lt_vol_size_list, lt_vol_type_list, lt_vol_enc_list, lt_vol_kms_id_list, lt_sg_list, lt_instance_tag_list, lt_volume_tag_list, lt_eni_tag_list = get_launch_template_version_info(ec2_client, lt_id, lt_version)

            # 시트에 데이터 쓰기
            variables = [
                lt_cluster_name, lt_ng_name, lt_name, lt_image, lt_instance_type, lt_tags,
                lt_vol_device_name_list, lt_vol_type_list, lt_vol_size_list, lt_vol_enc_list, lt_vol_kms_id_list, lt_sg_list,
                lt_instance_tag_list, lt_volume_tag_list, lt_eni_tag_list
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