# 리소스 ID를 Name 태그로 변환하는 모듈
# VPC 또는 서브넷 등이 Default를 사용할 때를 고려할 필요가 있음...

# EC2 Name 태그 추출
def ec2_info(ec2_client, ec2_id):
    response = ec2_client.describe_instances(
        InstanceIds=[
            ec2_id,
        ]
    )

    if len(response['Reservations'][0]['Instances'][0]['Tags']) == 0:
        ec2_name = '-'
    else:
        for tag in response['Reservations'][0]['Instances'][0]['Tags']:
            try:
                if tag['Key'] == 'Name':
                    ec2_name = tag['Value']
            except:
                pass
    return ec2_name

# Name 태그 추출
def name_tag_info(service_tags):
    # 태그 미존재 시 예외 처리
    if len(service_tags) == 0:
        service_tag = '-'
    else:
        for tag in service_tags:
            if tag['Key'] == 'Name':
                service_tag = tag['Value']
            else:
                service_tag = '-'
    return service_tag

# Tag 변환
def tag_info(service_tags):
    service_tag_list = []
    if len(service_tags) == 0:
        service_tag_list = '-'
    else:
        for tag in service_tags:
            service_tag_list.append(f"{tag['Key']} : {tag['Value']}")

        service_tag_list = '\n'.join(service_tag_list)

    return  service_tag_list

# VPC 변환
def vpc_info(ec2_client, vpc_id):
    response = ec2_client.describe_vpcs(
        VpcIds=[
            vpc_id,
        ]
    )
    # default VPC 의 경우, 태그가 존재하지 않는다. '172.31.0.0/16' 가 고정인 것으로 확인된다.
    if 'Tags' in response['Vpcs'][0]:
        for tag in response['Vpcs'][0]['Tags']:
            if tag.get('Key') == 'Name':
                vpc_name = tag.get('Value')
    else: 
        vpc_name = 'N/A'
    return  vpc_name

# 보안 그룹 변환
def sg_info(ec2_client, sg_id):
    response = ec2_client.describe_security_groups(
        GroupIds=[
            sg_id
        ]
    )
    
    return response['SecurityGroups'][0]['GroupName']

# 서브넷 변환
def subnet_info(ec2_client, subnet_id):
    response = ec2_client.describe_subnets(
        SubnetIds=[
            subnet_id
        ]
    )
    # default VPC 서브넷의 경우, 태그가 존재하지 않는다.
    if 'Tags' in response['Subnets'][0]:
        for tag in response['Subnets'][0]['Tags']:
            if tag.get('Key') == 'Name':
                subnet_name = tag.get('Value')
    else:
        subnet_name = 'N/A'
    return subnet_name