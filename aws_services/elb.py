from botocore.exceptions import ClientError
from conf import convert
from conf.sheet_style import front_header_font,header_font,header_fill,header_alignment,content_alignment,multiple_content_alignment,content_border,header_border

# 리스너 대상 그룹 조회
def describe_target_groups(elbv2_client, target_group_arn):
    response = elbv2_client.describe_target_groups(TargetGroupArns=[target_group_arn])

    return response['TargetGroups'][0]

# 리스너 대상 그룹 상태 조회
def describe_target_health(elbv2_client, target_group_arn):
    response = elbv2_client.describe_target_health(TargetGroupArn=target_group_arn)
    
    return response['TargetHealthDescriptions']
    

# 로드 밸런서의 리스너 조회
def describe_listener(elbv2_client, ec2_client ,elb_arn, elb_type):
    response = elbv2_client.describe_listeners(LoadBalancerArn=elb_arn)
    
    elb_listener_port = []
    elb_listener_protocol = []
    elb_listener_type = []
    elb_listener_rule = []
    elb_listener_rule_condition = []
    elb_listener_rule_condition_val = []
    elb_target_group_name = []
    elb_target_group_protocol = []
    elb_target_group_port = []
    elb_target_group_health_protocol = []
    elb_target_group_health_port = []
    elb_target_group_health_path = []
    elb_target_group_target = []
    elb_target_group_target_az = []
    elb_target_group_target_health = []

    for listener in response['Listeners']:
        listener_arn = listener['ListenerArn']
        if elb_type == 'gateway':
            elb_listener_protocol.append('-')
            elb_listener_port.append('-')
            elb_listener_type.append('-')
            elb_listener_rule.append('-')
        else:
            elb_listener_protocol.append(listener['Protocol'])
            elb_listener_port.append(str(listener['Port']))
            for listener_actions in listener['DefaultActions']:
                elb_listener_type.append(listener_actions['Type'])
                
                ###
                elb_target_group_info = describe_target_groups(elbv2_client, listener_actions['ForwardConfig']['TargetGroups'][0]['TargetGroupArn'])
                elb_target_group_name.append(elb_target_group_info['TargetGroupName'])
                elb_target_group_protocol.append(elb_target_group_info['Protocol'])
                elb_target_group_port.append(str(elb_target_group_info['Port']))
                elb_target_group_health_protocol.append(elb_target_group_info['HealthCheckProtocol'])
                elb_target_group_health_port.append(elb_target_group_info['HealthCheckPort'])
                try:
                    elb_target_group_health_path.append(elb_target_group_info['HealthCheckPath'])
                except:
                    elb_target_group_health_path.append('-')

                ###
                elb_target_group_health_info = describe_target_health(elbv2_client, listener_actions['ForwardConfig']['TargetGroups'][0]['TargetGroupArn'])

                for target_group in elb_target_group_health_info:
                    
                    # elb_target_group_target.append(target_group['Target']['Id'])
                    # convert.ec2_info(ec2_client,target_group['Target']['Id'])
                    # elb_target_group_target.append(convert.ec2_info(ec2_client, target_group['Target']['Id'])+'('+target_group['Target']['Id']+')')
                    try:
                        elb_target_group_target.append(convert.ec2_info(ec2_client, target_group['Target']['Id'])+'('+target_group['Target']['Id']+')')
                    except:
                        elb_target_group_target.append(target_group['Target']['Id'])
                    
                    try:
                        elb_target_group_target_az.append(target_group['Target']['AvailabilityZone'])
                    except:
                        elb_target_group_target_az.append('-')
                    elb_target_group_target_health.append(target_group['TargetHealth']['State'])

                ###
                elb_listener_rule_condition_info = describe_listener_rule(elbv2_client, listener_arn)
                for key,value in elb_listener_rule_condition_info.items():
                    elb_listener_rule_condition.append(key)
                    elb_listener_rule_condition_val.append(','.join(value))

        if len(response['Listeners']) > 2:
            # 타켓그룹 정보 구분자
            elb_target_group_target.append(' ')
            elb_target_group_target_az.append(' ')
            elb_target_group_target_health.append(' ')
    
    elb_listener_port = '\n'.join(elb_listener_port)
    elb_listener_protocol = '\n'.join(elb_listener_protocol)
    elb_listener_type = '\n'.join(elb_listener_type)
    elb_listener_rule_condition = '\n'.join(elb_listener_rule_condition)
    elb_listener_rule_condition_val = '\n'.join(elb_listener_rule_condition_val)
    elb_target_group_name = '\n'.join(elb_target_group_name)
    elb_target_group_protocol = '\n'.join(elb_target_group_protocol)
    elb_target_group_port = '\n'.join(elb_target_group_port)
    elb_target_group_health_protocol = '\n'.join(elb_target_group_health_protocol)
    elb_target_group_health_port = '\n'.join(elb_target_group_health_port)
    elb_target_group_health_path = '\n'.join(elb_target_group_health_path)
    elb_target_group_target = '\n'.join(elb_target_group_target)
    elb_target_group_target_az = '\n'.join(elb_target_group_target_az)
    elb_target_group_target_health = '\n'.join( elb_target_group_target_health)

    return elb_listener_port, elb_listener_protocol, elb_listener_type, elb_listener_rule_condition, elb_listener_rule_condition_val, elb_target_group_name, elb_target_group_protocol, elb_target_group_port, elb_target_group_health_protocol, elb_target_group_health_port, elb_target_group_health_path, elb_target_group_target, elb_target_group_target_az, elb_target_group_target_health

def describe_listener_rule(elbv2_client, listener_arn):
    response = elbv2_client.describe_rules(ListenerArn=listener_arn)
    
    # 리스너 규칙 정보가 저장될 딕셔너리 선언
    listener_rule_condition_info = {}

    for listener_rule in response['Rules']:
        if len(listener_rule['Conditions']) == 0:
            listener_rule_condition_info['-'] = ['-']
        if len(listener_rule['Conditions']) == 1:
            for condition in listener_rule['Conditions']:
                # listener_rule_condition_info 값을 리스트로 사전 선언
                listener_rule_condition_info[condition['Field']] = []
                # 리스너 규칙 조건에 따른 분리 작업
                if condition['Field'] == 'http-header':
                    # 규칙 조건에 대한 값이 여러 개일 경우 추가 작업
                    for value in condition['HttpHeaderConfig']['Values']:
                        listener_rule_condition_info[condition['Field']].append(value)
                elif condition['Field'] == 'http-request-method':
                    for value in condition['HttpRequestMethodConfig']['Values']:
                        listener_rule_condition_info[condition['Field']].append(value)
                elif condition['Field'] == 'host-header':
                    for value in condition['HostHeaderConfig']['Values']:
                        listener_rule_condition_info[condition['Field']].append(value)
                elif condition['Field'] == 'path-pattern':
                    for value in condition['PathPatternConfig']['Values']:
                        listener_rule_condition_info[condition['Field']].append(value)
                elif condition['Field'] == 'query-string':
                    for value in condition['QueryStringConfig']['Values']:
                        listener_rule_condition_info[condition['Field']].append(value)
                elif condition['Field'] == 'source-ip':
                    for value in condition['SourceIpConfig']['Values']:
                        listener_rule_condition_info[condition['Field']].append(value)
    
    return listener_rule_condition_info


def export_elb_info_to_excel(workbook, elbv2_client, ec2_client):
    elb_info = elbv2_client.describe_load_balancers()

    # ELB 시트 생성
    worksheet = workbook.create_sheet('ELB')
    
    # ELB 열 정보 추가
    worksheet.append(['ELB'])
    worksheet.cell(1, 1).font = front_header_font

    elb_headers = [
        'Nmae', 'Type', 'Scheme','DNS',
        'VPC', 'Subnet', 'SecurityGroup', 'Listener Protocol', 'Listener Port', 'Listener Type', 'Listener Condition', 'Listener Condition Value',
        'Target Group Name', 'Target Group Protocol', 'Target Group Port', 'Health Protocol', 'Health Port', 'Health Path', 'Target', 'Target AZ', 'Target Status'
    ]

    for col_num, header in enumerate(elb_headers,1):
        worksheet.cell(2, col_num, value=header).font = header_font
        worksheet.cell(2, col_num, value=header).fill = header_fill
        worksheet.cell(2, col_num, value=header).alignment = header_alignment
        worksheet.cell(2, col_num, value=header).border = header_border
    
    # 현재 행 위치
    header_row = worksheet.max_row
    
    # auto_filter 적용
    worksheet.auto_filter.ref = f"A{header_row}:{chr(64 + len(elb_headers))}{header_row}"

    for elb in elb_info['LoadBalancers']:
        elb_name = elb['LoadBalancerName']
        elb_arn = elb['LoadBalancerArn']
        elb_type = elb['Type']
        elb_vpc = convert.vpc_info(ec2_client, elb['VpcId'])
        
        # GWLB의 경우 체계(Scheme)와 DNSName 정보가 없음
        if 'Scheme' and 'DNSName' in elb:
            elb_scheme = elb['Scheme']
            elb_dns_name = elb['DNSName']
        else:
            elb_scheme = '-'
            elb_dns_name = '-'
        
        elb_subnet_name = []
        elb_sg_name = []

        # ELB 중 보안 그룹 존재 여부 확인
        if 'SecurityGroups' in elb:
            for sg_id in elb['SecurityGroups']:
                elb_sg_name.append(convert.sg_info(ec2_client, sg_id))
        # GWLB의 경우 보안 그룹이 존재하지 않음
        elif elb_type == 'gateway':
            elb_sg_name = '-'
        
        # ELB Listener 
        elb_listener_port, elb_listener_protocol, elb_listener_type, elb_listener_rule_condition, elb_listener_rule_condition_val, elb_target_group_name, elb_target_group_protocol, elb_target_group_port, elb_target_group_health_protocol, elb_target_group_health_port, elb_target_group_health_path, elb_target_group_target, elb_target_group_target_az, elb_target_group_target_health = describe_listener(elbv2_client, ec2_client, elb_arn, elb_type)
        
        # ELB Subnet
        for subnet_id in [az['SubnetId'] for az in elb['AvailabilityZones']]:
            elb_subnet_name.append(convert.subnet_info(ec2_client, subnet_id))
        
        if elb_subnet_name != '-':
            elb_subnet_name = '\n'.join(elb_subnet_name)
        if elb_sg_name != '-':
            elb_sg_name = '\n'.join(elb_sg_name)
    
        # 시트에 데이터 쓰기
        variables = [
            elb_name, elb_type, elb_scheme, elb_dns_name,
            elb_vpc, elb_subnet_name, elb_sg_name,
            elb_listener_protocol, elb_listener_port, elb_listener_type, elb_listener_rule_condition, elb_listener_rule_condition_val,
            elb_target_group_name, elb_target_group_protocol, str(elb_target_group_port), elb_target_group_health_protocol, elb_target_group_health_port, elb_target_group_health_path,
            elb_target_group_target, elb_target_group_target_az, elb_target_group_target_health
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