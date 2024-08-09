# 프로토콜 정의 : https://www.iana.org/assignments/protocol-numbers/protocol-numbers.xhtml
from botocore.exceptions import ClientError
from conf import convert
from conf.sheet_style import front_header_font,header_font,header_fill,header_alignment,content_alignment,multiple_content_alignment,content_border,header_border

def target_ip(worksheet, sg_rule, sg_name, sg_id, sg_vpc, sg_tags, sg_description, sg_rule_port, sg_rule_protocol, sg_rule_traffic):
    if 'IpRanges' in sg_rule: 
        sg_rule_target_sg = '-' # 대상 SG - 처리
        for ip_range in sg_rule['IpRanges']:
            if 'Description' in ip_range:
                sg_rule_description = ip_range['Description']
            else:
                sg_rule_description = '-'
            
            sg_rule_target_ip_range = ip_range['CidrIp'] # 대상 IP Range
            
             # 시트에 데이터 쓰기
            variables = [
                sg_vpc, sg_name, sg_id, sg_tags, sg_description, sg_rule_traffic, sg_rule_protocol, sg_rule_port, sg_rule_target_ip_range, sg_rule_target_sg, sg_rule_description
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


def target_sg(worksheet, ec2_client, sg_rule, sg_name, sg_id, sg_vpc, sg_tags, sg_description, sg_rule_port, sg_rule_protocol, sg_rule_traffic):
    if 'UserIdGroupPairs' in sg_rule:
        sg_rule_target_ip_range = '-' # 대상 IP Range - 처리
        for user_id_group in sg_rule['UserIdGroupPairs']:
            if 'Description' in user_id_group:
                sg_rule_description = user_id_group['Description']
            else:
                sg_rule_description = '-'

            sg_rule_target_sg = f"{convert.sg_info(ec2_client, user_id_group['GroupId'])}({user_id_group['GroupId']})" # 대상 SG
            # 시트에 데이터 쓰기
            variables = [
                sg_vpc, sg_name, sg_id, sg_tags, sg_description, sg_rule_traffic, sg_rule_protocol, sg_rule_port, sg_rule_target_ip_range, sg_rule_target_sg, sg_rule_description
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

def export_sg_info_to_excel(workbook, ec2_client):
    sg_info = ec2_client.describe_security_groups()
    
    # SG 시트 생성
    worksheet = workbook.create_sheet('SG')
    
    # SG 열 정보 추가
    worksheet.append(['SG'])
    worksheet.cell(1, 1).font = front_header_font

    sg_headers = [
    # SG 기본 정보
    'VPC', 'Security Group', 'Security Group ID', 'Security Group Tag', 'Security Group Description', 'Inbound/Outbound',
    # SG Rule 정보
    'Protocol', 'Port Range', 'Target IP Range', 'Target Security Group', 'Security Group Rule Description'
    ]

    for col_num, header in enumerate(sg_headers,1):
        worksheet.cell(2, col_num, value=header).font = header_font
        worksheet.cell(2, col_num, value=header).fill = header_fill
        worksheet.cell(2, col_num, value=header).alignment = header_alignment
        worksheet.cell(2, col_num, value=header).border = header_border
    
    # 현재 행 위치
    header_row = worksheet.max_row
    
    # auto_filter 적용
    worksheet.auto_filter.ref = f"A{header_row}:{chr(64 + len(sg_headers))}{header_row}"

    #보안그룹 수만큼 반복
    for sg in sg_info['SecurityGroups']:
        sg_name = sg['GroupName']
        sg_id = sg['GroupId']
        sg_vpc = convert.vpc_info(ec2_client, sg['VpcId'])
        sg_description = sg['Description']
        
        if sg.get('Tags'):
            sg_tags = []
            for tag in sg['Tags']:
                sg_tags.append(f"{tag['Key']} : {tag['Value']}")
            sg_tags = '\n'.join(sg_tags)
        else:
            sg_tags = '-'
        
        # SG Rule 존재 확인 및 Ingres/Egress Rule 분리
        if 'IpPermissions' in sg:
            rule_traffic = 'Inbound'
            
            for ingress_rule in sg['IpPermissions']:
                # FromPort 존재할 시, -> TCP, UDP, ICMP... 등등 포함
                if 'FromPort' in ingress_rule:
                    # Rule Port 분리(All vs From/ToPort)
                    ingress_rule_protocol = ingress_rule['IpProtocol']
                    
                    if ingress_rule.get('FromPort') == -1:
                        ingress_rule_port = 'All'
                    else:
                        ingress_rule_port = f"{ingress_rule['FromPort']}-{ingress_rule['ToPort']}"
                    # Rule target 분리(IpRanges VS UserIdGroupPairs)
                    target_ip(worksheet, ingress_rule, sg_name, sg_id, sg_vpc, sg_tags, sg_description, ingress_rule_port, ingress_rule_protocol, rule_traffic)
                    target_sg(worksheet, ec2_client, ingress_rule, sg_name, sg_id, sg_vpc, sg_tags, sg_description, ingress_rule_port, ingress_rule_protocol, rule_traffic)
                
                # FromPort 가 존재하지 않을 시 -> 모든 트래픽(프로토콜)/사용자 지정 프로토콜 사용 기준으로 분리
                if 'FromPort' not in ingress_rule:
                    # 모든 트래픽(프로토콜) 허용은 즉, 모든 포트 허용한다는 의미...
                    if ingress_rule['IpProtocol'] == '-1':
                        ingress_rule_port = 'All'
                        ingress_rule_protocol = 'All'
                        
                        #Rule target 분리(IpRanges VS UserIdGroupPairs)
                        target_ip(worksheet, ingress_rule, sg_name, sg_id, sg_vpc, sg_tags, sg_description, ingress_rule_port, ingress_rule_protocol, rule_traffic)
                        target_sg(worksheet, ec2_client, ingress_rule, sg_name, sg_id, sg_vpc, sg_tags, sg_description, ingress_rule_port, ingress_rule_protocol, rule_traffic)
                    
                    # 사용자 지정 프로토콜 사용 시, 모든 포트 허용으로 표시됨...
                    else:
                        ingress_rule_protocol = ingress_rule['IpProtocol']
                        ingress_rule_port = 'All'
                        
                        #Rule target 분리(IpRanges VS UserIdGroupPairs)
                        target_ip(worksheet, ingress_rule, sg_name, sg_id, sg_vpc, sg_tags, sg_description, ingress_rule_port, ingress_rule_protocol, rule_traffic)
                        target_sg(worksheet, ec2_client, ingress_rule, sg_name, sg_id, sg_vpc, sg_tags, sg_description, ingress_rule_port, ingress_rule_protocol, rule_traffic)
        
        if 'IpPermissionsEgress' in sg:
            rule_traffic = 'Outbound'
            for egress_rule in sg['IpPermissionsEgress']:
                #FromPort 존재할 시, -> TCP, UDP, ICMP... 등등 포함
                if 'FromPort' in egress_rule:
                    # Rule Port 분리(All vs From/ToPort)
                    egress_rule_protocol = egress_rule['IpProtocol']
                    
                    if egress_rule.get('FromPort') == -1:
                        egress_rule_port = 'All'
                    else:
                        egress_rule_port = f"{egress_rule['FromPort']}-{egress_rule['ToPort']}"
                    # Rule target 분리(IpRanges VS UserIdGroupPairs)
                    target_ip(worksheet, egress_rule, sg_name, sg_id, sg_vpc, sg_tags, sg_description, egress_rule_port, egress_rule_protocol, rule_traffic)
                    target_sg(worksheet, ec2_client, egress_rule, sg_name, sg_id, sg_vpc, sg_tags, sg_description, egress_rule_port, egress_rule_protocol, rule_traffic)
                
                # FromPort 가 존재하지 않을 시 -> 모든 트래픽(프로토콜)/사용자 지정 프로토콜 사용 기준으로 분리
                if 'FromPort' not in egress_rule:
                    # 모든 트래픽(프로토콜) 허용은 즉, 모든 포트 허용한다는 의미...
                    if egress_rule['IpProtocol'] == '-1':
                        egress_rule_port = 'All'
                        egress_rule_protocol = 'All'
                        
                        #Rule target 분리(IpRanges VS UserIdGroupPairs)
                        target_ip(worksheet, egress_rule, sg_name, sg_id, sg_vpc, sg_tags, sg_description, egress_rule_port, egress_rule_protocol, rule_traffic)
                        target_sg(worksheet, ec2_client, egress_rule, sg_name, sg_id, sg_vpc, sg_tags, sg_description, egress_rule_port, egress_rule_protocol, rule_traffic)
                    
                    # 사용자 지정 프로토콜 사용 시, 모든 포트 허용으로 표시됨...
                    else:
                        egress_rule_protocol = egress_rule['IpProtocol']
                        egress_rule_port = 'All'
                        
                        #Rule target 분리(IpRanges VS UserIdGroupPairs)
                        target_ip(worksheet, egress_rule, sg_name, sg_id, sg_vpc, sg_tags, sg_description, egress_rule_port, egress_rule_protocol, rule_traffic)
                        target_sg(worksheet, ec2_client, egress_rule, sg_name, sg_id, sg_vpc, sg_tags, sg_description, egress_rule_port, egress_rule_protocol, rule_traffic)