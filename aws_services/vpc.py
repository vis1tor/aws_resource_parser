# 시작 템플릿 추가 필요
from botocore.exceptions import ClientError
from tqdm import tqdm
from conf import convert
from conf.sheet_style import front_header_font,header_font,header_fill,header_alignment,content_alignment,multiple_content_alignment,content_border,header_border

# VPC 속성 확인 함수
def get_vpc_attribute_info(ec2_client, vpc_id):
    dns_support_option = ec2_client.describe_vpc_attribute(Attribute='enableDnsSupport',VpcId=vpc_id)
    dns_hostname_option = ec2_client.describe_vpc_attribute(Attribute='enableDnsHostnames',VpcId=vpc_id)
    dns_networkmetric_option = ec2_client.describe_vpc_attribute(Attribute='enableNetworkAddressUsageMetrics',VpcId=vpc_id)
    
    return str(dns_support_option['EnableDnsSupport']['Value']), str(dns_hostname_option['EnableDnsHostnames']['Value']), str(dns_networkmetric_option['EnableNetworkAddressUsageMetrics']['Value'])

# VPC Flow Log 확인 함수
def get_vpc_flow_info(ec2_client, vpc_id):
    response = ec2_client.describe_flow_logs(Filters=[{'Name': 'resource-id', 'Values': [vpc_id]}])
    
    if len(response['FlowLogs']) > 0:
        for flow_log in response['FlowLogs']:
            vpc_flow_log_name = convert.name_tag_info(flow_log['Tags'])
            vpc_flow_log_type = flow_log['LogDestinationType']
            vpc_flow_log_dst = flow_log['LogDestination']
    else:
        vpc_flow_log_name = '-'
        vpc_flow_log_type = '-'
        vpc_flow_log_dst = '-'
    return vpc_flow_log_name, vpc_flow_log_type, vpc_flow_log_dst


def export_vpc_info_to_excel(workbook, ec2_client):
#====================================== VPC Section ======================================
    # VPC 목록 조회
    vpc_info = ec2_client.describe_vpcs()

    # VPC 시트 생성
    worksheet = workbook.create_sheet('VPC')
    
    # VPC 열 정보 추가
    worksheet.append(['VPC'])
    worksheet.cell(1, 1).font = front_header_font

    vpc_headers = [
        # VPC 기본 정보
        'VPC Name', 'VPC ID', 'Main CIDR', 'Sub CIDR',
        # VPC 태그 정보
        'Tags',
        # VPC 속성 정보'
        'DNS Support','DNS Hostname', 'Network Address Usage Metrics',
        # VPC Flow Log 정보
        'Flow Log Name', 'Flow Log Type', 'Flow Log Dst'
    ]

    for col_num, header in enumerate(vpc_headers,1):
        worksheet.cell(2, col_num, value=header).font = header_font
        worksheet.cell(2, col_num, value=header).fill = header_fill
        worksheet.cell(2, col_num, value=header).alignment = header_alignment
        worksheet.cell(2, col_num, value=header).border = header_border

    # 현재 행 위치
    header_row = worksheet.max_row
    
    # auto_filter 적용
    worksheet.auto_filter.ref = f"A{header_row}:{chr(64 + len(vpc_headers))}{header_row}"

    # VPC 정보 조회
    for vpc in tqdm(vpc_info['Vpcs'], desc="VPC 정보 파싱 중..."):
        vpc_name = convert.name_tag_info(vpc['Tags'])
        vpc_id = vpc['VpcId']
        vpc_main_cidr = vpc['CidrBlock']
        
        # VPC 보조 IP 조회
        vpc_sub_cidr_list = []
        try:
            for index in range(1, len(vpc['CidrBlockAssociationSet'])): # sub_cidr 길이 만큼 반복
                vpc_sub_cidr_list.append(vpc['CidrBlockAssociationSet'][index]['CidrBlock'])
        except:
            pass
        vpc_sub_cidr_list.sort()
        if len(vpc_sub_cidr_list) > 0 :
            vpc_sub_cidr = '\n'.join(vpc_sub_cidr_list)
        else:
            vpc_sub_cidr = '-'
        
        vpc_tags = convert.tag_info(vpc['Tags'])
        
        vpc_dns_support_option, vpc_dns_hostname_option, vpc_dns_networkmetric_option = get_vpc_attribute_info(ec2_client, vpc_id)
        vpc_flow_log_name, vpc_flow_log_type, vpc_flow_log_dst = get_vpc_flow_info(ec2_client, vpc_id)

        variables = [
            # VPC 기본 정보
            vpc_name, vpc_id, vpc_main_cidr, vpc_sub_cidr,
            # VPC 태그 정보
            vpc_tags,
            # VPC 속성 정보
            vpc_dns_support_option, vpc_dns_hostname_option, vpc_dns_networkmetric_option,
            # VPC Flow Log 정보
            vpc_flow_log_name, vpc_flow_log_type, vpc_flow_log_dst

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