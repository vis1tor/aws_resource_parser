# 시작 템플릿 추가 필요
from botocore.exceptions import ClientError
from conf import convert
from conf.sheet_style import front_header_font,header_font,header_fill,header_alignment,content_alignment,multiple_content_alignment,content_border,header_border

def export_subnet_info_to_excel(workbook, ec2_client):
#====================================== VPC Section ======================================
    # Subnet 목록 조회
    subnet_info = ec2_client.describe_subnets()

    # VPC 시트 생성
    worksheet = workbook.create_sheet('Subnet')
    
    # VPC 열 정보 추가
    worksheet.append(['Subnet'])
    worksheet.cell(1, 1).font = front_header_font

    subnet_headers = [
        # 서브넷 기본 정보
        'VPC Name', 'Subnet Name', 'Subnet ID', 'CIDR', 'Zone',
        # 서브넷 속성 정보
        'DNS64 Option', 'DNS A Record Option', 'DNS AAAA Record Option',
        # 서브넷 태그 정보
        'Tags'
    ]

    for col_num, header in enumerate(subnet_headers,1):
        worksheet.cell(2, col_num, value=header).font = header_font
        worksheet.cell(2, col_num, value=header).fill = header_fill
        worksheet.cell(2, col_num, value=header).alignment = header_alignment
        worksheet.cell(2, col_num, value=header).border = header_border  
    
    # 현재 행 위치
    header_row = worksheet.max_row

    # auto_filter 적용
    worksheet.auto_filter.ref = f"A{header_row}:{chr(64 + len(subnet_headers))}{header_row}"

    # Subnet 정보 조회
    for subnet in subnet_info['Subnets']:
        subnet_az = subnet['AvailabilityZone']
        subnet_cidr = subnet['CidrBlock']
        subnet_vpc = convert.vpc_info(ec2_client, subnet['VpcId'])
        
        # Tags 없는 경우 존재
        try:
            subnet_name = convert.name_tag_info(subnet['Tags'])
            subnet_tags = convert.tag_info(subnet['Tags'])
        except:
            subnet_name = '-'
            subnet_tags = '-'
        
        subnet_id = subnet['SubnetId']
        subnet_dns64_option = str(subnet['EnableDns64'])
        subnet_dns_a_record_option = str(subnet['PrivateDnsNameOptionsOnLaunch']['EnableResourceNameDnsARecord'])
        subnet_dns_aaaa_record_option = str(subnet['PrivateDnsNameOptionsOnLaunch']['EnableResourceNameDnsAAAARecord'])

        variables = [
            # 서브넷 기본 정보
            subnet_vpc, subnet_name, subnet_id, subnet_cidr, subnet_az,
            # 서브넷 속성 정보
            subnet_dns64_option, subnet_dns_a_record_option, subnet_dns_aaaa_record_option,
            # 서브넷 태그 정보
            subnet_tags
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