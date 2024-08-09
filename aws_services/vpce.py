# 시작 템플릿 추가 필요
from botocore.exceptions import ClientError
from conf import convert
from conf.sheet_style import front_header_font,header_font,header_fill,header_alignment,content_alignment,multiple_content_alignment,content_border,header_border


def export_vpce_info_to_excel(workbook, ec2_client):
#====================================== VPCE Section ======================================
    # VPCE 목록 조회
    vpce_info = ec2_client.describe_vpc_endpoints()

    # VPCE 시트 생성
    worksheet = workbook.create_sheet('VPCE')
    
    # EKS 열 정보 추가
    worksheet.append(['VPCE'])
    worksheet.cell(1, 1).font = front_header_font

    vpce_headers = [
        # VPCE 기본 정보
        'Name', 'ID', 'Type', 'Service', 'Tags',
        # VPCE 네트워크 정보
        'VPC', 'Subnet', 'Security Group',
        # VPCE DNS 정보
        'Private DNS Enabled', 'Private DNS Only for Inbound Endpoint', 'DNS Name'
    ]

    for col_num, header in enumerate(vpce_headers,1):
        worksheet.cell(2, col_num, value=header).font = header_font
        worksheet.cell(2, col_num, value=header).fill = header_fill
        worksheet.cell(2, col_num, value=header).alignment = header_alignment
        worksheet.cell(2, col_num, value=header).border = header_border

    # 현재 행 위치
    header_row = worksheet.max_row
    
    # auto_filter 적용
    worksheet.auto_filter.ref = f"A{header_row}:{chr(64 + len(vpce_headers))}{header_row}"

    for vpce in vpce_info['VpcEndpoints']:
        
        vpce_name = convert.name_tag_info(vpce['Tags'])
        vpce_id = vpce['VpcEndpointId']
        vpce_type = vpce['VpcEndpointType']
        vpce_svc = vpce['ServiceName']
        vpce_vpc = convert.vpc_info(ec2_client,vpce['VpcId']) + '('+ vpce['VpcId'] +')'
        vpce_tags = convert.tag_info(vpce['Tags'])

        # 서브넷 정리
        vpce_subnet_list = []
        for subnet in vpce['SubnetIds']:
            vpce_subnet_list.append(convert.subnet_info(ec2_client,subnet) + '('+ subnet +')')

        vpce_subnet_list.sort()
        vpce_subnet_list = '\n'.join(vpce_subnet_list)
        
        # 보안그룹 정리
        vpce_sg_list = []
        for sg in vpce['Groups']:
            vpce_sg_list.append(sg['GroupName'] + '('+ sg['GroupId'] +')')
            
        vpce_sg_list.sort()
        vpce_sg_list = '\n'.join(vpce_sg_list)
        
        # DNS 관련
        vpce_pri_dns_enabled = str(vpce['PrivateDnsEnabled'])
        
        try: # 인바운드 엔드포인트에 대해서만 프라이빗 DNS 활성화 확인
            vpce_inbound_only_pri_dns_enabled = vpce['PrivateDnsOnlyForInboundResolverEndpoint']
        except:
            vpce_inbound_only_pri_dns_enabled = '-'
        
        vpce_dns_name_list = []
        for dns in vpce['DnsEntries']:
            vpce_dns_name_list.append(dns['DnsName'])

        vpce_dns_name_list.sort()
        vpce_dns_name_list = '\n'.join(vpce_dns_name_list)
        

        variables = [
            # VPCE 기본 정보
            vpce_name, vpce_id, vpce_type, vpce_svc, vpce_tags,
            # VPCE 네트워크 정보
            vpce_vpc, vpce_subnet_list, vpce_sg_list,
            # VPCE DNS 정보
            vpce_pri_dns_enabled, vpce_inbound_only_pri_dns_enabled, vpce_dns_name_list
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