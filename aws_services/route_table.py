# 시작 템플릿 추가 필요
from botocore.exceptions import ClientError
from tqdm import tqdm
from conf import convert
from conf.sheet_style import front_header_font,header_font,header_fill,header_alignment,content_alignment,multiple_content_alignment,content_border,header_border

def export_rt_info_to_excel(workbook, ec2_client):
#====================================== VPC Section ======================================
    # Routing Table 목록 조회
    rt_info = ec2_client.describe_route_tables()

    # VPC 시트 생성
    worksheet = workbook.create_sheet('RT')
    
    # VPC 열 정보 추가
    worksheet.append(['Routing Table'])
    worksheet.cell(1, 1).font = front_header_font

    rt_headers = [
        # 서브넷 기본 정보
        'VPC Name', 'Routing Table Name', 'Routing Table ID', 'Associated Subnets'
    ]

    for col_num, header in enumerate(rt_headers,1):
        worksheet.cell(2, col_num, value=header).font = header_font
        worksheet.cell(2, col_num, value=header).fill = header_fill
        worksheet.cell(2, col_num, value=header).alignment = header_alignment
        worksheet.cell(2, col_num, value=header).border = header_border  
    
    # 현재 행 위치
    header_row = worksheet.max_row

    # auto_filter 적용
    worksheet.auto_filter.ref = f"A{header_row}:{chr(64 + len(rt_headers))}{header_row}"

    # Routing Table 정보 조회
    for rt in tqdm(rt_info['RouteTables'], desc="Routing Table 정보 파싱 중..."):
        rt_name = convert.name_tag_info(rt['Tags'])
        rt_id = rt['RouteTableId'] 
        rt_vpc = convert.vpc_info(ec2_client, rt['VpcId'])

        # Routing Table 연결 서브넷 조회
        rt_asso_subnet_list = []
        try:
            for asso_subnet in rt['Associations']:
                rt_asso_subnet_list.append(convert.subnet_info(ec2_client, asso_subnet['SubnetId']))
        except:
            pass
        rt_asso_subnet_list.sort()
        if len(rt_asso_subnet_list) > 0 :
            rt_asso_subnet = '\n'.join(rt_asso_subnet_list)
        else:
            rt_asso_subnet = '-'
        
        # Routing Table Route 조회
        rt_route_dst_cidr_list = []
        rt_route_dst_target_list = []
        try:
            for route in rt['Routes']:
                rt_route_dst_cidr_list.append(route['DestinationCidrBlock'])
                rt_route_dst_target_list.append(route['GatewayId'])
        except:
            pass
        # rt_route_dst_cidr_list.sort()

        if len(rt_route_dst_cidr_list) > 0 :
            rt_route_dst_cidr = '\n'.join(rt_route_dst_cidr_list)
        else:
            rt_route_dst_cidr = '-'

        if len(rt_route_dst_target_list) > 0 :
            rt_route_dst_target = '\n'.join(rt_route_dst_target_list)
        else:
            rt_route_dst_target = '-'
        
        variables = [
            # Routing Table 기본 정보
            rt_vpc, rt_name, rt_id, rt_asso_subnet, rt_route_dst_cidr , rt_route_dst_target
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