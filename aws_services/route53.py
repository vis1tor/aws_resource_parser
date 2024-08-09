from botocore.exceptions import ClientError
from conf.sheet_style import front_header_font,header_font,header_fill,header_alignment,content_alignment,multiple_content_alignment,content_border,header_border

def export_route53_info_to_excel(workbook, route53_client):
    route53_list = route53_client.list_hosted_zones()

    # Route53 시트 생성
    worksheet = workbook.create_sheet('Route53')
    
    # Route53 열 정보 추가
    worksheet.append(['Route53'])
    worksheet.cell(1, 1).font = front_header_font

    route53_headers = [
        # Route53 기본 정보
        'Zone', 'Record', 'Record Type', 'Record Values'
    ]

    for col_num, header in enumerate(route53_headers,1):
        worksheet.cell(2, col_num, value=header).font = header_font
        worksheet.cell(2, col_num, value=header).fill = header_fill
        worksheet.cell(2, col_num, value=header).alignment = header_alignment
        worksheet.cell(2, col_num, value=header).border = header_border

    # 현재 행 위치
    header_row = worksheet.max_row
    
    # auto_filter 적용
    worksheet.auto_filter.ref = f"A{header_row}:{chr(64 + len(route53_headers))}{header_row}"

    for hosted_zone in route53_list['HostedZones']:

        zone_name = hosted_zone['Name'][:-1]
        zone_record_info = route53_client.list_resource_record_sets(HostedZoneId=hosted_zone['Id'])
        for record in zone_record_info['ResourceRecordSets']:
            if zone_name != record['Name']: # Hosted Zone 이름과 Record 이름 비교로 NS, SOA 레코드 분리 처리
                record_name = record['Name'][:-1]
                record_type = record['Type']
                record_values = []
                if record_type == 'A':
                    if 'AliasTarget' in record:
                        record_values.append(record['AliasTarget']['DNSName'][:-1])
                    if 'ResourceRecords' in record:
                        for value in record['ResourceRecords']:
                            for key,value in value.items():
                                record_values.append(value)
                if record_type == 'CNAME':
                    for value in record['ResourceRecords']:
                            for key,value in value.items():
                                record_values.append(value)
                record_values = '\n'.join(record_values)

                variables = [
                    # Route53 기본 정보
                    zone_name, record_name, record_type, record_values
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