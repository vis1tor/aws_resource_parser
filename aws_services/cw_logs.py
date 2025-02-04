from botocore.exceptions import ClientError
from conf.sheet_style import front_header_font,header_font,header_fill,header_alignment,content_alignment,multiple_content_alignment,content_border,header_border

def get_lg_tag(cw_logs_client, cw_lg_name):
    response = cw_logs_client.list_tags_log_group(logGroupName=cw_lg_name)
    
    # Tag 존재 여부 확인
    if len(response['tags']) == 0:
        cw_lg_tag_list = '-'
    else:
        cw_lg_tag_list = []
        
        for tag_key, tag_value in response['tags'].items():
            cw_lg_tag_list.append(f"{tag_key} : {tag_value}")

        cw_lg_tag_list = '\n'.join(cw_lg_tag_list)

    return cw_lg_tag_list


def convert_retention_day(retention_days):
    retention_periods = {
        1: "1일", 3: "3일", 5: "5일", 7: "1주", 14: "2주", 
        30: "1개월", 60: "2개월", 90: "3개월", 120: "4개월", 150: "5개월", 180: "6개월", 365: "12개월", 400: "13개월", 545: "18개월",
        731: "2년", 1096: "3년", 1827: "5년", 2192: "6년", 2557: "7년", 2922: "8년", 3288: "9년", 3653: "10년"
    }
    if retention_days in retention_periods:
        convert_retention_days = retention_periods[retention_days]
    else:
        print("Unsupported retention period.")
    
    return convert_retention_days


def export_cw_logs_info_to_excel(workbook, cw_logs_client):
    paginator = cw_logs_client.get_paginator('describe_log_groups')
    
    # CW Log group 시트 생성
    worksheet = workbook.create_sheet('CW Log Group')
    
    # CW Log group 열 정보 추가
    worksheet.append(['CloudWatch Log Group'])
    worksheet.cell(1, 1).font = front_header_font

    cw_log_group_headers = [
    # CW Log group 기본 정보
    'Log Group', 'Encryption', 'Encryption Key', 'Retention', 'Tags'
    ]
    for col_num, header in enumerate(cw_log_group_headers,1):
        worksheet.cell(2, col_num, value=header).font = header_font
        worksheet.cell(2, col_num, value=header).fill = header_fill
        worksheet.cell(2, col_num, value=header).alignment = header_alignment
        worksheet.cell(2, col_num, value=header).border = header_border
    
    # 현재 행 위치
    header_row = worksheet.max_row
    
    # auto_filter 적용
    worksheet.auto_filter.ref = f"A{header_row}:{chr(64 + len(cw_log_group_headers))}{header_row}"

    for cw_logs_info in paginator.paginate():
        for cw_log in cw_logs_info['logGroups']:
            cw_lg_name = cw_log['logGroupName']
            
            # 암호화 여부 확인
            if 'kmsKeyId' in cw_log:
                cw_lg_enc = 'Enabled'
                cw_lg_kms_key_id = cw_log['kmsKeyId']
            else:
                cw_lg_enc = 'Disabled'
                cw_lg_kms_key_id = '-'

            # 보존 기간 설정 및 보존 기간일 변환
            if 'retentionInDays' in cw_log:
                cw_lg_retention = convert_retention_day(cw_log['retentionInDays'])
            else:
                cw_lg_retention = '-'

            # 태그 확인
            cw_lg_tag = get_lg_tag(cw_logs_client, cw_lg_name)

            variables = [
                cw_lg_name, cw_lg_enc, cw_lg_kms_key_id, cw_lg_retention, cw_lg_tag
            ]

            # 시트에 데이터 쓰기
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