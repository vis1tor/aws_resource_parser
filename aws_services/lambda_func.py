from botocore.exceptions import ClientError
from conf import convert
from conf.sheet_style import front_header_font,header_font,header_fill,header_alignment,content_alignment,multiple_content_alignment,content_border,header_border

# Lambda 상세 정보 조회
def get_lambda_function_info(lambda_client, lambda_name):
    response = lambda_client.get_function(FunctionName=lambda_name)
    
    lambda_role = response['Configuration']['Role'].split('/')[-1]
    lambda_runtime = response['Configuration']['Runtime']
    lambda_arch = '\n'.join(response['Configuration']['Architectures'])
    lambda_handler = response['Configuration']['Handler']
    lambda_timeout = str(response['Configuration']['Timeout'])
    lambda_mem_size = str(response['Configuration']['MemorySize'])
    lambda_log_format = response['Configuration']['LoggingConfig']['LogFormat']
    lambda_log_group = response['Configuration']['LoggingConfig']['LogGroup']
    lambda_descrition = response['Configuration']['Description']

    # Lambda Tag는 딕셔너리 형태 
    try:
        lambda_tag = convert.dic_tag_info(response['Tags'])
    except:
        pass
        lambda_tag = '-'

    return lambda_role, lambda_runtime, lambda_arch, lambda_handler, lambda_timeout, lambda_mem_size, lambda_log_format, lambda_log_group, lambda_descrition, lambda_tag

def export_lambda_info_to_excel(workbook, lambda_client, ec2_client):
#====================================== Lambda Section ======================================
    # Lambda 목록 조회
    lambda_info = lambda_client.list_functions()

    # Lambda 시트 생성
    worksheet = workbook.create_sheet('Lambda')
    
    # Lambda 열 정보 추가
    worksheet.append(['Lambda'])
    worksheet.cell(1, 1).font = front_header_font

    lambda_headers = [
        # Lambda 기본 정보
        'Lambda Name', 'Lambda Role', 'Lambda Runtime', 'Lambda Arch', 'Lambda Handler', 'Lambda Timeout',
        'Lambda Mem Size', 'Lambda Log Format', 'Lambda Log Group', 'Lambda Descrition',
        # Lambda 태그 정보
        'Tags'
    ]

    for col_num, header in enumerate(lambda_headers,1):
        worksheet.cell(2, col_num, value=header).font = header_font
        worksheet.cell(2, col_num, value=header).fill = header_fill
        worksheet.cell(2, col_num, value=header).alignment = header_alignment
        worksheet.cell(2, col_num, value=header).border = header_border

    # 현재 행 위치
    header_row = worksheet.max_row
    
    # auto_filter 적용
    worksheet.auto_filter.ref = f"A{header_row}:{chr(64 + len(lambda_headers))}{header_row}"

    # Lambda 정보 조회
    for lambda_func in lambda_info['Functions']:
        lambda_name = lambda_func['FunctionName']

        lambda_role, lambda_runtime, lambda_arch, lambda_handler, lambda_timeout, lambda_mem_size, lambda_log_format, lambda_log_group, lambda_descrition, lambda_tag = get_lambda_function_info(lambda_client, lambda_name)

        
        variables = [
            # Lambda 기본 정보
            lambda_name, lambda_role, lambda_runtime, str(lambda_arch), lambda_handler, lambda_timeout,
            lambda_mem_size, lambda_log_format, lambda_log_group, lambda_descrition,
            # Lambda 태그 정보
            lambda_tag
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