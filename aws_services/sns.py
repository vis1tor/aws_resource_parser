from botocore.exceptions import ClientError
from conf.sheet_style import front_header_font,header_font,header_fill,header_alignment,content_alignment,multiple_content_alignment,content_border,header_border

def get_sns_topic_tags(sns_client, sns_arn):
    response = sns_client.list_tags_for_resource(ResourceArn=sns_arn)
    
    sns_topic_tag_list = []

    # SNS 태그 미존재 시 예외 처리
    if len(response['Tags']) == 0:
        sns_topic_tag_list = '-'
    else:
        for tag in response['Tags']:
            sns_topic_tag_list.append(f"{tag['Key']} : {tag['Value']}")
        
        sns_topic_tag_list = '\n'.join(sns_topic_tag_list)
    
    return sns_topic_tag_list

def get_sns_topic_attributes(sns_client, sns_arn):
    response = sns_client.get_topic_attributes(TopicArn=sns_arn)
    
    sns_topic_name = response['Attributes']['DisplayName']
    sns_topic_arn = response['Attributes']['TopicArn']
    try:
        sns_topic_kms_id = response['Attributes']['KmsMasterKeyId']
    except:
        sns_topic_kms_id = '-'
    
    return sns_topic_name, sns_topic_arn, sns_topic_kms_id

def export_sns_info_to_excel(workbook, sns_client):
    sns_info = sns_client.list_topics()

    # SNS 시트 생성
    worksheet = workbook.create_sheet('SNS')
    
    # SNS 열 정보 추가
    worksheet.append(['SNS'])
    worksheet.cell(1, 1).font = front_header_font

    sns_headers = [
    # Volume 정보
    'SNS Name', 'SNS ARN', 'KMS ID', 'Tags'
    ]

    for col_num, header in enumerate(sns_headers,1):
        worksheet.cell(2, col_num, value=header).font = header_font
        worksheet.cell(2, col_num, value=header).fill = header_fill
        worksheet.cell(2, col_num, value=header).alignment = header_alignment
        worksheet.cell(2, col_num, value=header).border = header_border
    
    # 현재 행 위치
    header_row = worksheet.max_row
    
    # auto_filter 적용
    worksheet.auto_filter.ref = f"A{header_row}:{chr(64 + len(sns_headers))}{header_row}"


    for sns_arn in sns_info['Topics']:
        sns_topic_name, sns_topic_arn, sns_topic_kms_id = get_sns_topic_attributes(sns_client, sns_arn['TopicArn'])
        sns_tags = get_sns_topic_tags(sns_client, sns_arn['TopicArn'])
        
        variables = [
            # EC2 기본 정보
            sns_topic_name, sns_topic_arn, sns_topic_kms_id, sns_tags
        ]

        worksheet.append(variables)
    
        # 모든 셀 텍스트 높이 가운데 맞춤
        for index, value in enumerate(variables, start=1):
            cell = worksheet.cell(row=worksheet.max_row, column=index)
            cell.alignment = content_alignment
            cell.border = content_border
            # '\n' 즉, 개행이 포함되어 있으면 즉, 셀 값이 다중값이면 텍스트 자동 줄바꿈
            if '\n' in value:
                cell = worksheet.cell(row=worksheet.max_row, column=index)
                cell.alignment = multiple_content_alignment
                cell.border = content_border