from botocore.exceptions import ClientError
from tqdm import tqdm
from conf import convert
from conf.sheet_style import front_header_font,header_font,header_fill,header_alignment,content_alignment,multiple_content_alignment,content_border,header_border

# ACM 상세 정보 조회
def get_certificate_info(acm_client, acm_arn):
    response = acm_client.describe_certificate(CertificateArn=acm_arn)
    
    acm_type = response['Certificate']['Type']
    acm_domain = response['Certificate']['DomainName']
    acm_sub_domain = '\n'.join(response['Certificate']['SubjectAlternativeNames'])
    acm_key_algorithm = response['Certificate']['KeyAlgorithm']
    acm_issuer = response['Certificate']['Issuer']
    
    return acm_type, acm_domain, acm_sub_domain, acm_key_algorithm, acm_issuer

# ACM 태그 정보 조회
def get_certificate_tag(acm_client, acm_arn):
    response = acm_client.list_tags_for_certificate(CertificateArn=acm_arn)
    
    acm_name = convert.name_tag_info(response['Tags'])
    acm_tag = convert.tag_info(response['Tags'])
    
    return acm_name, acm_tag

def export_acm_info_to_excel(workbook, acm_client):
#====================================== ACM Section ======================================
    # ACM 목록 조회
    acm_info = acm_client.list_certificates()

    # ACM 시트 생성
    worksheet = workbook.create_sheet('ACM')
    
    # ACM 열 정보 추가
    worksheet.append(['ACM'])
    worksheet.cell(1, 1).font = front_header_font

    acm_headers = [
        # ACM 기본 정보
        'ACM Name', 'ACM Type', 'ACM Domain', 'ACM Sub Domain', 'ACM Key Algorithm', 'ACM Issuer', 'ACM ARN',
        # ACM 태그 정보
        'Tags'
    ]

    for col_num, header in enumerate(acm_headers,1):
        worksheet.cell(2, col_num, value=header).font = header_font
        worksheet.cell(2, col_num, value=header).fill = header_fill
        worksheet.cell(2, col_num, value=header).alignment = header_alignment
        worksheet.cell(2, col_num, value=header).border = header_border

    # 현재 행 위치
    header_row = worksheet.max_row
    
    # auto_filter 적용
    worksheet.auto_filter.ref = f"A{header_row}:{chr(64 + len(acm_headers))}{header_row}"

    # ACM 정보 조회
    for acm in tqdm(acm_info['CertificateSummaryList'], desc="ACM 정보 파싱 중..."):
        acm_arn = acm['CertificateArn']
        acm_name, acm_tag = get_certificate_tag(acm_client, acm_arn)
        acm_type, acm_domain, acm_sub_domain, acm_key_algorithm, acm_issuer = get_certificate_info(acm_client, acm_arn)
        
        variables = [
            # ACM 기본 정보
            acm_name, acm_type, acm_domain, acm_sub_domain, acm_key_algorithm, acm_issuer, acm_arn,
            # ACM 태그 정보
            acm_tag
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