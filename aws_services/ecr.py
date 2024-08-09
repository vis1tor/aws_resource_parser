from conf.sheet_style import front_header_font,header_font,header_fill,header_alignment,content_alignment,multiple_content_alignment,content_border,header_border

def get_ecr_tags(ecr_client, repositoryArn):
    response = ecr_client.list_tags_for_resource(resourceArn=repositoryArn)
    ecr_tag_list = []

    # ECR 태그 미존재 시 예외 처리
    if len(response['tags']) == 0:
        ecr_tag_list = '-'
    else:
        for tag in response['tags']:
            ecr_tag_list.append(f"{tag['Key']} : {tag['Value']}")

        ecr_tag_list = '\n'.join(ecr_tag_list)

    return ecr_tag_list


def export_ecr_info_to_excel(workbook, private_ecr_client, public_ecr_client):
    private_ecr_info = private_ecr_client.describe_repositories()
    public_ecr_info = public_ecr_client.describe_repositories()
    
    # ECR 시트 생성
    worksheet = workbook.create_sheet('ECR')
    
    # Private ECR 열 정보 추가
    worksheet.append(['Private ECR'])
    worksheet.cell(1, 1).font = front_header_font
    
    private_ecr_headers = ['Repository', 'URI', 'Immutability', 'Scan Frequency', 'Encryption Type', 'Encryption Key', 'Tags']
    
    for col_num, header in enumerate(private_ecr_headers,1):
        worksheet.cell(2, col_num, value=header).font = header_font
        worksheet.cell(2, col_num, value=header).fill = header_fill
        worksheet.cell(2, col_num, value=header).alignment = header_alignment
        worksheet.cell(2, col_num, value=header).border = header_border

    # 현재 행 위치
    header_row = worksheet.max_row
    
    # auto_filter 적용
    worksheet.auto_filter.ref = f"A{header_row}:{chr(64 + len(private_ecr_headers))}{header_row}"

    # Private ECR 정보를 엑셀에 쓰기
    for repo in private_ecr_info['repositories']:
        repository_name = repo['repositoryName']
        repository_uri = repo['repositoryUri']
        repository_immutability = repo['imageTagMutability']
        repository_imagescan = repo['imageScanningConfiguration']['scanOnPush']
        repository_enc_type = repo['encryptionConfiguration']['encryptionType']
        repository_tags = get_ecr_tags(private_ecr_client, repo['repositoryArn'])

        # AES256 일때 kmsKey 예외처리
        if not repo['encryptionConfiguration']['encryptionType'] == 'KMS':
            repository_enc_key = '-'
        else:
            repository_enc_key = repo['encryptionConfiguration']['kmsKey']

        # 시트에 데이터 쓰기
        # repository_imagescan 값은 bool 이라 str으로 변환해야 바로 아래 for 반복에서 에러 발생 안함.
        variables = [
            repository_name, repository_uri, repository_immutability, str(repository_imagescan), repository_enc_type, repository_enc_key, repository_tags
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


    # Private ECR 열 정보 추가
    worksheet.append([None])
    worksheet.append(['Public ECR'])
    max_row = worksheet.max_row
    worksheet.cell(max_row, 1).font = front_header_font

    public_ecr_headers = ['Repository', 'URI', 'Tags']
    for col_num, header in enumerate(public_ecr_headers,1):
        worksheet.cell(max_row+1, col_num, value=header).font = header_font
        worksheet.cell(max_row+1, col_num, value=header).fill = header_fill
        worksheet.cell(max_row+1, col_num, value=header).alignment = header_alignment
        worksheet.cell(max_row+1, col_num, value=header).border = header_border

    # Public ECR 정보를 엑셀에 쓰기
    for repo in public_ecr_info['repositories']:
        repository_name = repo['repositoryName']
        repository_uri = repo['repositoryUri']
        repository_tags = get_ecr_tags(public_ecr_client, repo['repositoryArn'])
        
        # 시트에 데이터 쓰기
        worksheet.append([repository_name, repository_uri, repository_tags])