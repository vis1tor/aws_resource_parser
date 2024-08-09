from conf.sheet_style import front_header_font,header_font,header_fill,header_alignment,content_alignment,multiple_content_alignment,content_border,header_border

def get_repository_tag(codecommit_client, repository_arn):
    response = codecommit_client.list_tags_for_resource(resourceArn=repository_arn)
    
    # Tag 존재 여부 확인
    if len(response['tags']) == 0:
        repository_tag_list = '-'
    else:
        repository_tag_list = []
    
        for tag_key, tag_value in response['tags'].items():
            repository_tag_list.append(f"{tag_key} : {tag_value}")

        repository_tag_list = '\n'.join(repository_tag_list)

    return repository_tag_list


def get_repository(codecommit_client, repository_name):
    response = codecommit_client.get_repository(repositoryName=repository_name)
    
    repository_id = response['repositoryMetadata']['repositoryId']
    repository_http_url = response['repositoryMetadata']['cloneUrlHttp']
    repository_arn = response['repositoryMetadata']['Arn']
    repository_enc_key = response['repositoryMetadata']['kmsKeyId']

    if 'repositoryDescription' not in response['repositoryMetadata']:
        repository_description = '-'
    else:
        repository_description = response['repositoryMetadata']['repositoryDescription']

    return repository_id, repository_http_url, repository_arn, repository_enc_key, repository_description


def export_codecommit_info_to_excel(workbook, codecommit_client):
    codecommit_info = codecommit_client.list_repositories()

    # CodeCommit 시트 생성
    worksheet = workbook.create_sheet('CodeCommit')
    
    # CodeCommit 열 정보 추가
    worksheet.append(['CodeCommit'])
    worksheet.cell(1, 1).font = front_header_font
    
    codecommit_headers = ['Repository', 'Repository ID', 'KMS KEY', 'HTTP URL' , 'Repository Description', 'Tags']

    for col_num, header in enumerate(codecommit_headers,1):
        worksheet.cell(2, col_num, value=header).font = header_font
        worksheet.cell(2, col_num, value=header).fill = header_fill
        worksheet.cell(2, col_num, value=header).alignment = header_alignment
        worksheet.cell(2, col_num, value=header).border = header_border

    # 현재 행 위치
    header_row = worksheet.max_row
    
    # auto_filter 적용
    worksheet.auto_filter.ref = f"A{header_row}:{chr(64 + len(codecommit_headers))}{header_row}"

    # Private CodeCommit 정보를 엑셀에 쓰기
    for repo in codecommit_info['repositories']:
        repository_name = repo['repositoryName']
        repository_id, repository_http_url, repository_arn, repository_enc_key, repository_description = get_repository(codecommit_client, repository_name)
        repository_tag = get_repository_tag(codecommit_client, repository_arn)

        variables = [
            repository_name, repository_id, repository_enc_key, repository_http_url, repository_description, repository_tag
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