from conf.sheet_style import front_header_font,header_font,header_fill,header_alignment,content_alignment,multiple_content_alignment,content_border,header_border

def get_codebuild_tag(codebuild_tags):
    # Tag 존재 여부 확인
    if len(codebuild_tags) == 0:
        codebuild_tag_list = '-'
    else:
        codebuild_tag_list = []
        for codebuild_tag in codebuild_tags:
            codebuild_tag_list.append(f"{codebuild_tag['key']} : {codebuild_tag['value']}")
        
        codebuild_tag_list = '\n'.join(codebuild_tag_list)
    
    return codebuild_tag_list


def export_codebuild_info_to_excel(workbook, codebuild_client):
    codebuild_project_list = codebuild_client.list_projects(sortBy='NAME')
    
    # CodeBuild 시트 생성
    worksheet = workbook.create_sheet('CodeBuild')
    
    # CodeBuild 열 정보 추가
    worksheet.append(['CodeBuild'])
    worksheet.cell(1, 1).font = front_header_font
    
    codebuild_headers = ['Build Project', 'Description', 'Tags', 'Service Role']

    for col_num, header in enumerate(codebuild_headers,1):
        worksheet.cell(2, col_num, value=header).font = header_font
        worksheet.cell(2, col_num, value=header).fill = header_fill
        worksheet.cell(2, col_num, value=header).alignment = header_alignment
        worksheet.cell(2, col_num, value=header).border = header_border

    # 현재 행 위치
    header_row = worksheet.max_row
    
    # auto_filter 적용
    worksheet.auto_filter.ref = f"A{header_row}:{chr(64 + len(codebuild_headers))}{header_row}"

    if len(codebuild_project_list['projects']) > 0:
        codebuild_info = codebuild_client.batch_get_projects(names=codebuild_project_list['projects'])

        # Private CodeBuild 정보를 엑셀에 쓰기
        for codebuild in codebuild_info['projects']:
            codebuild_name = codebuild['name']
            codebuild_tag = get_codebuild_tag(codebuild['tags'])
            
            if 'description' in codebuild:
                codebuild_description = codebuild['description']
            else:
                codebuild_description = '-'
            
            codebuild_servicerole = codebuild['serviceRole']#.split('/',2)[2]

            variables = [
                codebuild_name, codebuild_description, codebuild_tag, codebuild_servicerole
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