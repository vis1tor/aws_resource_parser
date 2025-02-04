from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

front_header_font = Font(bold=True, size=12)  # 상단 헤더 글꼴
header_font = Font(bold=True, size=10)  # 헤더 글꼴
header_fill = PatternFill(start_color='d0d0d0', end_color='d0d0d0', fill_type='solid')  # 헤더 배경색
header_alignment = Alignment(horizontal='center', vertical='center') # 헤더 가운데 정렬
content_alignment = Alignment(vertical='center', wrap_text=False) # 셀 텍스트 높이 가운데 맞춤
multiple_content_alignment = Alignment(vertical='center', wrap_text=True) # 셀 텍스트 다중값일 때, 텍스트 높이 가운데 맞춤 및 텍스트 자동 줄바꿈
content_border = Border( 
    top=Side(border_style='thin', color='000000'),
    right=Side(border_style='thin', color='000000'),
    bottom=Side(border_style='thin', color='000000'),
    left=Side(border_style='thin', color='000000')
    ) # 콘텐츠 테두리
header_border = Border( 
    top=Side(border_style='thin', color='000000'),
    right=Side(border_style='thin', color='000000'),
    bottom=Side(border_style='thin', color='000000'),
    left=Side(border_style='thin', color='000000')
    ) # 헤더 테두리