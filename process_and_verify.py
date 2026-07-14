import os
import re
from datetime import datetime

def find_today_reports():
    today_str = datetime.today().strftime('%Y-%m-%d')
    
    # 작업 디렉토리 및 불필요한 작업파일 디렉토리 탐색
    search_dirs = ['.', './불필요한 작업파일']
    
    domestic_md = None
    foreign_md = None
    
    for s_dir in search_dirs:
        if not os.path.exists(s_dir):
            continue
        for file in os.listdir(s_dir):
            if file.startswith(today_str) and file.endswith('.md'):
                full_path = os.path.join(s_dir, file)
                if '해외' in file:
                    foreign_md = full_path
                else:
                    domestic_md = full_path
                    
    # 만약 오늘자 파일이 없다면 가장 최신 파일들로 대체 시도
    if not domestic_md or not foreign_md:
        print("Today's markdown files not fully found. Searching for most recent ones...")
        all_mds = []
        for s_dir in search_dirs:
            if not os.path.exists(s_dir):
                continue
            for file in os.listdir(s_dir):
                if file.endswith('.md'):
                    full_path = os.path.join(s_dir, file)
                    all_mds.append((full_path, os.path.getmtime(full_path)))
        
        # 수정 시간 역순 정렬
        all_mds.sort(key=lambda x: x[1], reverse=True)
        
        for path, _ in all_mds:
            if '해외' in path and not foreign_md:
                foreign_md = path
            elif '해외' not in path and not domestic_md:
                domestic_md = path
                
    return domestic_md, foreign_md

def markdown_to_html(md_content):
    # 간단한 마크다운 -> HTML 변환 (네이버 블로그용 서식화)
    lines = md_content.split('\n')
    html_lines = []
    
    in_table = False
    table_headers = []
    table_rows = []
    
    for line in lines:
        line_trimmed = line.strip()
        
        # 테이블 처리
        if line_trimmed.startswith('|'):
            in_table = True
            # 구분선 행 무시 (|---|---|)
            if re.match(r'^\s*\|?\s*:\s*[-=]+', line_trimmed) or '---' in line_trimmed:
                continue
            
            # 열 분리
            cols = [c.strip() for c in line_trimmed.split('|')[1:-1]]
            if not table_headers:
                table_headers = cols
            else:
                table_rows.append(cols)
            continue
        else:
            if in_table:
                # 테이블 빌드 및 삽입
                table_html = "<table style='border-collapse: collapse; width: 100%; border: 1px solid #ddd; margin: 15px 0;'>"
                # 헤더
                table_html += "<tr style='background-color: #f2f2f2; text-align: left;'>"
                for h in table_headers:
                    table_html += f"<th style='border: 1px solid #ddd; padding: 10px; font-weight: bold;'>{h}</th>"
                table_html += "</tr>"
                # 데이터
                for row in table_rows:
                    table_html += "<tr>"
                    for val in row:
                        # 링크 변환
                        val_cleaned = re.sub(r'\[([^\]]+)\]\(([^)]+)\)', r'<a href="\2" target="_blank" style="color: #0073e6; text-decoration: none;">\1</a>', val)
                        table_html += f"<td style='border: 1px solid #ddd; padding: 10px;'>{val_cleaned}</td>"
                    table_html += "</tr>"
                table_html += "</table>"
                html_lines.append(table_html)
                
                # 초기화
                in_table = False
                table_headers = []
                table_rows = []
                
        # 제목 처리
        if line_trimmed.startswith('# '):
            html_lines.append(f"<h1 style='font-size: 24px; color: #111; border-bottom: 2px solid #222; padding-bottom: 8px; margin-top: 30px;'>{line_trimmed[2:]}</h1>")
        elif line_trimmed.startswith('## '):
            html_lines.append(f"<h2 style='font-size: 20px; color: #0066cc; margin-top: 25px; border-bottom: 1px solid #ddd; padding-bottom: 5px;'>{line_trimmed[3:]}</h2>")
        elif line_trimmed.startswith('### '):
            html_lines.append(f"<h3 style='font-size: 16px; color: #333; margin-top: 20px;'>{line_trimmed[4:]}</h3>")
        # 수평선
        elif line_trimmed == '---':
            html_lines.append("<hr style='border: 0; height: 1px; background: #ccc; margin: 20px 0;'>")
        # 빈 줄
        elif not line_trimmed:
            html_lines.append("<p>&nbsp;</p>")
        # 일반 본문
        else:
            # 굵은 글씨 변환
            clean_line = re.sub(r'\*\*([^*]+)\*\*', r'<strong>\1</strong>', line_trimmed)
            # 링크 변환
            clean_line = re.sub(r'\[([^\]]+)\]\(([^)]+)\)', r'<a href="\2" target="_blank" style="color: #0073e6; text-decoration: none;">\1</a>', clean_line)
            html_lines.append(f"<p style='line-height: 1.6; color: #444; margin: 8px 0;'>{clean_line}</p>")
            
    # 남아 있는 테이블 처리
    if in_table and table_headers:
        table_html = "<table style='border-collapse: collapse; width: 100%; border: 1px solid #ddd; margin: 15px 0;'>"
        table_html += "<tr style='background-color: #f2f2f2; text-align: left;'>"
        for h in table_headers:
            table_html += f"<th style='border: 1px solid #ddd; padding: 10px; font-weight: bold;'>{h}</th>"
        table_html += "</tr>"
        for row in table_rows:
            table_html += "<tr>"
            for val in row:
                val_cleaned = re.sub(r'\[([^\]]+)\]\(([^)]+)\)', r'<a href="\2" target="_blank" style="color: #0073e6; text-decoration: none;">\1</a>', val)
                table_html += f"<td style='border: 1px solid #ddd; padding: 10px;'>{val_cleaned}</td>"
            table_html += "</tr>"
        table_html += "</table>"
        html_lines.append(table_html)
        
    return '\n'.join(html_lines)

def main():
    print("Searching for domestic and foreign box office report markdown files...")
    dom_path, for_path = find_today_reports()
    
    if not dom_path:
        print("Warning: Domestic report markdown not found!")
    if not for_path:
        print("Warning: Foreign report markdown not found!")
        
    merged_html = ""
    merged_html += "<div style='font-family: \"Nanum Gothic\", sans-serif; max-width: 800px; margin: 0 auto;'>"
    merged_html += f"<p style='color: #777; font-size: 14px;'>자동 생성 일시: {datetime.today().strftime('%Y-%m-%d %H:%M:%S')}</p>"
    
    # 1. 국내 박스오피스 본문 가공
    if dom_path:
        print(f"Reading domestic report: {dom_path}")
        with open(dom_path, 'r', encoding='utf-8') as f:
            content = f.read()
        merged_html += "<!-- Domestic Box Office -->"
        merged_html += markdown_to_html(content)
        
    merged_html += "<br><hr style='border: 0; border-top: 2px dashed #999; margin: 40px 0;'><br>"
    
    # 2. 해외 박스오피스 본문 가공
    if for_path:
        print(f"Reading foreign report: {for_path}")
        with open(for_path, 'r', encoding='utf-8') as f:
            content = f.read()
        merged_html += "<!-- Foreign Box Office -->"
        merged_html += markdown_to_html(content)
        
    # 사실관계 및 출처 표기 하단 고정 추가
    merged_html += "<br><hr style='border: 0; height: 1px; background: #ccc; margin: 30px 0;'>"
    merged_html += "<div style='background-color: #f9f9f9; padding: 15px; border-left: 4px solid #4caf50; font-size: 13px; color: #555;'>"
    merged_html += "<p style='margin: 0; font-weight: bold;'>⚠️ 사실관계 검증 및 출처 안내</p>"
    merged_html += "<p style='margin: 5px 0 0 0;'>본 보고서의 데이터는 영화진흥위원회 통합전산망(KOBIS) 및 해외 박스오피스 모조(Box Office Mojo) 등 공인된 데이터 소스를 기반으로 AI가 자동 수집 및 크로스체크하여 작성되었습니다.</p>"
    merged_html += "</div>"
    merged_html += "</div>"
    
    # 최종 결과 저장
    output_path = "blog_post.html"
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(merged_html)
        
    print(f"Successfully generated merged blog HTML at: {output_path}")

if __name__ == "__main__":
    main()
