from docx import Document
import pandas as pd
import re

def extract_document_structure(docx_path, excel_path):
    doc = Document(docx_path)
    elements = []  # 見出し、表、図の情報を格納
    
    # 各レベルの現在の番号を管理
    current_numbers = {1: 0, 2: 0, 3: 0, 4: 0, 5: 0}
    current_headings = {1: '', 2: '', 3: '', 4: '', 5: ''}  # 現在の見出しを保持
    
    def get_current_section_number(level):
        number_parts = []
        for i in range(1, level + 1):
            if current_numbers[i] > 0:
                number_parts.append(str(current_numbers[i]))
        return '.'.join(number_parts)
    
    def is_caption(text):
        # 表や図のキャプションかどうかを判定
        patterns = [
            r'^表\s*\d+',
            r'^図\s*\d+',
            r'^Tab\w*\.*\s*\d+',
            r'^Fig\.*\s*\d+'
        ]
        return any(re.match(pattern, text.strip()) for pattern in patterns)
    
    current_level = 0
    
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if not text:  # 空の段落をスキップ
            continue
            
        element_info = None
        
        if paragraph.style.name.startswith('Heading'):
            # 見出しの処理
            level = int(paragraph.style.name.split()[-1])
            current_level = level
            
            # 現在のレベルの番号を増やす
            current_numbers[level] += 1
            # より深いレベルの番号をリセット
            for i in range(level + 1, 6):
                current_numbers[i] = 0
            
            section_number = get_current_section_number(level)
            full_text = f"{section_number} {text}"
            
            # 現在の見出しを更新
            current_headings[level] = full_text
            for i in range(level + 1, 6):
                current_headings[i] = ''
                
            element_info = {
                'type': 'heading',
                'level': level,
                'number': section_number,
                'text': text,
                'full_text': full_text,
                'parent_heading': current_headings[level-1] if level > 1 else ''
            }
            
        elif is_caption(text):
            # 表・図の処理
            element_type = '表' if text.startswith(('表', 'Tab')) else '図'
            element_info = {
                'type': element_type,
                'level': current_level + 1,  # 現在の見出しの1つ下のレベル
                'number': '',  # 表・図独自の番号は維持
                'text': text,
                'full_text': text,
                'parent_heading': current_headings[current_level]
            }
        
        if element_info:
            elements.append(element_info)
    
    # DataFrameの作成
    df1 = pd.DataFrame(elements)
    
    # レベル別のDataFrame作成
    max_level = max(df1['level'])
    rows = []
    current_elements = [''] * max_level
    
    for _, row in df1.iterrows():
        level = row['level']
        text = row['full_text']
        
        # 現在のレベルの要素を更新
        current_elements[level-1] = text
        
        # より深いレベルの要素をクリア
        for i in range(level, max_level):
            current_elements[i] = ''
        
        rows.append(current_elements[:])
    
    # レベル別DataFrameの作成
    columns = [f'Level {i}' for i in range(1, max_level + 1)]
    df2 = pd.DataFrame(rows, columns=columns)
    
    # Excelファイルに保存
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        df1.to_excel(writer, sheet_name='全要素一覧', index=False)
        df2.to_excel(writer, sheet_name='レベル別構造', index=False)

# 使用例
if __name__ == "__main__":
    word_file = "input.docx"
    excel_file = "document_structure.xlsx"
    
    try:
        extract_document_structure(word_file, excel_file)
        print("文書構造の抽出が完了しました。")
    except Exception as e:
        print(f"エラーが発生しました: {str(e)}")
