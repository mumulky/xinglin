try:
    from docx import Document
    
    doc = Document('2.7改.docx')
    text = []
    
    print("\n--- 文档内容 ---")
    for para in doc.paragraphs:
        if para.text.strip():
            print(para.text)
            text.append(para.text)
    
    print("\n--- 提取完成 ---")
except Exception as e:
    print(f"错误: {e}")
