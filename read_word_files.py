import docx

def read_docx(file_path):
    try:
        doc = docx.Document(file_path)
        content = []
        for para in doc.paragraphs:
            if para.text.strip():
                content.append(para.text)
        return content
    except Exception as e:
        return [f"Error reading file: {str(e)}"]

# 读取两个Word文档
file1 = '演讲稿改版.docx'
file2 = '2.7改.docx'

print(f"\n--- {file1} 内容 ---\n")
content1 = read_docx(file1)
for line in content1:
    print(line)

print(f"\n--- {file2} 内容 ---\n")
content2 = read_docx(file2)
for line in content2:
    print(line)
