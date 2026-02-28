import docx
import PyPDF2
import os

# 读取Word文档
def read_word(file_path):
    try:
        doc = docx.Document(file_path)
        text = []
        for para in doc.paragraphs:
            if para.text.strip():
                text.append(para.text)
        return '\n'.join(text)
    except Exception as e:
        return f"Error reading Word file: {str(e)}"

# 读取PDF文档
def read_pdf(file_path):
    try:
        with open(file_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            text = []
            for page_num in range(len(reader.pages)):
                page = reader.pages[page_num]
                text.append(page.extract_text())
            return '\n'.join(text)
    except Exception as e:
        return f"Error reading PDF file: {str(e)}"

# 主函数
def main():
    files = ['2.7改.docx', '演讲稿改版.docx', '杏林薪火队..pdf']
    for file in files:
        if file.endswith('.docx'):
            content = read_word(file)
        elif file.endswith('.pdf'):
            content = read_pdf(file)
        else:
            content = "Unsupported file format"
        
        print(f"\n=== {file} ===")
        print(content[:2000])  # 只打印前2000个字符，避免输出过多

if __name__ == "__main__":
    main()
