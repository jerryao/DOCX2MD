import re
from werkzeug.utils import secure_filename

def safe_filename(filename):
    """自定义的安全文件名处理函数，保留中文字符"""
    # 保留中文字符、字母、数字和一些安全字符，并移除不安全字符
    safe_chars = re.sub(r'[^\w\u4e00-\u9fa5\-\.]', '_', filename)
    return safe_chars

# 测试文件名
test_filenames = [
    '绿证全覆盖.docx',
    '测试文档123.docx',
    'document with spaces.docx',
    '带有特殊符号!@#$%^&*().docx',
    'mixed-中文-english.docx'
]

print("文件名测试结果对比:")
print("-" * 60)
print("原始文件名 | werkzeug secure_filename | 自定义 safe_filename")
print("-" * 60)

for filename in test_filenames:
    werkzeug_result = secure_filename(filename)
    custom_result = safe_filename(filename)
    print(f"{filename} | {werkzeug_result} | {custom_result}")

print("-" * 60)

# 测试转换后的MD文件名
for filename in test_filenames:
    custom_result = safe_filename(filename)
    md_filename = custom_result.rsplit('.', 1)[0] + '.md'
    print(f"DOCX: {filename} -> MD: {md_filename}") 