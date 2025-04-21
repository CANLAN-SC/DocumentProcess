import os
from docx import Document
from docxcompose.composer import Composer

# 配置路径
input_folder = '专利'
word_folder = os.path.join(input_folder, input_folder + '_docx')
output_path = os.path.join(input_folder, os.path.basename(input_folder) + '_合并.docx')

# 找到所有docx文件并排序
docx_files = sorted(
    [f for f in os.listdir(word_folder) if f.lower().endswith('.docx')],
    key=lambda x: x.lower()
)

# 创建主文档
master = Document(os.path.join(word_folder, docx_files[0]))
composer = Composer(master)

# 逐个附加其余文档
for docx_file in docx_files[1:]:
    sub_doc_path = os.path.join(word_folder, docx_file)
    sub_doc = Document(sub_doc_path)
    composer.append(sub_doc)

# 保存合并结果
composer.save(output_path)
print(f"文档合并完成，保留所有图片：{output_path}")
