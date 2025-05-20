"""
file_utils.py

提供与文件操作相关的实用函数，包括 Word 文档处理、路径检查等。
"""

# 修复 W0718、C0415 和 W0212 警告
# 示例：细化异常捕获，移动导入到顶部，添加 pylint disable 注释
from docx import Document


class WordDocument:
    """用于打开和读取Word文档内容的类

    Attributes:
        file_path (str): Word文件的路径
        doc (Document): 文档对象
        content (str): 文档内容
    """

    def __init__(self, file_path):
        """初始化Word文档对象

        Args:
            file_path (str): Word文件的路径
        """
        self.file_path = file_path
        self.doc = None
        self.content = None

    def open(self):
        """打开Word文档并加载内容

        Returns:
            bool: 打开成功返回True，否则返回False
        """
        try:
            self.doc = Document(self.file_path)
            return True
        except FileNotFoundError as e:
            print(f"Error opening file: {e}")
            return False

    def remove_headers(self):
        """删除文档中所有页眉内容

        Returns:
            bool: 删除成功返回True，否则返回False
        """
        try:
            if self.doc:
                for section in self.doc.sections:
                    header = section.header
                    # 从后往前删除，避免索引问题
                    for i in range(len(header.paragraphs) - 1, -1, -1):
                        p = header.paragraphs[i]
                        p._element.getparent().remove(p._element)
                        p._element._element = None  # 确保完全删除
            return True
        except Exception as e:
            print(f"Error removing headers: {e}")
            return False

    def remove_footers(self):
        """删除文档中所有页脚内容

        Returns:
            bool: 删除成功返回True，否则返回False
        """

        try:
            if self.doc:
                for section in self.doc.sections:
                    footer = section.footer
                    # 清除所有段落的内容
                    # 从后往前删除，避免索引问题
                for i in range(len(footer.paragraphs) - 1, -1, -1):
                    p = footer.paragraphs[i]
                    p._element.getparent().remove(p._element)
                    p._element._element = None  # 确保完全删除
            return True
        except Exception as e:
            print(f"Error removing footers: {e}")
            return False

    def add_page_numbers(self):
        """在文档页脚添加页码

        Returns:
            bool: 添加页码成功返回True，否则返回False
        """
        try:
            if self.doc:
                from docx.enum.text import WD_ALIGN_PARAGRAPH
                from docx.oxml.ns import qn
                from docx.oxml import OxmlElement

                for section in self.doc.sections:
                    footer = section.footer
                    if len(footer.paragraphs) == 0:
                        paragraph = footer.add_paragraph()
                    else:
                        paragraph = footer.paragraphs[0]
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

                    # 创建页码字段
                    run = paragraph.add_run()
                    run.text = "第"

                    # 创建PAGE字段
                    page_fld = OxmlElement("w:fldSimple")
                    page_fld.set(qn("w:instr"), r"PAGE \* MERGEFORMAT")
                    paragraph._element.append(page_fld)

                    run = paragraph.add_run()
                    run.text = "页/共"

                    # 创建NUMPAGES字段
                    pages_fld = OxmlElement("w:fldSimple")
                    pages_fld.set(qn("w:instr"), r"NUMPAGES \* MERGEFORMAT")
                    paragraph._element.append(pages_fld)

                    run = paragraph.add_run()
                    run.text = "页"
            return True
        except Exception as e:
            print(f"Error adding page numbers: {e}")
            return False

    def save(self, output_path=None):
        """保存文档到指定路径

        Args:
            output_path (str, optional): 输出文件路径。如果未指定，将覆盖原始文件。
                Defaults to None.

        Returns:
            bool: 保存成功返回True，否则返回False
        """
        try:
            if self.doc:
                # 如果未指定输出路径，则使用原始文件路径
                if output_path is None:
                    output_path = self.file_path

                # 保存文档
                self.doc.save(output_path)
                return True
            return False
        except Exception as e:
            print(f"Error saving file: {e}")
            return False
