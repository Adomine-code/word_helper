import unittest
import os
from src.utils.file_utils import WordDocument


"""
test_word_document.py

测试 word_document 模块中的功能，包括文档内容读取与修改。
"""


class TestWordDocument(unittest.TestCase):
    """测试WordDocument类的功能"""

    def setUp(self):
        """在每个测试方法前创建一个测试文档"""
        from docx import Document

        # 创建测试目录（如果不存在）
        self.test_dir = "test_files"
        if not os.path.exists(self.test_dir):
            os.makedirs(self.test_dir)

        # 创建测试用的Word文档
        self.test_file = os.path.join(self.test_dir, "test.docx")
        doc = Document()
        doc.add_paragraph("This is a test document.")
        doc.save(self.test_file)

    def test_open_success(self):
        """测试成功打开文档的情况"""
        doc = WordDocument(self.test_file)
        result = doc.open()
        self.assertTrue(result)

    def test_open_failure(self):
        """测试无法打开文档的情况"""
        doc = WordDocument("non_existent_file.docx")
        result = doc.open()
        self.assertFalse(result)

    def test_get_content_after_open(self):
        """测试打开文档后获取内容的功能"""
        doc = WordDocument(self.test_file)
        doc.open()
        content = doc.get_content()
        self.assertIsNotNone(content)
        self.assertIn("This is a test document.", content)

    def test_get_content_before_open(self):
        """测试未打开文档时获取内容的功能"""
        doc = WordDocument(self.test_file)
        content = doc.get_content()
        self.assertIsNone(content)

    def test_remove_headers(self):
        """测试删除页眉功能"""
        doc = WordDocument(self.test_file)
        doc.open()
        # 添加测试页眉内容
        for section in doc.doc.sections:
            header = section.header
            paragraph = header.paragraphs[0]
            paragraph.text = "Test Header"

        # 验证页眉内容存在
        for section in doc.doc.sections:
            self.assertEqual(section.header.paragraphs[0].text, "Test Header")

        # 删除页眉并验证
        result = doc.remove_headers()
        self.assertTrue(result)
        for section in doc.doc.sections:
            # 验证页眉内容是否不存在
            self.assertEqual(section.header.paragraphs, [])

    def test_remove_footers(self):
        """测试删除页脚功能"""
        doc = WordDocument(self.test_file)
        doc.open()
        # 添加测试页脚内容
        for section in doc.doc.sections:
            footer = section.footer
            paragraph = footer.paragraphs[0]
            paragraph.text = "Test Footer"

        # 验证页脚内容存在
        for section in doc.doc.sections:
            self.assertEqual(section.footer.paragraphs[0].text, "Test Footer")

        # 删除页脚并验证
        result = doc.remove_footers()
        self.assertTrue(result)
        for section in doc.doc.sections:
            # 验证页脚内容是否不存在
            self.assertEqual(section.footer.paragraphs, [])

    def test_add_page_numbers(self):
        """测试添加页码功能"""
        doc = WordDocument(self.test_file)
        doc.open()

        # 添加页码
        result = doc.add_page_numbers()
        self.assertTrue(result)

        # 验证页脚内容
        for section in doc.doc.sections:
            footer = section.footer
            self.assertTrue("第" in footer.paragraphs[0].text)
            self.assertTrue("页/共" in footer.paragraphs[0].text)
            self.assertTrue("页" in footer.paragraphs[0].text)

    def test_save(self):
        """测试保存文档功能"""
        doc = WordDocument(self.test_file)
        doc.open()

        # 修改文档内容
        doc.doc.add_paragraph("This is a new paragraph.")

        # 定义保存路径
        output_file = os.path.join(self.test_dir, "test_output.docx")

        # 保存文档
        result = doc.save(output_file)
        self.assertTrue(result)

        # 验证文件是否存在
        self.assertTrue(os.path.exists(output_file))

        # 验证保存后的文档内容
        saved_doc = WordDocument(output_file)
        saved_doc.open()
        content = saved_doc.get_content()
        self.assertIsNotNone(content)
        self.assertIn("This is a new paragraph", content)
