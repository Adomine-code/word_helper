import unittest
import os
from src.utils.file_utils import WordDocument

class TestWordDocument(unittest.TestCase):
    """测试WordDocument类的功能"""

    def setUp(self):
        """在每个测试方法前创建一个测试文档"""
        from docx import Document
        
        # 创建测试目录（如果不存在）
        self.test_dir = 'test_files'
        if not os.path.exists(self.test_dir):
            os.makedirs(self.test_dir)
        
        # 创建测试用的Word文档
        self.test_file = os.path.join(self.test_dir, 'test.docx')
        doc = Document()
        doc.add_paragraph('This is a test document.')
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
        self.assertIn('This is a test document.', content)

    def test_get_content_before_open(self):
        """测试未打开文档时获取内容的功能"""
        doc = WordDocument(self.test_file)
        content = doc.get_content()
        self.assertIsNone(content)