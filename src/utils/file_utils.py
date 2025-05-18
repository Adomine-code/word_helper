
class WordDocument:
    """用于打开和读取Word文档内容的类"""
    
    def __init__(self, file_path):
        """初始化Word文档对象
        
        Args:
            file_path (str): Word文件的路径
        """
        self.file_path = file_path
        self.doc = None
        self.content = None
    
    def open(self):
        """打开Word文档并加载内容"""
        try:
            from docx import Document
            self.doc = Document(self.file_path)
            return True
        except Exception as e:
            print(f"Error opening file: {e}")
            return False
    
    def get_content(self):
        """获取文档内容
        
        Returns:
            str: 文档内容，如果未打开文件则返回None
        """
        if self.doc:
            self.content = '\n'.join(paragraph.text for paragraph in self.doc.paragraphs)
        return self.content
    