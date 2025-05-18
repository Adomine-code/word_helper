# Word Helper API 参考文档

## WordDocument 类

### 描述
`WordDocument` 是一个用于处理 Microsoft Word (.docx) 文件的类。它提供了打开 Word 文档和提取内容的功能。

### 构造函数

```python
class WordDocument(file_path: str)
```

- **file_path**: Word 文档文件的路径。

### 方法

#### open()

```python
def open() -> bool
```

尝试打开由 `file_path` 指定的 Word 文档。

- **返回值**: 如果成功打开文件则返回 `True`，否则返回 `False`。

#### get_content()

```python
def get_content() -> Optional[str]
```

获取已打开 Word 文档的文本内容。

- **返回值**: 如果文档已打开，则返回文档内容字符串；否则返回 `None`。