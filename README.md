```bash
# 重新构建镜像
sudo docker compose up  --build
```

# 文件格式转换API

这是一个基于Python和Flask的文件格式转换后端服务，提供API接口支持各种文档格式的转换。

## 支持的转换格式

1. PDF转其他格式：
   - PDF转Word (.docx)
   - PDF转Excel (.xlsx)
   - PDF转PowerPoint (.pptx)
   - PDF转图片 (.jpg/.png)

2. 其他格式转PDF：
   - 图片转PDF (.jpg/.jpeg/.png)
   - Word转PDF (.doc/.docx)
   - Excel转PDF (.xls/.xlsx)
   - PowerPoint转PDF (.ppt/.pptx)
   - 文本转PDF (.txt)
   - Markdown转PDF (.md)

## 安装和运行

### 安装依赖

```bash
pip install -r requirements.txt
```

注意：部分功能可能需要额外安装系统依赖：

- LibreOffice：用于Office文档的转换
- Poppler：用于PDF转图片功能

### 运行服务

```bash
python app.py
```

默认情况下，服务将在 http://0.0.0.0:5000 上运行。

### 命令行参数

- `--host`: 指定主机地址，默认为 0.0.0.0
- `--port`: 指定端口，默认为 5000
- `--debug`: 启用调试模式

例如：
```bash
python app.py --host 127.0.0.1 --port 8000 --debug
```

## API接口说明

### 1. 查看支持的格式
GET /api/formats
### 2. 转换文件
POST /api/convert


表单参数：
- `file`：要转换的文件
- `from_format`：源格式（可选，默认从文件扩展名判断）
- `to_format`：目标格式
- `quality`：质量等级（1=低，2=中，3=高，默认为2）

返回JSON：
```json
{
  "success": true,
  "file_id": "uuid...",
  "original_name": "example.pdf",
  "from_format": "pdf",
  "to_format": "docx",
  "file_size": 12345,
  "result_url": "/api/download/uuid...docx",
  "converted_time": "2023-05-18 14:30:45"
}
```

### 3. 下载转换后的文件
GET /api/download/{filename}

## 与微信小程序联调

在微信小程序中，使用 `wx.uploadFile` 调用本API进行文件转换，使用 `wx.downloadFile` 下载转换后的文件。具体实现请参考微信小程序开发文档。

## 注意事项

- 本服务默认会清理24小时以上的上传文件和转换结果
- 请确保服务器有足够的存储空间和处理能力
- 部分转换功能可能依赖第三方软件，如LibreOffice
