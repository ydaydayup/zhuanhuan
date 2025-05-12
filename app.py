from flask import Flask, request, jsonify, send_file, send_from_directory
from werkzeug.utils import secure_filename
import os
import time
import uuid
import converters
from flask_cors import CORS
from utilities import cleanup_old_files
import logging
import json

app = Flask(__name__)
CORS(app)  # 允许跨域请求

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# 文件目录配置
UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
RESULT_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'results')
METADATA_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'metadata')
ALLOWED_EXTENSIONS = {'pdf', 'docx', 'doc', 'xlsx', 'xls', 'pptx', 'ppt', 'jpg', 'jpeg', 'png', 'txt', 'md', 'dwg'}

# 创建必要的目录
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)
os.makedirs(METADATA_FOLDER, exist_ok=True)


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def save_metadata(file_id, metadata):
    """保存文件元数据，包括原始文件名"""
    metadata_file = os.path.join(METADATA_FOLDER, f"{file_id}.json")
    with open(metadata_file, 'w', encoding='utf-8') as f:
        json.dump(metadata, f, ensure_ascii=False, indent=2)
    return metadata_file


def get_metadata(file_id):
    """获取文件元数据"""
    metadata_file = os.path.join(METADATA_FOLDER, f"{file_id}.json")
    if os.path.exists(metadata_file):
        with open(metadata_file, 'r', encoding='utf-8') as f:
            return json.load(f)
    return None


@app.route('/')
def index():
    return jsonify({
        'status': 'running',
        'api_version': '1.0.0',
        'endpoints': {
            '/api/convert': 'POST - 转换文件',
            '/api/download/{file_id}': 'GET - 下载转换后的文件',
            '/api/formats': 'GET - 获取支持的格式列表',
            '/api/list-files': 'GET - 列出上传和结果目录中的文件',
            '/test-upload': 'GET - 提供简单的文件上传测试页面'
        }
    })


@app.route('/api/formats', methods=['GET'])
def get_formats():
    """获取支持的转换格式列表"""
    supported_formats = {
        'pdf': ['docx', 'xlsx', 'pptx', 'jpg', 'png', 'scannable_pdf', 'scanned_pdf', 'searchable_pdf', 'dwg', 'dxf', 'cad'],
        'jpg': ['pdf'],
        'jpeg': ['pdf'],
        'png': ['pdf'],
        'docx': ['pdf'],
        'doc': ['pdf'],
        'xlsx': ['pdf'],
        'xls': ['pdf'],
        'pptx': ['pdf'],
        'ppt': ['pdf'],
        'txt': ['pdf'],
        'md': ['pdf']
    }
    return jsonify(supported_formats)


@app.route('/api/system-check', methods=['GET'])
def system_check():
    """检查系统依赖和转换器状态"""
    try:
        # 检查目录
        directories = {
            "upload_dir": {"path": UPLOAD_FOLDER, "exists": os.path.exists(UPLOAD_FOLDER), "writable": os.access(UPLOAD_FOLDER, os.W_OK)},
            "result_dir": {"path": RESULT_FOLDER, "exists": os.path.exists(RESULT_FOLDER), "writable": os.access(RESULT_FOLDER, os.W_OK)}
        }
        
        # 检查转换器依赖
        converter_status = {}
        
        # 检查PDF转Word依赖
        try:
            from pdf2docx import Converter
            converter_status["pdf2docx"] = "已安装"
        except ImportError:
            converter_status["pdf2docx"] = "未安装"
            
        # 检查PDF转Excel依赖
        try:
            import tabula
            import pandas as pd
            converter_status["tabula_pandas"] = "已安装"
        except ImportError:
            converter_status["tabula_pandas"] = "未安装"
            
        # 检查PDF转图片依赖
        try:
            from pdf2image import convert_from_path
            converter_status["pdf2image"] = "已安装"
        except ImportError:
            converter_status["pdf2image"] = "未安装"
            
        # 检查PIL
        try:
            from PIL import Image
            converter_status["PIL"] = "已安装"
        except ImportError:
            converter_status["PIL"] = "未安装"
            
        # 检查OCR依赖
        try:
            import pytesseract
            converter_status["pytesseract"] = "已安装"
            try:
                version = pytesseract.get_tesseract_version()
                converter_status["tesseract_version"] = str(version)
            except:
                converter_status["tesseract_version"] = "未检测到"
        except ImportError:
            converter_status["pytesseract"] = "未安装"
        
        # 检查中文字体
        chinese_fonts = []
        for font_path in [
            r"C:\Windows\Fonts\simhei.ttf",  # 黑体
            r"C:\Windows\Fonts\simsun.ttc",  # 宋体
            r"C:\Windows\Fonts\msyh.ttc",    # 微软雅黑
        ]:
            if os.path.exists(font_path):
                chinese_fonts.append(os.path.basename(font_path))
        
        converter_status["chinese_fonts"] = chinese_fonts
            
        return jsonify({
            "status": "运行中",
            "directories": directories,
            "converters": converter_status
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route('/api/test-pdf', methods=['GET'])
def test_pdf():
    """测试生成包含中文的PDF"""
    try:
        # 创建一个简单的测试PDF
        from reportlab.pdfgen import canvas
        from reportlab.lib.pagesizes import letter
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        import uuid
        
        # 创建测试文件路径
        test_file_id = str(uuid.uuid4())
        test_filename = "测试中文文档.pdf"  # 使用一个明确的中文文件名进行测试
        
        # 为测试文件创建子目录
        test_dir = os.path.join(RESULT_FOLDER, test_file_id)
        os.makedirs(test_dir, exist_ok=True)
        test_output_path = os.path.join(test_dir, test_filename)
        
        # 创建一个中文测试文本
        test_text = "这是一个中文测试文本。这个PDF应该能正确显示中文字符。"
        
        # 创建PDF
        c = canvas.Canvas(test_output_path, pagesize=letter)
        
        # 尝试注册中文字体
        font_registered = False
        for font_path in [
            r"C:\Windows\Fonts\simhei.ttf",    # 黑体
            r"C:\Windows\Fonts\simsun.ttc",    # 宋体
            r"C:\Windows\Fonts\msyh.ttc"       # 微软雅黑
        ]:
            if os.path.exists(font_path):
                try:
                    font_name = os.path.basename(font_path).split('.')[0]
                    if "simhei" in font_path.lower():
                        pdfmetrics.registerFont(TTFont('SimHei', font_path))
                        c.setFont('SimHei', 14)
                        font_registered = True
                        break
                    elif "simsun" in font_path.lower():
                        pdfmetrics.registerFont(TTFont('SimSun', font_path))
                        c.setFont('SimSun', 14)
                        font_registered = True
                        break
                    elif "msyh" in font_path.lower():
                        pdfmetrics.registerFont(TTFont('MicrosoftYaHei', font_path))
                        c.setFont('MicrosoftYaHei', 14)
                        font_registered = True
                        break
                except Exception as e:
                    logger.warning(f"注册字体失败: {str(e)}")
        
        if not font_registered:
            c.setFont('Helvetica', 14)
            test_text = "Chinese characters test - font registration failed."
        
        # 添加文本
        c.drawString(72, 720, test_text)
        c.drawString(72, 700, "This is a test PDF with Chinese characters.")
        
        # 添加一些额外内容
        c.setFont('Helvetica', 12)
        c.drawString(72, 650, "Font registration success: " + str(font_registered))
        
        # 尝试生成可搜索的PDF内容
        if font_registered:
            c.drawString(72, 600, "这个PDF应该是可搜索的")
            c.drawString(72, 580, "This PDF should be searchable")
        
        # 保存PDF
        c.save()
        
        # 保存元数据以确保能正确下载
        metadata = {
            'original_filename': test_filename,
            'uploaded_filename': test_filename,
            'output_filename': test_filename,
            'file_id': test_file_id,
            'upload_time': time.strftime('%Y-%m-%d %H:%M:%S'),
            'file_size': os.path.getsize(test_output_path)
        }
        save_metadata(test_file_id, metadata)
        
        # 使用新的URL格式
        result_url = f"/api/download/{test_file_id}"
        
        return jsonify({
            "success": True,
            "message": "测试PDF已生成",
            "font_registered": font_registered,
            "download_url": result_url,
            "filename": test_filename
        })
    except Exception as e:
        logger.error(f"生成测试PDF失败: {str(e)}")
        return jsonify({"error": str(e)}), 500


@app.route('/api/convert', methods=['POST'])
def convert_file():
    """文件转换API端点"""
    if 'file' not in request.files:
        return jsonify({'error': '没有文件'}), 400

    file = request.files['file']
    from_format = request.form.get('from_format', '').lower()
    to_format = request.form.get('to_format', '').lower()
    quality = int(request.form.get('quality', 2))  # 1=低, 2=中, 3=高
    
    # 获取前端传递的原始文件名参数
    original_filename = request.form.get('original_filename')

    if file.filename == '':
        return jsonify({'error': '没有选择文件'}), 400

    if not allowed_file(file.filename):
        return jsonify({'error': '不支持的文件格式'}), 400

    # 原始文件名和文件对象详情的调试输出
    logger.info(f"接收到的原始请求文件名: {file.filename}")
    logger.info(f"前端传递的原始文件名: {original_filename}")
    logger.info(f"文件对象详情: {file}")
    
    # 如果前端提供了原始文件名，则使用它；否则使用上传文件的文件名
    if original_filename and original_filename.strip():
        # 确保原始文件名有正确的扩展名
        orig_ext = original_filename.rsplit('.', 1)[1].lower() if '.' in original_filename else ''
        file_ext = file.filename.rsplit('.', 1)[1].lower() if '.' in file.filename else ''
        
        # 如果原始文件名没有扩展名或扩展名与实际文件不符，添加正确的扩展名
        if not orig_ext or orig_ext != file_ext:
            original_filename = f"{original_filename.rsplit('.', 1)[0] if '.' in original_filename else original_filename}.{file_ext}"
            logger.info(f"调整后的原始文件名(添加扩展名): {original_filename}")
    else:
        # 如果没有提供原始文件名，则使用上传的文件名
        original_filename = file.filename
    
    # 生成唯一文件ID用于内部追踪
    file_id = str(uuid.uuid4())
    file_extension = original_filename.rsplit('.', 1)[1].lower()

    if not from_format:
        from_format = file_extension

    # 检查文件名是否已经是随机字符串（来自截图观察到的模式）
    # 如果是，我们需要在元数据中记录两种情况
    meta_original_filename = original_filename
    
    # 保存文件元数据
    metadata = {
        'original_filename': meta_original_filename,
        'uploaded_filename': file.filename,
        'file_id': file_id,
        'upload_time': time.strftime('%Y-%m-%d %H:%M:%S'),
        'from_format': from_format,
        'to_format': to_format,
        'quality': quality
    }
    metadata_file = save_metadata(file_id, metadata)
    logger.info(f"已保存文件元数据: {metadata_file}")

    # 在上传目录中使用原始文件名保存文件
    # 为防止文件名冲突，在服务器端使用子目录(使用file_id创建临时工作目录)
    temp_upload_dir = os.path.join(UPLOAD_FOLDER, file_id)
    os.makedirs(temp_upload_dir, exist_ok=True)
    upload_path = os.path.join(temp_upload_dir, original_filename)
    file.save(upload_path)
    
    # 验证文件是否使用原始文件名保存成功
    logger.info(f"保存的文件路径: {upload_path}")
    if os.path.exists(upload_path):
        logger.info(f"文件成功保存为原始文件名，文件大小: {os.path.getsize(upload_path)} 字节")
    else:
        logger.error(f"文件保存失败，路径不存在: {upload_path}")
        # 列出上传目录内容，检查实际保存的文件名
        directory_contents = list_directory(temp_upload_dir)
        logger.info(f"上传目录内容: {directory_contents}")

    logger.info(f"已接收文件: {original_filename}, 类型: {from_format}, 转换目标: {to_format}")

    try:
        # 清理较旧的文件
        cleanup_old_files(UPLOAD_FOLDER, max_age_hours=24)
        cleanup_old_files(RESULT_FOLDER, max_age_hours=24)
        cleanup_old_files(METADATA_FOLDER, max_age_hours=24)

        # 获取原始文件名（不含扩展名）
        original_name_without_ext = os.path.splitext(original_filename)[0]
        
        # 获取目标格式的扩展名
        target_ext = to_format
        if to_format in ['scannable_pdf', 'scanned_pdf', 'searchable_pdf']:
            target_ext = 'pdf'
        
        # 为结果文件创建一个子目录，防止文件名冲突
        temp_result_dir = os.path.join(RESULT_FOLDER, file_id)
        os.makedirs(temp_result_dir, exist_ok=True)
        
        # 使用原始文件名构建输出文件名
        output_filename = f"{original_name_without_ext}.{target_ext}"
        output_path = os.path.join(temp_result_dir, output_filename)

        # 根据转换类型调用相应的转换函数
        try:
            result = converters.convert_file(
                upload_path,
                to_format,
                output_path,
                quality,
                original_filename  # 传递原始文件名给转换函数
            )
        except Exception as conv_error:
            logger.error(f"文件转换过程中出错: {str(conv_error)}", exc_info=True)
            return jsonify({'error': f"文件转换失败: {str(conv_error)}"}), 500

        # 获取转换后的路径
        if not isinstance(result, dict) or "output_path" not in result:
            error_msg = f"无效的转换结果: {result}"
            logger.error(error_msg)
            return jsonify({'error': error_msg}), 500
            
        result_path = result["output_path"]
        
        # 检查结果文件是否存在
        if not os.path.exists(result_path):
            error_msg = f"转换后的文件不存在: {result_path}"
            logger.error(error_msg)
            return jsonify({'error': error_msg}), 500

        # 获取文件大小
        file_size = os.path.getsize(result_path)
        
        # 如果转换函数生成的文件与预期的输出路径不同，则需要复制到预期的位置并使用原始文件名
        if result_path != output_path:
            logger.info(f"转换函数生成了不同的输出路径: {result_path}，将复制到预期位置: {output_path}")
            import shutil
            # 确保目标目录存在
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            # 复制文件
            shutil.copy2(result_path, output_path)
            # 更新结果路径
            result_path = output_path

        # 更新元数据
        metadata['output_path'] = output_path
        metadata['output_filename'] = output_filename 
        metadata['file_size'] = file_size
        metadata['conversion_time'] = time.strftime('%Y-%m-%d %H:%M:%S')
        save_metadata(file_id, metadata)

        # 生成可访问的URL - 使用file_id作为唯一标识符
        result_url = f"/api/download/{file_id}"

        logger.info(f"转换成功: {original_filename} -> {output_filename}")
        
        # 对于特殊格式，返回正确的格式信息
        response_format = to_format
        if to_format in ['scannable_pdf', 'scanned_pdf', 'searchable_pdf']:
            response_format = 'pdf'
            logger.info(f"特殊格式处理: {to_format} 将在响应中显示为 {response_format}")

        return jsonify({
            'success': True,
            'file_id': file_id,
            'original_name': original_filename,
            'from_format': from_format,
            'to_format': response_format,
            'file_size': file_size,
            'result_url': result_url,
            'converted_time': time.strftime('%Y-%m-%d %H:%M:%S')
        })

    except Exception as e:
        logger.error(f"转换失败: {str(e)}", exc_info=True)
        return jsonify({'error': str(e)}), 500


@app.route('/api/download/<file_id>', methods=['GET'])
def download_file(file_id):
    """下载转换后的文件，使用元数据中保存的原始文件名作为下载名称"""
    try:
        # 获取文件元数据
        metadata = get_metadata(file_id)
        if not metadata:
            logger.error(f"未找到文件元数据: {file_id}")
            return jsonify({'error': '文件不存在或元数据丢失'}), 404
        
        # 获取原始文件名（优先使用元数据中的原始文件名）
        original_filename = metadata.get('original_filename')
        uploaded_filename = metadata.get('uploaded_filename')
        output_filename = metadata.get('output_filename')
        
        if not original_filename or not output_filename:
            logger.error(f"元数据中缺少文件名信息: {metadata}")
            return jsonify({'error': '文件元数据不完整'}), 500
            
        # 构建文件的完整路径
        file_path = os.path.join(RESULT_FOLDER, file_id, output_filename)
        
        if not os.path.exists(file_path):
            logger.error(f"请求的文件不存在: {file_path}")
            return jsonify({'error': '文件不存在'}), 404
        
        # 获取原始文件名的扩展名和输出文件扩展名
        original_name_without_ext = os.path.splitext(original_filename)[0]
        target_ext = os.path.splitext(output_filename)[1]
        
        # 构建下载时使用的文件名 - 使用原始文件名加上目标格式的扩展名
        download_filename = f"{original_name_without_ext}{target_ext}"
        
        # 记录原始文件名和下载文件名
        logger.info(f"下载文件: {file_path}")
        logger.info(f"元数据中的原始文件名: {original_filename}")
        logger.info(f"元数据中的上传文件名: {uploaded_filename}")
        logger.info(f"下载时使用的文件名: {download_filename}")
        
        # 不使用 send_from_directory，改为直接发送文件并设置正确的头部
        from flask import send_file, Response
        import mimetypes
        
        # 获取文件的 MIME 类型
        mime_type, _ = mimetypes.guess_type(file_path)
        if not mime_type:
            mime_type = 'application/octet-stream'  # 默认二进制流
        
        # 确保文件名编码正确
        import urllib.parse
        encoded_filename = urllib.parse.quote(download_filename)
        
        # 使用 send_file 并手动设置 Content-Disposition 头
        response = send_file(
            file_path,
            mimetype=mime_type,
            as_attachment=True
        )
        
        # 为了兼容不同浏览器，同时提供 ASCII 和 UTF-8 编码的文件名
        response.headers.set(
            'Content-Disposition', 
            f'attachment; filename="{encoded_filename}"; filename*=UTF-8\'\'{encoded_filename}'
        )
        
        # 添加缓存控制头，禁止缓存
        response.headers.set('Cache-Control', 'no-cache, no-store, must-revalidate')
        response.headers.set('Pragma', 'no-cache')
        response.headers.set('Expires', '0')
        
        return response
    except Exception as e:
        logger.error(f"下载文件时出错: {str(e)}")
        return jsonify({'error': f"下载文件失败: {str(e)}"}), 500


# 添加辅助函数列出目录内容
def list_directory(directory_path):
    """列出指定目录中的所有文件和子目录"""
    contents = []
    try:
        for item in os.listdir(directory_path):
            item_path = os.path.join(directory_path, item)
            if os.path.isfile(item_path):
                contents.append(f"[file] {item} ({os.path.getsize(item_path)} bytes)")
            elif os.path.isdir(item_path):
                contents.append(f"[dir] {item}")
    except Exception as e:
        logger.error(f"列出目录内容时出错: {str(e)}")
    return contents

@app.route('/api/list-files', methods=['GET'])
def list_files():
    """API端点用于列出上传和结果目录中的文件"""
    upload_contents = list_directory(UPLOAD_FOLDER)
    result_contents = list_directory(RESULT_FOLDER)
    
    # 遍历上传目录的子目录，最多列出5个
    upload_subdirs = []
    for item in os.listdir(UPLOAD_FOLDER):
        item_path = os.path.join(UPLOAD_FOLDER, item)
        if os.path.isdir(item_path):
            upload_subdirs.append({
                "name": item,
                "contents": list_directory(item_path)
            })
            if len(upload_subdirs) >= 5:
                break
    
    # 遍历结果目录的子目录，最多列出5个
    result_subdirs = []
    for item in os.listdir(RESULT_FOLDER):
        item_path = os.path.join(RESULT_FOLDER, item)
        if os.path.isdir(item_path):
            result_subdirs.append({
                "name": item,
                "contents": list_directory(item_path)
            })
            if len(result_subdirs) >= 5:
                break
    
    return jsonify({
        "upload_directory": UPLOAD_FOLDER,
        "upload_contents": upload_contents,
        "upload_subdirs": upload_subdirs,
        "result_directory": RESULT_FOLDER,
        "result_contents": result_contents,
        "result_subdirs": result_subdirs
    })


@app.route('/test-upload')
def test_upload_page():
    """提供一个简单的文件上传测试页面"""
    html = """
    <!DOCTYPE html>
    <html>
    <head>
        <title>文件上传测试</title>
        <meta charset="UTF-8">
        <style>
            body { font-family: Arial, sans-serif; margin: 20px; }
            .container { max-width: 800px; margin: 0 auto; }
            .form-group { margin-bottom: 15px; }
            label { display: block; margin-bottom: 5px; }
            .btn { padding: 8px 16px; background-color: #4CAF50; color: white; border: none; cursor: pointer; }
            .result { margin-top: 20px; padding: 10px; border: 1px solid #ddd; background-color: #f9f9f9; }
        </style>
    </head>
    <body>
        <div class="container">
            <h1>文件上传测试</h1>
            <form id="uploadForm" enctype="multipart/form-data">
                <div class="form-group">
                    <label for="file">选择文件:</label>
                    <input type="file" id="file" name="file" required>
                </div>
                <div class="form-group">
                    <label for="toFormat">目标格式:</label>
                    <select id="toFormat" name="to_format">
                        <option value="pdf">PDF</option>
                        <option value="docx">DOCX</option>
                        <option value="xlsx">XLSX</option>
                        <option value="searchable_pdf">可搜索PDF</option>
                    </select>
                </div>
                <div class="form-group">
                    <label for="quality">质量:</label>
                    <select id="quality" name="quality">
                        <option value="1">低</option>
                        <option value="2" selected>中</option>
                        <option value="3">高</option>
                    </select>
                </div>
                <button type="submit" class="btn">上传并转换</button>
            </form>
            <div id="result" class="result" style="display:none;"></div>
            <div id="debug" class="result" style="display:none;">
                <h3>调试信息</h3>
                <pre id="debugInfo"></pre>
                <div id="fileInfo"></div>
            </div>
        </div>

        <script>
            document.getElementById('uploadForm').addEventListener('submit', function(e) {
                e.preventDefault();
                
                // 显示文件名信息
                var fileInput = document.getElementById('file');
                var debugInfo = document.getElementById('debugInfo');
                var fileInfoDiv = document.getElementById('fileInfo');
                var file = fileInput.files[0];
                
                if (file) {
                    debugInfo.textContent = 'File selected: ' + file.name + '\\n';
                    debugInfo.textContent += 'File size: ' + file.size + ' bytes\\n';
                    debugInfo.textContent += 'File type: ' + file.type + '\\n';
                    
                    // 显示文件对象的完整信息
                    fileInfoDiv.innerHTML = '<h4>文件对象完整信息</h4>';
                    for (var prop in file) {
                        if (typeof file[prop] !== 'function') {
                            fileInfoDiv.innerHTML += prop + ': ' + file[prop] + '<br>';
                        }
                    }
                    
                    document.getElementById('debug').style.display = 'block';
                }
                
                var formData = new FormData(this);
                
                // 添加原始文件名参数
                if (file) {
                    formData.append('original_filename', file.name);
                    debugInfo.textContent += '\\nAdded original_filename: ' + file.name;
                }
                
                var resultDiv = document.getElementById('result');
                
                resultDiv.innerHTML = '<p>上传中，请稍候...</p>';
                resultDiv.style.display = 'block';
                
                fetch('/api/convert', {
                    method: 'POST',
                    body: formData
                })
                .then(response => response.json())
                .then(data => {
                    if (data.success) {
                        // 显示转换成功信息
                        resultDiv.innerHTML = '<h3>转换成功!</h3>' +
                            '<p>原始文件名: ' + data.original_name + '</p>' +
                            '<p>文件大小: ' + data.file_size + ' 字节</p>' +
                            '<p><a href="' + data.result_url + '" target="_blank">下载转换后的文件</a></p>';
                            
                        // 添加更多调试信息
                        debugInfo.textContent += '\\nServer Response:\\n' + JSON.stringify(data, null, 2);
                    } else {
                        resultDiv.innerHTML = '<h3>转换失败</h3><p>' + data.error + '</p>';
                    }
                    document.getElementById('debug').style.display = 'block';
                })
                .catch(error => {
                    resultDiv.innerHTML = '<h3>错误</h3><p>' + error + '</p>';
                    console.error('Error:', error);
                });
            });
        </script>
    </body>
    </html>
    """
    return html


if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(description='文件转换API服务')
    parser.add_argument('--host', type=str, default='0.0.0.0', help='服务器主机地址')
    parser.add_argument('--port', type=int, default=5000, help='服务器端口')
    parser.add_argument('--debug', action='store_true', help='开启调试模式')

    args = parser.parse_args()

    logger.info(f"服务启动在 http://{args.host}:{args.port}")
    app.run(host=args.host, port=args.port, debug=args.debug)