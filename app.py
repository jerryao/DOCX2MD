import os
import tempfile
import uuid
import zipfile
import logging
import shutil
import re
from io import BytesIO
from pathlib import Path
from datetime import datetime
from flask import Response

from flask import Flask, render_template, request, send_file, redirect, url_for, flash, jsonify
from flask_wtf.csrf import CSRFProtect
from werkzeug.utils import secure_filename

from markitdown import MarkItDown

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler('app.log', encoding='utf-8')
    ]
)
logger = logging.getLogger(__name__)

app = Flask(__name__)
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'default-dev-key-replace-in-production')
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB 上传限制
app.config['UPLOAD_FOLDER'] = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
app.config['DOWNLOAD_FOLDER'] = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'downloads')

csrf = CSRFProtect(app)

# 确保上传和下载目录存在
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['DOWNLOAD_FOLDER'], exist_ok=True)

# 清理旧下载文件 (保留最近24小时的文件)
def cleanup_old_files():
    try:
        now = datetime.now()
        download_dir = app.config['DOWNLOAD_FOLDER']
        count = 0
        for file in os.listdir(download_dir):
            file_path = os.path.join(download_dir, file)
            if os.path.isfile(file_path):
                file_time = datetime.fromtimestamp(os.path.getctime(file_path))
                if (now - file_time).total_seconds() > 86400:  # 24小时
                    os.remove(file_path)
                    count += 1
        if count > 0:
            logger.info(f"已清理 {count} 个过期下载文件")
    except Exception as e:
        logger.error(f"清理旧文件时出错: {str(e)}")

# 允许的文件扩展名
ALLOWED_EXTENSIONS = {'docx'}

def allowed_file(filename):
    """检查文件扩展名是否允许"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def safe_filename(filename):
    """自定义的安全文件名处理函数，保留中文字符"""
    # 保留中文字符、字母、数字和一些安全字符，并移除不安全字符
    safe_chars = re.sub(r'[^\w\u4e00-\u9fa5\-\.]', '_', filename)
    return safe_chars

@app.route('/', methods=['GET', 'POST'])
def index():
    """主页和文件上传处理"""
    # 在页面加载时清理旧文件
    if request.method == 'GET':
        cleanup_old_files()
        
    if request.method == 'POST':
        # 检查是否有文件提交
        if 'files[]' not in request.files:
            flash('未选择任何文件', 'error')
            return redirect(request.url)
        
        files = request.files.getlist('files[]')
        
        # 检查是否选择了至少一个文件
        if not files or files[0].filename == '':
            flash('未选择任何文件', 'error')
            return redirect(request.url)
        
        # 过滤出有效的文件
        valid_files = [file for file in files if file and allowed_file(file.filename)]
        
        if not valid_files:
            flash('没有有效的DOCX文件', 'error')
            return redirect(request.url)
        
        # 获取导出选项
        export_type = request.form.get('export_type', 'zip')
        custom_dir = None
        
        if export_type == 'directory' and request.form.get('custom_directory'):
            custom_dir = request.form.get('custom_directory')
            # 验证目录路径是否有效
            if not os.path.exists(custom_dir) or not os.path.isdir(custom_dir):
                try:
                    os.makedirs(custom_dir, exist_ok=True)
                except Exception as e:
                    flash(f'创建指定的导出目录失败: {str(e)}', 'error')
                    return redirect(request.url)
        
        logger.info(f"开始转换 {len(valid_files)} 个文件，导出类型: {export_type}")
        
        # 使用临时目录处理文件
        with tempfile.TemporaryDirectory() as temp_dir:
            # 创建 MarkItDown 实例
            converter = MarkItDown()
            converted_files = []
            error_files = []
            
            # 确保文件名唯一性
            used_filenames = set()
            
            # 保存上传的文件并转换
            for file in valid_files:
                # 使用自定义函数，保留中文字符
                original_filename = safe_filename(file.filename)
                logger.info(f"处理文件: 原始文件名 = {file.filename}, 处理后文件名 = {original_filename}")
                
                input_file_path = os.path.join(temp_dir, original_filename)
                file.save(input_file_path)
                
                # 确定输出文件名 (保持与原文件名一致，只更改扩展名)
                output_filename = original_filename.rsplit('.', 1)[0] + '.md'
                
                # 确保文件名唯一
                if output_filename in used_filenames:
                    # 添加时间戳以确保唯一性
                    base_name = output_filename.rsplit('.', 1)[0]
                    timestamp = datetime.now().strftime("%H%M%S")
                    output_filename = f"{base_name}_{timestamp}.md"
                
                used_filenames.add(output_filename)
                output_file_path = os.path.join(temp_dir, output_filename)
                
                try:
                    # 使用 MarkItDown 进行转换
                    result = converter.convert(input_file_path)
                    
                    # 保存结果
                    with open(output_file_path, 'w', encoding='utf-8') as f:
                        f.write(result.text_content)
                    
                    # 如果指定了自定义目录，则复制到该目录
                    if custom_dir:
                        dest_path = os.path.join(custom_dir, output_filename)
                        shutil.copy2(output_file_path, dest_path)
                    
                    converted_files.append((original_filename, output_filename))
                    logger.info(f"成功转换文件: {original_filename} -> {output_filename}")
                except Exception as e:
                    error_message = str(e)
                    error_files.append((original_filename, error_message))
                    logger.error(f"转换文件失败: {original_filename} - {error_message}")
                    flash(f'转换文件 {original_filename} 时出错: {error_message}', 'error')
            
            if not converted_files:
                flash('没有文件被成功转换', 'error')
                return redirect(request.url)
            
            # 如果有错误文件但是也有成功文件，显示部分成功信息
            if error_files and converted_files:
                flash(f'成功转换了 {len(converted_files)} 个文件，{len(error_files)} 个文件失败', 'warning')
            elif converted_files:
                flash(f'成功转换了 {len(converted_files)} 个文件', 'success')
            
            # 如果用户选择了自定义目录，直接返回成功消息
            if export_type == 'directory' and custom_dir:
                flash(f'已将转换后的文件保存到: {custom_dir}', 'success')
                return redirect(request.url)
            
            # 否则创建ZIP文件下载
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            zip_filename = f'markdown_files_{timestamp}.zip'
            zip_path = os.path.join(app.config['DOWNLOAD_FOLDER'], zip_filename)
            
            with zipfile.ZipFile(zip_path, 'w') as zipf:
                for _, output_filename in converted_files:
                    file_path = os.path.join(temp_dir, output_filename)
                    zipf.write(file_path, arcname=output_filename)
            
            logger.info(f"已创建 ZIP 文件: {zip_filename}")
            
            # 返回ZIP文件
            return send_file(zip_path, 
                            mimetype='application/zip',
                            as_attachment=True, 
                            download_name=zip_filename)
    
    # GET 请求 - 显示上传页面
    return render_template('index.html', now=datetime.now())

@app.route('/api/convert', methods=['POST'])
def api_convert():
    """API 端点，用于处理异步文件上传和转换"""
    if 'files[]' not in request.files:
        return jsonify({"error": "未选择任何文件"}), 400
    
    files = request.files.getlist('files[]')
    
    if not files or files[0].filename == '':
        return jsonify({"error": "未选择任何文件"}), 400
    
    valid_files = [file for file in files if file and allowed_file(file.filename)]
    
    if not valid_files:
        return jsonify({"error": "没有有效的DOCX文件"}), 400
    
    # 获取导出选项
    export_type = request.form.get('export_type', 'zip')
    custom_dir = None
    
    if export_type == 'directory' and request.form.get('custom_directory'):
        custom_dir = request.form.get('custom_directory')
        # 验证目录路径是否有效
        if not os.path.exists(custom_dir) or not os.path.isdir(custom_dir):
            try:
                os.makedirs(custom_dir, exist_ok=True)
            except Exception as e:
                return jsonify({"error": f"创建指定的导出目录失败: {str(e)}"}), 400
    
    logger.info(f"API 开始转换 {len(valid_files)} 个文件，导出类型: {export_type}")
    
    # 使用临时目录处理文件
    with tempfile.TemporaryDirectory() as temp_dir:
        converter = MarkItDown()
        converted_files = []
        error_files = []
        
        # 确保文件名唯一性
        used_filenames = set()
        
        for file in valid_files:
            # 使用自定义函数，保留中文字符
            original_filename = safe_filename(file.filename)
            logger.info(f"API 处理文件: 原始文件名 = {file.filename}, 处理后文件名 = {original_filename}")
            
            input_file_path = os.path.join(temp_dir, original_filename)
            file.save(input_file_path)
            
            # 确定输出文件名 (保持与原文件名一致，只更改扩展名)
            output_filename = original_filename.rsplit('.', 1)[0] + '.md'
            
            # 确保文件名唯一
            if output_filename in used_filenames:
                # 添加时间戳以确保唯一性
                base_name = output_filename.rsplit('.', 1)[0]
                timestamp = datetime.now().strftime("%H%M%S")
                output_filename = f"{base_name}_{timestamp}.md"
            
            used_filenames.add(output_filename)
            output_file_path = os.path.join(temp_dir, output_filename)
            
            try:
                result = converter.convert(input_file_path)
                
                with open(output_file_path, 'w', encoding='utf-8') as f:
                    f.write(result.text_content)
                
                # 如果指定了自定义目录，则复制到该目录
                if custom_dir:
                    dest_path = os.path.join(custom_dir, output_filename)
                    shutil.copy2(output_file_path, dest_path)
                
                converted_files.append({"original": original_filename, "converted": output_filename})
                logger.info(f"API 成功转换文件: {original_filename} -> {output_filename}")
            except Exception as e:
                error_message = str(e)
                error_files.append({"file": original_filename, "error": error_message})
                logger.error(f"API 转换文件失败: {original_filename} - {error_message}")
        
        if not converted_files:
            return jsonify({"error": "没有文件被成功转换", "failed_files": error_files}), 400
        
        # 如果用户选择了自定义目录，返回保存位置信息
        if export_type == 'directory' and custom_dir:
            return jsonify({
                "success": True,
                "message": f"成功转换了 {len(converted_files)} 个文件" + (f"，{len(error_files)} 个文件失败" if error_files else ""),
                "converted_count": len(converted_files),
                "error_count": len(error_files),
                "export_type": "directory",
                "directory_path": custom_dir,
                "failed_files": error_files
            })
        
        # 否则创建ZIP文件
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        zip_filename = f'markdown_files_{timestamp}.zip'
        zip_path = os.path.join(app.config['DOWNLOAD_FOLDER'], zip_filename)
        
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for file_info in converted_files:
                file_path = os.path.join(temp_dir, file_info["converted"])
                zipf.write(file_path, arcname=file_info["converted"])
        
        logger.info(f"API 已创建 ZIP 文件: {zip_filename}")
        
        return jsonify({
            "success": True,
            "message": f"成功转换了 {len(converted_files)} 个文件" + (f"，{len(error_files)} 个文件失败" if error_files else ""),
            "converted_count": len(converted_files),
            "error_count": len(error_files),
            "export_type": "zip",
            "download_url": url_for('download_file', filename=zip_filename),
            "failed_files": error_files
        })

@app.route('/download/<filename>', methods=['GET'])
def download_file(filename):
    """下载已转换的ZIP文件"""
    zip_path = os.path.join(app.config['DOWNLOAD_FOLDER'], filename)
    if not os.path.exists(zip_path):
        logger.warning(f"请求下载不存在的文件: {filename}")
        flash('文件不存在或已过期', 'error')
        return redirect(url_for('index'))
    
    logger.info(f"下载文件: {filename}")
    return send_file(zip_path, 
                    mimetype='application/zip',
                    as_attachment=True, 
                    download_name=filename)

@app.errorhandler(413)
def request_entity_too_large(error):
    """处理文件过大的错误"""
    logger.warning("上传文件过大错误")
    flash('上传的文件过大。最大允许大小为50MB。', 'error')
    return redirect(url_for('index')), 413

@app.errorhandler(Exception)
def handle_exception(e):
    """全局异常处理"""
    logger.error(f"应用程序错误: {str(e)}", exc_info=True)
    flash('服务器内部错误，请稍后重试或联系管理员。', 'error')
    return redirect(url_for('index')), 500

@app.template_filter('filesizeformat')
def filesizeformat_filter(value):
    """格式化文件大小"""
    if value is None:
        return '0 bytes'
    
    value = float(value)
    for unit in ['bytes', 'KB', 'MB', 'GB', 'TB']:
        if value < 1024.0:
            return f"{value:.1f} {unit}"
        value /= 1024.0

# 添加处理favicon.ico的路由
@app.route('/favicon.ico')
def favicon():
    """提供网站图标"""
    try:
        return send_file(os.path.join(app.static_folder, 'favicon.ico'),
                      mimetype='image/vnd.microsoft.icon')
    except Exception as e:
        logger.warning(f"未能加载favicon: {str(e)}")
        # 返回一个1x1像素的透明GIF
        transparent_gif = b'GIF89a\x01\x00\x01\x00\x80\x00\x00\x00\x00\x00\x00\x00\x00!\xf9\x04\x01\x00\x00\x00\x00,\x00\x00\x00\x00\x01\x00\x01\x00\x00\x02\x02D\x01\x00;'
        return Response(transparent_gif, mimetype='image/gif')

if __name__ == '__main__':
    logger.info("应用程序启动")
    app.run(debug=True, host='0.0.0.0', port=5000) 