"""
Everything to PDF Converter - Flask Web应用
"""

import os
import sys
import uuid
import tempfile
import webbrowser
import threading
from pathlib import Path

from flask import Flask, render_template, request, jsonify, send_file

from converter import (
    get_libreoffice_status,
    get_supported_extensions,
    convert_files_to_pdf,
    get_file_type
)

# 获取资源路径（支持PyInstaller打包）
def get_resource_path(relative_path):
    """获取资源文件的绝对路径"""
    if hasattr(sys, '_MEIPASS'):
        # PyInstaller打包后的路径
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), relative_path)


# 创建Flask应用
app = Flask(
    __name__,
    template_folder=get_resource_path('templates'),
    static_folder=get_resource_path('static')
)

# 配置
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024  # 最大500MB
app.config['UPLOAD_FOLDER'] = tempfile.mkdtemp()

# 存储上传的文件信息
uploaded_files = {}


@app.route('/')
def index():
    """主页"""
    has_libreoffice, lo_path = get_libreoffice_status()
    extensions = get_supported_extensions()
    
    return render_template(
        'index.html',
        has_libreoffice=has_libreoffice,
        libreoffice_path=lo_path,
        supported_extensions=extensions
    )


@app.route('/api/status')
def api_status():
    """获取系统状态"""
    has_libreoffice, lo_path = get_libreoffice_status()
    
    return jsonify({
        'has_libreoffice': has_libreoffice,
        'libreoffice_path': lo_path,
        'supported_extensions': get_supported_extensions()
    })


@app.route('/api/upload', methods=['POST'])
def api_upload():
    """上传文件"""
    if 'files' not in request.files:
        return jsonify({'error': '没有文件上传'}), 400
    
    files = request.files.getlist('files')
    results = []
    
    for file in files:
        if file.filename:
            # 生成唯一ID
            file_id = str(uuid.uuid4())
            
            # 保存文件
            original_name = file.filename
            safe_name = f"{file_id}_{original_name}"
            save_path = os.path.join(app.config['UPLOAD_FOLDER'], safe_name)
            file.save(save_path)
            
            # 记录文件信息
            file_type = get_file_type(save_path)
            file_info = {
                'id': file_id,
                'name': original_name,
                'path': save_path,
                'type': file_type,
                'size': os.path.getsize(save_path)
            }
            uploaded_files[file_id] = file_info
            
            results.append({
                'id': file_id,
                'name': original_name,
                'type': file_type,
                'size': file_info['size']
            })
    
    return jsonify({'files': results})


@app.route('/api/remove/<file_id>', methods=['DELETE'])
def api_remove(file_id):
    """删除已上传的文件"""
    if file_id in uploaded_files:
        file_info = uploaded_files.pop(file_id)
        try:
            os.remove(file_info['path'])
        except Exception:
            pass
        return jsonify({'success': True})
    
    return jsonify({'error': '文件不存在'}), 404


@app.route('/api/clear', methods=['POST'])
def api_clear():
    """清空所有上传的文件"""
    for file_id, file_info in list(uploaded_files.items()):
        try:
            os.remove(file_info['path'])
        except Exception:
            pass
    
    uploaded_files.clear()
    return jsonify({'success': True})


@app.route('/api/convert', methods=['POST'])
def api_convert():
    """转换文件为PDF"""
    data = request.get_json()
    
    if not data or 'file_ids' not in data:
        return jsonify({'error': '没有指定文件'}), 400
    
    file_ids = data['file_ids']
    
    # 获取文件路径列表（按顺序）
    file_paths = []
    for file_id in file_ids:
        if file_id in uploaded_files:
            file_paths.append(uploaded_files[file_id]['path'])
        else:
            return jsonify({'error': f'文件不存在: {file_id}'}), 404
    
    if not file_paths:
        return jsonify({'error': '没有有效的文件'}), 400
    
    # 生成输出文件
    output_filename = f"converted_{uuid.uuid4().hex[:8]}.pdf"
    output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
    
    # 获取LibreOffice路径
    has_lo, lo_path = get_libreoffice_status()
    
    # 执行转换
    success, message = convert_files_to_pdf(
        file_paths,
        output_path,
        libreoffice_path=lo_path
    )
    
    if success:
        return jsonify({
            'success': True,
            'message': message,
            'download_url': f'/api/download/{output_filename}'
        })
    else:
        return jsonify({'error': message}), 500


@app.route('/api/download/<filename>')
def api_download(filename):
    """下载转换后的PDF"""
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    
    if os.path.exists(file_path):
        return send_file(
            file_path,
            as_attachment=True,
            download_name='converted.pdf',
            mimetype='application/pdf'
        )
    
    return jsonify({'error': '文件不存在'}), 404


def open_browser(port):
    """在默认浏览器中打开应用"""
    try:
        webbrowser.open(f'http://127.0.0.1:{port}')
    except Exception as e:
        print(f"无法自动打开浏览器: {e}")
        print(f"请手动在浏览器中访问: http://127.0.0.1:{port}")


def main():
    """主函数"""
    port = 5000
    
    # 检查端口是否可用
    import socket
    while True:
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.bind(('127.0.0.1', port))
                break
        except OSError:
            port += 1
            if port > 5100:
                print("无法找到可用端口")
                sys.exit(1)
    
    print(f"Everything to PDF Converter")
    print(f"正在启动服务器: http://127.0.0.1:{port}")
    
    # 检查LibreOffice状态
    has_lo, lo_path = get_libreoffice_status()
    if has_lo:
        print(f"LibreOffice 已检测到: {lo_path}")
    else:
        print("未检测到 LibreOffice，Office文档将使用基础模式转换")
    
    # 延迟打开浏览器
    threading.Timer(1.5, open_browser, args=[port]).start()
    
    # 启动Flask服务器
    app.run(host='127.0.0.1', port=port, debug=False, threaded=True)


if __name__ == '__main__':
    main()
