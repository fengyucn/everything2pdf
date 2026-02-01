"""
Everything to PDF Converter - 核心转换逻辑
支持图片、Office文档、PDF的转换和合并
"""

import os
import io
import subprocess
import tempfile
import shutil
from pathlib import Path
from typing import List, Tuple, Optional

import fitz  # PyMuPDF
from PIL import Image
import img2pdf

# 尝试导入Office文档处理库（用于后备方案）
try:
    from docx import Document
    from docx.shared import Inches
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    from pptx import Presentation
    from pptx.util import Inches as PptxInches
    HAS_PPTX = True
except ImportError:
    HAS_PPTX = False

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import inch
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    HAS_REPORTLAB = True
except ImportError:
    HAS_REPORTLAB = False

# 支持的文件格式
IMAGE_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.tif', '.webp'}
OFFICE_EXTENSIONS = {'.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx'}
PDF_EXTENSION = '.pdf'

# 转换DPI设置
DEFAULT_DPI = 150


def detect_libreoffice() -> Optional[str]:
    """检测系统中的LibreOffice安装路径"""
    # Linux常见路径
    linux_paths = [
        '/usr/bin/libreoffice',
        '/usr/bin/soffice',
        '/usr/local/bin/libreoffice',
        '/usr/local/bin/soffice',
        '/snap/bin/libreoffice',
    ]
    
    # Windows常见路径
    windows_paths = [
        r'C:\Program Files\LibreOffice\program\soffice.exe',
        r'C:\Program Files (x86)\LibreOffice\program\soffice.exe',
    ]
    
    # 根据操作系统选择路径列表
    if os.name == 'nt':
        paths = windows_paths
    else:
        paths = linux_paths
    
    # 检查已知路径
    for path in paths:
        if os.path.isfile(path):
            return path
    
    # 尝试使用which/where命令查找
    try:
        if os.name == 'nt':
            result = subprocess.run(['where', 'soffice'], capture_output=True, text=True)
        else:
            result = subprocess.run(['which', 'libreoffice'], capture_output=True, text=True)
        
        if result.returncode == 0 and result.stdout.strip():
            return result.stdout.strip().split('\n')[0]
    except Exception:
        pass
    
    return None


def get_file_type(filepath: str) -> str:
    """获取文件类型"""
    ext = Path(filepath).suffix.lower()
    if ext in IMAGE_EXTENSIONS:
        return 'image'
    elif ext in OFFICE_EXTENSIONS:
        return 'office'
    elif ext == PDF_EXTENSION:
        return 'pdf'
    else:
        return 'unknown'


def convert_image_to_pdf_bytes(image_path: str) -> bytes:
    """将单个图片转换为PDF字节"""
    with open(image_path, 'rb') as f:
        img_bytes = f.read()
    
    # 检查图片格式，某些格式需要先转换
    try:
        img = Image.open(io.BytesIO(img_bytes))
        if img.mode in ('RGBA', 'P'):
            # 转换为RGB模式
            rgb_img = Image.new('RGB', img.size, (255, 255, 255))
            if img.mode == 'RGBA':
                rgb_img.paste(img, mask=img.split()[3])
            else:
                rgb_img.paste(img)
            
            buf = io.BytesIO()
            rgb_img.save(buf, format='JPEG', quality=95)
            img_bytes = buf.getvalue()
        elif img.format == 'GIF':
            # GIF转JPEG
            rgb_img = img.convert('RGB')
            buf = io.BytesIO()
            rgb_img.save(buf, format='JPEG', quality=95)
            img_bytes = buf.getvalue()
    except Exception:
        pass
    
    # 使用img2pdf转换
    pdf_bytes = img2pdf.convert(img_bytes)
    return pdf_bytes


def pdf_to_images(pdf_path: str, dpi: int = DEFAULT_DPI) -> List[bytes]:
    """将PDF转换为图片列表（PNG字节）"""
    images = []
    doc = fitz.open(pdf_path)
    
    zoom = dpi / 72  # 72是PDF默认DPI
    matrix = fitz.Matrix(zoom, zoom)
    
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        pix = page.get_pixmap(matrix=matrix)
        img_bytes = pix.tobytes('png')
        images.append(img_bytes)
    
    doc.close()
    return images


def convert_office_with_libreoffice(office_path: str, libreoffice_path: str, output_dir: str) -> Optional[str]:
    """使用LibreOffice转换Office文档为PDF"""
    try:
        cmd = [
            libreoffice_path,
            '--headless',
            '--convert-to', 'pdf',
            '--outdir', output_dir,
            office_path
        ]
        
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
        
        if result.returncode == 0:
            # 找到生成的PDF文件
            base_name = Path(office_path).stem
            pdf_path = os.path.join(output_dir, f'{base_name}.pdf')
            if os.path.exists(pdf_path):
                return pdf_path
    except Exception as e:
        print(f"LibreOffice转换失败: {e}")
    
    return None


def convert_docx_with_python(docx_path: str, output_dir: str) -> Optional[str]:
    """使用纯Python转换Word文档为PDF（后备方案）"""
    if not HAS_DOCX or not HAS_REPORTLAB:
        return None
    
    try:
        doc = Document(docx_path)
        output_path = os.path.join(output_dir, f'{Path(docx_path).stem}.pdf')
        
        # 尝试注册中文字体
        try:
            # Linux常见中文字体路径
            font_paths = [
                '/usr/share/fonts/truetype/wqy/wqy-microhei.ttc',
                '/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf',
                '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc',
                'C:/Windows/Fonts/simhei.ttf',
                'C:/Windows/Fonts/msyh.ttc',
            ]
            for font_path in font_paths:
                if os.path.exists(font_path):
                    pdfmetrics.registerFont(TTFont('ChineseFont', font_path))
                    break
        except Exception:
            pass
        
        pdf_doc = SimpleDocTemplate(output_path, pagesize=A4)
        styles = getSampleStyleSheet()
        
        # 创建支持中文的样式
        try:
            chinese_style = ParagraphStyle(
                'ChineseStyle',
                parent=styles['Normal'],
                fontName='ChineseFont',
                fontSize=12,
                leading=16,
            )
        except Exception:
            chinese_style = styles['Normal']
        
        story = []
        
        for para in doc.paragraphs:
            if para.text.strip():
                try:
                    p = Paragraph(para.text, chinese_style)
                    story.append(p)
                    story.append(Spacer(1, 6))
                except Exception:
                    # 如果段落渲染失败，跳过
                    pass
        
        if story:
            pdf_doc.build(story)
            return output_path
        
    except Exception as e:
        print(f"Python Word转换失败: {e}")
    
    return None


def convert_xlsx_with_python(xlsx_path: str, output_dir: str) -> Optional[str]:
    """使用纯Python转换Excel文档为PDF（后备方案）"""
    if not HAS_OPENPYXL or not HAS_REPORTLAB:
        return None
    
    try:
        wb = openpyxl.load_workbook(xlsx_path, data_only=True)
        output_path = os.path.join(output_dir, f'{Path(xlsx_path).stem}.pdf')
        
        pdf_doc = SimpleDocTemplate(output_path, pagesize=A4)
        story = []
        styles = getSampleStyleSheet()
        
        for sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]
            
            # 添加工作表标题
            story.append(Paragraph(f"Sheet: {sheet_name}", styles['Heading2']))
            story.append(Spacer(1, 12))
            
            # 获取数据范围
            data = []
            for row in sheet.iter_rows(values_only=True):
                row_data = [str(cell) if cell is not None else '' for cell in row]
                if any(row_data):  # 跳过空行
                    data.append(row_data)
            
            if data:
                # 创建表格
                table = Table(data)
                table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTSIZE', (0, 0), (-1, -1), 8),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ]))
                story.append(table)
            
            story.append(Spacer(1, 24))
        
        if story:
            pdf_doc.build(story)
            return output_path
        
    except Exception as e:
        print(f"Python Excel转换失败: {e}")
    
    return None


def convert_pptx_with_python(pptx_path: str, output_dir: str) -> Optional[str]:
    """使用纯Python转换PPT文档为PDF（后备方案）"""
    if not HAS_PPTX or not HAS_REPORTLAB:
        return None
    
    try:
        prs = Presentation(pptx_path)
        output_path = os.path.join(output_dir, f'{Path(pptx_path).stem}.pdf')
        
        pdf_doc = SimpleDocTemplate(output_path, pagesize=A4)
        story = []
        styles = getSampleStyleSheet()
        
        for slide_num, slide in enumerate(prs.slides, 1):
            story.append(Paragraph(f"Slide {slide_num}", styles['Heading2']))
            story.append(Spacer(1, 12))
            
            for shape in slide.shapes:
                if hasattr(shape, "text") and shape.text.strip():
                    try:
                        p = Paragraph(shape.text, styles['Normal'])
                        story.append(p)
                        story.append(Spacer(1, 6))
                    except Exception:
                        pass
            
            story.append(Spacer(1, 24))
        
        if story:
            pdf_doc.build(story)
            return output_path
        
    except Exception as e:
        print(f"Python PPT转换失败: {e}")
    
    return None


def convert_office_to_pdf(office_path: str, libreoffice_path: Optional[str], output_dir: str) -> Tuple[Optional[str], str]:
    """
    转换Office文档为PDF
    返回: (PDF路径, 使用的方法)
    """
    ext = Path(office_path).suffix.lower()
    
    # 优先使用LibreOffice
    if libreoffice_path:
        result = convert_office_with_libreoffice(office_path, libreoffice_path, output_dir)
        if result:
            return result, 'libreoffice'
    
    # 后备Python方案
    if ext in {'.doc', '.docx'}:
        result = convert_docx_with_python(office_path, output_dir)
        if result:
            return result, 'python'
    elif ext in {'.xls', '.xlsx'}:
        result = convert_xlsx_with_python(office_path, output_dir)
        if result:
            return result, 'python'
    elif ext in {'.ppt', '.pptx'}:
        result = convert_pptx_with_python(office_path, output_dir)
        if result:
            return result, 'python'
    
    return None, 'failed'


def images_to_pdf(image_bytes_list: List[bytes], output_path: str):
    """将图片字节列表合并为单个PDF"""
    doc = fitz.open()
    
    for img_bytes in image_bytes_list:
        # 从图片字节创建PDF页面
        img = fitz.open(stream=img_bytes, filetype='png')
        rect = img[0].rect
        
        # 创建新页面，大小与图片相同
        page = doc.new_page(width=rect.width, height=rect.height)
        page.insert_image(rect, stream=img_bytes)
        img.close()
    
    doc.save(output_path)
    doc.close()


def convert_files_to_pdf(
    file_paths: List[str],
    output_path: str,
    libreoffice_path: Optional[str] = None,
    dpi: int = DEFAULT_DPI,
    progress_callback=None
) -> Tuple[bool, str]:
    """
    将多个文件转换并合并为单个PDF
    
    Args:
        file_paths: 文件路径列表（按顺序）
        output_path: 输出PDF路径
        libreoffice_path: LibreOffice路径（可选）
        dpi: 输出图片DPI
        progress_callback: 进度回调函数 callback(current, total, message)
    
    Returns:
        (成功标志, 消息)
    """
    if not file_paths:
        return False, "没有选择文件"
    
    all_images = []
    temp_dir = tempfile.mkdtemp()
    
    try:
        total = len(file_paths)
        
        for idx, file_path in enumerate(file_paths):
            if progress_callback:
                progress_callback(idx, total, f"处理: {Path(file_path).name}")
            
            file_type = get_file_type(file_path)
            
            if file_type == 'image':
                # 图片：直接转为PDF字节，再转为图片
                try:
                    pdf_bytes = convert_image_to_pdf_bytes(file_path)
                    # 用PyMuPDF读取PDF字节并转为图片
                    temp_pdf = fitz.open(stream=pdf_bytes, filetype='pdf')
                    zoom = dpi / 72
                    matrix = fitz.Matrix(zoom, zoom)
                    for page in temp_pdf:
                        pix = page.get_pixmap(matrix=matrix)
                        all_images.append(pix.tobytes('png'))
                    temp_pdf.close()
                except Exception as e:
                    # 如果img2pdf失败，尝试直接用PIL
                    img = Image.open(file_path)
                    if img.mode != 'RGB':
                        img = img.convert('RGB')
                    buf = io.BytesIO()
                    img.save(buf, format='PNG')
                    all_images.append(buf.getvalue())
            
            elif file_type == 'office':
                # Office文档：转为PDF再转图片
                pdf_path, method = convert_office_to_pdf(file_path, libreoffice_path, temp_dir)
                if pdf_path:
                    images = pdf_to_images(pdf_path, dpi)
                    all_images.extend(images)
                else:
                    return False, f"无法转换文件: {Path(file_path).name}"
            
            elif file_type == 'pdf':
                # PDF：转为图片
                images = pdf_to_images(file_path, dpi)
                all_images.extend(images)
            
            else:
                return False, f"不支持的文件格式: {Path(file_path).name}"
        
        if progress_callback:
            progress_callback(total, total, "正在生成PDF...")
        
        # 合并所有图片为PDF
        if all_images:
            images_to_pdf(all_images, output_path)
            return True, f"成功转换 {len(file_paths)} 个文件"
        else:
            return False, "没有可转换的内容"
    
    except Exception as e:
        return False, f"转换失败: {str(e)}"
    
    finally:
        # 清理临时目录
        try:
            shutil.rmtree(temp_dir)
        except Exception:
            pass


# 全局状态
_libreoffice_path = None
_libreoffice_checked = False


def get_libreoffice_status() -> Tuple[bool, Optional[str]]:
    """获取LibreOffice状态"""
    global _libreoffice_path, _libreoffice_checked
    
    if not _libreoffice_checked:
        _libreoffice_path = detect_libreoffice()
        _libreoffice_checked = True
    
    return _libreoffice_path is not None, _libreoffice_path


def get_supported_extensions() -> List[str]:
    """获取支持的文件扩展名列表"""
    extensions = list(IMAGE_EXTENSIONS) + list(OFFICE_EXTENSIONS) + [PDF_EXTENSION]
    return sorted(extensions)
