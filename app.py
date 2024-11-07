import sys
import os

# 添加项目根目录到 Python 路径
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from datetime import datetime, timedelta
import glob
from flask import Flask, request, jsonify, send_file, render_template, make_response
from werkzeug.utils import secure_filename
import json
from config import (
    UPLOAD_FOLDER, 
    RESULT_FOLDER, 
    DOWNLOAD_FOLDER,
    ALLOWED_EXTENSIONS
)

# 现在应该能正确导入这些模块了
from functions.vlookup import VlookupProcessor
from functions.concatenate import ConcatenateProcessor
from functions.pivot import PivotProcessor
from functions.format import ExcelProcessor
from functions.import_processor import ImportProcessor

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['RESULT_FOLDER'] = RESULT_FOLDER
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER

def init_app():
    """初始化应用"""
    try:
        # 确保所有必要的目录存在
        os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
        os.makedirs(app.config['RESULT_FOLDER'], exist_ok=True)
        os.makedirs(app.config['DOWNLOAD_FOLDER'], exist_ok=True)
        print(f"目录已创建/确认：")
        print(f"UPLOAD_FOLDER: {app.config['UPLOAD_FOLDER']}")
        print(f"RESULT_FOLDER: {app.config['RESULT_FOLDER']}")
        print(f"DOWNLOAD_FOLDER: {app.config['DOWNLOAD_FOLDER']}")
    except Exception as e:
        print(f"初始化目录时出错: {str(e)}")

_is_initialized = False

@app.before_request
def initialize():
    global _is_initialized
    if not _is_initialized:
        init_app()
        _is_initialized = True

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/get-sheets', methods=['POST'])
def get_sheets():
    try:
        if 'file' not in request.files:
            return jsonify({'success': False, 'error': '未找到文件'})
        
        file = request.files['file']
        if not file.filename:
            return jsonify({'success': False, 'error': '文件名不能为空'})
            
        if not allowed_file(file.filename):
            return jsonify({'success': False, 'error': '不支持的文件格式'})
        
        # 保存文件
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)
        
        # 获取工作表名称
        processor = VlookupProcessor(file_path, file_path, '')
        sheets = processor.get_sheets(file_path)
        
        # 清理临时文件
        os.remove(file_path)
        
        return jsonify({'success': True, 'sheets': sheets})
        
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@app.route('/api/get-columns', methods=['POST'])
def get_columns():
    try:
        if 'file' not in request.files:
            return jsonify({'success': False, 'error': '未找到文件'})
        
        file = request.files['file']
        sheet = request.form.get('sheet')
        
        if not file.filename:
            return jsonify({'success': False, 'error': '文件名不能为空'})
            
        if not allowed_file(file.filename):
            return jsonify({'success': False, 'error': '不支持的文件格式'})
        
        if not sheet:
            return jsonify({'success': False, 'error': '未指定工作表'})
        
        # 保存文件
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)
        
        # 根据请求来源选择不同的处理器
        if 'concatenateFile' in request.form:
            processor = ConcatenateProcessor(file_path, '')
        elif 'pivotFile' in request.form:
            processor = PivotProcessor(file_path, '')
        else:
            processor = VlookupProcessor(file_path, file_path, '')
            
        sheet_info = processor.get_sheet_info(file_path, sheet)
        
        # 清理临时文件
        os.remove(file_path)
        
        return jsonify({
            'success': True,
            'columns': sheet_info
        })
        
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@app.route('/api/vlookup', methods=['POST'])
def vlookup():
    try:
        # 检查文件
        if 'mainFile' not in request.files:
            return jsonify({'success': False, 'error': '请上传主数据表文件'})
        
        main_file = request.files['mainFile']
        if not main_file.filename:
            return jsonify({'success': False, 'error': '主数据表文件名不能为空'})

        # 获取查找表文件（可能与主表是同一个文件）
        lookup_file = request.files.get('lookupFile', main_file)
        
        # 获取所有参数
        main_sheet = request.form.get('mainSheet')
        main_match_type = request.form.get('mainMatchType')
        main_columns = json.loads(request.form.get('mainColumns', '[]'))
        
        lookup_sheet = request.form.get('lookupSheet')
        lookup_match_type = request.form.get('lookupMatchType')
        lookup_match_columns = json.loads(request.form.get('lookupMatchColumns', '[]'))
        return_type = request.form.get('returnType')
        return_columns = json.loads(request.form.get('returnColumns', '[]'))

        # 验证参数
        if not all([
            main_sheet, main_match_type, main_columns,
            lookup_sheet, lookup_match_type, 
            lookup_match_columns, return_type, return_columns
        ]):
            return jsonify({'success': False, 'error': '缺少必要的参数'})

        if not (allowed_file(main_file.filename) and allowed_file(lookup_file.filename)):
            return jsonify({'success': False, 'error': '不支持的文件格式'})

        # 生成唯一的结果文件名
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        result_filename = f'result_{timestamp}.xlsx'
        result_path = os.path.join(app.config['RESULT_FOLDER'], result_filename)

        # 保存文件
        main_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(main_file.filename))
        main_file.save(main_path)

        if lookup_file != main_file:
            # 如果是不同文件，则保存查找表文件
            lookup_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(lookup_file.filename))
            lookup_file.save(lookup_path)
        else:
            # 如果是同一个文件，使用相同的路径
            lookup_path = main_path

        # 处理VLOOKUP
        processor = VlookupProcessor(main_path, lookup_path, result_path)
        success, error_message = processor.process(
            main_sheet=main_sheet,
            main_match_type=main_match_type,
            main_columns=main_columns,
            lookup_sheet=lookup_sheet,
            lookup_match_type=lookup_match_type,
            lookup_match_columns=lookup_match_columns,
            return_type=return_type,
            return_columns=return_columns
        )
        
        # 清理临时文件
        processor.cleanup()
        
        if success:
            return jsonify({'success': True})
        else:
            return jsonify({'success': False, 'error': error_message})
            
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@app.route('/api/concatenate', methods=['POST'])
def concatenate():
    try:
        if 'file' not in request.files:
            return jsonify({'success': False, 'error': '请上传文件'})
        
        file = request.files['file']
        sheet = request.form.get('sheet')
        columns_data = json.loads(request.form.get('columnsData', '[]'))
        
        if not all([file.filename, sheet, columns_data]):
            return jsonify({'success': False, 'error': '缺少必要的参数'})
            
        if not allowed_file(file.filename):
            return jsonify({'success': False, 'error': '不支持的文件格式'})
        
        # 生成唯一的结果文件名
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        result_filename = f'result_{timestamp}.xlsx'
        
        # 保存文件
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file.filename))
        result_path = os.path.join(app.config['RESULT_FOLDER'], result_filename)
        
        file.save(file_path)
        
        # 理合并
        processor = ConcatenateProcessor(file_path, result_path)
        success, error_message = processor.process(sheet, columns_data)
        
        # 清理临时文件
        processor.cleanup()
        
        if success:
            return jsonify({'success': True})
        else:
            return jsonify({'success': False, 'error': error_message})
            
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@app.route('/api/pivot', methods=['POST'])
def pivot():
    try:
        if 'file' not in request.files:
            return jsonify({'success': False, 'error': '请上传文件'})
        
        file = request.files['file']
        sheet = request.form.get('sheet')
        config = json.loads(request.form.get('config', '{}'))
        
        if not all([file.filename, sheet, config]):
            return jsonify({'success': False, 'error': '缺少要的参数'})
            
        if not allowed_file(file.filename):
            return jsonify({'success': False, 'error': '不支持的文件格式'})
        
        # 验证配置
        if not config.get('rows'):
            return jsonify({'success': False, 'error': '请至少选择一个行标签'})
        if not config.get('values'):
            return jsonify({'success': False, 'error': '请至少选择一个值字段'})
        
        # 生成唯一的结果文件名
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        result_filename = f'result_{timestamp}.xlsx'
        
        # 保存文件
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file.filename))
        result_path = os.path.join(app.config['RESULT_FOLDER'], result_filename)
        
        file.save(file_path)
        
        # 处理透视表
        processor = PivotProcessor(file_path, result_path)
        success, error_message = processor.process(sheet, config)
        
        # 清理临时文件
        processor.cleanup()
        
        if success:
            return jsonify({'success': True})
        else:
            return jsonify({'success': False, 'error': error_message})
            
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@app.route('/api/download-result')
def download_result():
    try:
        # 获取最新的结果文件
        result_files = glob.glob(os.path.join(app.config['RESULT_FOLDER'], 'result_*.xlsx'))
        if not result_files:
            return jsonify({'success': False, 'error': '没有可下载的结果文件'})
        
        # 按文件名排序（因为包含时间戳），获取最新的文件
        latest_result = max(result_files)
        
        return send_file(
            latest_result,
            as_attachment=True,
            download_name='result.xlsx'
        )
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@app.route('/api/check-column-type', methods=['POST'])
def check_column_type():
    try:
        data = request.json
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], data['file'])
        sheet = data['sheet']
        column = data['column']

        # 读取Excel文件
        df = pd.read_excel(file_path, sheet_name=sheet)
        
        # 获取列索引
        col_idx = ord(column.upper()) - ord('A')
        
        # 获取列数据
        col_data = df.iloc[:, col_idx]
        
        # 检查是否可以转换为数值
        try:
            pd.to_numeric(col_data, errors='raise')
            return jsonify({'type': 'numeric'})
        except:
            return jsonify({'type': 'text'})

    except Exception as e:
        return jsonify({'error': str(e)}), 400

@app.route('/api/split-excel', methods=['POST'])
def split_excel():
    try:
        if 'file' not in request.files:
            return jsonify({'success': False, 'error': '请上传文件'})
        
        file = request.files['file']
        if not file.filename:
            return jsonify({'success': False, 'error': '文件名不能为空'})
            
        if not allowed_file(file.filename):
            return jsonify({'success': False, 'error': '不支持的文件格式'})
        
        split_all = request.form.get('splitAll', 'true').lower() == 'true'
        selected_sheets = json.loads(request.form.get('sheets', '[]'))
        
        # 创建临时文件路径
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        temp_filename = f"temp_{timestamp}_{secure_filename(file.filename)}"
        temp_filepath = os.path.join(app.config['UPLOAD_FOLDER'], temp_filename)
        
        try:
            # 保存上传的文件
            file.save(temp_filepath)
            
            # 创建 ExcelProcessor 实例
            processor = ExcelProcessor()
            
            # 重新打开文件
            with open(temp_filepath, 'rb') as f:
                file_obj = FileStorage(
                    stream=f,
                    filename=file.filename,
                    content_type=file.content_type
                )
                # 处理文件拆分
                result_files = processor.split_excel(file_obj, split_all, selected_sheets)
            
            return jsonify({
                'success': True,
                'files': result_files
            })
            
        except Exception as e:
            return jsonify({'success': False, 'error': str(e)})
            
        finally:
            # 清理临时文件
            try:
                if os.path.exists(temp_filepath):
                    os.remove(temp_filepath)
            except Exception as e:
                print(f"清理临时文件失败: {str(e)}")
        
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@app.route('/api/merge-excel', methods=['POST'])
def merge_excel():
    try:
        if 'files[]' not in request.files:
            return jsonify({'success': False, 'error': '请上传文件'})
        
        files = request.files.getlist('files[]')
        if not files or not any(file.filename for file in files):
            return jsonify({'success': False, 'error': '请选择至少一个文件'})
            
        for file in files:
            if not allowed_file(file.filename):
                return jsonify({'success': False, 'error': f'不支持的文件格式: {file.filename}'})
        
        merge_all = request.form.get('mergeAll', 'true').lower() == 'true'
        add_source = request.form.get('addSource', 'true').lower() == 'true'
        merge_mode = request.form.get('mergeMode', 'sheets')
        output_filename = request.form.get('filename', '合并文件')
        
        # 创建 ExcelProcessor 实例
        processor = ExcelProcessor()
        
        # 处理文件合并
        result = processor.merge_excel(
            files,
            merge_all=merge_all,
            add_source=add_source,
            output_filename=output_filename,
            merge_mode=merge_mode
        )
        
        return jsonify({
            'success': True,
            'filename': result['filename'],
            'downloadUrl': result['downloadUrl']
        })
        
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@app.route('/download/<filename>')
def download_file(filename):
    try:
        # 从下载目录获取文件
        download_path = os.path.join(app.config['DOWNLOAD_FOLDER'], filename)
        
        if not os.path.exists(download_path):
            return jsonify({'error': '文件不存在'}), 404
            
        return send_file(
            download_path,
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# 添加定期清理临时文件的功能
def cleanup_temp_files():
    """定期清理临时文件"""
    try:
        current_time = datetime.now()
        expiry_time = current_time - timedelta(seconds=config.RESULT_EXPIRY)
        
        # 清理上传目录
        for filename in os.listdir(app.config['UPLOAD_FOLDER']):
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            if os.path.getmtime(file_path) < expiry_time.timestamp():
                try:
                    os.remove(file_path)
                except Exception as e:
                    print(f"清理临时文件失败: {str(e)}")
        
        # 清理结果目录
        for filename in os.listdir(app.config['RESULT_FOLDER']):
            file_path = os.path.join(app.config['RESULT_FOLDER'], filename)
            if os.path.getmtime(file_path) < expiry_time.timestamp():
                try:
                    os.remove(file_path)
                except Exception as e:
                    print(f"清理结果文件失败: {str(e)}")
                    
    except Exception as e:
        print(f"清理文件时出错: {str(e)}")

@app.route('/api/preview-data', methods=['POST'])
def preview_data():
    try:
        print("接收到预览请求")
        if 'file' not in request.files:
            print("请求中没有文件")
            return jsonify({'error': '请选择文件'}), 400
        
        file = request.files['file']
        if not file.filename:
            print("文件名为空")
            return jsonify({'error': '文件名不能为空'}), 400
            
        print(f"处理文件: {file.filename}")
        
        source_type = request.form.get('sourceType')
        start_row = int(request.form.get('startRow', 1))
        header_row = int(request.form.get('headerRow', 1))
        auto_split = request.form.get('autoSplit', 'true').lower() == 'true'
        delimiter = request.form.get('delimiter')

        print(f"参数: source_type={source_type}, start_row={start_row}, header_row={header_row}, auto_split={auto_split}, delimiter={delimiter}")

        processor = ImportProcessor(
            upload_folder=app.config['UPLOAD_FOLDER'],
            download_folder=app.config['DOWNLOAD_FOLDER']
        )
        
        preview_data = processor.preview_data(
            file=file,
            source_type=source_type,
            delimiter=delimiter,
            start_row=start_row,
            header_row=header_row,
            auto_split=auto_split
        )
        
        return jsonify({'data': preview_data})
    except Exception as e:
        print(f"预览错误: {str(e)}")
        return jsonify({'error': str(e)}), 400

@app.route('/api/import-data', methods=['POST'])
def import_data():
    try:
        files = request.files.getlist('files[]')
        source_type = request.form.get('sourceType')
        start_row = int(request.form.get('startRow', 1))
        header_row = int(request.form.get('headerRow', 1))
        auto_split = request.form.get('autoSplit', 'true').lower() == 'true'
        delimiter = request.form.get('delimiter')

        # 修改这里：传入必要的文件夹路径
        processor = ImportProcessor(
            upload_folder=app.config['UPLOAD_FOLDER'],
            download_folder=app.config['DOWNLOAD_FOLDER']
        )
        result = processor.import_data(
            files=files,
            source_type=source_type,
            delimiter=delimiter,
            start_row=start_row,
            header_row=header_row,
            auto_split=auto_split
        )
        
        return jsonify(result)
    except Exception as e:
        return jsonify({'error': str(e)}), 400

@app.after_request
def add_security_headers(response):
    """添加安全相关的响应头"""
    response.headers['Content-Security-Policy'] = (
        "default-src 'self'; "
        "script-src 'self' 'unsafe-inline' 'unsafe-eval'; "
        "style-src 'self' 'unsafe-inline'; "
        "font-src 'self' data: https:; "
        "img-src 'self' data: https:; "
        "media-src 'self' data: https:; "  # 添加 media-src
        "connect-src 'self'"
    )
    return response

# 确保所有必要的目录都正确创建
@app.before_request
def check_file_upload():
    """检查文件上传请求"""
    if request.method == 'POST' and request.files:
        print("\n=== 文件上传请求检查 ===")
        print(f"端点: {request.endpoint}")
        print(f"文件: {list(request.files.keys())}")
        for key, file in request.files.items():
            print(f"文件 {key}: {file.filename} ({type(file)})")
        print("======================\n")

# 在应用启动时调用初始化函数
if __name__ == '__main__':
    with app.app_context():
        init_app()
    app.run(debug=True)