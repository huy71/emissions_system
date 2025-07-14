from flask import Flask, jsonify, request, send_from_directory
import pandas as pd
import os
from flask_cors import CORS
from datetime import datetime
import re

app = Flask(__name__)
CORS(app)

EXCEL_FILE = 'SoLieuKhiThai.xlsx'

def load_df():
    import pandas as pd
    import os

    if os.path.exists(EXCEL_FILE):
        df_raw = pd.read_excel(EXCEL_FILE, header=None)
        required_cols = ['Thời gian', 'CO((mg/Nm3))', 'SO2_1((mg/Nm3))']
        header_row = None
        for i, row in df_raw.iterrows():
            if all(col in row.values for col in required_cols):
                header_row = i
                break
        if header_row is not None:
            df = pd.read_excel(EXCEL_FILE, header=header_row)
            # Chỉ giữ lại các dòng có giá trị hợp lệ ở cột 'Thời gian'
            df = df[df['Thời gian'].apply(lambda x: pd.notnull(x) and str(x).strip() != '' and str(x).count('/') == 2)]
            return df
        else:
            return pd.DataFrame()
    else:
        return pd.DataFrame()

@app.route('/api/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    if not file.filename.lower().endswith(('.xlsx', '.xls')):
        return jsonify({'error': 'File must be .xlsx or .xls'}), 400
    file.save(EXCEL_FILE)
    return jsonify({'success': True})

@app.route('/api/data')
def get_data():
    try:
        df = load_df()
        if df.empty:
            return jsonify({'data': [], 'total': 0, 'page': 1, 'total_pages': 0})

        # Đảm bảo cột 'Thời gian' là datetime với format dd/mm/yyyy
        if 'Thời gian' in df.columns:
            df['Thời gian'] = pd.to_datetime(df['Thời gian'], format='%d/%m/%Y %H:%M:%S', errors='coerce')
            if df['Thời gian'].isna().any():
                df['Thời gian'] = pd.to_datetime(df['Thời gian'], format='%d/%m/%Y', errors='coerce')

        from_time = request.args.get('from_time')
        to_time = request.args.get('to_time')
        print('from_time:', from_time, 'to_time:', to_time)
        print('df["Thời gian"] sample:', df['Thời gian'].head(3).tolist())

        def convert_date_str(date_str):
            import re
            if date_str:
                # Nếu có dạng yyyy-mm-dd HH:MM:SS hoặc yyyy-mm-dd
                match = re.match(r'(\d{4})-(\d{2})-(\d{2})', date_str)
                if match:
                    y, m, d = match.groups()
                    return f"{d}/{m}/{y}"
            return date_str
        from_time_fmt = convert_date_str(from_time)
        to_time_fmt = convert_date_str(to_time)

        # So sánh chỉ theo ngày, dùng dayfirst=True
        if from_time_fmt:
            from_time_dt = pd.to_datetime(from_time_fmt, dayfirst=True, errors='coerce')
            df = df[df['Thời gian'].dt.date >= from_time_dt.date()]
        if to_time_fmt:
            to_time_dt = pd.to_datetime(to_time_fmt, dayfirst=True, errors='coerce')
            df = df[df['Thời gian'].dt.date <= to_time_dt.date()]

        print('Filtered df:', df.head(3).to_dict())

        # Lọc theo chỉ số và cấp độ vượt chuẩn
        metric = request.args.get('metric')
        level = request.args.get('level')
        value_min = request.args.get('value_min')
        value_max = request.args.get('value_max')
        qc_dict = {
            'CO((mg/Nm3))': 900,
            'SO2_1((mg/Nm3))': 450,
            'NOX_1((mg/Nm3))': 765,
            'O2_1(%)': 21,
            'Q_1(m3/h)': 9999999999,
            'Temp_1((oC))': 200,
            'Dust_1((mg/Nm3))': 180,
            'Pkq': None
        }
        if metric and metric in df.columns:
            if level and level != 'all' and qc_dict.get(metric):
                qc = qc_dict[metric]
                ratio = df[metric] / qc
                if level == 'Đạt QC':
                    df = df[ratio < 1]
                elif level == 'Cấp 1':
                    df = df[(ratio >= 1.1) & (ratio < 2)]
                elif level == 'Cấp 2':
                    df = df[(ratio >= 2) & (ratio < 5)]
                elif level == 'Cấp 3':
                    df = df[(ratio >= 5) & (ratio < 10)]
                elif level == 'Cấp 4':
                    df = df[ratio >= 10]
            # Lọc theo giá trị cụ thể hoặc khoảng giá trị
            if value_min:
                df = df[df[metric] >= float(value_min)]
            if value_max:
                df = df[df[metric] <= float(value_max)]

        page = int(request.args.get('page', 1))
        page_size = int(request.args.get('page_size', 20))
        total = len(df)

        start = (page - 1) * page_size
        end = start + page_size
        data = df.iloc[start:end].to_dict(orient='records')
        return jsonify({
            'data': data,
            'total': total,
            'page': page,
            'page_size': page_size,
            'total_pages': (total + page_size - 1) // page_size
        })
    except Exception as e:
        return jsonify({'data': [], 'total': 0, 'page': 1, 'total_pages': 0, 'error': str(e)}), 500

@app.route('/api/stats')
def get_stats():
    df = load_df()
    stats = {}
    for col in df.columns:
        if df[col].dtype.kind in 'biufc' and col != 'Thời gian':
            stats[col] = {
                'min': float(df[col].min()),
                'max': float(df[col].max()),
                'avg': float(df[col].mean())
            }
    return jsonify(stats)

@app.route('/api/summary')
def get_summary():
    df = load_df()
    qc_dict = {
        'CO((mg/Nm3))': 900,
        'SO2_1((mg/Nm3))': 450,
        'NOX_1((mg/Nm3))': 765,
        'O2_1(%)': 21,
        'Q_1(m3/h)': 9999999999,
        'Temp_1((oC))': 200,
        'Dust_1((mg/Nm3))': 180,
        'Pkq': None
    }
    result = {}
    total = len(df)
    for col, qc in qc_dict.items():
        if col not in df.columns or qc is None:
            continue
        values = df[col]
        ratio = values / qc
        levels = {
            'Đạt QC': (ratio < 1),
            'Cấp 1': (ratio >= 1.1) & (ratio < 2),
            'Cấp 2': (ratio >= 2) & (ratio < 5),
            'Cấp 3': (ratio >= 5) & (ratio < 10),
            'Cấp 4': (ratio >= 10)
        }
        level_stats = {}
        for level, cond in levels.items():
            count = cond.sum()
            percent = round(count / total * 100, 1) if total > 0 else 0
            level_stats[level] = {'count': int(count), 'percent': percent}
        min_val = float(values.min())
        max_val = float(values.max())
        avg_val = float(values.mean())
        result[col] = {
            'levels': level_stats,
            'min': min_val,
            'max': max_val,
            'avg': avg_val,
            'qc': qc
        }
    return jsonify({'total': total, 'summary': result})

@app.route('/api/delete', methods=['POST'])
def delete_data():
    try:
        # Xóa file Excel nếu tồn tại
        if os.path.exists(EXCEL_FILE):
            os.remove(EXCEL_FILE)
        return jsonify({'success': True, 'message': 'Đã xóa toàn bộ dữ liệu'})
    except Exception as e:
        return jsonify({'success': False, 'error': f'Lỗi khi xóa dữ liệu: {str(e)}'}), 500

@app.route('/')
def index():
    return send_from_directory('', 'index.html')

@app.route('/test')
def test_page():
    return send_from_directory('', 'test_delete.html')

@app.route('/<path:path>')
def static_files(path):
    return send_from_directory('', path)

if __name__ == '__main__':
    app.run(debug=True) 