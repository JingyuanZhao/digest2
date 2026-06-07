#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
from flask import Flask, render_template, request, jsonify
import tempfile

# 尝试导入 digest2，如果失败则标记
try:
    from digest2 import Digest2, parse_mpc80
    from digest2.observation import parse_ades_psv, parse_ades_xml, _parse_ades_optical, _parse_ades_psv_row
    HAS_DIGEST2 = True
except ImportError:
    HAS_DIGEST2 = False
    Digest2 = None
    parse_mpc80 = None
    parse_ades_psv = None
    parse_ades_xml = None
    _parse_ades_optical = None
    _parse_ades_psv_row = None

try:
    from docx import Document
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

app = Flask(__name__)

@app.route('/')
def index():
    digest2_version = '2.8.0'
    try:
        import importlib.metadata as metadata
        try:
            digest2_version = metadata.version('digest2')
        except metadata.PackageNotFoundError:
            try:
                import digest2
                digest2_version = getattr(digest2, '__version__', getattr(digest2, 'VERSION', '2.8.0'))
            except ImportError:
                pass
    except Exception:
        pass
    
    return render_template('index.html', digest2_version=digest2_version)

@app.route('/types')
def types():
    return render_template('types.html')

@app.route('/about')
def about():
    digest2_version = '2.8.0'
    try:
        import importlib.metadata as metadata
        try:
            digest2_version = metadata.version('digest2')
        except metadata.PackageNotFoundError:
            try:
                import digest2
                digest2_version = getattr(digest2, '__version__', getattr(digest2, 'VERSION', '2.8.0'))
            except ImportError:
                pass
    except Exception:
        pass
    
    return render_template('about.html', digest2_version=digest2_version)

@app.route('/download')
def download():
    return render_template('download.html')

@app.route('/analyze', methods=['POST'])
def analyze():
    if not HAS_DIGEST2:
        return jsonify({'error': 'digest2 模块未安装，无法进行分析'}), 500
    
    try:
        data = request.form
        observations = data.get('observations', '').strip()
        orbit_class = data.get('orbit_class', 'all')
        
        if not observations:
            return jsonify({'error': '请输入观测数据'}), 400
        
        results = []
        
        is_ades_xml = observations.startswith('<?xml') or observations.startswith('<ades>') or '<optical>' in observations[:1000]
        
        lines = observations.split('\n')
        is_ades_psv = False
        has_psv_header = False
        has_data_with_pipes = False
        
        for line in lines:
            stripped_line = line.strip()
            if not stripped_line:
                continue
            
            if stripped_line.startswith('!') and '|' in stripped_line:
                has_psv_header = True
                break
            
            if stripped_line.count('|') >= 2:
                has_data_with_pipes = True
                if len(stripped_line) > 12 and stripped_line[12] == '|':
                    if stripped_line.count('|') > 1:
                        has_data_with_pipes = True
                    else:
                        has_data_with_pipes = False
                break
        
        is_ades_psv = has_psv_header or has_data_with_pipes
        
        if is_ades_xml:
            results = _analyze_ades_xml_input(observations)
        elif is_ades_psv:
            results = _analyze_ades_psv_input(observations)
        else:
            results = _analyze_mpc80_input(observations)
        
        if orbit_class == 'neo':
            results = [r for r in results if r['NEO'] > 0]
        
        total_observations = sum(r.get('observations', 0) for r in results)
        
        return jsonify({'success': True, 'results': results, 'total_observations': total_observations})
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500

def _analyze_ades_xml_input(input_data):
    results = []
    from digest2.observation import parse_ades_xml
    
    with tempfile.NamedTemporaryFile(mode='w', suffix='.xml', delete=False, encoding='utf-8') as f:
        f.write(input_data)
        temp_path = f.name
    
    try:
        tracklets = parse_ades_xml(temp_path)
        
        for desig in tracklets:
            tracklets[desig].sort(key=lambda o: o.mjd)
        
        with Digest2() as d2:
            for desig, obs_list in tracklets.items():
                try:
                    result = d2.classify_tracklet(obs_list, is_ades=True)
                    results.append(_format_result(result, desig))
                except Exception:
                    continue
    
    finally:
        os.unlink(temp_path)
    
    return results

def _analyze_ades_psv_input(input_data):
    results = []
    from digest2.observation import parse_ades_psv
    import tempfile
    import os
    
    with tempfile.NamedTemporaryFile(mode='w', suffix='.psv', delete=False, encoding='utf-8') as f:
        f.write(input_data)
        temp_path = f.name
    
    try:
        tracklets = parse_ades_psv(temp_path)
        
        for desig in tracklets:
            tracklets[desig].sort(key=lambda o: o.mjd)
        
        with Digest2() as d2:
            for desig, obs_list in tracklets.items():
                try:
                    result = d2.classify_tracklet(obs_list, is_ades=True)
                    results.append(_format_result(result, desig))
                except Exception:
                    continue
    finally:
        os.unlink(temp_path)
    
    return results

def _date_to_mjd(year, month, day):
    """将日期转换为MJD（使用digest2库的标准实现）"""
    from digest2.observation import _date_to_mjd as digest_mjd
    return digest_mjd(year, month, day)

def _parse_observation_line(line):
    """
    解析观测数据行，支持多种格式：
    1. 标准 MPC80 格式（80列）
    2. 简化格式（如 H251538 C2017 01 26.46663 ...）
    """
    from digest2.observation import Observation
    import re

    line = line.strip()
    if not line:
        return None, None

    # 尝试标准 MPC80 格式
    if len(line) >= 80:
        fixed_line = line
        note2 = line[14] if len(line) > 14 else ' '
        if note2 not in ('C', 'S', 'B'):
            if len(line) > 15 and line[15] in ('C', 'S', 'B'):
                fixed_line = line[:14] + line[15:]
            elif len(line) > 13 and line[12] == '|' and line[13] in ('C', 'S', 'B'):
                fixed_line = line[:12] + ' ' + line[13:]
        obs = parse_mpc80(fixed_line)
        if obs:
            desig = line[:12].strip() or "Unknown"
            return desig, obs
    
    # 尝试解析简化格式
    try:
        line_clean = line.strip()
        line_clean_no_asterisk = line_clean.replace('*', ' ')
        
        obscode = line_clean[-3:] if len(line_clean) >= 3 else '500'
        
        line_clean = re.sub(r'(\d+\.\d{3})([+-])', r'\1 \2', line_clean_no_asterisk)
        line_clean = re.sub(r'(\d+\.\d{6})(\d{2})', r'\1 \2', line_clean)
        
        parts = line_clean.split()
        if len(parts) < 9:
            return None, None
        
        desig = parts[0]
        
        date_field = parts[1]
        year_match = re.search(r'\d{4}', date_field)
        if not year_match:
            return None, None
        year = int(year_match.group(0))
        month = int(parts[2])
        day = float(parts[3])
        
        ra_h = float(parts[4])
        ra_m = float(parts[5])
        ra_s = float(parts[6])
        ra_deg = (ra_h + ra_m/60 + ra_s/3600) * 15
        
        dec_sign = 1 if parts[7][0] == '+' else -1
        dec_d = abs(float(parts[7]))
        dec_m = float(parts[8])
        dec_s = float(parts[9]) if len(parts) > 9 else 0.0
        dec_deg = dec_sign * (dec_d + dec_m/60 + dec_s/3600)
        
        mag = None
        band = 'V'
        
        if len(parts) >= 11:
            mag_field = parts[10]
            if mag_field.replace('.', '', 1).isdigit() and mag_field.count('.') <= 1:
                mag = float(mag_field)
                if len(parts) >= 12:
                    if len(parts[11]) == 1 and parts[11].isalpha():
                        band = parts[11].upper()
                    elif len(parts[11]) == 2 and parts[11][0].isalpha() and parts[11][1].isdigit():
                        band = parts[11][0].upper()
            else:
                mag_match = re.search(r'^(\d+\.\d+)([A-Za-z])(.*)$', mag_field)
                if mag_match:
                    mag = float(mag_match.group(1))
                    band = mag_match.group(2).upper()
                else:
                    mag = -1.0
        
        obs = Observation(
            mjd=_date_to_mjd(year, month, day),
            ra=ra_deg,
            dec=dec_deg,
            mag=mag,
            band=band,
            obscode=obscode
        )
        
        return desig, obs
    
    except Exception:
        return None, None

def _analyze_mpc80_input(input_data):
    results = []
    observations = []
    
    for line in input_data.split('\n'):
        line = line.rstrip()
        if line:
            try:
                desig, obs = _parse_observation_line(line)
                if obs:
                    observations.append((desig, obs))
            except Exception:
                continue
    
    tracklets = {}
    for desig, obs in observations:
        if desig not in tracklets:
            tracklets[desig] = []
        tracklets[desig].append(obs)
    
    for desig in tracklets:
        tracklets[desig].sort(key=lambda o: o.mjd)
    
    with Digest2() as d2:
        for desig, obs_list in tracklets.items():
            try:
                result = d2.classify_tracklet(obs_list)
                results.append(_format_result(result, desig, len(obs_list)))
            except Exception:
                continue
    
    return results

def _format_score(score):
    """格式化评分，显示一位小数（与GUI版本一致）"""
    return f"{score:.1f}"

def _format_result(result, desig, num_observations=0):
    scores = result.noid
    other_possibilities = []
    
    # 定义最小显示阈值（避免显示极小的浮点数值）
    min_score = 0.05
    
    # 只显示有意义评分的轨道类型（评分 >= min_score）
    if scores.MC >= min_score:
        other_possibilities.append(f"(MC {_format_score(scores.MC)})")
    if scores.MB1 >= min_score:
        other_possibilities.append(f"(MB1 {_format_score(scores.MB1)})")
    if scores.MB2 >= min_score:
        other_possibilities.append(f"(MB2 {_format_score(scores.MB2)})")
    if scores.MB3 >= min_score:
        other_possibilities.append(f"(MB3 {_format_score(scores.MB3)})")
    if scores.Hun >= min_score:
        other_possibilities.append(f"(Hil {_format_score(scores.Hun)})")
    if scores.Pho >= min_score:
        other_possibilities.append(f"(Pho {_format_score(scores.Pho)})")
    if scores.Pal >= min_score:
        other_possibilities.append(f"(Pal {_format_score(scores.Pal)})")
    if scores.Han >= min_score:
        other_possibilities.append(f"(Han {_format_score(scores.Han)})")
    if scores.Hil >= min_score:
        other_possibilities.append(f"(Hil {_format_score(scores.Hil)})")
    if scores.JTr >= min_score:
        other_possibilities.append(f"(JTr {_format_score(scores.JTr)})")
    if scores.JFC >= min_score:
        other_possibilities.append(f"(JFC {_format_score(scores.JFC)})")
    
    # 用空格连接
    other_possibilities_str = ' '.join(other_possibilities) if other_possibilities else ''
    
    # 获取RMS值，如果result.rms不可用，尝试从其他属性获取
    rms_value = getattr(result, 'rms', None)
    if rms_value is None or rms_value == 0:
        rms_value = getattr(result, 'orbit_rms', 0.0)
    
    return {
        'designation': desig,
        'RMS': f"{rms_value:.2f}",
        'Int': _format_score(scores.Int),
        'NEO': _format_score(scores.NEO),
        'N22': _format_score(scores.N22),
        'N18': _format_score(scores.N18),
        'Other_Possibilities': other_possibilities_str,
        'observations': num_observations
    }

def _filter_psv_content(content):
    """过滤PSV格式内容，跳过注释行和元数据行"""
    lines = []
    for line in content.split('\n'):
        stripped = line.rstrip('\n')
        # 跳过注释行 (# 开头)
        if stripped.startswith('#'):
            continue
        # 跳过元数据行 (! 开头但不包含 |)
        if stripped.startswith('!') and '|' not in stripped:
            continue
        # 保留表头行 (! 开头且包含 |) 和数据行
        lines.append(stripped)
    return '\n'.join(lines)

def _filter_mpc80_content(content):
    """过滤MPC80格式内容，只保留至少有两条观测的天体"""
    lines = content.split('\n')
    
    tracklets = {}
    seen_lines = set()
    
    for line in lines:
        line = line.rstrip()
        if not line:
            continue
        
        desig = line[:12].strip() if len(line) >= 12 else "Unknown"
        
        if line not in seen_lines:
            seen_lines.add(line)
            if desig not in tracklets:
                tracklets[desig] = []
            tracklets[desig].append(line)
    
    filtered = []
    for desig, obs_lines in tracklets.items():
        if len(obs_lines) >= 2:
            filtered.extend(obs_lines)
    
    return '\n'.join(filtered)

def _detect_and_filter_content(content):
    """自动检测内容格式并进行相应的过滤"""
    is_ades_xml = content.startswith('<?xml') or content.startswith('<ades>') or '<optical>' in content[:1000]
    
    if is_ades_xml:
        return content
    
    lines = content.strip().split('\n')
    has_psv_header = False
    has_data_with_pipes = False
    
    for line in lines:
        stripped_line = line.strip()
        if not stripped_line:
            continue
        
        if stripped_line.startswith('!') and '|' in stripped_line:
            has_psv_header = True
            break
        
        if stripped_line.count('|') >= 2:
            has_data_with_pipes = True
            if len(stripped_line) > 12 and stripped_line[12] == '|':
                if stripped_line.count('|') > 1:
                    has_data_with_pipes = True
                else:
                    has_data_with_pipes = False
            break
    
    if has_psv_header or has_data_with_pipes:
        return _filter_psv_content(content)
    else:
        return _filter_mpc80_content(content)

@app.route('/upload-docx', methods=['POST'])
def upload_docx():
    try:
        if not HAS_DOCX:
            return jsonify({'success': False, 'error': '未安装 python-docx 库，无法处理 docx 文件'}), 500
        
        if 'file' not in request.files:
            return jsonify({'success': False, 'error': '未选择文件'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'success': False, 'error': '未选择文件'}), 400
        
        ext = file.filename.split('.')[-1].lower()
        if ext not in ['docx', 'doc']:
            return jsonify({'success': False, 'error': '不支持的文件格式'}), 400
        
        # 保存临时文件
        with tempfile.NamedTemporaryFile(mode='wb', suffix='.docx', delete=False) as f:
            file.save(f.name)
            temp_path = f.name
        
        try:
            # 读取docx内容
            doc = Document(temp_path)
            text_lines = []
            for para in doc.paragraphs:
                if para.text.strip():
                    text_lines.append(para.text)
            
            content = '\n'.join(text_lines)
            filtered_content = _detect_and_filter_content(content)
            return jsonify({'success': True, 'content': filtered_content})
        
        finally:
            os.unlink(temp_path)
    
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/upload-file', methods=['POST'])
def upload_file():
    try:
        if 'file' not in request.files:
            return jsonify({'success': False, 'error': '未选择文件'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'success': False, 'error': '未选择文件'}), 400
        
        content = file.read().decode('utf-8', errors='replace')
        filtered_content = _detect_and_filter_content(content)
        
        return jsonify({'success': True, 'content': filtered_content})
    
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    debug_mode = os.environ.get('FLASK_ENV') != 'production'
    app.run(debug=debug_mode, host='0.0.0.0', port=port)

# WSGI entry point for gunicorn
application = app