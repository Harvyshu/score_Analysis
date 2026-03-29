import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import requests
import json
import os
import shutil
from datetime import datetime
import re
from PIL import Image, ImageTk
import zipfile
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import threading  # 新增：处理流式请求的线程
import time

# 配置matplotlib中文显示
plt.rcParams['font.sans-serif'] = ['SimHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

# 全局变量
CURRENT_STU = None
STU_DIR = None
DATA_DIR = os.path.join(os.getcwd(), 'data')
TEMPLATE_DIR = os.path.join(os.getcwd(), 'template')
TMP_DIR = os.path.join(DATA_DIR, 'tmp')
DEFAULT_AVATAR = os.path.join(os.getcwd(), 'default_avatar.png')
CONFIG_PATH = os.path.join(os.getcwd(), 'config.json')

# 加载配置文件
try:
    with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
        CONFIG = json.load(f)
except:
    CONFIG = {
        "base_url": "https://api.deepseek.com/chat/completions",
        "dp_model": "deepseek-reasoner",
        "speed": 1,
        "role": "user",
        "max_tokens": 6144,
        "temperature": 1,
        "top_p": 0.7,
        "top_k": 50,
        "frequency_penalty": 0.5,
        "n": 1,
        "dp_key": "",
        "stream": True,  # 改为true启用流式返回
        "isTools": False,
        "role_prompt": "你是一位资深的教育学专家，一线老师，也是一位资深的儿童心理学专家，请你根据{dataxlsx_content}的数据分析一下孩子的各科的学习趋势，并且给最近一次的成绩做一次深入的分析，给出鼓励和建议",
        "seed": 4999999999,
        "timeout": 60,
        "retry_times": 3
    }


# 校验Excel文件是否有效
def is_valid_excel(file_path):
    if not os.path.exists(file_path):
        return False
    if not file_path.lower().endswith('.xlsx'):
        return False
    try:
        with zipfile.ZipFile(file_path, 'r') as zf:
            required_files = ['[Content_Types].xml', 'xl/workbook.xml']
            file_list = zf.namelist()
            return all(f in file_list for f in required_files)
    except:
        return False


# 重建损坏的Excel文件
def rebuild_excel(file_path):
    df_empty = pd.DataFrame({
        'date': [], '语文_score': [], '语文_content': [],
        '数学_score': [], '数学_content': [], '英语_score': [],
        '英语_content': [], '物理_score': [], '物理_content': [],
        '化学_score': [], '化学_content': [], '生物_score': [],
        '生物_content': [], '历史_score': [], '历史_content': [],
        '地理_score': [], '地理_content': [], '政治_score': [],
        '政治_content': []
    })
    df_empty.to_excel(file_path, index=False, engine='openpyxl')
    return True


# 初始化目录
def init_dirs():
    for dir_path in [DATA_DIR, TEMPLATE_DIR, TMP_DIR, os.path.join(TEMPLATE_DIR, 'reports')]:
        if not os.path.exists(dir_path):
            os.makedirs(dir_path)
    template_excel = os.path.join(TEMPLATE_DIR, 'data.xlsx')
    if not is_valid_excel(template_excel):
        rebuild_excel(template_excel)


# 创建学生空间
def create_stu_space(stu_name):
    if stu_name == 'tmp':
        messagebox.showerror('错误', '学生名不能为"tmp"！')
        return None
    stu_dir = os.path.join(DATA_DIR, stu_name)
    if os.path.exists(stu_dir):
        res = messagebox.askyesno('提示', f'已存在{stu_name}空间，是否覆盖创建？（原数据会丢失）')
        if not res:
            return stu_dir
        shutil.rmtree(stu_dir)
    shutil.copytree(TEMPLATE_DIR, stu_dir)
    excel_path = os.path.join(stu_dir, 'data.xlsx')
    try:
        df_template = pd.read_excel(excel_path, engine='openpyxl')
        with pd.ExcelWriter(excel_path, engine='openpyxl', mode='w') as writer:
            df_template.to_excel(writer, sheet_name=stu_name, index=False)
    except Exception as e:
        messagebox.warning('提示', f'修改Sheet名失败：{str(e)}，不影响数据使用')
    return stu_dir


# 通用日期格式化函数
def format_date_str(date_str):
    if (pd.isna(date_str) or
            date_str in ['NaT', 'nan', 'None', '', 'NaN'] or
            str(date_str).strip() == '' or
            'nan' in str(date_str).lower()):
        return ''

    date_str = str(date_str).strip()
    try:
        formats = ['%Y-%m-%d', '%Y/%m/%d', '%Y.%m.%d', '%m-%d', '%m/%d', '%m.%d']
        for fmt in formats:
            try:
                dt = datetime.strptime(date_str, fmt)
                return dt.strftime('%m/%d')
            except:
                continue
        return date_str[:5] if len(date_str) >= 5 else date_str
    except:
        return date_str[:5] if len(date_str) >= 5 else date_str


# 加载成绩数据
def load_stu_data(stu_dir):
    excel_path = os.path.join(stu_dir, 'data.xlsx')
    if not os.path.exists(excel_path):
        messagebox.showerror('错误', f'Excel文件不存在：{excel_path}')
        return pd.DataFrame()
    if not is_valid_excel(excel_path):
        res = messagebox.askyesno('文件损坏', f'检测到{excel_path}文件格式异常/损坏，是否自动重建？（原数据会丢失）')
        if res:
            rebuild_excel(excel_path)
        else:
            return pd.DataFrame()
    try:
        df = pd.read_excel(excel_path, engine='openpyxl')
        if 'date' not in df.columns:
            df['date'] = ''
        df['date'] = df['date'].astype(str)

        df['date_formatted'] = df['date'].apply(format_date_str)
        df['date_final'] = df['date_formatted'].fillna('')
        for idx in df.index:
            if df.loc[idx, 'date_final'] == '' or 'nan' in df.loc[idx, 'date_final'].lower():
                df.loc[idx, 'date_final'] = f'第{idx + 1}次'

        df['date_final'] = df['date_final'].astype(str)
        return df
    except Exception as e:
        messagebox.showerror('读取失败', f'读取Excel出错：{str(e)}\n将自动重建空数据文件')
        rebuild_excel(excel_path)
        return pd.read_excel(excel_path, engine='openpyxl')


# 保存成绩数据
def save_stu_data(stu_dir, df):
    excel_path = os.path.join(stu_dir, 'data.xlsx')
    for col in ['date_original', 'date_formatted', 'date_final']:
        if col in df.columns:
            df = df.drop(columns=[col])
    df.to_excel(excel_path, index=False, engine='openpyxl')
    if not os.path.exists(TMP_DIR):
        os.makedirs(TMP_DIR)
    shutil.copy(excel_path, os.path.join(TMP_DIR, 'data.xlsx'))


# 绘制成绩趋势图
def draw_score_chart(df, canvas):
    for widget in canvas.winfo_children():
        widget.destroy()

    score_cols = [col for col in df.columns if '_score' in col]
    df = df.dropna(subset=score_cols, how='all').reset_index(drop=True)
    if len(df) == 0:
        tk.Label(canvas, text='未加载成绩数据，暂无图表', font=('宋体', 12)).pack(pady=80)
        return

    df_latest = df.tail(8).reset_index(drop=True)
    total_rows = len(df_latest)

    if 'date_final' in df_latest.columns:
        x_dates = df_latest['date_final'].tolist()
    else:
        x_dates = [f'第{i + 1}次' for i in range(total_rows)]

    clean_x_dates = []
    for i, x in enumerate(x_dates):
        x_str = str(x).strip()
        if x_str == '' or 'nan' in x_str.lower() or x_str == 'nan':
            clean_x_dates.append(f'第{i + 1}次')
        else:
            clean_x_dates.append(x_str)

    all_subjects = ['语文', '数学', '英语', '物理', '化学', '生物', '历史', '地理', '政治']
    valid_subjects = []
    subject_data = {}

    for sub in all_subjects:
        score_col = f'{sub}_score'
        if score_col in df_latest.columns:
            scores = df_latest[score_col].tolist()
            all_y = []
            for score in scores:
                if pd.notna(score) and isinstance(score, (int, float)):
                    all_y.append(score)
                else:
                    all_y.append(np.nan)
            if len([y for y in all_y if not np.isnan(y)]) > 0:
                valid_subjects.append(sub)
                subject_data[sub] = (clean_x_dates, all_y)

    fig, ax = plt.subplots(figsize=(10, 4), dpi=80)
    colors = ['blue', 'orange', 'green', 'purple', 'brown', 'pink', 'gray', 'cyan', 'red']
    for idx, sub in enumerate(valid_subjects):
        all_x, all_y = subject_data[sub]
        ax.plot(all_x, all_y, marker='o', label=sub, color=colors[idx],
                linewidth=1.5, markersize=4, markerfacecolor='white', markeredgewidth=1.5)

    ax.set_xlabel('测试日期/次数', fontsize=10)
    ax.set_ylabel('成绩', fontsize=10)
    ax.set_title('成绩趋势图（最近8次）', fontsize=12, fontweight='bold')
    ax.grid(True, alpha=0.3)
    ax.set_ylim(80, 100)

    ax.set_xticks(clean_x_dates)
    ax.set_xticklabels(clean_x_dates, rotation=45, ha='right')
    ax.set_xlim(left=-0.5, right=len(clean_x_dates) - 0.5)

    if valid_subjects:
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=min(len(valid_subjects), 5),
                  fontsize=8, framealpha=0.8, fancybox=True, shadow=False)

    plt.tight_layout(rect=[0, 0.1, 1, 1])

    chart_path = os.path.join(TMP_DIR, 'score_chart.png')
    fig.savefig(chart_path, bbox_inches='tight', pad_inches=0.2)
    plt.close(fig)

    try:
        img = tk.PhotoImage(file=chart_path)
        canvas.img = img
        img_label = tk.Label(canvas, image=img)
        img_label.pack(anchor='center', pady=5)
    except Exception as e:
        tk.Label(canvas, text=f'图表生成失败：{str(e)}', font=('宋体', 12)).pack(pady=80)


# 新增：UI更新函数（线程安全）
def update_report_ui(content, is_append=True):
    """
    安全更新分析报告区域的内容
    :param content: 要显示的内容
    :param is_append: True=追加内容，False=覆盖内容
    """

    def update():
        if not is_append:
            text_report.delete(1.0, tk.END)
        # 解析markdown并显示
        parsed_content = parse_markdown(content)
        text_report.insert(tk.END, parsed_content)
        # 滚动到最新内容
        text_report.see(tk.END)
        text_report.update_idletasks()

    # 使用after方法确保在主线程更新UI
    root.after(0, update)


# 新增：流式API调用函数
def call_deepseek_api_stream(prompt):
    """处理流式返回的API调用"""
    if not CONFIG['dp_key'] or CONFIG['dp_key'].strip() == '':
        err_msg = '请先在config.json中填写有效的DeepSeek API密钥！'
        update_report_ui(err_msg, False)
        messagebox.showerror('错误', err_msg)
        return err_msg

    session = requests.Session()
    retry_strategy = Retry(
        total=CONFIG.get('retry_times', 3),
        backoff_factor=1,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["POST"]
    )
    adapter = HTTPAdapter(max_retries=retry_strategy)
    session.mount("https://", adapter)
    session.mount("http://", adapter)

    url = CONFIG['base_url']
    headers = {
        'Content-Type': 'application/json',
        'Authorization': f'Bearer {CONFIG["dp_key"].strip()}'
    }
    data = {
        'model': CONFIG['dp_model'],
        'messages': [{'role': CONFIG['role'], 'content': prompt}],
        'max_tokens': CONFIG['max_tokens'],
        'temperature': CONFIG['temperature'],
        'top_p': CONFIG['top_p'],
        'top_k': CONFIG['top_k'],
        'frequency_penalty': CONFIG['frequency_penalty'],
        'n': 1,
        'stream': True,  # 强制流式
        'seed': CONFIG['seed']
    }

    full_content = ""
    try:
        # 发送流式请求
        response = session.post(
            url,
            headers=headers,
            json=data,
            timeout=CONFIG.get('timeout', 60),
            stream=True  # 启用流式响应
        )
        response.raise_for_status()

        # 清空原有内容，显示加载提示
        update_report_ui("正在生成分析报告...\n", False)

        # 逐行解析流式数据
        for line in response.iter_lines():
            if line:
                line = line.decode('utf-8').strip()
                # 过滤掉非数据行（如event: ping）
                if line.startswith('data: '):
                    data_str = line[6:]
                    if data_str == '[DONE]':  # 流式结束标记
                        break
                    try:
                        # 解析单条流式数据
                        chunk = json.loads(data_str)
                        if 'choices' in chunk and len(chunk['choices']) > 0:
                            delta = chunk['choices'][0].get('delta', {})
                            content = delta.get('content', '')
                            if content:
                                full_content += content
                                # 实时更新UI（追加内容）
                                update_report_ui(content)
                    except json.JSONDecodeError:
                        continue

        # 流式结束后保存完整报告
        return full_content

    except requests.exceptions.Timeout:
        err_msg = f'API请求超时（已重试{CONFIG.get("retry_times", 3)}次）\n建议：检查网络连接/稍后再试'
        update_report_ui(err_msg, False)
        messagebox.showerror('大模型调用失败', err_msg)
        return err_msg
    except requests.exceptions.ConnectionError:
        err_msg = 'API连接失败\n建议：检查网络连接/代理设置'
        update_report_ui(err_msg, False)
        messagebox.showerror('大模型调用失败', err_msg)
        return err_msg
    except Exception as e:
        err_msg = f'大模型分析失败：{str(e)}\n建议：检查API密钥/账户余额'
        update_report_ui(err_msg, False)
        messagebox.showerror('大模型调用失败', err_msg)
        return err_msg


# 原有API调用函数（非流式）
def call_deepseek_api_non_stream(prompt):
    """处理非流式返回的API调用"""
    if not CONFIG['dp_key'] or CONFIG['dp_key'].strip() == '':
        err_msg = '请先在config.json中填写有效的DeepSeek API密钥！'
        messagebox.showerror('错误', err_msg)
        return err_msg

    session = requests.Session()
    retry_strategy = Retry(
        total=CONFIG.get('retry_times', 3),
        backoff_factor=1,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["POST"]
    )
    adapter = HTTPAdapter(max_retries=retry_strategy)
    session.mount("https://", adapter)
    session.mount("http://", adapter)

    url = CONFIG['base_url']
    headers = {
        'Content-Type': 'application/json',
        'Authorization': f'Bearer {CONFIG["dp_key"].strip()}'
    }
    data = {
        'model': CONFIG['dp_model'],
        'messages': [{'role': CONFIG['role'], 'content': prompt}],
        'max_tokens': CONFIG['max_tokens'],
        'temperature': CONFIG['temperature'],
        'top_p': CONFIG['top_p'],
        'top_k': CONFIG['top_k'],
        'frequency_penalty': CONFIG['frequency_penalty'],
        'n': 1,
        'stream': False,
        'seed': CONFIG['seed']
    }

    try:
        response = session.post(
            url,
            headers=headers,
            json=data,
            timeout=CONFIG.get('timeout', 60)
        )
        response.raise_for_status()
        res_json = response.json()
        content = res_json['choices'][0]['message']['content']
        update_report_ui(content, False)
        return content
    except requests.exceptions.Timeout:
        err_msg = f'API请求超时（已重试{CONFIG.get("retry_times", 3)}次）\n建议：检查网络连接/稍后再试'
        messagebox.showerror('大模型调用失败', err_msg)
        return err_msg
    except requests.exceptions.ConnectionError:
        err_msg = 'API连接失败\n建议：检查网络连接/代理设置'
        messagebox.showerror('大模型调用失败', err_msg)
        return err_msg
    except Exception as e:
        err_msg = f'大模型分析失败：{str(e)}\n建议：检查API密钥/账户余额'
        messagebox.showerror('大模型调用失败', err_msg)
        return err_msg


# 统一API调用入口
def call_deepseek_api(prompt):
    if CONFIG.get('stream', False):
        # 流式调用需要在子线程中执行，避免阻塞UI
        thread = threading.Thread(target=call_deepseek_api_stream, args=(prompt,))
        thread.daemon = True  # 守护线程，关闭程序时自动退出
        thread.start()
        return "STREAMING"  # 标记为流式处理中
    else:
        return call_deepseek_api_non_stream(prompt)


# 生成分析报告
def generate_report(df, stu_dir):
    data_content = df.to_string(index=False)
    role_prompt = CONFIG['role_prompt'].replace('{dataxlsx_content}', data_content)
    report_content = call_deepseek_api(role_prompt)

    # 流式处理时，报告内容在子线程中逐步保存
    if report_content != "STREAMING":
        report_dir = os.path.join(stu_dir, 'reports')
        report_name = f'{datetime.now().strftime("%Y%m%d_%H%M%S")}成绩分析报告.md'
        report_path = os.path.join(report_dir, report_name)
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write(report_content)

        if os.name == 'nt':
            os.startfile(report_path)
        else:
            os.system(f'open {report_path}')
    return report_content


# 解析markdown
def parse_markdown(md_text):
    table_pattern = re.compile(r'^\|.*\|$', flags=re.M)
    table_lines = table_pattern.findall(md_text)
    if table_lines:
        clean_table = []
        for line in table_lines:
            if '---' not in line:
                cells = [cell.strip() for cell in line.strip('|').split('|')]
                cell_widths = [10, 8, 60]
                formatted_cells = []
                for i, cell in enumerate(cells):
                    if i < len(cell_widths):
                        formatted_cells.append(cell.ljust(cell_widths[i])[:cell_widths[i]])
                clean_table.append(' | '.join(formatted_cells))
        md_text = table_pattern.sub('\n'.join(clean_table), md_text)

    md_text = re.sub(r'^## (.*)$', r'\n【\1】\n', md_text, flags=re.M)
    md_text = re.sub(r'^### (.*)$', r'— \1 —', md_text, flags=re.M)
    md_text = re.sub(r'\*\*(.*?)\*\*', r'【\1】', md_text)
    md_text = re.sub(r'\*(.*?)\*', r'\1', md_text)
    md_text = re.sub(r'---+', r'================================', md_text)
    md_text = re.sub(r'\n{3,}', '\n\n', md_text)
    return md_text


# 成绩录入窗口
def input_score_window():
    global STU_DIR, CURRENT_STU
    if CURRENT_STU is None:
        res = messagebox.askyesno('提示', '未加载学生，是否先创建学生空间？')
        if not res:
            return
        stu_name = simpledialog.askstring('创建学生', '请输入学生姓名：')
        if not stu_name:
            return
        STU_DIR = create_stu_space(stu_name)
        CURRENT_STU = stu_name
        messagebox.showinfo('成功', f'已创建{stu_name}学生空间')
        load_avatar()

    input_win = tk.Toplevel(root)
    input_win.title('成绩录入')
    input_win.geometry('800x500')
    input_win.resizable(False, False)

    frame_table = ttk.Frame(input_win)
    frame_table.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

    cols = ['科目', '分数', '考核内容']
    for idx, col in enumerate(cols):
        ttk.Label(frame_table, text=col, font=('黑体', 10, 'bold')).grid(row=0, column=idx, padx=5, pady=5)

    entry_widgets = []
    for row in range(1, 4):
        sub_entry = ttk.Entry(frame_table, width=10)
        score_entry = ttk.Entry(frame_table, width=10)
        content_entry = ttk.Entry(frame_table, width=50)
        sub_entry.grid(row=row, column=0, padx=5, pady=5)
        score_entry.grid(row=row, column=1, padx=5, pady=5)
        content_entry.grid(row=row, column=2, padx=5, pady=5)
        entry_widgets.append([sub_entry, score_entry, content_entry])

    def add_row():
        row = len(entry_widgets) + 1
        sub_entry = ttk.Entry(frame_table, width=10)
        score_entry = ttk.Entry(frame_table, width=10)
        content_entry = ttk.Entry(frame_table, width=50)
        sub_entry.grid(row=row, column=0, padx=5, pady=5)
        score_entry.grid(row=row, column=1, padx=5, pady=5)
        content_entry.grid(row=row, column=2, padx=5, pady=5)
        entry_widgets.append([sub_entry, score_entry, content_entry])

    def del_row():
        if len(entry_widgets) <= 3:
            messagebox.showwarning('提示', '至少保留3行输入框！')
            return
        row_widgets = entry_widgets.pop()
        for w in row_widgets:
            w.destroy()

    frame_btn_row = ttk.Frame(input_win)
    frame_btn_row.pack(pady=5)
    ttk.Button(frame_btn_row, text='+ 新增行', command=add_row, style='Green.TButton').grid(row=0, column=0, padx=5)
    ttk.Button(frame_btn_row, text='- 删除行', command=del_row, style='Red.TButton').grid(row=0, column=1, padx=5)

    def clear_input():
        if messagebox.askyesno('提示', '是否清空所有输入内容？'):
            for row_widgets in entry_widgets:
                for w in row_widgets:
                    w.delete(0, tk.END)

    def submit_score():
        input_data = []
        for row_widgets in entry_widgets:
            sub = row_widgets[0].get().strip()
            score = row_widgets[1].get().strip()
            content = row_widgets[2].get().strip()
            if sub and score:
                try:
                    score = float(score)
                    input_data.append({'subject': sub, 'score': score, 'content': content})
                except ValueError:
                    messagebox.showerror('错误', f'科目{sub}的分数必须是数字！')
                    return
        if not input_data:
            messagebox.showwarning('提示', '请填写至少1条有效成绩数据！')
            return

        df = load_stu_data(STU_DIR)
        new_row = {'date': datetime.now().strftime('%Y-%m-%d')}
        for item in input_data:
            sub = item['subject']
            new_row[f'{sub}_score'] = item['score']
            new_row[f'{sub}_content'] = item['content']

        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        save_stu_data(STU_DIR, df)
        messagebox.showinfo('成功', '成绩数据提交成功！')

        df_new = load_stu_data(STU_DIR)
        draw_score_chart(df_new, canvas_chart)

        generate_report(df_new, STU_DIR)
        input_win.destroy()

    frame_btn_op = ttk.Frame(input_win)
    frame_btn_op.pack(pady=10)
    ttk.Button(frame_btn_op, text='清除', command=clear_input).grid(row=0, column=0, padx=20)
    ttk.Button(frame_btn_op, text='提交', command=submit_score, style='Blue.TButton').grid(row=0, column=1, padx=20)

    style = ttk.Style(input_win)
    style.configure('Green.TButton', foreground='green')
    style.configure('Red.TButton', foreground='red')
    style.configure('Blue.TButton', foreground='blue', font=('黑体', 10, 'bold'))


# 加载学生窗口
def load_stu_window():
    stu_dir = filedialog.askdirectory(title='选择学生目录', initialdir=DATA_DIR)
    if not stu_dir or 'tmp' in os.path.basename(stu_dir):
        return
    excel_path = os.path.join(stu_dir, 'data.xlsx')
    report_dir = os.path.join(stu_dir, 'reports')
    if not os.path.exists(excel_path) or not os.path.exists(report_dir):
        messagebox.showerror('错误', '选择的目录不是有效学生空间！')
        return
    global CURRENT_STU, STU_DIR
    CURRENT_STU = os.path.basename(stu_dir)
    STU_DIR = stu_dir
    load_avatar()
    try:
        df = load_stu_data(STU_DIR)
        draw_score_chart(df, canvas_chart)
        load_latest_report()
        messagebox.showinfo('成功', f'已加载{CURRENT_STU}的学生数据')
    except zipfile.BadZipFile:
        messagebox.showerror('加载失败', 'Excel文件损坏（不是有效的ZIP归档），请重建文件')
    except Exception as e:
        messagebox.showerror('加载失败', f'加载学生数据出错：{str(e)}')


# 加载头像
def load_avatar():
    avatar_frame_width = 240
    avatar_frame_height = 300

    if not STU_DIR:
        avatar_path = DEFAULT_AVATAR
    else:
        avatar_path = os.path.join(STU_DIR, 'pic/avatar.jpg')
        if not os.path.exists(avatar_path):
            avatar_path = DEFAULT_AVATAR

    try:
        img = Image.open(avatar_path)
        img = img.resize((avatar_frame_width, avatar_frame_height), Image.Resampling.LANCZOS)
        photo = ImageTk.PhotoImage(img)
        label_avatar.config(image=photo)
        label_avatar.img = photo
    except Exception as e:
        label_avatar.config(text='头像加载失败', image='')
        print(f'头像加载失败：{e}')

    label_stu_name.config(text=f'当前学生：{CURRENT_STU if CURRENT_STU else "未加载"}')


# 加载最新报告
def load_latest_report():
    if not STU_DIR:
        update_report_ui('未加载学生，暂无报告', False)
        return
    report_dir = os.path.join(STU_DIR, 'reports')
    if not os.path.exists(report_dir):
        update_report_ui('暂无分析报告', False)
        return
    report_files = [f for f in os.listdir(report_dir) if f.endswith('.md')]
    if not report_files:
        update_report_ui('暂无分析报告', False)
        return
    report_files.sort(reverse=True)
    latest_report = os.path.join(report_dir, report_files[0])
    try:
        with open(latest_report, 'r', encoding='utf-8') as f:
            content = f.read()
        update_report_ui(content, False)
    except:
        update_report_ui('报告文件读取失败', False)


# 展示报告内容（兼容新的UI更新函数）
def show_report(content):
    update_report_ui(content, False)


# 主UI初始化
def init_main_ui():
    global root, label_avatar, label_stu_name, canvas_chart, text_report

    root = tk.Tk()
    root.title('学生成绩统计与分析工具')
    root.geometry('1250x800')
    root.resizable(True, True)

    frame_row1 = ttk.Frame(root)
    frame_row1.pack(fill=tk.X, padx=10, pady=10)

    frame_avatar = ttk.LabelFrame(frame_row1, text='学生信息', width=240, height=400)
    frame_avatar.pack(side=tk.LEFT, fill=tk.NONE, padx=5)
    frame_avatar.pack_propagate(False)

    frame_avatar_img = ttk.Frame(frame_avatar, width=240, height=320)
    frame_avatar_img.pack(fill=tk.BOTH, expand=True)
    frame_avatar_img.pack_propagate(False)

    label_avatar = ttk.Label(frame_avatar_img, anchor='center')
    label_avatar.pack(fill=tk.BOTH, expand=True)

    label_stu_name = ttk.Label(frame_avatar, text='当前学生：未加载', font=('黑体', 10), wraplength=220)
    label_stu_name.pack(pady=5)

    frame_btn_stu = ttk.Frame(frame_avatar)
    frame_btn_stu.pack(pady=5, fill=tk.X, padx=10)
    btn_entry = ttk.Button(frame_btn_stu, text='录入', command=input_score_window, width=8)
    btn_entry.pack(side=tk.LEFT, expand=True, padx=5)
    btn_load = ttk.Button(frame_btn_stu, text='加载', command=load_stu_window, width=8)
    btn_load.pack(side=tk.RIGHT, expand=True, padx=5)

    frame_report = ttk.LabelFrame(frame_row1, text='分析报告', width=960, height=400)
    frame_report.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5)
    frame_report.pack_propagate(False)

    scroll_report = ttk.Scrollbar(frame_report)
    scroll_report.pack(side=tk.RIGHT, fill=tk.Y)
    # 全局text_report变量，供流式更新使用
    text_report = tk.Text(frame_report, yscrollcommand=scroll_report.set, font=('宋体', 11), wrap=tk.WORD)
    text_report.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
    scroll_report.config(command=text_report.yview)
    text_report.insert(tk.END, '未加载学生，暂无报告')
    text_report.config(state=tk.NORMAL)

    frame_row2 = ttk.Frame(root)
    frame_row2.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    frame_chart = ttk.LabelFrame(frame_row2, text='成绩趋势图', height=350)
    frame_chart.pack(fill=tk.BOTH, expand=True)
    frame_chart.pack_propagate(False)

    canvas_chart = ttk.Frame(frame_chart)
    canvas_chart.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
    tk.Label(canvas_chart, text='未加载成绩数据，暂无图表', font=('宋体', 12)).pack(pady=80)

    load_avatar()

    return root


# 主程序入口
if __name__ == '__main__':
    try:
        from PIL import Image, ImageTk
    except ImportError:
        import subprocess
        import sys

        subprocess.check_call([sys.executable, "-m", "pip", "install", "Pillow"])
        from PIL import Image, ImageTk

    init_dirs()
    root = init_main_ui()
    root.mainloop()