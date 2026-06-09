#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Digest2 GUI - 小行星轨道分类工具图形界面
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import sys
import webbrowser


def resource_path(relative_path: str) -> str:
    """Return an absolute path to a resource, supporting PyInstaller onefile mode."""
    base_path = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
    return os.path.join(base_path, relative_path)


# 添加虚拟环境路径（如果直接运行）
venv_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'venv', 'Lib', 'site-packages')
if os.path.exists(venv_path) and venv_path not in sys.path:
    sys.path.insert(0, venv_path)

def show_error_with_link(title, message, link_text=None, link_url=None):
    """显示带有可点击链接的错误对话框"""
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    
    dialog = tk.Toplevel(root)
    dialog.title(title)
    dialog.geometry("480x280")
    dialog.resizable(False, False)
    dialog.transient(root)
    dialog.grab_set()
    
    # 设置图标
    try:
        icon_img = tk.PhotoImage(file=resource_path('digest2_icon_256.png'))
        dialog.iconphoto(True, icon_img)
    except:
        pass
    
    frame = ttk.Frame(dialog, padding=20)
    frame.pack(fill=tk.BOTH, expand=True)
    
    # 错误图标
    icon_label = ttk.Label(frame, text="⚠️", font=('Arial', 32))
    icon_label.pack(pady=(0, 10))
    
    # 错误消息
    msg_label = ttk.Label(frame, text=message, justify=tk.CENTER, wraplength=400)
    msg_label.pack(pady=(0, 15))
    
    # 可点击链接
    if link_text and link_url:
        link_label = ttk.Label(frame, text=link_text, foreground='#1a73e8', cursor='hand2')
        link_label.pack(pady=(0, 15))
        link_label.bind('<Button-1>', lambda e: webbrowser.open(link_url))
    
    # 确定按钮
    ok_btn = ttk.Button(frame, text="确定", command=lambda: (root.destroy(), sys.exit(1)))
    ok_btn.pack(pady=(10, 0))
    
    # 居中显示
    dialog.update_idletasks()
    x = (dialog.winfo_screenwidth() - dialog.winfo_width()) // 2
    y = (dialog.winfo_screenheight() - dialog.winfo_height()) // 2
    dialog.geometry(f"+{x}+{y}")
    
    root.mainloop()


try:
    from digest2 import Digest2, parse_mpc80
except ImportError as e:
    error_msg = str(e)
    if "DLL load failed" in error_msg or "找不到指定的模块" in error_msg:
        show_error_with_link(
            "缺少运行时库",
            "程序无法启动，可能缺少 Visual C++ 运行时库。\n\n请点击下方链接下载并安装：",
            "👉 https://aka.ms/vs/17/release/vc_redist.x64.exe",
            "https://aka.ms/vs/17/release/vc_redist.x64.exe"
        )
    else:
        messagebox.showerror(
            "错误",
            "找不到 digest2 模块。\n\n"
            "请确保已安装 digest2：\n"
            "1. 运行：venv\\Scripts\\pip install digest2\n"
            "2. 或使用虚拟环境运行此程序"
        )
        sys.exit(1)

# 尝试导入 python-docx（用于读取 Word 文档）
try:
    from docx import Document
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False


class Digest2GUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Asterorbit 小行星轨道分类评估工具")
        self.root.minsize(800, 600)
        self.set_window_icon()
        
        # 设置样式
        self.style = ttk.Style()
        self.style.configure('TButton', font=('微软雅黑', 10))
        self.style.configure('TLabel', font=('微软雅黑', 10))
        self.style.configure('Header.TLabel', font=('微软雅黑', 12, 'bold'))
        
        # 用于存储结果和天体名称的映射（因为 ClassificationResult 是 frozen 的）
        self._desig_map = {}
        
        self.create_widgets()
        
    def create_widgets(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(0, weight=1)
        
        # 设置窗口大小
        self.root.geometry("1200x800")
        
        # 创建标签页控件
        # 创建Notebook并设置统一标签宽度和居中对齐
        style = ttk.Style()
        style.configure('TNotebook.Tab', width=10, anchor='center')
        
        # 分析结果表格样式（表头微软雅黑加粗，内容宋体）
        style.configure('ResultTreeview.Treeview',
                       font=('宋体', 10))
        style.configure('ResultTreeview.Treeview.Heading', 
                       background='#f5f5f5', 
                       foreground='#333333',
                       font=('微软雅黑', 10, 'bold'))
        style.map('ResultTreeview.Treeview',
                  foreground=[('selected', '#ffffff')],
                  background=[('selected', '#0078d4'), ('active', '#f0f0f0')])
        
        # 类型说明表格样式（更大字体）
        style.configure('DescTreeview.Treeview',
                       font=('黑体', 11),
                       rowheight=30)
        style.configure('DescTreeview.Treeview.Heading', 
                       background='#f5f5f5', 
                       foreground='#333333',
                       font=('微软雅黑', 11, 'bold'))
        style.map('DescTreeview.Treeview',
                  foreground=[('selected', '#ffffff')],
                  background=[('selected', '#0078d4'), ('active', '#f0f0f0')])
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 创建类型评估页面
        self.create_evaluation_tab()
        
        # 创建类型说明页面
        self.create_description_tab()
        
        # 创建关于页面
        self.create_about_tab()

    def set_window_icon(self):
        try:
            icon_path = resource_path('digest2_icon_256.png')
            icon_img = tk.PhotoImage(file=icon_path)
            self.root.iconphoto(True, icon_img)
        except Exception:
            pass
    
    def create_evaluation_tab(self):
        """创建类型评估页面"""
        eval_frame = ttk.Frame(self.notebook)
        self.notebook.add(eval_frame, text='类型评估')
        
        # 配置网格权重
        eval_frame.columnconfigure(0, weight=1)
        eval_frame.rowconfigure(2, weight=1)  # 输入框区域可扩展
        eval_frame.rowconfigure(4, weight=1)  # 结果显示区域可扩展
        
        # 说明文字
        desc_label = ttk.Label(
            eval_frame,
            text="选择文件或输入 MPC 80列、ADES 格式的观测数据，程序将自动分析并评估轨道类型。",
            wraplength=800
        )
        desc_label.grid(row=0, column=0, columnspan=3, pady=(0, 10), sticky=tk.W)
        
        # 文件选择区域
        file_frame = ttk.LabelFrame(eval_frame, text="文件选择", padding="10")
        file_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)
        
        ttk.Label(file_frame, text="观测文件：").grid(row=0, column=0, sticky=tk.W)
        
        self.file_path_var = tk.StringVar()
        self.file_entry = ttk.Entry(file_frame, textvariable=self.file_path_var)
        self.file_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5)
        
        self.browse_btn = ttk.Button(file_frame, text="浏览...", command=self.browse_file)
        self.browse_btn.grid(row=0, column=2, padx=5)
        
        # 手动输入区域
        input_frame = ttk.LabelFrame(eval_frame, text="观测数据", padding="10")
        input_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        input_frame.columnconfigure(0, weight=1)
        input_frame.rowconfigure(1, weight=1)
        
        ttk.Label(input_frame, text="输入 MPC 80列、ADES 格式的观测数据：").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        
        self.input_text = scrolledtext.ScrolledText(
            input_frame,
            wrap=tk.NONE,
            width=80,
            height=8,
            font=('宋体', 11)
        )
        self.input_text.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        # 为观测数据输入框添加右键菜单
        self.input_text.bind('<Button-3>', self.show_text_context_menu)
        
        # 按钮区域
        btn_frame = ttk.Frame(eval_frame)
        btn_frame.grid(row=4, column=0, columnspan=3, pady=10)
        
        self.analyze_btn = ttk.Button(
            btn_frame,
            text="开始分析",
            command=self.analyze,
            width=20
        )
        self.analyze_btn.pack(side=tk.LEFT, padx=5)

        self.clear_data_btn = ttk.Button(
            btn_frame,
            text="清空观测数据",
            command=self.clear_input_data,
            width=20
        )
        self.clear_data_btn.pack(side=tk.LEFT, padx=5)
        
        self.download_obscode_btn = ttk.Button(
            btn_frame,
            text="下载MPC天文台站代码",
            command=self.download_obscode,
            width=20
        )
        self.download_obscode_btn.pack(side=tk.LEFT, padx=5)
        
        # 结果显示区域（row=4）
        result_frame = ttk.LabelFrame(eval_frame, text="分析结果", padding="10")
        result_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        result_frame.columnconfigure(0, weight=1)
        result_frame.rowconfigure(0, weight=1)
        
        # 创建树形视图显示结果
        columns = ('designation', 'rms', 'int', 'neo', 'n22', 'n18', 'mc', 'mb1', 'mb2', 'mb3', 'hun', 'pho', 'pal', 'han', 'hil', 'jtr', 'jfc')
        self.tree = ttk.Treeview(
            result_frame,
            columns=columns,
            show='headings',
            height=10,
            style='ResultTreeview.Treeview',
            selectmode='extended'
        )
        self.result_drag_start = None  # 分析结果表格拖拽起始行
        # 为树形视图添加右键菜单（支持复制）
        self.tree.bind('<Button-3>', self.show_tree_context_menu)
        # 绑定拖拽选择事件
        self.tree.bind('<Button-1>', self.on_result_tree_press)
        self.tree.bind('<B1-Motion>', self.on_result_tree_drag)

        # 配置标签样式：NEO高亮（黄色背景）和加粗
        self.tree.tag_configure('neo_highlight', background='#FFD700')
        self.tree.tag_configure('neo_bold', font=('宋体', 10, 'bold'))
        
        # 设置列标题
        self.tree.heading('designation', text='天体名称')
        self.tree.heading('rms', text='RMS')
        self.tree.heading('int', text='Int')
        self.tree.heading('neo', text='NEO')
        self.tree.heading('n22', text='N22')
        self.tree.heading('n18', text='N18')
        self.tree.heading('mc', text='MC')
        self.tree.heading('mb1', text='MB1')
        self.tree.heading('mb2', text='MB2')
        self.tree.heading('mb3', text='MB3')
        self.tree.heading('hun', text='Hun')
        self.tree.heading('pho', text='Pho')
        self.tree.heading('pal', text='Pal')
        self.tree.heading('han', text='Han')
        self.tree.heading('hil', text='Hil')
        self.tree.heading('jtr', text='JTr')
        self.tree.heading('jfc', text='JFC')
        
        # 设置列宽（为小数显示增加宽度）
        self.tree.column('designation', width=120, minwidth=100)
        self.tree.column('rms', width=60, minwidth=50, anchor='center')
        self.tree.column('int', width=55, minwidth=45, anchor='center')
        for col in columns[3:]:
            self.tree.column(col, width=55, minwidth=45, anchor='center')
        
        # 滚动条
        scrollbar_y = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar_x = ttk.Scrollbar(result_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar_y.grid(row=0, column=1, sticky=(tk.N, tk.S))
        scrollbar_x.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # 状态栏（row=6）
        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        status_bar = ttk.Label(eval_frame, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(5, 0))
    
    def create_description_tab(self):
        """创建类型说明页面"""
        desc_frame = ttk.Frame(self.notebook)
        self.notebook.add(desc_frame, text='类型说明')
        
        # 创建主框架
        main_frame = ttk.Frame(desc_frame)
        main_frame.pack(padx=15, pady=15, fill=tk.BOTH, expand=True)
        
        # 添加标题
        title_label = ttk.Label(main_frame, text="轨道类型说明", font=('微软雅黑', 16, 'bold'), foreground='#333')
        title_label.pack(anchor=tk.W, pady=(0, 8))
        
        # 添加绿色下划线（与关于页面一致）
        title_line = ttk.Frame(main_frame, height=2, style='TitleLine.TFrame')
        title_line.pack(anchor=tk.W, fill=tk.X, pady=(0, 15))
        
        # 创建Treeview容器
        tree_frame = ttk.Frame(main_frame, relief='solid', borderwidth=1)
        tree_frame.pack(fill=tk.X, expand=False)
        
        # 创建Treeview组件
        columns = ('abbrev', 'fullname', 'chinese', 'definition')
        self.desc_tree = ttk.Treeview(
            tree_frame,
            columns=columns,
            show='headings',
            height=15,
            style='DescTreeview.Treeview',
            selectmode='extended'
        )
        self.drag_start = None  # 拖拽起始行
        
        # 移除交替背景色
        
        # 设置列标题
        self.desc_tree.heading('abbrev', text='缩写')
        self.desc_tree.heading('fullname', text='英文全称')
        self.desc_tree.heading('chinese', text='中文含义')
        self.desc_tree.heading('definition', text='Digest2定义')
        
        # 设置列宽
        self.desc_tree.column('abbrev', width=30, minwidth=30)
        self.desc_tree.column('fullname', width=180, minwidth=180)
        self.desc_tree.column('chinese', width=120, minwidth=100)
        self.desc_tree.column('definition', width=700, minwidth=300)
        
        # 轨道类型说明（使用 Digest2 官方定义及标准天文学参数）
        orbit_types = [
            ('Int', 'MPC Interesting', 'MPC关注天体', '满足以下任一条件：𝑞 < 1.3 AU，𝑄 > 10 AU，𝑒 ≥ 0.5，𝑖 ≥ 40°'),
            ('NEO', 'Near-Earth Object', '近地天体', '𝑞 < 1.3 AU'),
            ('N22', 'Intermediate-size NEO', '中等大小近地天体', '𝐷 > 140 m。𝑞 < 1.3 AU，𝐻 < 22.5'),
            ('N18', 'Large Near-Earth Object', '大型近地天体', '𝐷 > 1.2 km。𝑞 < 1.3 AU，𝐻 < 18.5'),
            ('MC', 'Mars Crosser', '越火小天体', '穿越火星轨道。1.3 AU < 𝑞 < 1.67 AU，𝑄 > 1.58 AU'),
            ('Hun', 'Hungaria Group', '匈牙利群', '以434号小行星匈牙利星为代表的小行星群。1.78 AU < 𝑎 < 2 AU，𝑒 < 0.18，16° < 𝑖 < 34°'),
            ('Pho', 'Phocaea Group', '福后星群', '以25号小行星福后星为代表的小行星群。𝑞 > 1.5 AU，2.2 AU < 𝑎 < 2.45 AU，20° < 𝑖 < 27°'),
            ('MB1', 'Inner Main Belt', '内主带', '𝑞 > 1.67 AU，2.1 AU < 𝑎 < 2.5 AU，𝑖 < 25a−45.5'),
            ('Pal', 'Pallas Group', '智神星群', '以2号小行星智神星为代表的小行星群。2.5 AU < 𝑎 < 2.8 AU，𝑒 < 0.35，24° < 𝑖 < 37°'),
            ('Han', 'Hansa Group', '汉萨群', '以480号小行星汉萨星为代表的小行星群。2.55 AU < 𝑎 < 2.72 AU，𝑒 < 0.25，20° < 𝑖 < 23.5°'),
            ('MB2', 'Middle Main Belt', '中主带', '2.5 AU < 𝑎 < 2.8 AU，𝑒 < 0.45，𝑖 < 20°'),
            ('MB3', 'Outer Main Belt', '外主带', '2.8 AU < 𝑎 < 3.25 AU，𝑒 < 0.4，𝑖 < (320𝑎-716)/9'),
            ('Hil', 'Hilda Group', '希尔达群', '以153号小行星希尔达星为代表的小行星群。3.9 AU < 𝑎 < 4.02 AU，𝑒 < 0.4，𝑖 < 18°'),
            ('JTr', 'Jupiter Trojan', '木星特洛伊群', '位于木星 L4、L5 拉格朗日点的小行星群，5.05 AU < 𝑎 < 5.35 AU，𝑒 < 0.22，𝑖 < 38°'),
            ('JFC', 'Jupiter Family Comet', '木星族彗星', '𝑞 > 1.3 AU，2 < 𝑇𝐽 < 3'),
        ]
        
        # 添加数据行
        for i, (abbrev, fullname, chinese, definition) in enumerate(orbit_types):
            self.desc_tree.insert('', tk.END, values=(abbrev, fullname, chinese, definition))
        
        # 滚动条
        scrollbar_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.desc_tree.yview)
        scrollbar_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.desc_tree.xview)
        self.desc_tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        
        self.desc_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar_y.grid(row=0, column=1, sticky=(tk.N, tk.S))
        scrollbar_x.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # 绑定拖拽选择事件
        self.desc_tree.bind('<Button-1>', self.on_desc_tree_press)
        self.desc_tree.bind('<B1-Motion>', self.on_desc_tree_drag)
        # 绑定右键菜单到类型说明表格
        self.desc_tree.bind('<Button-3>', self.show_desc_tree_context_menu)
        
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)
        
        # 添加参数说明区域
        legend_frame = ttk.Frame(main_frame)
        legend_frame.pack(fill=tk.X, pady=(10, 0))
        
        legend_title = ttk.Label(legend_frame, text='参数说明：', 
                                font=('微软雅黑', 10, 'bold'))
        legend_title.pack(anchor=tk.W, pady=(0, 5))
        
        # 四列布局
        legend_columns = ttk.Frame(legend_frame)
        legend_columns.pack(fill=tk.X)
        
        all_params = [
            ('𝑞', '近日距'),
            ('𝑄', '远日距'),
            ('𝑎', '半长轴'),
            ('𝑒', '轨道偏心率'),
            ('𝑖', '轨道倾角'),
            ('𝐻', '绝对星等'),
            ('𝑇𝐽', '相对于木星的蒂塞朗参数'),
            ('𝐷', '直径'),
        ]
        
        # 创建4列，每列2个
        for i in range(4):
            col_frame = ttk.Frame(legend_columns)
            col_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0 if i == 0 else 10, 0))
            
            # 每列添加2个参数
            for j in range(2):
                idx = i * 2 + j
                if idx < len(all_params):
                    symbol, meaning = all_params[idx]
                    row = ttk.Frame(col_frame)
                    row.pack(fill=tk.X, pady=1)
                    ttk.Label(row, text=symbol, font=('Segoe UI', 10), width=3).pack(side=tk.LEFT)
                    ttk.Label(row, text='= ' + meaning, font=('微软雅黑', 10)).pack(side=tk.LEFT)
    
    def insert_definition_with_italic(self, text_widget, definition, bg_tag):
        """插入定义文本，使用数学斜体符号"""
        import re
        # 字母替换为数学斜体符号 - 先处理TJ，再处理单个字母
        italic_map = {
            'TJ': '𝑇𝐽',
            'q': '𝑞', 'a': '𝑎', 'i': '𝑖', 'e': '𝑒',
            'Q': '𝑄', 'H': '𝐻', 'D': '𝐷', 'T': '𝑇', 'J': '𝐽'
        }
        
        result = definition
        for normal, italic in italic_map.items():
            # 使用正则表达式只替换单独出现的字母（单词边界）
            result = re.sub(r'\b' + re.escape(normal) + r'\b', italic, result)
        
        text_widget.insert(tk.END, result, ('definition', bg_tag))
    
    def create_about_tab(self):
        """创建关于页面"""
        about_frame = ttk.Frame(self.notebook, borderwidth=0)
        self.notebook.add(about_frame, text='关于')
        
        # 创建内容容器 - 简化布局
        content_frame = ttk.Frame(about_frame, padding=(15, 15))
        content_frame.pack(fill=tk.BOTH, expand=True)
        content_frame.configure(style='Transparent.TFrame')
        
        # 设置背景色
        about_frame.configure(style='Transparent.TFrame')
        
        # 添加标题
        title_label = ttk.Label(content_frame, text="关于", font=('微软雅黑', 14, 'bold'), foreground='#333')
        title_label.pack(anchor=tk.W, pady=(0, 8))
        
        # 添加绿色下划线（模拟网页版样式）
        title_line = ttk.Frame(content_frame, height=2, style='TitleLine.TFrame')
        title_line.pack(anchor=tk.W, fill=tk.X, pady=(0, 15))
        
        # 添加声明信息（合并版权和声明）- 使用Text控件确保正确换行
        about_text_widget = tk.Text(content_frame, wrap=tk.WORD, height=2, 
                                  font=('微软雅黑', 11), bg='#f5f5f5', relief='flat',
                                  foreground='#333', spacing1=3, spacing2=2, spacing3=3,
                                  borderwidth=0, highlightthickness=0,
                                  insertwidth=0)
        about_text_widget.pack(fill=tk.X, anchor=tk.W, pady=(0, 15))
        about_text_widget.insert(tk.END, "本应用由赵经远基于小行星中心（MPC）官方开源的Digest2源代码构建，非MPC官方项目。Digest2源代码作者：Sonia Keys、Carl Hergenrother、Robert McNaught、David Asher，源代码中的ADES支持由 Richard Cloete 和 Peter Vereš添加。")
        
        # 为about_text_widget添加右键菜单
        about_text_widget.bind('<Button-3>', lambda e: self.show_about_context_menu(e, about_text_widget))
        
        # 添加格式、网页版、版本信息（合并到一个Text控件）
        import digest2
        import importlib.metadata as metadata
        digest2_version = '未知版本'
        try:
            digest2_version = metadata.version('digest2')
        except metadata.PackageNotFoundError:
            digest2_version = getattr(digest2, '__version__', getattr(digest2, 'VERSION', '未知版本'))
        
        info_text = tk.Text(content_frame, wrap=tk.WORD, height=4, 
                            font=('微软雅黑', 10), bg='#f5f5f5', relief='flat',
                            foreground='#666', spacing1=0, spacing2=10, spacing3=0,
                            borderwidth=0, highlightthickness=0,
                            insertwidth=0)
        info_text.pack(anchor=tk.W, pady=(0, 20))
        info_text.insert(tk.END, "支持的数据格式：MPC 80列、ADES XML、ADES PSV\n")
        info_text.insert(tk.END, "支持的文件格式：obs、txt、dat、docx、doc、xml、psv\n")
        info_text.insert(tk.END, "网页版：")
        info_text.insert(tk.END, "https://asterorbit-digest2.hf.space/", 'link')
        info_text.insert(tk.END, f"\nDigest2 版本：{digest2_version}")
        info_text.tag_config('link', foreground='#1a73e8', underline=True)
        info_text.bind('<Button-3>', lambda e: self.show_about_context_menu(e, info_text))
        
        def open_web_link(event):
            webbrowser.open("https://asterorbit-digest2.hf.space/")
        info_text.tag_bind('link', '<Button-1>', open_web_link)
        
        # 鼠标悬停时显示手型光标（链接上）或I型光标（文字上）
        def on_info_mouse_move(event):
            tags = info_text.tag_names(tk.CURRENT)
            if 'link' in tags:
                info_text.config(cursor='hand2')
            else:
                info_text.config(cursor='xterm')
        info_text.bind('<Motion>', on_info_mouse_move)
        
        # 添加参考资料标题
        ref_title_label = ttk.Label(content_frame, text="参考资料", font=('微软雅黑', 11, 'bold'), foreground='#333')
        ref_title_label.pack(anchor=tk.W, pady=(15, 12))
        
        # 创建参考资料容器（带灰色边框）
        ref_container = ttk.Frame(content_frame, relief='solid', borderwidth=1, style='RefContainer.TFrame')
        ref_container.pack(fill=tk.X, padx=10, pady=10)
        
        # 使用Text控件显示参考资料（支持超链接）
        ref_text_widget = tk.Text(ref_container, wrap=tk.WORD, height=9, 
                                  font=('微软雅黑', 10), bg='#f5f5f5', relief='flat',
                                  foreground='#444', spacing1=5, spacing2=5, spacing3=10,
                                  borderwidth=0, highlightthickness=0,
                                  cursor='xterm',
                                  insertwidth=0)
        ref_text_widget.pack(fill=tk.X, padx=10, pady=10)
        
        # 添加参考资料内容（带超链接）
        ref_text_widget.insert(tk.END, "[1] KEYS S, VEREŠ P, PAYNE M J, et al. The digest2 NEO Classification Code[J]. Publications of the Astronomical Society of the Pacific, 2019, 131(1000): 1-18.\n")
        ref_text_widget.insert(tk.END, "[2] VEREŠ P, CLOETE R, WERYK R, et al. Improvement of Digest2 NEO Classification Code—utilizing the Astrometry Data Exchange Standard[J]. Publications of the Astronomical Society of the Pacific, 2023, 135(1052).\n")
        ref_text_widget.insert(tk.END, "[3] VEREŠ P, CLOETE R, PAYNE M J, et al. Improving the discovery of near-Earth objects with machine-learning methods[J]. Astronomy & Astrophysics, 2025, 698: A242.\n")
        ref_text_widget.insert(tk.END, "[4] M. P. C. Staff. EDITORIAL NOTICE: New digest2 repository and package: M.P.E.C. 2026-E23[EB/OL]. Cambridge, MA: Minor Planet Center, 2026-03-05. ")
        ref_text_widget.insert(tk.END, "https://www.minorplanetcenter.net/mpec/K26/K26E23.html", 'link')
        ref_text_widget.insert(tk.END, "\n")
        ref_text_widget.insert(tk.END, "[5] M. P. C. Staff. EDITORIAL NOTICE: digest2 population model update: M.P.E.C. 2026-K125[EB/OL]. Cambridge, MA: Minor Planet Center, 2026-05-29. ")
        ref_text_widget.insert(tk.END, "https://www.minorplanetcenter.net/mpec/K26/K26KC5.html", 'link')
        ref_text_widget.insert(tk.END, "\n")
        ref_text_widget.insert(tk.END, "[6] GitHub. Digest2 source code[EB/OL]. ")
        ref_text_widget.insert(tk.END, "https://github.com/Smithsonian/mpc-public/tree/main/digest2", 'link')
        ref_text_widget.insert(tk.END, "\n")
        ref_text_widget.insert(tk.END, "[7] MPC Public Documentation Hub. Digest2 tutorials[EB/OL]. ")
        ref_text_widget.insert(tk.END, "https://docs.minorplanetcenter.net/tutorials/iod_tutorials/?h=digest2", 'link')
        
        # 配置超链接样式
        ref_text_widget.tag_config('link', foreground='#1a73e8', underline=True)
        
        # 配置选中超链接时的样式（白色文字，蓝色背景，保持下划线）
        ref_text_widget.tag_config('link_sel', foreground='#ffffff', background='#1a73e8', underline=True)
        
        # 配置默认选中样式（不加下划线）
        ref_text_widget.tag_config('sel', foreground='white', background='#0078d4', underline=False)
        
        # 超链接点击事件处理
        def open_link(event):
            text = ref_text_widget.get('current linestart', 'current lineend')
            import re
            urls = re.findall(r'https?://\S+', text)
            if urls:
                webbrowser.open(urls[0])
        
        # 鼠标悬停时显示手型光标（链接上）或I型光标（文字上）
        def on_mouse_move(event):
            tags = ref_text_widget.tag_names(tk.CURRENT)
            if 'link' in tags:
                ref_text_widget.config(cursor='hand2')
            else:
                ref_text_widget.config(cursor='xterm')
        
        # 选中内容变化时更新链接样式
        def on_select(event):
            # 清除之前的选中样式
            ref_text_widget.tag_remove('link_sel', '1.0', tk.END)
            # 获取选中范围
            try:
                sel_start = ref_text_widget.index(tk.SEL_FIRST)
                sel_end = ref_text_widget.index(tk.SEL_LAST)
                # 只为选中区域内的链接添加选中样式
                # 遍历所有链接范围，只选中与选择区域重叠的部分
                for link_start, link_end in zip(
                    ref_text_widget.tag_ranges('link')[0::2],
                    ref_text_widget.tag_ranges('link')[1::2]
                ):
                    # 检查链接范围是否与选择区域重叠
                    if ref_text_widget.compare(link_start, '<', sel_end) and \
                       ref_text_widget.compare(link_end, '>', sel_start):
                        ref_text_widget.tag_add('link_sel', link_start, link_end)
            except tk.TclError:
                pass  # 没有选中的文本
        
        ref_text_widget.tag_bind('link', '<Button-1>', open_link)
        ref_text_widget.bind('<Motion>', on_mouse_move)
        ref_text_widget.bind('<<Selection>>', on_select)
        
        # 为ref_text_widget添加右键菜单
        ref_text_widget.bind('<Button-3>', lambda e: self.show_about_context_menu(e, ref_text_widget))
        
        # 配置样式
        style = ttk.Style()
        style.configure('TitleLine.TFrame', background='#4CAF50')
        style.configure('Transparent.TFrame', background='#f5f5f5')
        style.configure('Transparent.TLabel', background='#f5f5f5')
        style.configure('RefContainer.TFrame', background='#f5f5f5', bordercolor='#e0e0e0')
        
        # 点击关于页面任意位置时清除文本选中状态
        def clear_text_selection(event):
            about_text_widget.tag_remove(tk.SEL, '1.0', tk.END)
            info_text.tag_remove(tk.SEL, '1.0', tk.END)
            ref_text_widget.tag_remove(tk.SEL, '1.0', tk.END)
            ref_text_widget.tag_remove('link_sel', '1.0', tk.END)
        
        about_frame.bind('<Button-1>', clear_text_selection)
        content_frame.bind('<Button-1>', clear_text_selection)
    
    def on_result_tree_press(self, event):
        """处理分析结果表格的鼠标按下事件，记录拖拽起始位置"""
        item = self.tree.identify_row(event.y)
        if item:
            # 如果点击已选中的行，则取消选中
            if item in self.tree.selection():
                self.tree.selection_remove(item)
                self.result_drag_start = None
            else:
                self.result_drag_start = item
                # 选中起始行
                self.tree.selection_set(item)
        else:
            self.result_drag_start = None
    
    def on_result_tree_drag(self, event):
        """处理分析结果表格的鼠标拖拽事件，实现拖拽多选"""
        if self.result_drag_start is None:
            return
        
        current_item = self.tree.identify_row(event.y)
        if current_item:
            # 获取所有行
            all_items = self.tree.get_children()
            start_index = all_items.index(self.result_drag_start)
            current_index = all_items.index(current_item)
            
            # 确定选择范围
            min_index = min(start_index, current_index)
            max_index = max(start_index, current_index)
            
            # 选中范围内的所有行
            self.tree.selection_set(all_items[min_index:max_index+1])
    
    def on_desc_tree_press(self, event):
        """处理类型说明表格的鼠标按下事件，记录拖拽起始位置"""
        item = self.desc_tree.identify_row(event.y)
        if item:
            # 如果点击已选中的行，则取消选中
            if item in self.desc_tree.selection():
                self.desc_tree.selection_remove(item)
                self.drag_start = None
            else:
                self.drag_start = item
                # 选中起始行
                self.desc_tree.selection_set(item)
        else:
            self.drag_start = None
    
    def on_desc_tree_drag(self, event):
        """处理类型说明表格的鼠标拖拽事件，实现拖拽多选"""
        if self.drag_start is None:
            return
        
        current_item = self.desc_tree.identify_row(event.y)
        if current_item:
            # 获取所有行
            all_items = self.desc_tree.get_children()
            start_index = all_items.index(self.drag_start)
            current_index = all_items.index(current_item)
            
            # 确定选择范围
            min_index = min(start_index, current_index)
            max_index = max(start_index, current_index)
            
            # 选中范围内的所有行
            self.desc_tree.selection_set(all_items[min_index:max_index+1])
    
    def show_desc_tree_context_menu(self, event):
        """显示类型说明表格的右键菜单"""
        # 获取鼠标所在的行
        row_id = self.desc_tree.identify_row(event.y)
        if row_id:
            # 如果没有选中行，先选中鼠标所在的行
            selected_items = self.desc_tree.selection()
            if not selected_items or row_id not in selected_items:
                self.desc_tree.selection_set(row_id)
        
        # 创建右键菜单
        context_menu = tk.Menu(self.root, tearoff=0)
        # 只有当有选中行时才显示"复制选中行"
        if self.desc_tree.selection():
            context_menu.add_command(label="复制选中行", command=self.copy_desc_tree_selected_item)
        # 总是显示"复制全部数据"
        context_menu.add_command(label="复制全部数据", command=self.copy_desc_tree_all_data)
        
        # 在鼠标位置显示菜单
        context_menu.tk_popup(event.x_root, event.y_root)
    
    def copy_desc_tree_selected_item(self):
        """复制类型说明表格选中行的数据到剪贴板"""
        selected_items = self.desc_tree.selection()
        if not selected_items:
            return
        
        # 获取所有选中行的数据
        rows = []
        for item in selected_items:
            values = self.desc_tree.item(item, 'values')
            rows.append('\t'.join(str(v) for v in values))
        
        # 格式化数据（制表符分隔，便于粘贴到表格）
        text = '\n'.join(rows)
        
        # 复制到剪贴板
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
    
    def copy_desc_tree_all_data(self):
        """复制类型说明表格所有数据到剪贴板"""
        # 获取所有行数据
        rows = []
        # 添加表头
        headers = ['缩写', '英文全称', '中文含义', 'Digest2定义']
        rows.append('\t'.join(headers))
        
        # 添加数据行
        for item in self.desc_tree.get_children():
            values = self.desc_tree.item(item, 'values')
            rows.append('\t'.join(str(v) for v in values))
        
        # 合并为文本
        text = '\n'.join(rows)
        
        # 复制到剪贴板
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
    
    def browse_file(self):
        filetypes = [
            ('观测文件', '*.obs *.xml *.psv *.txt *.dat'),
            ('MPC 80列格式', '*.obs'),
            ('ADES XML格式', '*.xml'),
            ('ADES PSV格式', '*.psv'),
            ('文本文件', '*.txt *.dat'),
            ('所有文件', '*.*')
        ]
        # 如果支持 docx，添加到文件类型
        if HAS_DOCX:
            filetypes.insert(1, ('Word 文档', '*.docx *.doc'))
            filetypes[0] = ('观测文件', '*.obs *.xml *.psv *.txt *.dat *.docx *.doc')

        filename = filedialog.askopenfilename(
            title="选择观测文件",
            filetypes=filetypes
        )
        if filename:
            self.file_path_var.set(filename)
            # 根据文件类型自动处理
            self._load_file_content(filename)

    def _load_file_content(self, filename):
        """根据文件类型加载内容到输入框"""
        ext = filename.lower().split('.')[-1] if '.' in filename else ''

        try:
            if ext in ['docx', 'doc'] and HAS_DOCX:
                # Word 文档
                self._load_docx_content(filename)
            elif ext in ['txt', 'dat', 'obs']:
                # 纯文本文件（包括 MPC80 格式的 .obs 文件）
                self._load_text_content(filename)
            elif ext == 'xml':
                # ADES XML 文件 - 解析并显示为可读格式
                self._load_ades_xml_content(filename)
            elif ext == 'psv':
                # ADES PSV 文件 - 解析并显示为可读格式
                self._load_ades_psv_content(filename)
            # 其他格式不自动加载
        except Exception as e:
            messagebox.showerror("错误", f"无法读取文件：\n{str(e)}")

    def _load_text_content(self, filename):
        """从文本文件中读取内容，自动检测格式（MPC80、ADES XML 或 ADES PSV）"""
        try:
            with open(filename, 'r', encoding='utf-8') as f:
                content = f.read()
        except UnicodeDecodeError:
            with open(filename, 'r', encoding='gbk') as f:
                content = f.read()

        # 检测是否为 ADES XML 格式
        if content.strip().startswith('<?xml') or content.strip().startswith('<ades>'):
            # 是 XML 格式，调用 XML 加载方法
            self._load_ades_xml_content(filename)
            return

        # 检测是否为 ADES PSV 格式
        lines = content.strip().split('\n')
        non_empty_lines = [l for l in lines if l.strip()]
        
        is_psv_format = False
        has_psv_header = False
        has_data_with_pipes = False
        
        for line in non_empty_lines:
            # 跳过注释行
            if line.startswith('#'):
                continue
            
            # 检查是否是 PSV 表头行（以 ! 开头且包含 |）
            if line.startswith('!') and '|' in line:
                has_psv_header = True
                break
            
            # 检查是否是数据行（包含多个 |）
            if line.count('|') >= 2:
                has_data_with_pipes = True
                # 额外检查：| 不是仅在第13列位置（MPC80 格式的项目代码位置）
                if len(line) > 12 and line[12] == '|':
                    # 检查是否有其他位置的 |
                    if line.count('|') > 1:
                        has_data_with_pipes = True
                    else:
                        has_data_with_pipes = False
                break
        
        is_psv_format = has_psv_header or has_data_with_pipes
        
        if is_psv_format:
            # 是 PSV 格式，调用 PSV 加载方法
            self._load_ades_psv_content(filename)
            return

        # 解析并过滤 MPC80 数据
        filtered_lines = self._filter_mpc80_data(content)

        self.input_text.delete(1.0, tk.END)
        self.input_text.insert(1.0, '\n'.join(filtered_lines))
        self.status_var.set(f"已加载 {len(filtered_lines)} 行有效 MPC80 数据（至少2条观测的天体）")

    def _filter_mpc80_data(self, content):
        """过滤 MPC80 数据，只保留至少有两条观测的天体"""
        lines = content.split('\n')

        # 收集所有行的天体名称（去重：相同数据只保留一条）
        tracklets = {}
        seen_lines = set()  # 用于去重

        for line in lines:
            line = line.rstrip()
            if not line:
                continue

            # 尝试解析 MPC80 格式获取天体名称
            desig = self._try_extract_designation(line)
            if desig:
                # 去重检查
                if line not in seen_lines:
                    seen_lines.add(line)
                    if desig not in tracklets:
                        tracklets[desig] = []
                    tracklets[desig].append(line)

        # 只保留至少有两条观测的天体
        filtered = []
        for desig, obs_lines in tracklets.items():
            if len(obs_lines) >= 2:
                filtered.extend(obs_lines)

        return filtered

    def _try_extract_designation(self, line):
        """尝试从行中提取天体名称，返回 None 如果不是 MPC80 格式"""
        # MPC80 格式：天体名称在前 12 列
        if len(line) < 12:
            return None

        # 提取前 12 列作为天体名称
        desig = line[:12].strip()

        # 检查是否看起来像 MPC80 格式
        if len(line) >= 80:
            # 标准 MPC80 格式：检查第14列或第15列是否有观测类型（C/S/B）
            note2 = line[14] if len(line) > 14 else ' '
            # 检查第14列或第15列是否包含观测类型
            has_obs_type = note2 in ('C', 'S', 'B') or \
                          (len(line) > 15 and line[15] in ('C', 'S', 'B')) or \
                          (len(line) > 13 and line[13] in ('C', 'S', 'B'))
            
            # 检查第13-14列是否有项目代码
            has_prog_code = line[12:14].strip()
            
            if has_obs_type or has_prog_code:
                return desig if desig else None

        # 尝试简化格式（空格分隔）
        parts = line.split()
        if len(parts) >= 2:
            # 检查第二个字段是否包含年份（如 C2019, 4C2019）
            import re
            if re.search(r'C\d{4}', parts[1]):
                return parts[0]

        return None

    def _load_docx_content(self, filename):
        """从 Word 文档中提取文本内容，自动检测格式（MPC80、ADES XML 或 ADES PSV）"""
        doc = Document(filename)
        text_lines = []
        for para in doc.paragraphs:
            if para.text.strip():
                text_lines.append(para.text)

        # 将内容合并
        content = '\n'.join(text_lines)

        # 检测是否为 ADES XML 格式
        if content.strip().startswith('<?xml') or content.strip().startswith('<ades>'):
            # 是 XML 格式，保存到临时文件并调用 XML 加载方法
            import tempfile
            import os
            try:
                with tempfile.NamedTemporaryFile(mode='w', suffix='.xml', delete=False, encoding='utf-8') as f:
                    f.write(content)
                    temp_path = f.name
                self._load_ades_xml_content(temp_path)
                os.unlink(temp_path)
                return
            except Exception as e:
                messagebox.showerror("错误", f"无法解析 Word 文档中的 XML 数据：\n{str(e)}")
                return

        # 检测是否为 ADES PSV 格式（包含 | 分隔符）
        lines = content.strip().split('\n')
        non_empty_lines = [l for l in lines if l.strip()]
        if non_empty_lines and '|' in non_empty_lines[0]:
            # 是 PSV 格式，保存到临时文件并调用 PSV 加载方法
            import tempfile
            import os
            try:
                with tempfile.NamedTemporaryFile(mode='w', suffix='.psv', delete=False, encoding='utf-8') as f:
                    f.write(content)
                    temp_path = f.name
                self._load_ades_psv_content(temp_path)
                os.unlink(temp_path)
                return
            except Exception as e:
                messagebox.showerror("错误", f"无法解析 Word 文档中的 PSV 数据：\n{str(e)}")
                return

        # 过滤 MPC80 数据
        filtered_lines = self._filter_mpc80_data(content)

        self.input_text.delete(1.0, tk.END)
        self.input_text.insert(1.0, '\n'.join(filtered_lines))
        self.status_var.set(f"已从 Word 文档提取 {len(filtered_lines)} 行有效 MPC80 数据（至少2条观测的天体）")

    def _load_ades_psv_content(self, filename):
        """从 ADES PSV 文件中读取表格内容（跳过注释和元数据）"""
        from digest2.observation import parse_ades_psv

        lines = []
        try:
            with open(filename, 'r', encoding='utf-8') as f:
                for line in f:
                    stripped = line.rstrip('\n')
                    # 跳过注释行 (# 开头)
                    if stripped.startswith('#'):
                        continue
                    # 跳过元数据行 (! 开头但不包含 |)
                    if stripped.startswith('!') and '|' not in stripped:
                        continue
                    # 保留表头行 (! 开头且包含 |) 和数据行
                    lines.append(stripped)
        except UnicodeDecodeError:
            with open(filename, 'r', encoding='gbk') as f:
                for line in f:
                    stripped = line.rstrip('\n')
                    if stripped.startswith('#'):
                        continue
                    if stripped.startswith('!') and '|' not in stripped:
                        continue
                    lines.append(stripped)
        except Exception as e:
            messagebox.showerror("错误", f"无法读取 ADES PSV 文件：\n{str(e)}")
            return

        # 解析文件获取统计信息
        try:
            tracklets = parse_ades_psv(filename)
            total_obs = sum(len(obs_list) for obs_list in tracklets.values())
            status_msg = f"已加载 ADES PSV 文件：{len(tracklets)} 个天体，共 {total_obs} 条观测 - 请直接点击'开始分析'"
        except:
            status_msg = "已加载 ADES PSV 文件 - 请直接点击'开始分析'"

        # 显示表格内容（表头和数据行）
        self.input_text.delete(1.0, tk.END)
        self.input_text.insert(1.0, '\n'.join(lines))
        self.status_var.set(status_msg)

    def _load_ades_xml_content(self, filename):
        """从 ADES XML 文件中读取原始内容"""
        from digest2.observation import parse_ades_xml

        raw_content = None
        try:
            # 读取原始文件内容
            with open(filename, 'r', encoding='utf-8') as f:
                raw_content = f.read()
        except UnicodeDecodeError:
            try:
                with open(filename, 'r', encoding='gbk') as f:
                    raw_content = f.read()
            except Exception as e:
                messagebox.showerror("错误", f"无法读取 ADES XML 文件：\n{str(e)}")
                return
        except Exception as e:
            messagebox.showerror("错误", f"无法读取 ADES XML 文件：\n{str(e)}")
            return

        if raw_content is None:
            return

        # 解析文件获取统计信息
        try:
            tracklets = parse_ades_xml(filename)
            total_obs = sum(len(obs_list) for obs_list in tracklets.values())
            status_msg = f"已加载 ADES XML 文件：{len(tracklets)} 个天体，共 {total_obs} 条观测 - 请直接点击'开始分析'"
        except:
            status_msg = "已加载 ADES XML 文件 - 请直接点击'开始分析'"

        # 显示原始内容
        self.input_text.delete(1.0, tk.END)
        self.input_text.insert(1.0, raw_content)
        self.status_var.set(status_msg)

    def _extract_designation(self, line):
        """从 MPC80 行中提取天体名称"""
        # MPC80 格式：
        # 1-5 列：临时编号（如果有）
        # 6-12 列：空格 + 临时编号或永久编号
        # 这里我们提取前12列并去除空格
        desig = line[:12].strip()
        if not desig:
            desig = "Unknown"
        return desig
    
    def _parse_observation_line(self, line):
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
            # 修复 NEOCP 格式问题：某些行第14列是项目代码（如0、K、|等），第15列才是观测类型（C/S/B）
            # 根据 MPC 格式，第14列应该是 note2（观测类型），所以我们需要修复这种情况
            fixed_line = line
            note2 = line[14] if len(line) > 14 else ' '
            if note2 not in ('C', 'S', 'B'):
                # 检查第15列是否是 C/S/B
                if len(line) > 15 and line[15] in ('C', 'S', 'B'):
                    # 将第14列替换为空格，让第15列移到第14列位置
                    fixed_line = line[:14] + line[15:]
                # 处理 |C 格式：第13列是 |，第14列是 C，但年份位置不对
                # 检查第12列是否是 | 且第13列是 C/S/B
                elif len(line) > 13 and line[12] == '|' and line[13] in ('C', 'S', 'B'):
                    # 将 | 替换为空格，让 C/S/B 移到正确的位置
                    fixed_line = line[:12] + ' ' + line[13:]
            obs = parse_mpc80(fixed_line)
            if obs:
                desig = self._extract_designation(line)
                return desig, obs
        
        # 尝试解析简化格式
        # 格式: DESIG[*] [PC]CYYYY MM DD.DDDDD HH MM SS.S +/-DD MM SS.S MAG BAND OBS
        # 例如: H251538* C2017 01 26.46663 09 51 37.61 -01 30 57.3 22.3 z T09
        #       K16Ga9M 4C2016 04 15.47312 13 59 50.38 +00 01...
        # * 表示发现标记（discovery asterisk）
        # P 表示项目代码（第14列）
        try:
            # 处理发现标记 *（如果有的话）
            line_clean = line.strip()
            has_discovery = '*' in line_clean
            line_clean_no_asterisk = line_clean.replace('*', ' ')
            
            # 从原始行末尾提取观测站代码（最后3个字符）
            obscode = line_clean[-3:] if len(line_clean) >= 3 else '500'
            
            # 修复可能缺少空格的情况
            import re
            # 修复 RA秒和Dec符号合并的情况：如 "50.866+00" -> "50.866 +00"
            line_clean = re.sub(r'(\d+\.\d{3})([+-])', r'\1 \2', line_clean_no_asterisk)
            # 修复日期和赤经时合并的情况：如 "15.72751102" -> "15.727511 02"
            # 日期是6位小数时和赤经时紧连着
            line_clean = re.sub(r'(\d+\.\d{6})(\d{2})', r'\1 \2', line_clean)
            
            # 分割字段
            parts = line_clean.split()
            # 接受至少10个字段（缺少星等和波段）或9个字段（缺少星等、波段和观测站）
            if len(parts) < 9:
                return None, None
            
            # 提取天体名称（第一个字段，可能包含发现标记）

            desig = parts[0]
            
            # 检查第二个字段是否包含年份信息
            # 支持多种格式：C2019, 4C2019, KB2025, K2025, etc.
            date_field = parts[1]
            # 查找年份（4位数字）
            year_match = re.search(r'\d{4}', date_field)
            if not year_match:
                return None, None
            year = int(year_match.group(0))
            month = int(parts[2])
            day = float(parts[3])
            
            # 计算 MJD
            # 简化计算，使用近似值
            from datetime import datetime
            dt = datetime(year, month, int(day))
            # MJD = JD - 2400000.5
            # 简化处理，这里需要更精确的计算
            
            # 解析 RA (HH MM SS.S)
            ra_h = float(parts[4])
            ra_m = float(parts[5])
            ra_s = float(parts[6])
            ra_deg = (ra_h + ra_m/60 + ra_s/3600) * 15  # 转换为度
            
            # 解析 Dec (+/-DD MM SS.S)
            dec_sign = 1 if parts[7][0] == '+' else -1
            dec_d = abs(float(parts[7]))
            dec_m = float(parts[8])
            dec_s = float(parts[9])
            dec_deg = dec_sign * (dec_d + dec_m/60 + dec_s/3600)
            
            # 星等和波段（处理两种格式：19.8 G 和 19.84G）
            mag = None
            band = 'V'
            
            if len(parts) >= 11:
                # 检查第11个字段是否是星等（可能包含波段）
                mag_field = parts[10]
                # 先检查是否是纯数字星等（不包含字母）
                if mag_field.replace('.', '', 1).isdigit() and mag_field.count('.') <= 1:
                    # 纯数字星等
                    mag = float(mag_field)
                    if len(parts) >= 12:
                        # 检查第12个字段是否是波段（单字母或类似g1, i1的格式）
                        if len(parts[11]) == 1 and parts[11].isalpha():
                            band = parts[11].upper()
                        elif len(parts[11]) == 2 and parts[11][0].isalpha() and parts[11][1].isdigit():
                            # 格式如 g1, i1, r1 等，第一个字符是波段，第二个是数字
                            band = parts[11][0].upper()
                else:
                    # 尝试匹配星等+波段格式：19.84G 或 22.65gW~9HCuO18
                    mag_match = re.search(r'^(\d+\.\d+)([A-Za-z])(.*)$', mag_field)
                    if mag_match:
                        mag = float(mag_match.group(1))
                        band = mag_match.group(2).upper()
                    else:
                        # 第11个字段不是星等，标记为缺失
                        mag = -1.0  # 使用 -1.0 表示缺失星等
            
            # 创建 Observation 对象
            obs = Observation(
                mjd=self._date_to_mjd(year, month, day),
                ra=ra_deg,
                dec=dec_deg,
                mag=mag,
                band=band,
                obscode=obscode,
                rms_ra=0.0,
                rms_dec=0.0,
                spacebased=False,
                earth_obs=[0.0, 0.0, 0.0]
            )
            return desig, obs
            
        except Exception as e:
            return None, None
    
    def _date_to_mjd(self, year, month, day):
        """将日期转换为 MJD（使用digest2库的标准实现）"""
        from digest2.observation import _date_to_mjd as digest_mjd
        return digest_mjd(year, month, day)
        
    def analyze(self):
        """开始分析 - 优先分析观测数据框中显示的内容"""
        input_data = self.input_text.get(1.0, tk.END).strip()
        file_path = self.file_path_var.get().strip()

        if input_data:
            # 优先分析观测数据框中显示的内容
            self.analyze_input()
        elif file_path:
            # 观测数据框为空，使用文件分析
            self.analyze_file()
        else:
            messagebox.showwarning("警告", "请输入观测数据或选择文件！")
    
    def analyze_file(self):
        """分析文件"""
        file_path = self.file_path_var.get().strip()
        
        if not file_path:
            messagebox.showwarning("警告", "请先选择观测文件！")
            return
            
        if not os.path.exists(file_path):
            messagebox.showerror("错误", f"文件不存在：\n{file_path}")
            return
            
        # 清空之前的结果
        self.clear_results()
        
        # 禁用按钮
        self.analyze_btn.config(state='disabled')
        self.browse_btn.config(state='disabled')
        self.status_var.set("正在分析...")
        self.root.update()
        
        try:
            with Digest2() as d2:
                results = d2.classify_file(file_path)
                
                if not results:
                    messagebox.showinfo("提示", "文件中没有找到有效的观测数据。")
                    return
                    
                self.display_results(results)
                total_obs = sum(len(result.tracklet) for result in results)
                self.status_var.set(f"分析完成，共 {len(results)} 个天体，{total_obs} 条数据")
                
        except Exception as e:
            error_msg = str(e)
            if "Cannot find observatory codes" in error_msg or "digest2.obscodes" in error_msg or "ObsCodes.html" in error_msg or "observatory codes file" in error_msg:
                messagebox.showerror("错误", "请先下载MPC天文台站代码")
            else:
                messagebox.showerror("分析错误", f"分析过程中出现错误：\n\n{error_msg}")
            self.status_var.set("分析失败")
        finally:
            # 恢复按钮
            self.analyze_btn.config(state='normal')
            self.browse_btn.config(state='normal')
    
    def analyze_input(self):
        """分析手动输入的数据"""
        input_data = self.input_text.get(1.0, tk.END).strip()

        if not input_data:
            messagebox.showwarning("警告", "请输入观测数据！")
            return

        # 清空之前的结果
        self.clear_results()

        # 禁用按钮
        self.analyze_btn.config(state='disabled')
        self.status_var.set("正在分析...")
        self.root.update()

        try:
            # 检测输入格式
            is_ades_xml = input_data.startswith('<?xml') or input_data.startswith('<ades>') or '<optical>' in input_data[:1000]
            
            # 检测 ADES PSV 格式
            # PSV 格式的特征：
            # 1. 存在以 ! 开头且包含 | 的表头行
            # 2. 数据行包含多个 | 分隔符
            lines = input_data.split('\n')
            is_ades_psv = False
            has_psv_header = False
            has_data_with_pipes = False
            
            for line in lines:
                stripped_line = line.strip()
                if not stripped_line:
                    continue
                
                # 检查是否是 PSV 表头行（以 ! 开头且包含 |）
                if stripped_line.startswith('!') and '|' in stripped_line:
                    has_psv_header = True
                    break
                
                # 检查是否是数据行（包含多个 |）
                if stripped_line.count('|') >= 2:
                    has_data_with_pipes = True
                    # 额外检查：| 不是仅在第13列位置（MPC80 格式的项目代码位置）
                    if len(stripped_line) > 12 and stripped_line[12] == '|':
                        # 检查是否有其他位置的 |
                        if stripped_line.count('|') > 1:
                            has_data_with_pipes = True
                        else:
                            has_data_with_pipes = False
                    break
            
            is_ades_psv = has_psv_header or has_data_with_pipes

            if is_ades_xml:
                # 使用 ADES XML 解析器
                self._analyze_ades_xml_input(input_data)
            elif is_ades_psv:
                # 使用 ADES PSV 解析器
                self._analyze_ades_psv_input(input_data)
            else:
                # 使用 MPC80 格式解析器
                self._analyze_mpc80_input(input_data)
        except Exception as e:
            error_msg = str(e)
            if "Cannot find observatory codes" in error_msg or "digest2.obscodes" in error_msg or "ObsCodes.html" in error_msg or "observatory codes file" in error_msg:
                messagebox.showerror("错误", "请先下载MPC天文台站代码")
            else:
                messagebox.showerror("分析错误", f"分析过程中出现错误：\n\n{error_msg}")
            self.status_var.set("分析失败")
        finally:
            # 恢复按钮
            self.analyze_btn.config(state='normal')

    def _analyze_ades_xml_input(self, input_data):
        """分析 ADES XML 格式的手动输入数据"""
        from digest2.observation import _parse_ades_optical
        from lxml import etree as ET
        import tempfile
        import os

        # 将输入数据写入临时文件
        try:
            with tempfile.NamedTemporaryFile(mode='w', suffix='.xml', delete=False, encoding='utf-8') as f:
                f.write(input_data)
                temp_path = f.name

            # 解析 XML 文件，按 permID/provID 分组
            tree = ET.parse(temp_path)
            root = tree.getroot()

            # 处理命名空间
            ns = ""
            if root.tag.startswith("{"):
                ns = root.tag.split("}")[0] + "}"

            # 按 permID/provID 分组存储观测数据
            tracklets: Dict[str, List] = {}

            # 查找所有 optical 元素
            for optical in root.iter(f"{ns}optical"):
                # 解析观测数据
                obs = _parse_ades_optical(optical, ns)
                if obs is None:
                    continue

                # 获取 permID 或 provID
                permID_el = optical.find(f"{ns}permID")
                provID_el = optical.find(f"{ns}provID")

                # 优先使用 permID，如果不存在则使用 provID
                desig = None
                if permID_el is not None and permID_el.text:
                    desig = permID_el.text.strip()
                elif provID_el is not None and provID_el.text:
                    desig = provID_el.text.strip()

                if desig:
                    if desig not in tracklets:
                        tracklets[desig] = []
                    tracklets[desig].append(obs)

            # 删除临时文件
            os.unlink(temp_path)

            if not tracklets:
                messagebox.showinfo("提示", "没有找到有效的观测数据。")
                self.analyze_btn.config(state='normal')
                return

            # 对每个天体的观测按时间排序
            for desig in tracklets:
                tracklets[desig].sort(key=lambda o: o.mjd)

            # 分析每个天体
            with Digest2() as d2:
                results = []
                skipped_tracklets = []
                self._desig_map = {}
                for desig, obs_list in tracklets.items():
                    try:
                        result = d2.classify_tracklet(obs_list, is_ades=True)
                        self._desig_map[id(result)] = desig
                        results.append(result)
                    except Exception as e:
                        if "need >=2 obs with motion and time span" in str(e):
                            skipped_tracklets.append(desig)
                        else:
                            raise

                if skipped_tracklets:
                    msg = f"以下 {len(skipped_tracklets)} 个天体因时间跨度太短被跳过：\n" + ", ".join(skipped_tracklets[:10])
                    if len(skipped_tracklets) > 10:
                        msg += f" 等共 {len(skipped_tracklets)} 个"
                    messagebox.showwarning("分析警告", msg)

                if results:
                    self.display_results(results, use_custom_desig=True)
                    total_obs = sum(len(tracklets[desig]) for desig in tracklets if desig not in skipped_tracklets)
                    self.status_var.set(f"分析完成，共 {len(results)} 个天体，{total_obs} 条数据")
                else:
                    messagebox.showinfo("提示", "没有足够时间跨度的天体可供分析。")

        except Exception as e:
            error_msg = str(e)
            if "Cannot find observatory codes" in error_msg or "digest2.obscodes" in error_msg or "ObsCodes.html" in error_msg or "observatory codes file" in error_msg:
                messagebox.showerror("错误", "请先下载MPC天文台站代码")
            else:
                messagebox.showerror("错误", f"无法解析 ADES XML 数据：\n{error_msg}")

        self.analyze_btn.config(state='normal')

    def _analyze_ades_psv_input(self, input_data):
        """分析 ADES PSV 格式的手动输入数据"""
        from digest2.observation import parse_ades_psv, _parse_ades_psv_row
        from io import StringIO
        import csv

        # 将输入数据按行分割，跳过表头行（以 ! 开头或包含列名）
        lines = []
        for line in input_data.split('\n'):
            line = line.rstrip()
            if not line:
                continue
            # 跳过以 ! 开头的行（ADES 表头标记）
            if line.startswith('!'):
                continue
            # 跳过看起来像表头的行（包含列名如 provID, mode, stn, obsTime 等）
            if 'provID' in line or 'obsTime' in line or ('mode' in line and 'stn' in line and 'ra' in line and 'dec' in line):
                continue
            lines.append(line)

        if not lines:
            messagebox.showinfo("提示", "没有找到有效的 ADES PSV 数据。")
            self.analyze_btn.config(state='normal')
            return

        # 解析 ADES PSV 数据
        tracklets = {}
        failed_lines = []

        for line in lines:
            parts = [p.strip() for p in line.split('|')]
            if parts and parts[-1] == '':
                parts = parts[:-1]

            # 尝试解析每一行
            try:
                # 构造字典格式
                headers = ['provID', 'mode', 'stn', 'obsTime', 'ra', 'dec', 'rmsRA', 'rmsDec',
                          'astCat', 'photCat', 'mag', 'rmsMag', 'band', 'photAp', 'logSNR',
                          'seeing', 'exp', 'rmsFit', 'nStars', 'ref', 'disc', 'subFmt',
                          'precTime', 'precRA', 'precDec', 'notes', 'remarks']
                row_dict = {}
                for i, part in enumerate(parts):
                    if i < len(headers):
                        row_dict[headers[i]] = part

                # 解析观测数据
                obs = _parse_ades_psv_row(row_dict)
                if obs is not None:
                    # 获取天体名称
                    desig = row_dict.get('provID', '') or row_dict.get('permID', '') or row_dict.get('trkSub', '')
                    if desig:
                        if desig not in tracklets:
                            tracklets[desig] = []
                        tracklets[desig].append(obs)
                else:
                    failed_lines.append(line[:50])
            except Exception:
                failed_lines.append(line[:50])

        if failed_lines:
            messagebox.showwarning("解析警告", f"以下 {len(failed_lines)} 行无法解析：\n" + "\n".join(failed_lines[:5]))

        if not tracklets:
            messagebox.showinfo("提示", "没有找到有效的观测数据。")
            self.analyze_btn.config(state='normal')
            return

        # 对每个天体的观测按时间排序
        for desig in tracklets:
            tracklets[desig].sort(key=lambda o: o.mjd)

        # 分析每个天体
        with Digest2() as d2:
            results = []
            skipped_tracklets = []
            self._desig_map = {}
            for desig, obs_list in tracklets.items():
                try:
                    result = d2.classify_tracklet(obs_list, is_ades=True)
                    self._desig_map[id(result)] = desig
                    results.append(result)
                except Exception as e:
                    if "need >=2 obs with motion and time span" in str(e):
                        skipped_tracklets.append(desig)
                    else:
                        raise

            if skipped_tracklets:
                msg = f"以下 {len(skipped_tracklets)} 个天体因时间跨度太短被跳过：\n" + ", ".join(skipped_tracklets[:10])
                if len(skipped_tracklets) > 10:
                    msg += f" 等共 {len(skipped_tracklets)} 个"
                messagebox.showwarning("分析警告", msg)

            if results:
                self.display_results(results, use_custom_desig=True)
                total_obs = sum(len(tracklets[desig]) for desig in tracklets if desig not in skipped_tracklets)
                self.status_var.set(f"分析完成，共 {len(results)} 个天体，{total_obs} 条数据")
            else:
                messagebox.showinfo("提示", "没有足够时间跨度的天体可供分析。")

        self.analyze_btn.config(state='normal')

    def _analyze_mpc80_input(self, input_data):
        """分析 MPC80 格式的手动输入数据"""
        # 解析观测数据
        observations = []
        failed_lines = []
        for line in input_data.split('\n'):
            line = line.rstrip()
            if line:
                desig, obs = self._parse_observation_line(line)
                if obs is None:
                    failed_lines.append(line[:50])
                else:
                    observations.append((desig, obs))

        if failed_lines:
            messagebox.showwarning("解析警告", f"以下 {len(failed_lines)} 行无法解析：\n" + "\n".join(failed_lines[:5]))

        if not observations:
            messagebox.showinfo("提示", "没有找到有效的观测数据。\n请确保输入的是 MPC 80列格式或简化格式。")
            self.analyze_btn.config(state='normal')
            return

        # 按天体名称分组
        tracklets = {}
        for desig, obs in observations:
            if desig not in tracklets:
                tracklets[desig] = []
            tracklets[desig].append(obs)

        # 对每个天体的观测按时间排序（MJD 递增）
        for desig in tracklets:
            tracklets[desig].sort(key=lambda o: o.mjd)

        # 分析每个天体
        with Digest2() as d2:
            results = []
            skipped_tracklets = []
            self._desig_map = {}  # 清空映射
            for desig, obs_list in tracklets.items():
                try:
                    result = d2.classify_tracklet(obs_list)
                    # 保存天体名称到映射表（使用 id 作为键）
                    self._desig_map[id(result)] = desig
                    results.append(result)
                except Exception as e:
                    if "need >=2 obs with motion and time span" in str(e):
                        skipped_tracklets.append(desig)
                    else:
                        raise  # 重新抛出其他异常

            if skipped_tracklets:
                msg = f"以下 {len(skipped_tracklets)} 个天体因时间跨度太短被跳过：\n" + ", ".join(skipped_tracklets[:10])
                if len(skipped_tracklets) > 10:
                    msg += f" 等共 {len(skipped_tracklets)} 个"
                messagebox.showwarning("分析警告", msg)

            if results:
                self.display_results(results, use_custom_desig=True)
                total_obs = sum(len(tracklets[desig]) for desig in tracklets if desig not in skipped_tracklets)
                self.status_var.set(f"分析完成，共 {len(results)} 个天体，{total_obs} 条数据")
            else:
                messagebox.showinfo("提示", "没有足够时间跨度的天体可供分析。\n每个天体需要至少2条观测，且时间跨度足够大。")

        self.analyze_btn.config(state='normal')

    def clear_input_data(self):
        """清空观测数据（输入框和文件路径）"""
        self.input_text.delete(1.0, tk.END)
        self.file_path_var.set("")
        self.status_var.set("观测数据已清空")
    
    def download_obscode(self):
        """下载MPC天文台站代码文件"""
        import urllib.request
        import os
        import sys
        
        url = "https://minorplanetcenter.net/iau/lists/ObsCodes.html"
        
        # 获取程序实际运行目录（支持打包成exe后的情况）
        if getattr(sys, 'frozen', False):
            # 运行在打包后的exe中
            app_dir = os.path.dirname(sys.executable)
        else:
            # 运行在Python脚本中
            app_dir = os.path.dirname(__file__)
        
        save_path = os.path.join(app_dir, "ObsCodes.html")
        
        # 设置超时时间为10秒
        timeout = 10
        
        try:
            self.status_var.set("正在下载天文台站代码...")
            self.root.update_idletasks()  # 更新UI显示
            
            # 创建请求对象并设置超时
            req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req, timeout=timeout) as response:
                with open(save_path, 'wb') as f:
                    f.write(response.read())
            
            self.status_var.set(f"下载成功！文件已保存到: {save_path}")
        except urllib.error.URLError as e:
            self.status_var.set(f"网络错误: {str(e)}")
        except timeout:
            self.status_var.set("下载超时，请稍后重试")
        except Exception as e:
            self.status_var.set(f"下载失败: {str(e)}")

    def display_results(self, results, use_custom_desig=False):
        """显示分析结果

        Args:
            results: 分析结果列表
            use_custom_desig: 是否使用自定义天体名称（用于手动输入模式）
        """
        self.results_data = results  # 保存结果供详情查看

        for r in results:
            # 获取各类评分
            scores = r.noid
            # 使用自定义天体名称（如果有的话）
            if use_custom_desig and id(r) in self._desig_map:
                desig = self._desig_map[id(r)]
            else:
                desig = r.designation

            # 判断是否需要高亮（NEO评分大于等于65）
            tags = ()
            if scores.NEO >= 65:
                tags = ('neo_highlight', 'neo_bold')

            self.tree.insert('', tk.END, values=(
                desig,
                f"{r.rms:.2f}",
                f"{scores.Int:.1f}",
                f"{scores.NEO:.1f}",
                f"{scores.N22:.1f}",
                f"{scores.N18:.1f}",
                f"{scores.MC:.1f}",
                f"{scores.MB1:.1f}",
                f"{scores.MB2:.1f}",
                f"{scores.MB3:.1f}",
                f"{scores.Hun:.1f}",
                f"{scores.Pho:.1f}",
                f"{scores.Pal:.1f}",
                f"{scores.Han:.1f}",
                f"{scores.Hil:.1f}",
                f"{scores.JTr:.1f}",
                f"{scores.JFC:.1f}"
            ), tags=tags)
            
    def clear_results(self):
        """清空结果"""
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.status_var.set("就绪")
        self.results_data = []
        self._desig_map = {}

    def show_text_context_menu(self, event):
        """显示文本输入框的右键菜单"""
        # 创建右键菜单
        context_menu = tk.Menu(self.root, tearoff=0)
        context_menu.add_command(label="剪切", command=self.cut_text)
        context_menu.add_command(label="复制", command=self.copy_text)
        context_menu.add_command(label="粘贴", command=self.paste_text)
        context_menu.add_command(label="全选", command=self.select_all_text)

        # 在鼠标位置显示菜单
        context_menu.tk_popup(event.x_root, event.y_root)

    def copy_text(self):
        """复制选中的文本到剪贴板"""
        try:
            # 先将焦点设置到输入框
            self.input_text.focus_set()
            selected_text = self.input_text.get(tk.SEL_FIRST, tk.SEL_LAST)
            if selected_text:
                self.root.clipboard_clear()
                self.root.clipboard_append(selected_text)
                self.status_var.set("已复制")
        except tk.TclError:
            pass  # 没有选中的文本

    def paste_text(self):
        """从剪贴板粘贴文本（替换选中的内容）"""
        try:
            clipboard_text = self.root.clipboard_get()
            # 先将焦点设置到输入框
            self.input_text.focus_set()
            # 检查是否有选中的文本
            try:
                self.input_text.delete(tk.SEL_FIRST, tk.SEL_LAST)
            except tk.TclError:
                pass  # 没有选中的文本
            # 在当前光标位置插入
            self.input_text.insert(tk.INSERT, clipboard_text)
            self.status_var.set("已粘贴")
        except tk.TclError:
            pass  # 剪贴板为空

    def cut_text(self):
        """剪切选中的文本"""
        try:
            # 先将焦点设置到输入框
            self.input_text.focus_set()
            selected_text = self.input_text.get(tk.SEL_FIRST, tk.SEL_LAST)
            if selected_text:
                self.root.clipboard_clear()
                self.root.clipboard_append(selected_text)
                self.input_text.delete(tk.SEL_FIRST, tk.SEL_LAST)
                self.status_var.set("已剪切")
        except tk.TclError:
            pass  # 没有选中的文本

    def select_all_text(self):
        """全选文本"""
        # 先将焦点设置到输入框
        self.input_text.focus_set()
        self.input_text.tag_add(tk.SEL, "1.0", tk.END)
        self.input_text.mark_set(tk.INSERT, "1.0")
        self.input_text.see(tk.INSERT)
    
    def show_about_context_menu(self, event, widget):
        """显示关于页面文本控件的右键菜单"""
        # 创建右键菜单
        context_menu = tk.Menu(self.root, tearoff=0)
        context_menu.add_command(label="复制", command=lambda: self.copy_about_text(widget))
        context_menu.add_command(label="全选", command=lambda: self.select_all_about_text(widget))
        
        # 在鼠标位置显示菜单
        context_menu.tk_popup(event.x_root, event.y_root)
    
    def copy_about_text(self, widget):
        """复制选中的文本到剪贴板"""
        try:
            selected_text = widget.get(tk.SEL_FIRST, tk.SEL_LAST)
            if selected_text:
                self.root.clipboard_clear()
                self.root.clipboard_append(selected_text)
                self.status_var.set("已复制")
        except tk.TclError:
            pass  # 没有选中的文本
    
    def select_all_about_text(self, widget):
        """全选文本"""
        widget.focus_set()
        widget.tag_add(tk.SEL, "1.0", tk.END)
        widget.mark_set(tk.INSERT, "1.0")
        widget.see(tk.INSERT)

    def show_tree_context_menu(self, event):
        """显示树形视图的右键菜单"""
        # 获取鼠标所在的行
        row_id = self.tree.identify_row(event.y)
        if row_id:
            # 如果没有选中行，先选中鼠标所在的行
            selected_items = self.tree.selection()
            if not selected_items or row_id not in selected_items:
                self.tree.selection_set(row_id)
        
        # 创建右键菜单
        context_menu = tk.Menu(self.root, tearoff=0)
        # 只有当有选中行时才显示"复制选中行"
        if self.tree.selection():
            context_menu.add_command(label="复制选中行", command=self.copy_selected_item)
        # 总是显示"复制全部数据"
        context_menu.add_command(label="复制全部数据", command=self.copy_all_data)
        
        # 在鼠标位置显示菜单
        context_menu.tk_popup(event.x_root, event.y_root)

    def copy_selected_item(self):
        """复制选中行的数据到剪贴板"""
        selected_items = self.tree.selection()
        if not selected_items:
            return
        
        # 获取所有选中行的数据
        rows = []
        for item in selected_items:
            values = self.tree.item(item, 'values')
            rows.append('\t'.join(str(v) for v in values))
        
        # 格式化数据（制表符分隔，便于粘贴到表格）
        text = '\n'.join(rows)
        
        # 复制到剪贴板
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        count = len(selected_items)
        self.status_var.set(f"已复制{count}行")

    def copy_all_data(self):
        """复制所有结果数据到剪贴板"""
        # 获取所有行数据
        rows = []
        # 添加表头
        headers = ['天体名称', 'RMS', 'Int', 'NEO', 'N22', 'N18', 'MC', 'MB1', 'MB2', 'MB3', 'Hun', 'Pho', 'Pal', 'Han', 'Hil', 'JTr', 'JFC']
        rows.append('\t'.join(headers))
        
        # 添加数据行
        for item in self.tree.get_children():
            values = self.tree.item(item, 'values')
            rows.append('\t'.join(str(v) for v in values))
        
        # 合并为文本
        text = '\n'.join(rows)
        
        # 复制到剪贴板
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.status_var.set("已复制全部数据")


def main():
    root = tk.Tk()
    app = Digest2GUI(root)
    root.mainloop()


if __name__ == '__main__':
    main()