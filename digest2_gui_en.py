#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Asterorbit GUI - Asteroid Orbit Classification Tool (English Version)
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


# Add virtual environment path (if running directly)
venv_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'venv', 'Lib', 'site-packages')
if os.path.exists(venv_path) and venv_path not in sys.path:
    sys.path.insert(0, venv_path)

def show_error_with_link(title, message, link_text=None, link_url=None):
    """Show error dialog with clickable link"""
    root = tk.Tk()
    root.withdraw()  # Hide main window
    
    dialog = tk.Toplevel(root)
    dialog.title(title)
    dialog.geometry("480x280")
    dialog.resizable(False, False)
    dialog.transient(root)
    dialog.grab_set()
    
    # Set icon
    try:
        icon_img = tk.PhotoImage(file=resource_path('digest2_icon_256.png'))
        dialog.iconphoto(True, icon_img)
    except:
        pass
    
    frame = ttk.Frame(dialog, padding=20)
    frame.pack(fill=tk.BOTH, expand=True)
    
    # Error icon
    icon_label = ttk.Label(frame, text="⚠️", font=('Arial', 32))
    icon_label.pack(pady=(0, 10))
    
    # Error message
    msg_label = ttk.Label(frame, text=message, justify=tk.CENTER, wraplength=400)
    msg_label.pack(pady=(0, 15))
    
    # Clickable link
    if link_text and link_url:
        link_label = ttk.Label(frame, text=link_text, foreground='#1a73e8', cursor='hand2')
        link_label.pack(pady=(0, 15))
        link_label.bind('<Button-1>', lambda e: webbrowser.open(link_url))
    
    # OK button
    ok_btn = ttk.Button(frame, text="OK", command=lambda: (root.destroy(), sys.exit(1)))
    ok_btn.pack(pady=(10, 0))
    
    # Center the dialog
    dialog.update_idletasks()
    x = (dialog.winfo_screenwidth() - dialog.winfo_width()) // 2
    y = (dialog.winfo_screenheight() - dialog.winfo_height()) // 2
    dialog.geometry(f"+{x}+{y}")
    
    root.mainloop()


try:
    from digest2 import Digest2, parse_mpc80
except ImportError as e:
    error_msg = str(e)
    if "Cannot find observatory codes" in error_msg or "digest2.obscodes" in error_msg or "ObsCodes.html" in error_msg or "observatory codes file" in error_msg:
        messagebox.showerror("Error", "Please download MPC observatory codes first")
    else:
        show_error_with_link(
            "Missing Runtime Library",
            "Program cannot start, may be missing Visual C++ Runtime Library.\n\nPlease click the link below to download and install:",
            "👉 https://aka.ms/vs/17/release/vc_redist.x64.exe",
            "https://aka.ms/vs/17/release/vc_redist.x64.exe"
        )
    sys.exit(1)

# Try to import python-docx (for reading Word documents)
try:
    from docx import Document
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False


class Digest2GUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Asterorbit - Asteroid Orbit Classification Tool")
        self.root.minsize(800, 600)
        self.set_window_icon()
        
        # Set styles
        self.style = ttk.Style()
        self.style.configure('TButton', font=('Segoe UI', 10))
        self.style.configure('TLabel', font=('Segoe UI', 10))
        self.style.configure('Header.TLabel', font=('Segoe UI', 12, 'bold'))
        
        # Store results and designation mapping (since ClassificationResult is frozen)
        self._desig_map = {}
        
        self.create_widgets()
        
    def create_widgets(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(0, weight=1)
        
        # Set window size
        self.root.geometry("1200x800")
        
        # Create Notebook and set appropriate tab width for English
        style = ttk.Style()
        style.configure('TNotebook.Tab', padding=[15, 5])
        
        # Create result table style (header Segoe UI bold, content Segoe UI)
        style.configure('ResultTreeview.Treeview',
                       font=('Segoe UI', 10))
        style.configure('ResultTreeview.Treeview.Heading', 
                       background='#f5f5f5', 
                       foreground='#333333',
                       font=('Segoe UI', 10, 'bold'))
        style.map('ResultTreeview.Treeview',
                 foreground=[('selected', '#333333')],
                 background=[('selected', '#e8f0fe'), ('active', '#f0f0f0')])
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Create evaluation tab
        self.create_evaluation_tab()
        
        # Create description tab
        self.create_description_tab()
        
        # Create about tab
        self.create_about_tab()

    def set_window_icon(self):
        try:
            icon_path = resource_path('digest2_icon_256.png')
            icon_img = tk.PhotoImage(file=icon_path)
            self.root.iconphoto(True, icon_img)
        except Exception:
            pass
    
    def create_evaluation_tab(self):
        """Create evaluation tab"""
        eval_frame = ttk.Frame(self.notebook)
        self.notebook.add(eval_frame, text='Classification')
        
        # Configure grid weights
        eval_frame.columnconfigure(0, weight=1)
        eval_frame.rowconfigure(3, weight=1)  # Input area can expand
        eval_frame.rowconfigure(5, weight=1)  # Result display area can expand
        
        # Title
        title_label = ttk.Label(
            eval_frame, 
            text="Asterorbit - Asteroid Orbit Classification Tool", 
            style='Header.TLabel',
            font=('Segoe UI', 16, 'bold')
        )
        title_label.grid(row=0, column=0, columnspan=3, pady=(10, 10))
        
        # Description text
        desc_label = ttk.Label(
            eval_frame,
            text="Select a file or enter observation data in MPC 80-column or ADES format, and the program will automatically analyze and evaluate orbit types.",
            wraplength=1100
        )
        desc_label.grid(row=1, column=0, columnspan=3, pady=(0, 15), sticky=tk.W)
        
        # File selection area
        file_frame = ttk.LabelFrame(eval_frame, text="File Selection", padding="10")
        file_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)
        
        ttk.Label(file_frame, text="Observation File:").grid(row=0, column=0, sticky=tk.W)
        
        self.file_path_var = tk.StringVar()
        self.file_entry = ttk.Entry(file_frame, textvariable=self.file_path_var)
        self.file_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5)
        
        self.browse_btn = ttk.Button(file_frame, text="Browse...", command=self.browse_file)
        self.browse_btn.grid(row=0, column=2, padx=5)
        
        # Manual input area
        input_frame = ttk.LabelFrame(eval_frame, text="Observation Data", padding="10")
        input_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        input_frame.columnconfigure(0, weight=1)
        input_frame.rowconfigure(1, weight=1)
        
        ttk.Label(input_frame, text="Enter observation data in MPC 80-column or ADES format:").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        
        self.input_text = scrolledtext.ScrolledText(
            input_frame,
            wrap=tk.NONE,
            width=80,
            height=8,
            font=('Segoe UI', 11)
        )
        self.input_text.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        # Add right-click menu for input text
        self.input_text.bind('<Button-3>', self.show_text_context_menu)
        
        # Button area
        btn_frame = ttk.Frame(eval_frame)
        btn_frame.grid(row=4, column=0, columnspan=3, pady=10)
        
        self.analyze_btn = ttk.Button(
            btn_frame,
            text="Start Analysis",
            command=self.analyze,
            width=20
        )
        self.analyze_btn.pack(side=tk.LEFT, padx=5)

        self.clear_data_btn = ttk.Button(
            btn_frame,
            text="Clear Data",
            command=self.clear_input_data,
            width=20
        )
        self.clear_data_btn.pack(side=tk.LEFT, padx=5)
        
        self.download_obscode_btn = ttk.Button(
            btn_frame,
            text="Download MPC Observatory Codes",
            command=self.download_obscode,
            width=35
        )
        self.download_obscode_btn.pack(side=tk.LEFT, padx=5)
        
        # Result display area (row=5)
        result_frame = ttk.LabelFrame(eval_frame, text="Analysis Results", padding="10")
        result_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        result_frame.columnconfigure(0, weight=1)
        result_frame.rowconfigure(0, weight=1)
        
        # Create treeview to display results
        columns = ('designation', 'rms', 'int', 'neo', 'n22', 'n18', 'mc', 'mb1', 'mb2', 'mb3', 'hun', 'pho', 'pal', 'han', 'hil', 'jtr', 'jfc')
        self.tree = ttk.Treeview(
            result_frame,
            columns=columns,
            show='headings',
            height=10,
            style='ResultTreeview.Treeview',
            selectmode='extended'
        )
        # Add right-click menu for treeview (supports copying)
        self.tree.bind('<Button-3>', self.show_tree_context_menu)
        # Bind left-click event to toggle selection
        self.tree.bind('<Button-1>', self.on_tree_click)

        # Configure tag styles: NEO highlight (yellow background) and bold
        self.tree.tag_configure('neo_highlight', background='#FFD700')
        self.tree.tag_configure('neo_bold', font=('Segoe UI', 10, 'bold'))
        
        # Set column headers
        self.tree.heading('designation', text='Designation')
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
        
        # Set column widths (increase width for decimal display)
        self.tree.column('designation', width=120, minwidth=100)
        self.tree.column('rms', width=60, minwidth=50, anchor='center')
        self.tree.column('int', width=55, minwidth=45, anchor='center')
        for col in columns[3:]:
            self.tree.column(col, width=55, minwidth=45, anchor='center')
        
        # Scrollbars
        scrollbar_y = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar_x = ttk.Scrollbar(result_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar_y.grid(row=0, column=1, sticky=(tk.N, tk.S))
        scrollbar_x.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # Status bar (row=6)
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        status_bar = ttk.Label(eval_frame, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(5, 0))
    
    def create_description_tab(self):
        """Create description tab"""
        desc_frame = ttk.Frame(self.notebook)
        self.notebook.add(desc_frame, text='Orbit Types')
        
        # Create main frame
        main_frame = ttk.Frame(desc_frame)
        main_frame.pack(padx=15, pady=15, fill=tk.BOTH, expand=True)
        
        # Add title
        title_label = ttk.Label(main_frame, text="Orbit Type Descriptions", font=('Segoe UI', 16, 'bold'), foreground='#333')
        title_label.pack(anchor=tk.W, pady=(0, 8))
        
        # Add green underline (consistent with about page)
        title_line = ttk.Frame(main_frame, height=2, style='TitleLine.TFrame')
        title_line.pack(anchor=tk.W, fill=tk.X, pady=(0, 15))
        
        # Create text container (supports rich text)
        text_frame = ttk.Frame(main_frame, relief='solid', borderwidth=1)
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create Text widget
        text_widget = tk.Text(text_frame, wrap=tk.WORD, font=('Segoe UI', 10), 
                             bg='#ffffff', relief='flat', spacing1=6, spacing2=4, spacing3=6,
                             borderwidth=0, highlightthickness=0)
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Set tab stops for column alignment
        text_widget.config(tabs=('2c', '9c'))
        
        # Configure tag styles
        text_widget.tag_config('header', font=('Segoe UI', 10, 'bold'), background='#f5f5f5')
        text_widget.tag_config('abbrev', font=('Segoe UI', 10, 'bold'))
        text_widget.tag_config('fullname', font=('Segoe UI', 10))
        text_widget.tag_config('definition', font=('Segoe UI', 10))
        text_widget.tag_config('italic', font=('Segoe UI', 10, 'italic'))
        text_widget.tag_config('row_bg1', background='#ffffff')
        text_widget.tag_config('row_bg2', background='#f9f9f9')
        
        # Orbit type descriptions (using Digest2 official definitions and standard astronomical parameters)
        orbit_types = [
            ('Int', 'MPC Interesting', 'Meets any of: q < 1.3 AU, Q > 10 AU, e >= 0.5, i >= 40°'),
            ('NEO', 'Near-Earth Object', 'q < 1.3 AU'),
            ('N22', 'Intermediate-size Near-Earth Object', 'D > 140 m. q < 1.3 AU, H < 22.5'),
            ('N18', 'Large Near-Earth Object', 'D > 1.2 km. q < 1.3 AU, H < 18.5'),
            ('MC', 'Mars Crosser', 'Crosses Mars orbit. 1.3 AU < q <1.67 AU, Q > 1.58 AU'),
            ('Hun', 'Hungaria Group', 'Group represented by 434 Hungaria. 1.78 AU < a <2 AU, e < 0.18, 16° < i <34°'),
            ('Pho', 'Phocaea Group', 'Group represented by 25 Phocaea. q > 1.5 AU, 2.2 AU < a <2.45 AU, 20° < i <27°'),
            ('MB1', 'Inner Main Belt', 'q > 1.67 AU, 2.1 AU < a <2.5 AU, i < ((a-2.1)/0.4)*10+7'),
            ('Pal', 'Pallas Group', 'Group represented by 2 Pallas. 2.5 AU < a <2.8 AU, e < 0.35, 24° < i <37°'),
            ('Han', 'Hansa Group', 'Group represented by 480 Hansa. 2.55 AU < a <2.72 AU, e < 0.25, 20° < i <23.5°'),
            ('MB2', 'Middle Main Belt', '2.5 AU < a <2.8 AU, e < 0.45, i < 20°'),
            ('MB3', 'Outer Main Belt', '2.8 AU < a <3.25 AU, e < 0.4, i < ((a-2.8)/0.45)*16+20'),
            ('Hil', 'Hilda Group', 'Group represented by 153 Hilda. 3.9 AU < a <4.02 AU, e < 0.4, i < 18°'),
            ('JTr', 'Jupiter Trojan', 'Asteroids in Jupiter\'s L4/L5 Lagrange points. 5.05 AU < a <5.35 AU, e < 0.22, i < 38°'),
            ('JFC', 'Jupiter Family Comet', 'q > 1.3 AU, 2 < TJ <3'),
        ]
        
        # Add header
        text_widget.insert(tk.END, '  Abbr\t', ('header', 'row_bg1'))
        text_widget.insert(tk.END, 'Full Name\t', ('header', 'row_bg1'))
        text_widget.insert(tk.END, 'Digest2 Definition\n', ('header', 'row_bg1'))
        
        # Add data rows
        for i, (abbrev, fullname, definition) in enumerate(orbit_types):
            bg_tag = 'row_bg1' if i % 2 == 0 else 'row_bg2'
            
            text_widget.insert(tk.END, f'  {abbrev}\t', ('abbrev', bg_tag))
            text_widget.insert(tk.END, f'{fullname}\t', ('fullname', bg_tag))
            
            # Insert definition and mark italic parts
            self.insert_definition_with_italic(text_widget, definition, bg_tag)
            text_widget.insert(tk.END, '\n')
        
        text_widget.config(state=tk.DISABLED)
        
        # Add scrollbar
        scrollbar_y = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=text_widget.yview)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        text_widget.configure(yscrollcommand=scrollbar_y.set)
        
        self.desc_text = text_widget
    
    def insert_definition_with_italic(self, text_widget, definition, bg_tag):
        """Insert definition text, apply italic style to specific letters"""
        import re
        # Pattern for letters that need italics
        pattern = r'\b([qaieQHDTJ])\b'
        parts = re.split(pattern, definition)
        
        for j, part in enumerate(parts):
            if j % 2 == 1 and part in 'qaieQHDTJ':
                text_widget.insert(tk.END, part, ('italic', bg_tag))
            else:
                text_widget.insert(tk.END, part, ('definition', bg_tag))
    
    def create_about_tab(self):
        """Create about tab"""
        about_frame = ttk.Frame(self.notebook, borderwidth=0)
        self.notebook.add(about_frame, text='About')
        
        # Create content container - simplified layout
        content_frame = ttk.Frame(about_frame, padding=(15, 15))
        content_frame.pack(fill=tk.BOTH, expand=True)
        content_frame.configure(style='Transparent.TFrame')
        
        # Set background color
        about_frame.configure(style='Transparent.TFrame')
        
        # Add title
        title_label = ttk.Label(content_frame, text="About", font=('Segoe UI', 14, 'bold'), foreground='#333')
        title_label.pack(anchor=tk.W, pady=(0, 8))
        
        # Add green underline (simulate web style)
        title_line = ttk.Frame(content_frame, height=2, style='TitleLine.TFrame')
        title_line.pack(anchor=tk.W, fill=tk.X, pady=(0, 15))
        
        # Add disclaimer information (merge copyright and disclaimer) - use Text widget to ensure proper wrapping
        about_text_widget = tk.Text(content_frame, wrap=tk.WORD, height=2, 
                                  font=('Segoe UI', 11), bg='#f5f5f5', relief='flat',
                                  foreground='#333', spacing1=3, spacing2=2, spacing3=3,
                                  borderwidth=0, highlightthickness=0)
        about_text_widget.pack(fill=tk.X, anchor=tk.W, pady=(0, 15))
        about_text_widget.insert(tk.END, "This application is built by Jingyuan Zhao based on the official open source Digest2 code from the Minor Planet Center, and is not an official MPC project. Digest2 source code authors: Sonia Keys, Carl Hergenrother, Robert McNaught, David Asher, with ADES support added by Richard Cloete and Peter Vereš.")
        about_text_widget.config(state=tk.DISABLED)
        
        # Add supported data formats
        data_format_label = ttk.Label(content_frame, text="Supported data formats: MPC 80-column, ADES XML, ADES PSV", 
                                      font=('Segoe UI', 10), foreground='#666', style='Transparent.TLabel')
        data_format_label.pack(anchor=tk.W, pady=(0, 10))
        
        # Add supported file formats
        file_format_label = ttk.Label(content_frame, text="Supported file formats: obs, txt, dat, docx, doc, xml, psv", 
                                      font=('Segoe UI', 10), foreground='#666', style='Transparent.TLabel')
        file_format_label.pack(anchor=tk.W, pady=(0, 10))
        
        # Add web version link
        web_label_container = ttk.Frame(content_frame, style='Transparent.TFrame')
        web_label_container.pack(anchor=tk.W, pady=(0, 10))
        
        web_text_label = ttk.Label(web_label_container, text="Web version: ", 
                                   font=('Segoe UI', 10), foreground='#666', style='Transparent.TLabel')
        web_text_label.pack(side=tk.LEFT)
        
        web_link_label = tk.Label(web_label_container, text="https://asterorbit-digest2.hf.space/en/", 
                                  font=('Segoe UI', 10, 'underline'), fg='#1a73e8', cursor='hand2', 
                                  bg='#f5f5f5')
        web_link_label.pack(side=tk.LEFT)
        
        def open_web_link(event):
            webbrowser.open("https://asterorbit-digest2.hf.space/en/")
        
        web_link_label.bind("<Button-1>", open_web_link)
        
        # Add version info
        import digest2
        import importlib.metadata as metadata
        digest2_version = 'Unknown version'
        try:
            digest2_version = metadata.version('digest2')
        except metadata.PackageNotFoundError:
            digest2_version = getattr(digest2, '__version__', getattr(digest2, 'VERSION', 'Unknown version'))
        
        version_label = ttk.Label(content_frame, text=f"Digest2 Version: {digest2_version}", 
                         font=('Segoe UI', 10), foreground='#666', style='Transparent.TLabel')
        version_label.pack(anchor=tk.W, pady=(0, 20))
        
        # Add references title
        ref_title_label = ttk.Label(content_frame, text="References", font=('Segoe UI', 11, 'bold'), foreground='#333')
        ref_title_label.pack(anchor=tk.W, pady=(15, 12))
        
        # Create references container (with gray border)
        ref_container = ttk.Frame(content_frame, relief='solid', borderwidth=1, style='RefContainer.TFrame')
        ref_container.pack(fill=tk.X, padx=10, pady=10)
        
        # Use Text widget to display references (supports hyperlinks)
        ref_text_widget = tk.Text(ref_container, wrap=tk.WORD, height=9, 
                                  font=('Segoe UI', 10), bg='#f5f5f5', relief='flat',
                                  foreground='#444', spacing1=5, spacing2=5, spacing3=10,
                                  borderwidth=0, highlightthickness=0)
        ref_text_widget.pack(fill=tk.X, padx=10, pady=10)
        
        # Add references content (with hyperlinks)
        ref_text_widget.insert(tk.END, "[1] Keys S, Vereš P, Payne M J, et al. The digest2 NEO Classification Code. Publications of the Astronomical Society of the Pacific, 2019, 131(1000): 1-18.\n")
        ref_text_widget.insert(tk.END, "[2] Vereš P, Cloete R, Weryk R, et al. Improvement of Digest2 NEO Classification Code—utilizing the Astrometry Data Exchange Standard. Publications of the Astronomical Society of the Pacific, 2023, 135(1052).\n")
        ref_text_widget.insert(tk.END, "[3] Vereš P, Cloete R, Payne M J, et al. Improving the discovery of near-Earth objects with machine-learning methods. Astronomy & Astrophysics, 2025, 698: A242.\n")
        ref_text_widget.insert(tk.END, "[4] M. P. C. Staff. Editorial Notice: New digest2 repository and package: M.P.E.C. 2026-E23. Cambridge, MA: Minor Planet Center, 2026-03-05. ")
        ref_text_widget.insert(tk.END, "https://www.minorplanetcenter.net/mpec/K26/K26E23.html", 'link')
        ref_text_widget.insert(tk.END, "\n")
        ref_text_widget.insert(tk.END, "[5] M. P. C. Staff. Editorial Notice: digest2 population model update: M.P.E.C. 2026-K125. Cambridge, MA: Minor Planet Center, 2026-05-29. ")
        ref_text_widget.insert(tk.END, "https://www.minorplanetcenter.net/mpec/K26/K26KC5.html", 'link')
        ref_text_widget.insert(tk.END, "\n")
        ref_text_widget.insert(tk.END, "[6] GitHub. Digest2 source code. ")
        ref_text_widget.insert(tk.END, "https://github.com/Smithsonian/mpc-public/tree/main/digest2", 'link')
        ref_text_widget.insert(tk.END, "\n")
        ref_text_widget.insert(tk.END, "[7] MPC Public Documentation Hub. Digest2 tutorials. ")
        ref_text_widget.insert(tk.END, "https://docs.minorplanetcenter.net/tutorials/iod_tutorials/?h=digest2", 'link')
        
        # Configure hyperlink styles
        ref_text_widget.tag_config('link', foreground='#1a73e8', underline=True)
        
        # Hyperlink click event handler
        def open_link(event):
            text = ref_text_widget.get('current linestart', 'current lineend')
            import re
            urls = re.findall(r'https?://\S+', text)
            if urls:
                webbrowser.open(urls[0])
        
        # Show hand cursor on mouse hover
        def on_mouse_move(event):
            tags = ref_text_widget.tag_names(tk.CURRENT)
            if 'link' in tags:
                ref_text_widget.config(cursor='hand2')
            else:
                ref_text_widget.config(cursor='')
        
        ref_text_widget.tag_bind('link', '<Button-1>', open_link)
        ref_text_widget.bind('<Motion>', on_mouse_move)
        ref_text_widget.config(state=tk.DISABLED)
        
        # Configure styles
        style = ttk.Style()
        style.configure('TitleLine.TFrame', background='#4CAF50')
        style.configure('Transparent.TFrame', background='#f5f5f5')
        style.configure('Transparent.TLabel', background='#f5f5f5')
        style.configure('RefContainer.TFrame', background='#f5f5f5', bordercolor='#e0e0e0')
    
    def on_tree_click(self, event):
        """Handle click event for result table to toggle selection"""
        # Get clicked row
        item = self.tree.identify_row(event.y)
        if not item:
            return
        
        # Get modifier key states
        ctrl_pressed = (event.state & 0x4) != 0  # Ctrl key
        shift_pressed = (event.state & 0x1) != 0  # Shift key
        
        if ctrl_pressed:
            # Ctrl+Click: toggle individual selection
            if item in self.tree.selection():
                self.tree.selection_remove(item)
            else:
                self.tree.selection_add(item)
            return 'break'  # Prevent default selection behavior
        elif shift_pressed:
            # Shift+Click: let Treeview handle default range selection
            return
        else:
            # Normal click
            selected = self.tree.selection()
            if item in selected:
                # If clicked row is already selected, deselect it
                self.tree.selection_remove(item)
                return 'break'
            else:
                # If clicked row is not selected, select only this row, deselect all others
                self.tree.selection_set(item)
                return 'break'
    
    def browse_file(self):
        filetypes = [
            ('Observation Files', '*.obs *.xml *.psv *.txt *.dat'),
            ('MPC 80-column Format', '*.obs'),
            ('ADES XML Format', '*.xml'),
            ('ADES PSV Format', '*.psv'),
            ('Text Files', '*.txt *.dat'),
            ('All Files', '*.*')
        ]
        # Add Word documents to filetypes if supported
        if HAS_DOCX:
            filetypes.insert(1, ('Word Documents', '*.docx *.doc'))
            filetypes[0] = ('Observation Files', '*.obs *.xml *.psv *.txt *.dat *.docx *.doc')

        filename = filedialog.askopenfilename(
            title="Select Observation File",
            filetypes=filetypes
        )
        if filename:
            self.file_path_var.set(filename)
            # Automatically process based on file type
            self._load_file_content(filename)

    def _load_file_content(self, filename):
        """Load content into input box based on file type"""
        ext = filename.lower().split('.')[-1] if '.' in filename else ''

        try:
            if ext in ['docx', 'doc'] and HAS_DOCX:
                # Word document
                self._load_docx_content(filename)
            elif ext in ['txt', 'dat', 'obs']:
                # Text file (including MPC80 format .obs files)
                self._load_text_content(filename)
            elif ext == 'xml':
                # ADES XML file - parse and display as readable format
                self._load_ades_xml_content(filename)
            elif ext == 'psv':
                # ADES PSV file - parse and display as readable format
                self._load_ades_psv_content(filename)
            # Other formats are not auto-loaded
        except Exception as e:
            messagebox.showerror("Error", f"Cannot read file:\n{str(e)}")

    def _load_text_content(self, filename):
        """Read content from text file, auto-detect format (MPC80, ADES XML, or ADES PSV)"""
        try:
            with open(filename, 'r', encoding='utf-8') as f:
                content = f.read()
        except UnicodeDecodeError:
            with open(filename, 'r', encoding='gbk') as f:
                content = f.read()

        # Check if ADES XML format
        if content.strip().startswith('<?xml') or content.strip().startswith('<ades'):
            # XML format, call XML loading method
            self._load_ades_xml_content(filename)
            return

        # Check if ADES PSV format
        lines = content.strip().split('\n')
        non_empty_lines = [l for l in lines if l.strip()]
        
        is_psv_format = False
        has_psv_header = False
        has_data_with_pipes = False
        
        for line in non_empty_lines:
            # Skip comment lines
            if line.startswith('#'):
                continue
            
            # Check if PSV header line (starts with ! and contains |)
            if line.startswith('!') and '|' in line:
                has_psv_header = True
                break
            
            # Check if data line (contains multiple |)
            if line.count('|') >= 2:
                has_data_with_pipes = True
                # Extra check: | not only at column 13 (MPC80 format program code position)
                if len(line) > 12 and line[12] == '|':
                    # Check if there are | at other positions
                    if line.count('|') > 1:
                        has_data_with_pipes = True
                    else:
                        has_data_with_pipes = False
                break
        
        is_psv_format = has_psv_header or has_data_with_pipes
        
        if is_psv_format:
            # PSV format, call PSV loading method
            self._load_ades_psv_content(filename)
            return

        # Parse and filter MPC80 data
        filtered_lines = self._filter_mpc80_data(content)

        self.input_text.delete(1.0, tk.END)
        self.input_text.insert(1.0, '\n'.join(filtered_lines))
        self.status_var.set(f"Loaded {len(filtered_lines)} valid MPC80 lines (objects with at least 2 observations)")

    def _filter_mpc80_data(self, content):
        """Filter MPC80 data, keeping only objects with at least 2 observations"""
        lines = content.split('\n')

        # Collect all line designations (deduplicate: keep only one identical line)
        tracklets = {}
        seen_lines = set()  # For deduplication

        for line in lines:
            line = line.rstrip()
            if not line:
                continue

            # Try to parse MPC80 format to get designation
            desig = self._try_extract_designation(line)
            if desig:
                # Deduplicate check
                if line not in seen_lines:
                    seen_lines.add(line)
                    if desig not in tracklets:
                        tracklets[desig] = []
                    tracklets[desig].append(line)

        # Keep only objects with at least 2 observations
        filtered = []
        for desig, obs_lines in tracklets.items():
            if len(obs_lines) >= 2:
                filtered.extend(obs_lines)

        return filtered

    def _try_extract_designation(self, line):
        """Try to extract designation from line, return None if not MPC80 format"""
        # MPC80 format: designation in first 12 columns
        if len(line) < 12:
            return None

        # Extract first 12 columns as designation
        desig = line[:12].strip()

        # Check if looks like MPC80 format
        if len(line) >= 80:
            # Standard MPC80 format: check if column 14 or 15 has observation type (C/S/B)
            note2 = line[14] if len(line) > 14 else ' '
            # Check if column 14 or 15 contains observation type
            has_obs_type = note2 in ('C', 'S', 'B') or \
                          (len(line) > 15 and line[15] in ('C', 'S', 'B')) or \
                          (len(line) > 13 and line[13] in ('C', 'S', 'B'))
            
            # Check if column 13-14 has program code
            has_prog_code = line[12:14].strip()
            
            if has_obs_type or has_prog_code:
                return desig if desig else None

        # Try simplified format (space-separated)
        parts = line.split()
        if len(parts) >= 2:
            # Check if second part contains year (like C2019, 4C2019)
            import re
            if re.search(r'C\d{4}', parts[1]):
                return parts[0]

        return None

    def _load_docx_content(self, filename):
        """Extract text content from Word document, auto-detect format (MPC80, ADES XML, or ADES PSV)"""
        doc = Document(filename)
        text_lines = []
        for para in doc.paragraphs:
            if para.text.strip():
                text_lines.append(para.text)

        # Merge content
        content = '\n'.join(text_lines)

        # Check if ADES XML format
        if content.strip().startswith('<?xml') or content.strip().startswith('<ades'):
            # XML format, save to temp file and call XML loading method
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
                messagebox.showerror("Error", f"Cannot parse XML data from Word document:\n{str(e)}")
                return

        # Check if ADES PSV format (contains | separator)
        lines = content.strip().split('\n')
        non_empty_lines = [l for l in lines if l.strip()]
        if non_empty_lines and '|' in non_empty_lines[0]:
            # PSV format, save to temp file and call PSV loading method
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
                messagebox.showerror("Error", f"Cannot parse PSV data from Word document:\n{str(e)}")
                return

        # Filter MPC80 data
        filtered_lines = self._filter_mpc80_data(content)

        self.input_text.delete(1.0, tk.END)
        self.input_text.insert(1.0, '\n'.join(filtered_lines))
        self.status_var.set(f"Extracted {len(filtered_lines)} valid MPC80 lines from Word document (objects with at least 2 observations)")

    def _load_ades_psv_content(self, filename):
        """Read table content from ADES PSV file (skip comments and metadata)"""
        from digest2.observation import parse_ades_psv

        lines = []
        try:
            with open(filename, 'r', encoding='utf-8') as f:
                for line in f:
                    stripped = line.rstrip('\n')
                    # Skip comment lines (# start)
                    if stripped.startswith('#'):
                        continue
                    # Skip metadata lines (! start but no |)
                    if stripped.startswith('!') and '|' not in stripped:
                        continue
                    # Keep header lines (! start and contains |) and data lines
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
            messagebox.showerror("Error", f"Cannot read ADES PSV file:\n{str(e)}")
            return

        # Parse file to get statistics
        try:
            tracklets = parse_ades_psv(filename)
            total_obs = sum(len(obs_list) for obs_list in tracklets.values())
            status_msg = f"Loaded ADES PSV file: {len(tracklets)} objects, {total_obs} observations - click 'Start Analysis' directly"
        except:
            status_msg = "Loaded ADES PSV file - click 'Start Analysis' directly"

        # Display table content (headers and data lines)
        self.input_text.delete(1.0, tk.END)
        self.input_text.insert(1.0, '\n'.join(lines))
        self.status_var.set(status_msg)

    def _load_ades_xml_content(self, filename):
        """Read raw content from ADES XML file"""
        from digest2.observation import parse_ades_xml

        raw_content = None
        try:
            # Read raw file content
            with open(filename, 'r', encoding='utf-8') as f:
                raw_content = f.read()
        except UnicodeDecodeError:
            try:
                with open(filename, 'r', encoding='gbk') as f:
                    raw_content = f.read()
            except Exception as e:
                messagebox.showerror("Error", f"Cannot read ADES XML file:\n{str(e)}")
                return
        except Exception as e:
            messagebox.showerror("Error", f"Cannot read ADES XML file:\n{str(e)}")
            return

        if raw_content is None:
            return

        # Parse file to get statistics
        try:
            tracklets = parse_ades_xml(filename)
            total_obs = sum(len(obs_list) for obs_list in tracklets.values())
            status_msg = f"Loaded ADES XML file: {len(tracklets)} objects, {total_obs} observations - click 'Start Analysis' directly"
        except:
            status_msg = "Loaded ADES XML file - click 'Start Analysis' directly"

        # Display raw content
        self.input_text.delete(1.0, tk.END)
        self.input_text.insert(1.0, raw_content)
        self.status_var.set(status_msg)

    def _extract_designation(self, line):
        """Extract designation from MPC80 line"""
        # MPC80 format:
        # Columns 1-5: provisional designation (if any)
        # Columns 6-12: space + provisional or permanent designation
        # Here we extract first 12 columns and trim spaces
        desig = line[:12].strip()
        if not desig:
            desig = "Unknown"
        return desig
    
    def _parse_observation_line(self, line):
        """
        Parse observation data line, support multiple formats:
        1. Standard MPC80 format (80 columns)
        2. Simplified format (like H251538 C2019 01 26.46663 ...)
        """
        from digest2.observation import Observation
        import re

        line = line.strip()
        if not line:
            return None, None

        # Try standard MPC80 format
        if len(line) >= 80:
            # Fix NEOCP format issue: some lines have program code at column 14 (like 0, K, |, etc.), observation type at column 15 (C/S/B)
            # According to MPC format, column 14 should be note2 (observation type), so we need to fix this case
            fixed_line = line
            note2 = line[14] if len(line) > 14 else ' '
            if note2 not in ('C', 'S', 'B'):
                # Check if column 15 is C/S/B
                if len(line) > 15 and line[15] in ('C', 'S', 'B'):
                    # Replace column 14 with space, move column 15 to column 14 position
                    fixed_line = line[:14] + line[15:]
                # Handle |C format: column 12 is |, column 13 is C/S/B, but year position is wrong
                # Check if column 12 is | and column 13 is C/S/B
                elif len(line) > 13 and line[12] == '|' and line[13] in ('C', 'S', 'B'):
                    # Replace | with space, move C/S/B to correct position
                    fixed_line = line[:12] + ' ' + line[13:]
            obs = parse_mpc80(fixed_line)
            if obs:
                desig = self._extract_designation(line)
                return desig, obs
        
        # Try simplified format
        # Format: DESIG[*] [PC]CYYYY MM DD.DDDDD HH MM SS.S +/-DD MM SS.S MAG BAND OBS
        # Example: H251538* C2019 01 26.46663 09 51 37.61 -01 30 57.3 22.3 z T09
        #          K16Ga9M 4C2016 04 15.47312 13 59 50.38 +00 01...
        # * indicates discovery asterisk
        # P indicates program code (column 14)
        try:
            # Handle discovery asterisk * (if any)
            line_clean = line.strip()
            has_discovery = '*' in line_clean
            line_clean_no_asterisk = line_clean.replace('*', ' ')
            
            # Extract observatory code from original line end (last 3 characters)
            obscode = line_clean[-3:] if len(line_clean) >= 3 else '500'
            
            # Fix possible missing spaces
            import re
            # Fix RA seconds and Dec sign merged: like "50.866+00" -> "50.866 +00"
            line_clean = re.sub(r'(\d+\.\d{3})([+-])', r'\1 \2', line_clean_no_asterisk)
            # Fix date and RA hour merged: like "15.72751102" -> "15.727511 02"
            # Date has 6 decimal places and is tightly connected to RA hour
            line_clean = re.sub(r'(\d+\.\d{6})(\d{2})', r'\1 \2', line_clean)
            
            # Split fields
            parts = line_clean.split()
            # Accept at least 10 fields (missing magnitude and band) or 9 fields (missing magnitude, band, and observatory)
            if len(parts) < 9:
                return None, None
            
            # Extract designation (first field, may contain discovery asterisk)

            desig = parts[0]
            
            # Check if second field contains year information
            # Support multiple formats: C2019, 4C2019, KB2025, K2025, etc.
            date_field = parts[1]
            # Find year (4 digits)
            year_match = re.search(r'\d{4}', date_field)
            if not year_match:
                return None, None
            year = int(year_match.group(0))
            month = int(parts[2])
            day = float(parts[3])
            
            # Calculate MJD
            # Simplified calculation, use approximate value
            from datetime import datetime
            dt = datetime(year, month, int(day))
            # MJD = JD - 2400000.5
            # Simplified handling, needs more precise calculation here
            
            # Parse RA (HH MM SS.S)
            ra_h = float(parts[4])
            ra_m = float(parts[5])
            ra_s = float(parts[6])
            ra_deg = (ra_h + ra_m/60 + ra_s/3600) * 15  # Convert to degrees
            
            # Parse Dec (+/-DD MM SS.S)
            dec_sign = 1 if parts[7][0] == '+' else -1
            dec_d = abs(float(parts[7]))
            dec_m = float(parts[8])
            dec_s = float(parts[9])
            dec_deg = dec_sign * (dec_d + dec_m/60 + dec_s/3600)
            
            # Magnitude and band (handle two formats: 19.8 G and 19.84G)
            mag = None
            band = 'V'
            
            if len(parts) >= 11:
                # Check if 11th field is magnitude (may contain band)
                mag_field = parts[10]
                # First check if pure numeric magnitude (no letters)
                if mag_field.replace('.', '', 1).isdigit() and mag_field.count('.') <= 1:
                    # Pure numeric magnitude
                    mag = float(mag_field)
                    if len(parts) >= 12:
                        # Check if 12th field is band (single letter or format like g1, i1)
                        if len(parts[11]) == 1 and parts[11].isalpha():
                            band = parts[11].upper()
                        elif len(parts[11]) == 2 and parts[11][0].isalpha() and parts[11][1].isdigit():
                            # Format like g1, i1, r1, etc., first character is band, second is number
                            band = parts[11][0].upper()
                else:
                    # Try to match magnitude+band format: 19.84G or 22.65gW~9HCuO18
                    mag_match = re.search(r'^(\d+\.\d+)([A-Za-z])(.*)$', mag_field)
                    if mag_match:
                        mag = float(mag_match.group(1))
                        band = mag_match.group(2).upper()
                    else:
                        # 11th field is not magnitude, mark as missing
                        mag = -1.0  # Use -1.0 to indicate missing magnitude
            
            # Create Observation object
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
        """Convert date to MJD (using digest2 library standard implementation)"""
        from digest2.observation import _date_to_mjd as digest_mjd
        return digest_mjd(year, month, day)
        
    def analyze(self):
        """Start analysis - prioritize analysis of observation data in input box"""
        input_data = self.input_text.get(1.0, tk.END).strip()
        file_path = self.file_path_var.get().strip()

        if input_data:
            # Prioritize analysis of observation data in input box
            self.analyze_input()
        elif file_path:
            # Input box is empty, use file analysis
            self.analyze_file()
        else:
            messagebox.showwarning("Warning", "Please enter observation data or select a file!")
    
    def analyze_file(self):
        """Analyze file"""
        file_path = self.file_path_var.get().strip()
        
        if not file_path:
            messagebox.showwarning("Warning", "Please select an observation file first!")
            return
            
        if not os.path.exists(file_path):
            messagebox.showerror("Error", f"File does not exist:\n{file_path}")
            return
            
        # Clear previous results
        self.clear_results()
        
        # Disable buttons
        self.analyze_btn.config(state='disabled')
        self.browse_btn.config(state='disabled')
        self.status_var.set("Analyzing...")
        self.root.update()
        
        try:
            with Digest2() as d2:
                results = d2.classify_file(file_path)
                
                if not results:
                    messagebox.showinfo("Info", "No valid observation data found in file.")
                    return
                    
                self.display_results(results)
                total_obs = sum(len(result.tracklet) for result in results)
                self.status_var.set(f"Analysis complete, {len(results)} objects, {total_obs} observations")
                
        except Exception as e:
            error_msg = str(e)
            if "Cannot find observatory codes" in error_msg or "digest2.obscodes" in error_msg or "ObsCodes.html" in error_msg or "observatory codes file" in error_msg:
                messagebox.showerror("Error", "Please download MPC observatory codes first")
            else:
                messagebox.showerror("Analysis Error", f"An error occurred during analysis:\n\n{error_msg}")
            self.status_var.set("Analysis failed")
        finally:
            # Re-enable buttons
            self.analyze_btn.config(state='normal')
            self.browse_btn.config(state='normal')
    
    def analyze_input(self):
        """Analyze manually entered data"""
        input_data = self.input_text.get(1.0, tk.END).strip()

        if not input_data:
            messagebox.showwarning("Warning", "Please enter observation data!")
            return

        # Clear previous results
        self.clear_results()

        # Disable buttons
        self.analyze_btn.config(state='disabled')
        self.status_var.set("Analyzing...")
        self.root.update()

        try:
            # Detect input format
            is_ades_xml = input_data.startswith('<?xml') or input_data.startswith('<ades') or '<optical' in input_data[:1000]
            
            # Detect ADES PSV format
            # PSV format characteristics:
            # 1. Has header lines starting with ! that contain |
            # 2. Data lines contain multiple | separators
            lines = input_data.split('\n')
            is_ades_psv = False
            has_psv_header = False
            has_data_with_pipes = False
            
            for line in lines:
                stripped_line = line.strip()
                if not stripped_line:
                    continue
                
                # Check if PSV header line (starts with ! and contains |)
                if stripped_line.startswith('!') and '|' in stripped_line:
                    has_psv_header = True
                    break
                
                # Check if data line (contains multiple |)
                if stripped_line.count('|') >= 2:
                    has_data_with_pipes = True
                    # Extra check: | not only at column 13 (MPC80 format program code position)
                    if len(stripped_line) > 12 and stripped_line[12] == '|':
                        # Check if there are | at other positions
                        if stripped_line.count('|') > 1:
                            has_data_with_pipes = True
                        else:
                            has_data_with_pipes = False
                    break
            
            is_ades_psv = has_psv_header or has_data_with_pipes

            if is_ades_xml:
                # Use ADES XML parser
                self._analyze_ades_xml_input(input_data)
            elif is_ades_psv:
                # Use ADES PSV parser
                self._analyze_ades_psv_input(input_data)
            else:
                # Use MPC80 format parser
                self._analyze_mpc80_input(input_data)
        except Exception as e:
            error_msg = str(e)
            if "Cannot find observatory codes" in error_msg or "digest2.obscodes" in error_msg or "ObsCodes.html" in error_msg or "observatory codes file" in error_msg:
                messagebox.showerror("Error", "Please download MPC observatory codes first")
            else:
                messagebox.showerror("Analysis Error", f"An error occurred during analysis:\n\n{error_msg}")
            self.status_var.set("Analysis failed")
        finally:
            # Re-enable buttons
            self.analyze_btn.config(state='normal')

    def _analyze_ades_xml_input(self, input_data):
        """Analyze manually entered ADES XML format data"""
        from digest2.observation import _parse_ades_optical
        from lxml import etree as ET
        import tempfile
        import os

        # Write input data to temp file
        try:
            with tempfile.NamedTemporaryFile(mode='w', suffix='.xml', delete=False, encoding='utf-8') as f:
                f.write(input_data)
                temp_path = f.name

            # Parse XML file, group by permID/provID
            tree = ET.parse(temp_path)
            root = tree.getroot()

            # Handle namespace
            ns = ""
            if root.tag.startswith("{"):
                ns = root.tag.split("}")[0] + "}"

            # Group observations by permID/provID
            tracklets = {}

            # Find all optical elements
            for optical in root.iter(f"{ns}optical"):
                # Parse observation data
                obs = _parse_ades_optical(optical, ns)
                if obs is None:
                    continue

                # Get permID or provID
                permID_el = optical.find(f"{ns}permID")
                provID_el = optical.find(f"{ns}provID")

                # Prioritize permID, use provID if not available
                desig = None
                if permID_el is not None and permID_el.text:
                    desig = permID_el.text.strip()
                elif provID_el is not None and provID_el.text:
                    desig = provID_el.text.strip()

                if desig:
                    if desig not in tracklets:
                        tracklets[desig] = []
                    tracklets[desig].append(obs)

            # Delete temp file
            os.unlink(temp_path)

            if not tracklets:
                messagebox.showinfo("Info", "No valid observation data found.")
                self.analyze_btn.config(state='normal')
                return

            # Sort observations by time for each object
            for desig in tracklets:
                tracklets[desig].sort(key=lambda o: o.mjd)

            # Analyze each object
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
                    msg = f"The following {len(skipped_tracklets)} objects were skipped due to insufficient time span:\n" + ", ".join(skipped_tracklets[:10])
                    if len(skipped_tracklets) > 10:
                        msg += f" and {len(skipped_tracklets)-10} more"
                    messagebox.showwarning("Analysis Warning", msg)

                if results:
                    self.display_results(results, use_custom_desig=True)
                    total_obs = sum(len(tracklets[desig]) for desig in tracklets if desig not in skipped_tracklets)
                    self.status_var.set(f"Analysis complete, {len(results)} objects, {total_obs} observations")
                else:
                    messagebox.showinfo("Info", "No objects with sufficient time span available for analysis.")

        except Exception as e:
            error_msg = str(e)
            if "Cannot find observatory codes" in error_msg or "digest2.obscodes" in error_msg or "ObsCodes.html" in error_msg or "observatory codes file" in error_msg:
                messagebox.showerror("Error", "Please download MPC observatory codes first")
            else:
                messagebox.showerror("Error", f"Cannot parse ADES XML data:\n{error_msg}")

        self.analyze_btn.config(state='normal')

    def _analyze_ades_psv_input(self, input_data):
        """Analyze manually entered ADES PSV format data"""
        from digest2.observation import parse_ades_psv, _parse_ades_psv_row
        from io import StringIO
        import csv

        # Split input data by lines, skip header lines (starting with ! or containing column names)
        lines = []
        for line in input_data.split('\n'):
            line = line.rstrip()
            if not line:
                continue
            # Skip lines starting with ! (ADES header marker)
            if line.startswith('!'):
                continue
            # Skip lines that look like headers (containing column names like provID, mode, stn, obsTime, etc.)
            if 'provID' in line or 'obsTime' in line or ('mode' in line and 'stn' in line and 'ra' in line and 'dec' in line):
                continue
            lines.append(line)

        if not lines:
            messagebox.showinfo("Info", "No valid ADES PSV data found.")
            self.analyze_btn.config(state='normal')
            return

        # Parse ADES PSV data
        tracklets = {}
        failed_lines = []

        for line in lines:
            parts = [p.strip() for p in line.split('|')]
            if parts and parts[-1] == '':
                parts = parts[:-1]

            # Try to parse each line
            try:
                # Construct dictionary format
                headers = ['provID', 'mode', 'stn', 'obsTime', 'ra', 'dec', 'rmsRA', 'rmsDec',
                          'astCat', 'photCat', 'mag', 'rmsMag', 'band', 'photAp', 'logSNR',
                          'seeing', 'exp', 'rmsFit', 'nStars', 'ref', 'disc', 'subFmt',
                          'precTime', 'precRA', 'precDec', 'notes', 'remarks']
                row_dict = {}
                for i, part in enumerate(parts):
                    if i < len(headers):
                        row_dict[headers[i]] = part

                # Parse observation data
                obs = _parse_ades_psv_row(row_dict)
                if obs is not None:
                    # Get designation
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
            messagebox.showwarning("Parse Warning", f"The following {len(failed_lines)} lines could not be parsed:\n" + "\n".join(failed_lines[:5]))

        if not tracklets:
            messagebox.showinfo("Info", "No valid observation data found.")
            self.analyze_btn.config(state='normal')
            return

        # Sort observations by time for each object
        for desig in tracklets:
            tracklets[desig].sort(key=lambda o: o.mjd)

        # Analyze each object
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
                msg = f"The following {len(skipped_tracklets)} objects were skipped due to insufficient time span:\n" + ", ".join(skipped_tracklets[:10])
                if len(skipped_tracklets) > 10:
                    msg += f" and {len(skipped_tracklets)-10} more"
                messagebox.showwarning("Analysis Warning", msg)

            if results:
                self.display_results(results, use_custom_desig=True)
                total_obs = sum(len(tracklets[desig]) for desig in tracklets if desig not in skipped_tracklets)
                self.status_var.set(f"Analysis complete, {len(results)} objects, {total_obs} observations")
            else:
                messagebox.showinfo("Info", "No objects with sufficient time span available for analysis.")

        self.analyze_btn.config(state='normal')

    def _analyze_mpc80_input(self, input_data):
        """Analyze manually entered MPC80 format data"""
        # Parse observation data
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
            messagebox.showwarning("Parse Warning", f"The following {len(failed_lines)} lines could not be parsed:\n" + "\n".join(failed_lines[:5]))

        if not observations:
            messagebox.showinfo("Info", "No valid observation data found.\nPlease ensure you entered MPC 80-column format or simplified format.")
            self.analyze_btn.config(state='normal')
            return

        # Group by designation
        tracklets = {}
        for desig, obs in observations:
            if desig not in tracklets:
                tracklets[desig] = []
            tracklets[desig].append(obs)

        # Sort observations by time for each object (increasing MJD)
        for desig in tracklets:
            tracklets[desig].sort(key=lambda o: o.mjd)

        # Analyze each object
        with Digest2() as d2:
            results = []
            skipped_tracklets = []
            self._desig_map = {}  # Clear mapping
            for desig, obs_list in tracklets.items():
                try:
                    result = d2.classify_tracklet(obs_list)
                    # Save designation to mapping (use id as key)
                    self._desig_map[id(result)] = desig
                    results.append(result)
                except Exception as e:
                    if "need >=2 obs with motion and time span" in str(e):
                        skipped_tracklets.append(desig)
                    else:
                        raise  # Re-raise other exceptions

            if skipped_tracklets:
                msg = f"The following {len(skipped_tracklets)} objects were skipped due to insufficient time span:\n" + ", ".join(skipped_tracklets[:10])
                if len(skipped_tracklets) > 10:
                    msg += f" and {len(skipped_tracklets)-10} more"
                messagebox.showwarning("Analysis Warning", msg)

            if results:
                self.display_results(results, use_custom_desig=True)
                total_obs = sum(len(tracklets[desig]) for desig in tracklets if desig not in skipped_tracklets)
                self.status_var.set(f"Analysis complete, {len(results)} objects, {total_obs} observations")
            else:
                messagebox.showinfo("Info", "No objects with sufficient time span available for analysis.\nEach object needs at least 2 observations with sufficient time span.")

        self.analyze_btn.config(state='normal')

    def clear_input_data(self):
        """Clear observation data (input box and file path)"""
        self.input_text.delete(1.0, tk.END)
        self.file_path_var.set("")
        self.status_var.set("Observation data cleared")
    
    def download_obscode(self):
        """Download MPC observatory codes file"""
        import urllib.request
        import os
        import sys
        
        url = "https://minorplanetcenter.net/iau/lists/ObsCodes.html"
        
        # Get actual program running directory (supports exe packaging)
        if getattr(sys, 'frozen', False):
            # Running in packaged exe
            app_dir = os.path.dirname(sys.executable)
        else:
            # Running in Python script
            app_dir = os.path.dirname(__file__)
        
        save_path = os.path.join(app_dir, "ObsCodes.html")
        
        # Set timeout to 10 seconds
        timeout = 10
        
        try:
            self.status_var.set("Downloading observatory codes...")
            self.root.update_idletasks()  # Update UI display
            
            # Create request object and set timeout
            req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req, timeout=timeout) as response:
                with open(save_path, 'wb') as f:
                    f.write(response.read())
            
            self.status_var.set(f"Download successful! File saved to: {save_path}")
        except urllib.error.URLError as e:
            self.status_var.set(f"Network error: {str(e)}")
        except timeout:
            self.status_var.set("Download timeout, please try again later")
        except Exception as e:
            self.status_var.set(f"Download failed: {str(e)}")

    def display_results(self, results, use_custom_desig=False):
        """Display analysis results

        Args:
            results: Analysis results list
            use_custom_desig: Whether to use custom designation (for manual input mode)
        """
        self.results_data = results  # Save results for detail view

        for r in results:
            # Get various scores
            scores = r.noid
            # Use custom designation (if available)
            if use_custom_desig and id(r) in self._desig_map:
                desig = self._desig_map[id(r)]
            else:
                desig = r.designation

            # Determine if highlighting needed (NEO score >= 65)
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
        """Clear results"""
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.status_var.set("Ready")
        self.results_data = []
        self._desig_map = {}

    def show_text_context_menu(self, event):
        """Show right-click menu for text input box"""
        # Create right-click menu
        context_menu = tk.Menu(self.root, tearoff=0)
        context_menu.add_command(label="Cut", command=self.cut_text)
        context_menu.add_command(label="Copy", command=self.copy_text)
        context_menu.add_command(label="Paste", command=self.paste_text)
        context_menu.add_command(label="Select All", command=self.select_all_text)

        # Show menu at mouse position
        context_menu.tk_popup(event.x_root, event.y_root)

    def copy_text(self):
        """Copy selected text to clipboard"""
        try:
            # First set focus to input box
            self.input_text.focus_set()
            selected_text = self.input_text.get(tk.SEL_FIRST, tk.SEL_LAST)
            if selected_text:
                self.root.clipboard_clear()
                self.root.clipboard_append(selected_text)
                self.status_var.set("Copied")
        except tk.TclError:
            pass  # No selected text

    def paste_text(self):
        """Paste text from clipboard (replace selected content)"""
        try:
            clipboard_text = self.root.clipboard_get()
            # First set focus to input box
            self.input_text.focus_set()
            # Check if there is selected text
            try:
                self.input_text.delete(tk.SEL_FIRST, tk.SEL_LAST)
            except tk.TclError:
                pass  # No selected text
            # Insert at current cursor position
            self.input_text.insert(tk.INSERT, clipboard_text)
            self.status_var.set("Pasted")
        except tk.TclError:
            pass  # Clipboard empty

    def cut_text(self):
        """Cut selected text"""
        try:
            # First set focus to input box
            self.input_text.focus_set()
            selected_text = self.input_text.get(tk.SEL_FIRST, tk.SEL_LAST)
            if selected_text:
                self.root.clipboard_clear()
                self.root.clipboard_append(selected_text)
                self.input_text.delete(tk.SEL_FIRST, tk.SEL_LAST)
                self.status_var.set("Cut")
        except tk.TclError:
            pass  # No selected text

    def select_all_text(self):
        """Select all text"""
        # First set focus to input box
        self.input_text.focus_set()
        self.input_text.tag_add(tk.SEL, "1.0", tk.END)
        self.input_text.mark_set(tk.INSERT, "1.0")
        self.input_text.see(tk.INSERT)

    def show_tree_context_menu(self, event):
        """Show right-click menu for tree view"""
        # Get row at mouse position
        row_id = self.tree.identify_row(event.y)
        if row_id:
            # If no rows selected, first select row at mouse
            selected_items = self.tree.selection()
            if not selected_items or row_id not in selected_items:
                self.tree.selection_set(row_id)
        
        # Create right-click menu
        context_menu = tk.Menu(self.root, tearoff=0)
        # Only show "Copy Selected" when there are selected rows
        if self.tree.selection():
            context_menu.add_command(label="Copy Selected", command=self.copy_selected_item)
        # Always show "Copy All"
        context_menu.add_command(label="Copy All", command=self.copy_all_data)
        
        # Show menu at mouse position
        context_menu.tk_popup(event.x_root, event.y_root)

    def copy_selected_item(self):
        """Copy selected row data to clipboard"""
        selected_items = self.tree.selection()
        if not selected_items:
            return
        
        # Get all selected row data
        rows = []
        for item in selected_items:
            values = self.tree.item(item, 'values')
            rows.append('\t'.join(str(v) for v in values))
        
        # Format data (tab-separated, easy to paste to spreadsheet)
        text = '\n'.join(rows)
        
        # Copy to clipboard
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        count = len(selected_items)
        self.status_var.set(f"Copied {count} rows")

    def copy_all_data(self):
        """Copy all result data to clipboard"""
        # Get all row data
        rows = []
        # Add headers
        headers = ['Designation', 'RMS', 'Int', 'NEO', 'N22', 'N18', 'MC', 'MB1', 'MB2', 'MB3', 'Hun', 'Pho', 'Pal', 'Han', 'Hil', 'JTr', 'JFC']
        rows.append('\t'.join(headers))
        
        # Add data rows
        for item in self.tree.get_children():
            values = self.tree.item(item, 'values')
            rows.append('\t'.join(str(v) for v in values))
        
        # Merge to text
        text = '\n'.join(rows)
        
        # Copy to clipboard
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.status_var.set("Copied all data")


def main():
    root = tk.Tk()
    app = Digest2GUI(root)
    root.mainloop()


if __name__ == '__main__':
    main()