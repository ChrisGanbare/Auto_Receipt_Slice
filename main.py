import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import fitz  # PyMuPDF
import pdfplumber  # 用于表格提取
import re
import os
import threading
from datetime import datetime
import csv
from operator import itemgetter
import queue

# --- Pre-compiled Regular Expressions for Performance and Maintainability ---
# Regex for a 20-digit receipt number
RECEIPT_NO_REGEX_20 = re.compile(r'(\d{20})')
# Regex for finding a 20-digit number after the label
RECEIPT_NO_LABEL_REGEX_20 = re.compile(r'回单编号[：:\s]*(\d{20})')

class ReceiptSplitterApp:
    """
    农行电子回单智能拆分工具主应用程序类
    
    提供图形用户界面，用于自动识别和拆分中国农业银行的电子回单PDF文件。
    主要功能包括：
    - 自动识别PDF中的回单区域
    - 提取回单信息（客户名称、回单编号、金额等）
    - 预览和编辑回单信息
    - 将多个回单拆分为独立的PDF文件
    - 生成处理日志
    
    使用tkinter构建GUI界面，使用PyMuPDF和pdfplumber处理PDF文件。
    """
    def __init__(self, root):
        """
        初始化应用程序主窗口和界面组件
        
        :param root: tkinter根窗口对象，用于创建应用程序的主窗口
        """
        self.root = root
        self.root.title("农行电子回单智能拆分工具 V1.0.0")
        self.root.geometry("1200x700")
        # 设置窗口图标（如果有图标文件）
        try:
            icon_path = os.path.join(os.path.dirname(__file__), 'icon.ico')
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except Exception:
            pass  # 如果图标文件不存在或加载失败，忽略错误

        self.source_file = ""
        self.doc = None
        self.preview_data = []
        self.preview_image = None
        self.preview_image_ref = None  # 保持图片引用，防止垃圾回收
        self.placeholder_text = "若付款方为我方公司，则取对手方(收款方)户名为客户名称，若留空则默认使用付款方户名作为客户名称"
        self.update_queue = queue.Queue()  # 用于线程安全的GUI更新
        self.check_queue()  # 启动队列检查

        frame_top = ttk.LabelFrame(root, text="操作面板", padding=10)
        frame_top.pack(fill="x", padx=10, pady=5)
        frame_top.columnconfigure(1, weight=1)

        self.btn_load = ttk.Button(frame_top, text="1. 选择PDF源文件", command=self.load_file)
        self.btn_load.grid(row=0, column=0, padx=(0, 5), sticky="w")
        self.lbl_file = ttk.Label(frame_top, text="未选择文件", foreground="gray", anchor="w")
        self.lbl_file.grid(row=0, column=1, padx=5, sticky="ew")
        self.btn_process = ttk.Button(frame_top, text="2. 开始拆分导出", command=self.start_processing, state="disabled")
        self.btn_process.grid(row=0, column=2, padx=(5, 0), sticky="e")

        # 电子回单本方公司户名选择区域（初始隐藏）
        self.local_company_frame = ttk.Frame(frame_top)
        self.lbl_local_company = ttk.Label(self.local_company_frame, text="电子回单本方公司户名（可选）:")
        self.lbl_local_company.grid(row=0, column=0, padx=(0, 5), sticky="w")
        self.combo_local_company = ttk.Combobox(self.local_company_frame, state="readonly", width=40)
        self.combo_local_company.grid(row=0, column=1, padx=5, sticky="ew")
        # 绑定选择事件，当选择非默认值时显示确认按钮
        self.combo_local_company.bind("<<ComboboxSelected>>", self.on_company_selected)
        self.local_company_frame.columnconfigure(1, weight=1)
        
        # 确认更新按钮（初始隐藏，只有选择了非默认值才显示）
        self.btn_confirm_company = ttk.Button(self.local_company_frame, text="确认更新", command=self.confirm_company_name)
        # 按钮初始不显示，通过grid_remove隐藏（保留布局信息）
        
        # 提示标签
        self.lbl_hint = ttk.Label(self.local_company_frame, 
                                  text="💡 提示：选择本方公司户名并确认更新后，系统将更新对应记录的客户名称为收款方户名；不选择则默认使用付款方户名作为客户名称", 
                                  foreground="blue", font=("Arial", 9))
        self.lbl_hint.grid(row=1, column=0, columnspan=2, padx=5, pady=(5, 0), sticky="w")

        main_pane = ttk.PanedWindow(root, orient=tk.HORIZONTAL)
        main_pane.pack(fill="both", expand=True, padx=10, pady=5)

        frame_left = ttk.LabelFrame(main_pane, text="解析预览 (单击查看原文, 双击可修改)", padding=10)
        main_pane.add(frame_left, weight=2)

        columns = ("seq", "name", "receipt_no", "amount", "status")
        self.tree = ttk.Treeview(frame_left, columns=columns, show="headings", selectmode="browse")
        self.tree.heading("seq", text="序号")
        self.tree.heading("name", text="客户名称")
        self.tree.heading("receipt_no", text="回单编号")
        self.tree.heading("amount", text="金额")
        self.tree.heading("status", text="状态")
        self.tree.column("seq", width=40, anchor="center")
        self.tree.column("name", width=200)
        self.tree.column("receipt_no", width=150)
        self.tree.column("amount", width=80, anchor="e")
        self.tree.column("status", width=60, anchor="center")

        self.tree.bind("<Double-1>", self.open_edit_window)
        self.tree.bind("<<TreeviewSelect>>", self.show_receipt_preview)

        scrollbar = ttk.Scrollbar(frame_left, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        frame_right = ttk.LabelFrame(main_pane, text="回单原文预览 (下方文本可直接选中复制)", padding=10)
        main_pane.add(frame_right, weight=3)

        # 创建内部上下拆分的 PanedWindow
        preview_splitter = ttk.PanedWindow(frame_right, orient=tk.VERTICAL)
        preview_splitter.pack(fill="both", expand=True)

        # --- 上半部分：图片预览 ---
        preview_container = ttk.Frame(preview_splitter)
        preview_splitter.add(preview_container, weight=4)  # 图片占主要部分
        
        # 创建Canvas用于显示图片和滚动
        self.preview_canvas = tk.Canvas(preview_container, bg="white", highlightthickness=0)
        
        # 创建垂直滚动条
        v_scrollbar = ttk.Scrollbar(preview_container, orient="vertical", command=self.preview_canvas.yview)
        self.preview_canvas.configure(yscrollcommand=v_scrollbar.set)
        
        # 创建水平滚动条
        h_scrollbar = ttk.Scrollbar(preview_container, orient="horizontal", command=self.preview_canvas.xview)
        self.preview_canvas.configure(xscrollcommand=h_scrollbar.set)
        
        # 布局：Canvas在中间，滚动条在边缘
        self.preview_canvas.grid(row=0, column=0, sticky="nsew")
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar.grid(row=1, column=0, sticky="ew")
        
        preview_container.grid_rowconfigure(0, weight=1)
        preview_container.grid_columnconfigure(0, weight=1)
        
        # 在Canvas上创建图片容器
        self.preview_image_container = self.preview_canvas.create_image(0, 0, anchor="nw")
        
        # 初始提示文本
        self.preview_canvas.create_text(200, 100, text="请在左侧选择一条记录以预览", anchor="center", fill="gray")

        # --- 下半部分：文本复制区 (新增) ---
        text_container = ttk.Frame(preview_splitter)
        preview_splitter.add(text_container, weight=1)  # 文本区占较小部分
        
        self.txt_extract = tk.Text(text_container, height=6, font=("Microsoft YaHei", 10), 
                                  undo=True, wrap="word", bg="#f8f9fa", 
                                  state="normal", selectbackground="#316AC5", 
                                  selectforeground="white")
        txt_scroll = ttk.Scrollbar(text_container, orient="vertical", command=self.txt_extract.yview)
        self.txt_extract.configure(yscrollcommand=txt_scroll.set)
        
        self.txt_extract.pack(side="left", fill="both", expand=True)
        txt_scroll.pack(side="right", fill="y")
        
        # 确保文本可以选择和复制（绑定右键菜单）
        def show_context_menu(event):
            self.txt_extract.focus_set()  # 弹出菜单前先获取焦点
            context_menu = tk.Menu(self.root, tearoff=0)
            context_menu.add_command(label="复制 (Ctrl+C)", command=lambda: self.txt_extract.event_generate("<<Copy>>"))
            context_menu.add_command(label="全选 (Ctrl+A)",
                                     command=lambda: self.txt_extract.tag_add("sel", "1.0", tk.END))
            context_menu.post(event.x_root, event.y_root)
        
        self.txt_extract.bind("<Button-3>", show_context_menu)  # 右键菜单
        
        # --- 新增：拦截所有修改操作，使其变为只读但可选中 ---
        def disable_editing(event):
            # 允许 Ctrl+C (复制) 和 Ctrl+A (全选)
            if event.state & 0x0004 and event.keysym.lower() in ('c', 'a'):
                return None
            # 拦截其他所有按键输入（退格、删除、回车、普通字母等）
            return "break"
        
        self.txt_extract.bind("<Key>", disable_editing)
        self.txt_extract.bind("<<Cut>>", lambda e: "break")  # 显式禁用剪切
        self.txt_extract.bind("<<Paste>>", lambda e: "break")  # 显式禁用粘贴
        
        # 初始提示
        self.txt_extract.insert("1.0", "选中左侧记录后，此处将显示可复制的原文文本...")
        
        # 绑定鼠标滚轮事件（支持垂直和水平滚动）
        def on_mousewheel(event):
            # 垂直滚动
            if event.delta:
                self.preview_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            else:
                # Linux系统
                if event.num == 4:
                    self.preview_canvas.yview_scroll(-1, "units")
                elif event.num == 5:
                    self.preview_canvas.yview_scroll(1, "units")
        
        def on_shift_mousewheel(event):
            # Shift+滚轮：水平滚动
            if event.delta:
                self.preview_canvas.xview_scroll(int(-1 * (event.delta / 120)), "units")
            else:
                # Linux系统
                if event.num == 4:
                    self.preview_canvas.xview_scroll(-1, "units")
                elif event.num == 5:
                    self.preview_canvas.xview_scroll(1, "units")
        
        # 绑定滚轮事件
        self.preview_canvas.bind("<MouseWheel>", on_mousewheel)
        self.preview_canvas.bind("<Shift-MouseWheel>", on_shift_mousewheel)
        # Linux系统
        self.preview_canvas.bind("<Button-4>", on_mousewheel)
        self.preview_canvas.bind("<Button-5>", on_mousewheel)
        
        # 设置Canvas可获得焦点，以便接收键盘事件
        self.preview_canvas.focus_set()

        frame_bottom = ttk.Frame(root)
        frame_bottom.pack(fill="x", side="bottom", padx=10, pady=5)
        frame_bottom.columnconfigure(0, weight=1)

        self.lbl_status = ttk.Label(frame_bottom, text="就绪", anchor="w")
        self.lbl_status.grid(row=0, column=0, sticky="ew")

        self.progress_bar = ttk.Progressbar(frame_bottom, orient="horizontal", mode="determinate")
        self.progress_bar.grid(row=0, column=1, sticky="e")
        
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        # 在 __init__ 底部修改
        def handle_root_click(event):
            # 只有点击的既不是下拉框，也不是文本预览框时，才将焦点转移回 root
            if event.widget != self.combo_local_company and event.widget != self.txt_extract:
                self.root.focus_set()

        self.root.bind("<Button-1>", handle_root_click)
        
        # 存储付款方和收款方户名信息
        self.payer_names = []  # 存储所有付款方户名
        self.receiver_names_map = {}  # 存储付款方户名到收款方户名的映射

    def check_queue(self):
        """
        检查队列中的GUI更新请求（线程安全）
        
        定期检查更新队列，执行从后台线程提交的GUI更新操作。
        每100毫秒检查一次，确保后台线程可以安全地更新界面。
        这是一个递归调用，通过root.after实现定时检查。
        """
        try:
            while True:
                callback, args = self.update_queue.get_nowait()
                callback(*args)
        except queue.Empty:
            pass
        finally:
            self.root.after(100, self.check_queue)  # 每100ms检查一次

    def safe_gui_update(self, callback, *args):
        """
        线程安全的GUI更新方法
        
        将GUI更新操作放入队列，由主线程的check_queue方法执行，确保线程安全。
        用于从后台线程安全地更新GUI界面。
        
        :param callback: 要执行的回调函数（应在主线程中执行）
        :param args: 传递给回调函数的参数
        """
        self.update_queue.put((callback, args))

    def on_closing(self):
        """
        关闭窗口时的清理工作
        
        在用户关闭程序窗口时调用，负责关闭PDF文档对象，
        释放资源，然后销毁主窗口。
        """
        try:
            if self.doc:
                self.doc.close()
        except Exception:
            pass
        self.root.destroy()

    def show_receipt_preview(self, event):
        """
        显示选中回单的预览（图片和文本）
        
        当用户在左侧列表中选择一条记录时触发，在右侧预览区域显示：
        1. 回单的图片预览（Canvas显示）
        2. 回单的可复制文本内容（Text组件显示）
        
        :param event: tkinter事件对象，由Treeview的<<TreeviewSelect>>事件触发
        """
        item_id = self.tree.focus()
        if not item_id or not self.doc:
            return
        
        # 获取选中项的序号
        try:
            seq = int(self.tree.item(item_id, 'values')[0])
        except (ValueError, IndexError):
            return
        
        # 优先通过item_id查找，如果没有则通过seq查找
        item_data = None
        for item in self.preview_data:
            if 'item_id' in item and item['item_id'] == item_id:
                item_data = item
                break
        
        # 如果通过item_id没找到，则通过seq查找
        if item_data is None:
            item_data = next((item for item in self.preview_data if item['seq'] == seq), None)
        
        if not item_data:
            return

        try:
            page = self.doc[item_data['page_idx']]
            # 使用rect坐标裁剪预览区域
            crop_rect = fitz.Rect(item_data['rect'])
            
            # 确保rect在页面范围内
            page_rect = page.rect
            crop_rect = crop_rect & page_rect
            
            # --- 1. 更新图片预览 (原有逻辑) ---
            # 生成裁剪后的预览图片
            pix = page.get_pixmap(dpi=150, clip=crop_rect)
            img_data = pix.tobytes("ppm")
            
            # 保存图片引用，防止被垃圾回收
            self.preview_image = tk.PhotoImage(data=img_data)
            self.preview_image_ref = self.preview_image  # 保持引用
            
            # 更新Canvas上的图片
            self.preview_canvas.delete("all")  # 清除之前的内容
            self.preview_image_container = self.preview_canvas.create_image(0, 0, anchor="nw", image=self.preview_image)
            
            # 更新Canvas的滚动区域
            self.preview_canvas.configure(scrollregion=self.preview_canvas.bbox("all"))

            # --- 2. 更新文本复制区 (新增逻辑) ---
            try:
                # 提取裁剪区域内的所有文本
                raw_text = page.get_text("text", clip=crop_rect)
                
                # 清理文本：去除多余空格和空行，方便用户选择
                if raw_text and raw_text.strip():
                    lines = [line.strip() for line in raw_text.split('\n') if line.strip()]
                    clean_text = "\n".join(lines)
                else:
                    clean_text = ""
                
                # 更新文本内容（insert方法不会触发Key事件，所以不受disable_editing影响）
                self.txt_extract.config(state="normal")
                self.txt_extract.delete("1.0", tk.END)
                if clean_text:
                    self.txt_extract.insert("1.0", clean_text)
                else:
                    self.txt_extract.insert("1.0", "（未提取到文本内容）")
                
                # 将光标移到开头，方便用户选择
                self.txt_extract.mark_set("insert", "1.0")
                self.txt_extract.see("1.0")
                
            except Exception as text_error:
                # 如果文本提取失败，显示错误信息
                import traceback
                error_detail = traceback.format_exc()
                self.txt_extract.config(state="normal")
                self.txt_extract.delete("1.0", tk.END)
                self.txt_extract.insert("1.0", f"文本提取失败: {str(text_error)}\n\n详细信息:\n{error_detail}")
                self.log(f"文本提取失败: {text_error}")
            
        except Exception as e:
            # 显示错误信息
            self.preview_canvas.delete("all")
            self.preview_canvas.create_text(200, 100, text=f"无法生成预览:\n{str(e)}", anchor="center", fill="red")
            self.preview_canvas.configure(scrollregion=self.preview_canvas.bbox("all"))
            self.log(f"生成预览失败: {e}")
            
            # 更新文本区域显示错误
            self.txt_extract.config(state="normal")
            self.txt_extract.delete("1.0", tk.END)
            self.txt_extract.insert("1.0", f"文本提取失败: {str(e)}")

    def open_edit_window(self, event):
        """
        打开编辑窗口，允许用户修改回单信息
        
        当用户双击左侧列表中的记录时触发，弹出编辑对话框，
        可以修改客户名称、回单编号和金额。
        
        :param event: tkinter事件对象，由Treeview的<Double-1>事件触发
        """
        item_id = self.tree.focus()
        if not item_id: return
        seq = int(self.tree.item(item_id, 'values')[0])
        item_to_edit = next((item for item in self.preview_data if item['seq'] == seq), None)
        if not item_to_edit: return

        edit_win = tk.Toplevel(self.root)
        edit_win.title("修改记录")
        edit_win.geometry("400x200")
        edit_win.transient(self.root)
        edit_win.grab_set()

        frame = ttk.Frame(edit_win, padding=15)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="客户名称:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        name_entry = ttk.Entry(frame)
        name_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        name_entry.insert(0, item_to_edit['name'])
        
        # 绑定粘贴事件（Ctrl+V），自动清理换行符
        def on_paste_name(event):
            # 获取剪贴板内容
            try:
                clipboard_text = self.root.clipboard_get()
                # 去除换行符和回车符，替换为空格
                cleaned_text = clipboard_text.replace('\n', ' ').replace('\r', ' ').replace('\t', ' ')
                # 将多个连续空格替换为单个空格
                cleaned_text = re.sub(r'\s+', ' ', cleaned_text).strip()
                # 插入清理后的文本
                name_entry.delete(0, tk.END)
                name_entry.insert(0, cleaned_text)
                return "break"  # 阻止默认粘贴行为
            except tk.TclError:
                # 如果剪贴板为空或无法获取，允许默认行为
                return None
        
        # 绑定Ctrl+V和粘贴事件
        name_entry.bind("<Control-v>", on_paste_name)
        name_entry.bind("<Control-V>", on_paste_name)
        name_entry.bind("<<Paste>>", on_paste_name)

        ttk.Label(frame, text="回单编号:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        no_entry = ttk.Entry(frame)
        no_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        no_entry.insert(0, item_to_edit['no'])

        ttk.Label(frame, text="金额:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        amt_entry = ttk.Entry(frame)
        amt_entry.grid(row=2, column=1, padx=5, pady=5, sticky="ew")
        amt_entry.insert(0, item_to_edit['amt'])

        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=3, column=0, columnspan=2, pady=10)
        save_btn = ttk.Button(btn_frame, text="保存", command=lambda: self.save_edits(edit_win, item_id, seq, name_entry.get(), no_entry.get(), amt_entry.get()))
        save_btn.pack(side="left", padx=10)
        cancel_btn = ttk.Button(btn_frame, text="取消", command=edit_win.destroy)
        cancel_btn.pack(side="left", padx=10)

    def save_edits(self, edit_win, item_id, seq, new_name, new_no, new_amt):
        """
        保存编辑后的回单信息
        
        验证并保存用户编辑的回单数据，更新内存中的数据和界面显示。
        会对金额格式进行验证，确保格式正确（如：123.45）。
        
        :param edit_win: 编辑窗口对象，保存后关闭此窗口
        :param item_id: 树视图中的项目ID
        :param seq: 回单序号
        :param new_name: 新的客户名称
        :param new_no: 新的回单编号
        :param new_amt: 新的金额（字符串格式，如"123.45"）
        """
        cleaned_name = self.clean_filename(new_name)
        cleaned_amt = new_amt.replace(",", "").strip()
        
        # 验证金额格式
        try:
            float(cleaned_amt)
            if not re.match(r'^\d+(\.\d{1,2})?$', cleaned_amt):
                messagebox.showwarning("警告", "金额格式不正确，应为数字（如：123.45）")
                return
        except ValueError:
            messagebox.showwarning("警告", "金额格式不正确，应为数字（如：123.45）")
            return
        
        for item in self.preview_data:
            if item['seq'] == seq:
                item['name'] = cleaned_name
                item['no'] = new_no
                item['amt'] = cleaned_amt
                break
        self.tree.item(item_id, values=(seq, cleaned_name, new_no, cleaned_amt, "已修正"))
        edit_win.destroy()
        self.log(f"序号 {seq} 的记录已更新。")

    def on_company_selected(self, event=None):
        """
        当下拉列表选择改变时触发
        
        当用户选择或更改"电子回单本方公司户名"下拉列表的选项时调用。
        如果选择了非默认值，显示"确认更新"按钮；如果选择默认值，隐藏该按钮。
        
        :param event: tkinter事件对象（可选），由ComboboxSelected事件触发
        """
        selected_value = self.combo_local_company.get()
        default_text = "使用付款方户名作为客户名称（默认值）"
        
        # 如果选择了非默认值，显示确认更新按钮
        if selected_value and selected_value != default_text:
            self.btn_confirm_company.grid(row=0, column=2, padx=(5, 0), sticky="e")
        else:
            # 如果选择的是默认值或清空，隐藏确认按钮
            self.btn_confirm_company.grid_remove()
    
    def confirm_company_name(self):
        """
        确认选择的公司户名，更新预览列表
        
        当用户点击"确认更新"按钮时调用，将所有匹配的记录的客户名称
        更新为对应的收款方户名，并更新状态为"已更新"。
        如果付款方是选中的公司户名，则将客户名称改为收款方户名。
        """
        selected_company = self.combo_local_company.get()
        default_text = "使用付款方户名作为客户名称（默认值）"
        
        if not selected_company or selected_company == default_text:
            messagebox.showwarning("提示", "请先选择电子回单本方公司户名（不能选择默认值）")
            return
        
        # 更新所有匹配的记录
        updated_count = 0
        for item in self.preview_data:
            # 如果付款方户名匹配选中的公司户名，则更新客户名称为收款方户名
            if 'payer_name' in item and item['payer_name'] == selected_company:
                if 'receiver_name' in item and item['receiver_name']:
                    new_name = self.clean_filename(item['receiver_name'])
                    item['name'] = new_name
                    # 更新树视图
                    if 'item_id' in item:
                        current_values = list(self.tree.item(item['item_id'], 'values'))
                        current_values[1] = new_name  # 更新客户名称
                        current_values[4] = "已更新"  # 更新状态
                        self.tree.item(item['item_id'], values=tuple(current_values))
                    updated_count += 1
        
        if updated_count > 0:
            self.log(f"已更新 {updated_count} 条记录的客户名称和状态")
            messagebox.showinfo("成功", f"已成功更新 {updated_count} 条记录！\n客户名称已更新为对应的收款方户名，状态已标记为'已更新'。")
            # 更新后隐藏确认按钮
            self.btn_confirm_company.grid_remove()
        else:
            messagebox.showinfo("提示", f"未找到付款方为'{selected_company}'的记录，无需更新。")

    def log(self, message):
        """
        在状态栏显示日志消息
        
        在界面底部状态栏显示带时间戳的消息，用于向用户反馈程序运行状态。
        
        :param message: 要显示的日志消息字符串
        """
        self.lbl_status.config(text=f"[{datetime.now().strftime('%H:%M:%S')}] {message}")
        self.root.update_idletasks()

    def load_file(self):
        """
        加载PDF文件并开始分析
        
        弹出文件选择对话框，让用户选择要处理的PDF文件。
        选择文件后，会打开PDF文档并在后台线程中开始分析回单内容。
        """
        file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if not file_path:
            return
        try:
            if self.doc:
                self.doc.close()
        except Exception:
            pass
        self.source_file = file_path
        self.doc = fitz.open(file_path)
        self.lbl_file.config(text=os.path.basename(file_path), foreground="black")
        # 显示公司户名选择区域（放在第二行，与"开始拆分导出"按钮分开，视觉上更清晰）
        self.local_company_frame.grid(row=1, column=0, columnspan=3, padx=0, pady=(10, 0), sticky="ew")
        # 确保确认按钮初始隐藏
        self.btn_confirm_company.grid_remove()
        # 在主线程中获取公司户名，避免线程安全问题
        local_company_name = self.combo_local_company.get().strip() if hasattr(self, 'combo_local_company') else ""
        self.log("正在分析文件，请稍候...")
        threading.Thread(target=self.analyze_pdf, args=(local_company_name,), daemon=True).start()

    def clean_filename(self, text):
        """
        清理文本，使其适合用作文件名
        
        去除换行符、回车符、制表符，以及Windows文件系统不允许的字符，
        确保生成的文件名合法且可读。
        
        :param text: 原始文本字符串
        :return: 清理后的文本字符串，去除非法字符和多余的空白
        """
        # 先去除换行符和回车符，再去除文件系统不允许的字符
        text = text.replace('\n', ' ').replace('\r', ' ').replace('\t', ' ')
        # 将多个连续空格替换为单个空格
        text = re.sub(r'\s+', ' ', text)
        return re.sub(r'[\\/*?:"<>|]', "", text).strip()

    def is_valid_abc_receipt(self, doc, check_limit=3):
        """
        极速检测是否为农行回单
        
        通过检查PDF前几页是否包含农行回单的特征关键词来判断。
        关键词包括："中国农业银行"、"电子回单"、"回单编号"。
        如果一页内匹配到2个以上关键词，判定为农行回单格式。
        
        :param doc: fitz.Document对象，要检查的PDF文档
        :param check_limit: 最多检查前几页，默认3页
        :return: 元组(bool, message)，(True, "验证通过") 或 (False, 错误信息)
        """
        # 1. 基础指纹关键词
        fingerprints = ["中国农业银行", "电子回单", "回单编号"]

        # 2. 如果文档总页数比限制少，则按实际页数检查
        actual_limit = min(len(doc), check_limit)

        found_any_feature = False
        for i in range(actual_limit):
            page_text = doc[i].get_text()
            # 统计匹配到的关键词数量
            match_count = sum(1 for word in fingerprints if word in page_text)

            # 如果一页内匹配到2个以上关键词，基本可以判定是目标格式
            if match_count >= 2:
                found_any_feature = True
                break

        if not found_any_feature:
            return False, f"在前 {actual_limit} 页中未检测到农行回单指纹标识。"

        return True, "验证通过"

    def analyze_pdf(self, local_company_name=""):
        """
        核心PDF解析逻辑：高精度定位回单区域并提取关键信息
        
        分析PDF文件的每一页，识别回单分隔线或回单编号标签来定位每个回单的位置，
        然后提取每个回单的关键信息：客户名称、回单编号（20位数字）、金额等。
        
        流程：
        1. 验证PDF是否为农行回单格式
        2. 逐页分析，识别回单区域（通过分隔线或标签位置）
        3. 对每个回单区域提取：付款方/收款方户名、回单编号、金额
        4. 根据本方公司户名判断客户名称（如果付款方是本公司，则用收款方作为客户）
        5. 将提取的数据添加到预览列表
        
        :param local_company_name: 本方公司户名，用于判断客户名称（在主线程中获取，避免线程安全问题）
        """
        # 使用线程安全的方式清空树视图
        self.safe_gui_update(self._clear_tree)

        # --- 新增：指纹校验逻辑 ---
        is_valid, msg = self.is_valid_abc_receipt(self.doc)
        if not is_valid:
            self.safe_gui_update(self._show_analysis_error, msg)
            return

        try:
            total_receipts = 0
            for page_idx, page in enumerate(self.doc):
                width, height = page.rect.width, page.rect.height
                paths = page.get_drawings()
                separator_tops = [p['rect'].y0 for p in paths if p['dashes'] and p['rect'].width > width * 0.8 and p['rect'].height < 2]
                boundaries = sorted(list(set([0] + separator_tops + [height])))
                receipt_rects = [fitz.Rect(0, boundaries[i] + 2, width, boundaries[i+1] - 2) 
                                for i in range(len(boundaries) - 1) 
                                if boundaries[i+1] - boundaries[i] > 150]
                
                # 如果没有识别到分隔线，尝试基于"回单编号"标签位置来分割
                if not receipt_rects or len(receipt_rects) == 1:
                    all_words = page.get_text("words")
                    receipt_no_labels = []
                    for w in all_words:
                        if "回单编号" in w[4]:
                            w_rect = fitz.Rect(w[:4])
                            receipt_no_labels.append(w_rect.y0)
                    
                    if len(receipt_no_labels) > 1:
                        # 基于"回单编号"标签位置重新分割
                        receipt_no_labels = sorted(set(receipt_no_labels))
                        # 为每个回单编号标签创建区域（从标签上方50像素到下一个标签上方50像素）
                        new_boundaries = [0]
                        for label_y in receipt_no_labels:
                            new_boundaries.append(label_y - 50)  # 标签上方50像素
                        new_boundaries.append(height)
                        new_boundaries = sorted(set(new_boundaries))
                        
                        # 创建新的回单区域
                        receipt_rects = []
                        for i in range(len(new_boundaries) - 1):
                            if new_boundaries[i+1] - new_boundaries[i] > 150:
                                receipt_rects.append(fitz.Rect(0, new_boundaries[i], width, new_boundaries[i+1]))
                
                if not receipt_rects and height > 150:
                    receipt_rects.append(page.rect)
                
                # 确保回单区域按y坐标排序
                receipt_rects.sort(key=lambda r: r.y0)

                for crop_rect in receipt_rects:
                    words = page.get_text("words", clip=crop_rect)
                    if not words: continue

                    def find_text_from_anchor(anchor_texts, search_width=300, x_offset=0, y_offset_v=3):
                        """
                        从锚点文本位置查找并提取后续的文本内容
                        
                        在PDF页面中查找指定的锚点文本（如"金额（小写）"），
                        然后在其右侧的搜索区域内提取文本内容。
                        
                        :param anchor_texts: 锚点文本列表，按优先级顺序查找
                        :param search_width: 搜索区域的宽度（像素），默认300
                        :param x_offset: X轴偏移量，默认0
                        :param y_offset_v: Y轴垂直方向的容差，默认3像素
                        :return: 找到的文本内容字符串，如果未找到则返回None
                        """
                        for anchor_text in anchor_texts:
                            anchor_words = [w for w in words if anchor_text in w[4]]
                            if not anchor_words: continue

                            anchor_rect = fitz.Rect(anchor_words[0][:4])
                            search_rect = fitz.Rect(
                                anchor_rect.x1 + x_offset,
                                anchor_rect.y0 - y_offset_v,
                                anchor_rect.x1 + search_width,
                                anchor_rect.y1 + y_offset_v
                            )

                            found_words = [w for w in words if fitz.Rect(w[:4]).intersects(search_rect)]
                            if not found_words: continue

                            # 修复：按x坐标排序，确保正确的阅读顺序
                            found_words.sort(key=itemgetter(0))
                            return " ".join(w[4] for w in found_words)
                        return None

                    def extract_name_only(anchor_texts, search_width=250, stop_keywords=None):
                        """
                        精确提取户名，遇到停止关键词时停止
                        
                        从PDF页面中提取付款方或收款方的户名信息。
                        通过查找锚点文本（如"付款方户名"），然后在同一行搜索户名，
                        遇到停止关键词（如"账号"、"金额"）时停止提取。
                        
                        :param anchor_texts: 锚点文本列表，如["付款方户名", "付款方"]
                        :param search_width: 搜索区域的宽度（像素），默认250
                        :param stop_keywords: 停止关键词列表，遇到这些词时停止提取，默认包含"账号"、"金额"等
                        :return: 提取到的户名字符串，如果未找到则返回None
                        """
                        if stop_keywords is None:
                            stop_keywords = ["账号", "账户", "开户行", "金额", "日期", "摘要", "用途", "备注", "回单编号"]
                        
                        for anchor_text in anchor_texts:
                            anchor_words = [w for w in words if anchor_text in w[4]]
                            if not anchor_words:
                                continue

                            anchor_rect = fitz.Rect(anchor_words[0][:4])
                            anchor_y = anchor_rect.y0
                            
                            # 查找同一行的冒号位置
                            search_start_x = anchor_rect.x1
                            colon_found = False
                            for w in words:
                                w_rect = fitz.Rect(w[:4])
                                w_text = w[4]
                                # 查找同一行的冒号
                                if abs(w_rect.y0 - anchor_y) < 5 and ("：" in w_text or ":" in w_text):
                                    if w_rect.x0 >= anchor_rect.x0:
                                        search_start_x = w_rect.x1
                                        colon_found = True
                                        break

                            # 如果没有找到冒号，从锚点文本结束位置开始
                            if not colon_found:
                                search_start_x = anchor_rect.x1

                            # 在同一行搜索户名，遇到停止关键词时停止
                            found_words = []
                            for w in words:
                                w_rect = fitz.Rect(w[:4])
                                w_text = w[4].strip()
                                
                                # 检查是否在同一行（y坐标相近，允许小误差）
                                if abs(w_rect.y0 - anchor_y) < 5:
                                    # 检查是否在搜索范围内（在冒号之后）
                                    if w_rect.x0 >= search_start_x and w_rect.x0 < search_start_x + search_width:
                                        # 遇到停止关键词时停止（检查完整词，避免误判）
                                        should_stop = False
                                        for kw in stop_keywords:
                                            # 检查是否是独立的词（前后是空格、标点或边界）
                                            # 对于中文，使用更宽松的匹配
                                            if kw in w_text:
                                                # 检查是否是完整词（前后是标点、空格或边界）
                                                kw_pos = w_text.find(kw)
                                                if kw_pos >= 0:
                                                    before = w_text[kw_pos-1] if kw_pos > 0 else ' '
                                                    after = w_text[kw_pos+len(kw)] if kw_pos+len(kw) < len(w_text) else ' '
                                                    # 如果前后是标点、空格或中文字符边界，认为是完整词
                                                    if before in [' ', '，', ',', '。', '.', '：', ':', '、', '（', '(', '）', ')'] or \
                                                       after in [' ', '，', ',', '。', '.', '：', ':', '、', '（', '(', '）', ')']:
                                                        should_stop = True
                                                        break
                                        if should_stop:
                                            break
                                        
                                        # 跳过冒号、空白和标点符号
                                        if w_text and w_text not in ["：", ":", " ", "，", ",", "。", "."]:
                                            found_words.append(w)
                            
                            if found_words:
                                # 按x坐标排序
                                found_words.sort(key=itemgetter(0))
                                # 提取文本并清理
                                name_text = " ".join(w[4] for w in found_words)
                                # 移除开头的冒号、空格等
                                name_text = re.sub(r'^[：:\s，,。.]+', '', name_text)
                                # 移除"户名 "或"户名"前缀
                                name_text = re.sub(r'^户名\s*', '', name_text)
                                # 再次检查停止关键词，确保截断
                                for kw in stop_keywords:
                                    if kw in name_text:
                                        kw_pos = name_text.find(kw)
                                        if kw_pos >= 0:
                                            # 检查是否是完整词
                                            before = name_text[kw_pos-1] if kw_pos > 0 else ' '
                                            after = name_text[kw_pos+len(kw)] if kw_pos+len(kw) < len(name_text) else ' '
                                            if before in [' ', '，', ',', '。', '.', '：', ':', '、', '（', '(', '）', ')'] or \
                                               after in [' ', '，', ',', '。', '.', '：', ':', '、', '（', '(', '）', ')']:
                                                name_text = name_text[:kw_pos].strip()
                                                break
                                # 清理末尾的标点
                                name_text = re.sub(r'[，,。.\s]+$', '', name_text)
                                if name_text:
                                    return name_text.strip()
                        
                        return None

                    # --- 数据提取与清洗 ---
                    payer_name_text = extract_name_only(["付款方户名", "付款方", "户名"], search_width=200) or ""
                    # 清理换行符和多余空格
                    payer_name = payer_name_text.replace('\n', ' ').replace('\r', ' ').replace('\t', ' ')
                    payer_name = re.sub(r'\s+', ' ', payer_name).strip() or "未知付款方"

                    receiver_name_text = extract_name_only(["收款方户名", "收款方", "户名"], search_width=200) or ""
                    # 清理换行符和多余空格
                    receiver_name = receiver_name_text.replace('\n', ' ').replace('\r', ' ').replace('\t', ' ')
                    receiver_name = re.sub(r'\s+', ' ', receiver_name).strip() or "未知收款方"

                    def extract_receipt_no_with_pdfplumber(page_idx, crop_rect):
                        """
                        使用pdfplumber提取回单编号，严格匹配20位数字
                        
                        优先使用pdfplumber库从表格中提取回单编号，如果表格提取失败，
                        则使用文本提取方式，通过正则表达式匹配20位数字。
                        
                        :param page_idx: PDF页面索引（从0开始）
                        :param crop_rect: 裁剪区域的矩形坐标（fitz.Rect对象）
                        :return: 20位数字的回单编号字符串，如果未找到则返回None
                        """
                        try:
                            with pdfplumber.open(self.source_file) as pdf:
                                if page_idx >= len(pdf.pages):
                                    return None
                                
                                page = pdf.pages[page_idx]

                                # 直接使用原始坐标，pdfplumber 也是默认左上角坐标系
                                bbox = (crop_rect.x0, crop_rect.y0, crop_rect.x1, crop_rect.y1)
                                
                                cropped_page = page.crop(bbox)
                                
                                # 方法1：提取表格
                                tables = cropped_page.extract_tables()
                                if tables:
                                    for table in tables:
                                        for row in table:
                                            row_text = " ".join([str(cell) if cell else "" for cell in row])
                                            if "回单编号" in row_text:
                                                for cell in row:
                                                    if cell:
                                                        cell_text = str(cell).strip()
                                                        match = RECEIPT_NO_REGEX_20.search(cell_text)
                                                        if match:
                                                            return match.group(1)
                                
                                # 方法2：如果表格提取失败，使用文本提取
                                text = cropped_page.extract_text()
                                if text:
                                    match = RECEIPT_NO_LABEL_REGEX_20.search(text)
                                    if match:
                                        return match.group(1)
                        except Exception:
                            pass
                        
                        return None
                    
                    def extract_receipt_no_with_pymupdf(anchor_texts, search_width=250, stop_keywords=None):
                        """
                        使用PyMuPDF提取回单编号，严格匹配20位数字
                        
                        从PDF页面中查找"回单编号"标签，然后在其右侧搜索区域内
                        提取数字，组合成20位数字的回单编号。
                        如果找到的数字长度不是20位，则返回None。
                        
                        :param anchor_texts: 锚点文本列表，通常为["回单编号"]
                        :param search_width: 搜索区域的宽度（像素），默认250
                        :param stop_keywords: 停止关键词列表，遇到这些词时停止搜索，默认包含"付款方"、"收款方"等
                        :return: 20位数字的回单编号字符串，如果未找到或长度不正确则返回None
                        """
                        if stop_keywords is None:
                            stop_keywords = ["付款方", "收款方", "账号", "账户", "开户行", "金额", "日期"]
                        
                        for anchor_text in anchor_texts:
                            anchor_words = [w for w in words if anchor_text in w[4]]
                            if not anchor_words:
                                continue

                            anchor_words.sort(key=lambda w: (w[1], w[0]))
                            anchor_word = anchor_words[0]
                            
                            anchor_rect = fitz.Rect(anchor_word[:4])
                            anchor_y = anchor_rect.y0
                            
                            y_tolerance = 3
                            
                            search_start_x = anchor_rect.x1
                            for w in words:
                                w_rect = fitz.Rect(w[:4])
                                if abs(w_rect.y0 - anchor_y) < y_tolerance and (":" in w[4] or "：" in w[4]):
                                    if w_rect.x0 >= anchor_rect.x0:
                                        search_start_x = w_rect.x1
                                        break

                            found_words = []
                            for w in words:
                                w_rect = fitz.Rect(w[:4])
                                w_text = w[4].strip()
                                
                                if abs(w_rect.y0 - anchor_y) > y_tolerance:
                                    continue
                                
                                if w_rect.x0 >= search_start_x and w_rect.x0 < search_start_x + search_width:
                                    if any(kw in w_text for kw in stop_keywords):
                                        break
                                    
                                    if w_text and re.match(r'^\d+$', w_text):
                                        found_words.append(w)
                            
                            if found_words:
                                found_words.sort(key=itemgetter(0))
                                no_text = "".join(w[4] for w in found_words)
                                no_text_clean = re.sub(r'[^\d]', '', no_text)
                                
                                if len(no_text_clean) == 20:
                                    return no_text_clean
                        
                        return None

                    # --- 提取流程 ---
                    # 1. 优先使用pdfplumber
                    r_no_text = extract_receipt_no_with_pdfplumber(page_idx, crop_rect)
                    
                    # 2. 如果失败，使用PyMuPDF
                    if not r_no_text:
                        r_no_text = extract_receipt_no_with_pymupdf(["回单编号"], search_width=250)
                    
                    # 3. 最后手段：在区域文本中直接搜索
                    if not r_no_text:
                        crop_text = page.get_text(clip=crop_rect)
                        if crop_text:
                            match = RECEIPT_NO_LABEL_REGEX_20.search(crop_text)
                            if match:
                                r_no_text = match.group(1)
                    
                    # 清理回单编号中的换行符和空格
                    if r_no_text:
                        r_no = r_no_text.replace('\n', '').replace('\r', '').replace('\t', '').replace(' ', '').strip()
                    else:
                        r_no = "未知编号"

                    r_amt_text = find_text_from_anchor(["金额（小写）"], search_width=150) or ""
                    # 清理换行符和多余空格
                    r_amt_text = r_amt_text.replace('\n', ' ').replace('\r', ' ').replace('\t', ' ')
                    r_amt_match = re.search(r'([0-9,]+\.\d{2})', r_amt_text)
                    r_amt = r_amt_match.group(1).replace(",", "") if r_amt_match else "0.00"

                    if r_amt == "0.00":
                        full_text = page.get_text(clip=crop_rect)
                        # 清理换行符
                        full_text = full_text.replace('\n', ' ').replace('\r', ' ').replace('\t', ' ')
                        amt_match = re.search(r'([0-9,]+\.\d{2})', full_text)
                        if amt_match: r_amt = amt_match.group(1).replace(",", "")

                    r_name = payer_name
                    if local_company_name and local_company_name in payer_name:
                        r_name = receiver_name
                    
                    total_receipts += 1
                    # 存储付款方和收款方户名信息
                    item_data = {
                        "page_idx": page_idx, 
                        "rect": list(crop_rect), 
                        "name": self.clean_filename(r_name), 
                        "no": r_no, 
                        "amt": r_amt, 
                        "seq": total_receipts,
                        "payer_name": payer_name,  # 存储原始付款方户名
                        "receiver_name": receiver_name  # 存储原始收款方户名
                    }
                    
                    status = "正常" if "未知" not in r_name and "未知" not in r_no else "需核对"
                    # 使用线程安全的方式插入数据和更新preview_data
                    self.safe_gui_update(self._insert_tree_item_with_data, item_data, total_receipts, r_no, r_amt, status)

            # 使用线程安全的方式更新状态
            self.safe_gui_update(self._update_analysis_complete, total_receipts)

        except Exception as e:
            error_msg = str(e)
            self.safe_gui_update(self._show_analysis_error, error_msg)

    def _clear_tree(self):
        """
        清空树视图和预览数据（在主线程中执行）
        
        清空所有已解析的回单数据，重置界面状态。
        用于在加载新文件前清理旧数据。
        """
        self.preview_data = []
        self.payer_names = []
        self.receiver_names_map = {}
        self.combo_local_company.set("")
        self.combo_local_company['values'] = []
        # 隐藏确认按钮
        self.btn_confirm_company.grid_remove()
        for item in self.tree.get_children():
            self.tree.delete(item)

    def _insert_tree_item_with_data(self, item_data, seq, receipt_no, amount, status):
        """
        插入树视图项并更新preview_data（在主线程中执行）
        
        将解析到的回单数据添加到预览列表和树视图中显示。
        会对数据进行清理，确保换行符等特殊字符被正确处理。
        
        :param item_data: 回单数据字典，包含page_idx、rect、name、no、amt等信息
        :param seq: 回单序号
        :param receipt_no: 回单编号
        :param amount: 金额字符串
        :param status: 状态字符串（如"正常"、"需核对"等）
        """
        # 确保seq一致（使用传入的seq参数，确保数据一致性）
        item_data['seq'] = seq
        # 确保所有字段都清理了换行符（双重保险）
        if 'name' in item_data:
            item_data['name'] = item_data['name'].replace('\n', ' ').replace('\r', ' ').replace('\t', ' ')
            item_data['name'] = re.sub(r'\s+', ' ', item_data['name']).strip()
        if 'no' in item_data:
            item_data['no'] = item_data['no'].replace('\n', '').replace('\r', '').replace('\t', '').replace(' ', '').strip()
        # 清理receipt_no和amount参数（从外部传入的）
        receipt_no = receipt_no.replace('\n', '').replace('\r', '').replace('\t', '').replace(' ', '').strip() if receipt_no else ""
        amount = amount.replace('\n', '').replace('\r', '').replace('\t', '').replace(' ', '').strip() if amount else ""
        
        # 使用item_data中的值，确保数据一致性
        final_name = item_data.get('name', '')
        final_no = item_data.get('no', receipt_no) if item_data.get('no') else receipt_no
        final_amt = item_data.get('amt', amount) if item_data.get('amt') else amount
        
        self.preview_data.append(item_data)
        # 将item_id存储到item_data中，方便后续查找
        item_id = self.tree.insert("", "end", values=(seq, final_name, final_no, final_amt, status))
        item_data['item_id'] = item_id

    def _update_analysis_complete(self, total_receipts):
        """
        更新分析完成状态（在主线程中执行）
        
        在PDF分析完成后调用，更新界面状态：
        1. 提取所有唯一的付款方户名，填充到下拉列表
        2. 更新状态栏显示分析结果
        3. 启用"开始拆分导出"按钮
        
        :param total_receipts: 总共识别到的回单数量
        """
        # 提取所有唯一的付款方户名
        payer_names_set = set()
        for item in self.preview_data:
            if 'payer_name' in item and item['payer_name'] and item['payer_name'] != "未知付款方":
                payer_names_set.add(item['payer_name'])
        
        # 更新下拉列表，添加默认选项
        self.payer_names = sorted(list(payer_names_set))
        default_text = "使用付款方户名作为客户名称（默认值）"
        combo_values = [default_text] + self.payer_names
        self.combo_local_company['values'] = combo_values
        # 设置默认选中第一项（默认值）
        self.combo_local_company.set(default_text)
        # 确保确认按钮隐藏
        self.btn_confirm_company.grid_remove()
        
        if self.payer_names:
            self.log(f"解析完成，共发现 {total_receipts} 条回单。检测到 {len(self.payer_names)} 个不同的付款方户名。可选择本方公司户名进行更新，或使用默认值。")
        else:
            self.log(f"解析完成，共发现 {total_receipts} 条回单。请核对后点击开始拆分。")
        
        if total_receipts > 0:
            self.btn_process.config(state="normal")

    def _show_analysis_error(self, error_msg):
        """
        显示分析错误（在主线程中执行）
        
        当PDF分析过程中发生错误时调用，在状态栏和消息框中显示错误信息。
        
        :param error_msg: 错误消息字符串
        """
        self.log(f"解析出错: {error_msg}")
        messagebox.showerror("错误", error_msg)

    def start_processing(self):
        """
        开始拆分和导出处理流程
        
        弹出目录选择对话框让用户选择保存位置，然后在后台线程中
        执行PDF拆分和保存操作。处理过程中会显示进度条。
        """
        output_dir = filedialog.askdirectory(title="选择保存位置")
        if not output_dir: return
        self.btn_process.config(state="disabled")
        self.progress_bar['value'] = 0
        self.progress_bar['maximum'] = len(self.preview_data)
        threading.Thread(target=self.process_and_save, args=(output_dir,), daemon=True).start()

    def process_and_save(self, output_dir):
        """
        处理所有回单并保存为独立的PDF文件
        
        在后台线程中执行，遍历所有识别到的回单，将每个回单裁剪并保存为独立的PDF文件。
        文件名格式：客户名称_回单编号_金额.pdf
        同时生成CSV格式的处理日志文件，记录每个文件的处理状态。
        
        :param output_dir: 输出目录路径，拆分后的PDF文件和日志文件将保存在此目录
        """
        # 检查文档是否有效
        if not self.doc or self.source_file == "":
            self.safe_gui_update(self._show_export_error, "文档未加载或已被关闭，请重新选择PDF文件")
            return
        
        log_filename = f"log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        log_filepath = os.path.join(output_dir, log_filename)
        log_header = ["原文件名", "拆分后文件名", "生成时间", "状态"]

        try:
            with open(log_filepath, 'w', newline='', encoding='utf-8-sig') as log_file:
                writer = csv.writer(log_file)
                writer.writerow(log_header)

                success_count = 0
                total_files = len(self.preview_data)
                source_basename = os.path.basename(self.source_file)

                for item in self.preview_data:
                    # 确保文件名安全（使用clean_filename处理）
                    safe_name = self.clean_filename(item.get('name', '未知'))
                    safe_no = item.get('no', '未知编号').replace('\\', '_').replace('/', '_')
                    safe_amt = item.get('amt', '0.00').replace('\\', '_').replace('/', '_')
                    filename = f"{safe_name}_{safe_no}_{safe_amt}.pdf"
                    save_path = os.path.join(output_dir, filename)
                    
                    try:
                        # 检查文档是否仍然有效
                        if not self.doc:
                            raise Exception("文档已被关闭")
                        
                        counter = 1
                        while os.path.exists(save_path):
                            filename = f"{safe_name}_{safe_no}_{safe_amt}_{counter}.pdf"
                            save_path = os.path.join(output_dir, filename)
                            counter += 1
                        
                        # 验证页面索引有效性
                        if item['page_idx'] >= len(self.doc):
                            raise Exception(f"页面索引 {item['page_idx']} 超出文档范围")
                        
                        new_doc = fitz.open()
                        new_doc.insert_pdf(self.doc, from_page=item['page_idx'], to_page=item['page_idx'])
                        new_page = new_doc[0]
                        new_page.set_cropbox(fitz.Rect(item['rect']))
                        new_doc.save(save_path)
                        new_doc.close()
                        
                        writer.writerow([source_basename, filename, datetime.now().strftime('%Y-%m-%d %H:%M:%S'), "成功"])
                        success_count += 1

                    except Exception as item_error:
                        error_msg = f"失败: {str(item_error)}"
                        writer.writerow([source_basename, filename, datetime.now().strftime('%Y-%m-%d %H:%M:%S'), error_msg])

                    finally:
                        # 使用线程安全的方式更新进度
                        current_progress = self.progress_bar['value'] + 1
                        self.safe_gui_update(self._update_progress, current_progress, total_files)

                # 使用线程安全的方式显示完成消息
                self.safe_gui_update(self._show_completion_message, success_count, log_filename, output_dir)

        except Exception as e:
            error_msg = str(e)
            self.safe_gui_update(self._show_export_error, error_msg)
        
        finally:
            # 使用线程安全的方式重置按钮和进度条
            self.safe_gui_update(self._reset_processing_ui)

    def _update_progress(self, current, total):
        """
        更新进度条（在主线程中执行）
        
        更新导出进度条的显示，同时更新状态栏消息。
        
        :param current: 当前已处理的文件数量
        :param total: 总共需要处理的文件数量
        """
        self.progress_bar['value'] = current
        self.log(f"正在导出... ({current}/{total})")

    def _show_completion_message(self, success_count, log_filename, output_dir):
        """
        显示完成消息（在主线程中执行）
        
        当所有回单处理完成后调用，显示成功消息并在Windows资源管理器中打开输出目录。
        
        :param success_count: 成功导出的文件数量
        :param log_filename: 生成的日志文件名
        :param output_dir: 输出目录路径
        """
        self.log(f"处理完成！成功导出 {success_count} 个文件。日志已保存至 {log_filename}")
        messagebox.showinfo("成功", f"已成功拆分并保存 {success_count} 个回单文件！\n日志文件已生成：{log_filename}")
        # 添加异常处理
        try:
            os.startfile(output_dir)
        except Exception as e:
            self.log(f"无法打开文件夹: {str(e)}")

    def _show_export_error(self, error_msg):
        """
        显示导出错误（在主线程中执行）
        
        当导出过程中发生错误时调用，在状态栏和消息框中显示错误信息。
        
        :param error_msg: 错误消息字符串
        """
        self.log(f"导出出错: {error_msg}")
        messagebox.showerror("导出错误", error_msg)

    def _reset_processing_ui(self):
        """
        重置处理UI（在主线程中执行）
        
        重置处理相关的UI元素，恢复按钮状态和进度条。
        在导出完成后或出错后调用。
        """
        self.btn_process.config(state="normal")
        self.progress_bar['value'] = 0


if __name__ == "__main__":
    root = tk.Tk()
    style = ttk.Style()
    style.theme_use('clam')
    app = ReceiptSplitterApp(root)
    root.mainloop()