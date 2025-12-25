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
    def __init__(self, root):
        self.root = root
        self.root.title("农行电子回单智能拆分工具 V1.0.0")
        self.root.geometry("1200x700")

        self.source_file = ""
        self.doc = None
        self.preview_data = []
        self.preview_image = None
        self.preview_image_ref = None  # 保持图片引用，防止垃圾回收
        self.placeholder_text = "输入我方公司户名全称，用于自动判断客户名称，留空则默认使用付款方户名作为客户名称"
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

        self.lbl_local_company = ttk.Label(frame_top, text="本方公司名:")
        self.lbl_local_company.grid(row=1, column=0, padx=(0, 5), pady=(10, 0), sticky="w")
        self.entry_local_company = ttk.Entry(frame_top)
        self.entry_local_company.grid(row=1, column=1, columnspan=2, padx=5, pady=(10, 0), sticky="ew")
        self.entry_local_company.insert(0, self.placeholder_text)
        self.entry_local_company.config(foreground="gray")
        self.entry_local_company.bind("<FocusIn>", self.clear_placeholder)
        self.entry_local_company.bind("<FocusOut>", self.restore_placeholder)

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

        frame_right = ttk.LabelFrame(main_pane, text="回单原文预览", padding=10)
        main_pane.add(frame_right, weight=3)

        # 创建Canvas和滚动条容器
        preview_container = ttk.Frame(frame_right)
        preview_container.pack(fill="both", expand=True)
        
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
            # 如果点击的不是输入框本身，才转移焦点
            if event.widget != self.entry_local_company:
                self.root.focus_set()

        self.root.bind("<Button-1>", handle_root_click)

    def check_queue(self):
        """检查队列中的GUI更新请求（线程安全）"""
        try:
            while True:
                callback, args = self.update_queue.get_nowait()
                callback(*args)
        except queue.Empty:
            pass
        finally:
            self.root.after(100, self.check_queue)  # 每100ms检查一次

    def safe_gui_update(self, callback, *args):
        """线程安全的GUI更新方法"""
        self.update_queue.put((callback, args))

    def on_closing(self):
        """关闭窗口时的清理工作"""
        try:
            if self.doc:
                self.doc.close()
        except Exception:
            pass
        self.root.destroy()

    def show_receipt_preview(self, event):
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
            
        except Exception as e:
            # 显示错误信息
            self.preview_canvas.delete("all")
            self.preview_canvas.create_text(200, 100, text=f"无法生成预览:\n{str(e)}", anchor="center", fill="red")
            self.preview_canvas.configure(scrollregion=self.preview_canvas.bbox("all"))
            self.log(f"生成预览失败: {e}")

    def open_edit_window(self, event):
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

    def clear_placeholder(self, event):
        """当获得焦点时：如果是占位符，则清空"""
        current_text = self.entry_local_company.get()
        if current_text == self.placeholder_text:
            self.entry_local_company.delete(0, "end")
            self.entry_local_company.config(foreground="black")

    def restore_placeholder(self, event):
        """当失去焦点时：如果为空，则恢复占位符"""
        # 调试打印
        print(f"DEBUG: FocusOut Triggered. Current content: '{self.entry_local_company.get()}'")

        current_text = self.entry_local_company.get().strip()
        if not current_text:
            self.entry_local_company.delete(0, "end")
            self.entry_local_company.insert(0, self.placeholder_text)
            self.entry_local_company.config(foreground="gray")
        else:
            # 如果有实际内容，确保颜色是黑色的
            self.entry_local_company.config(foreground="black")

    def log(self, message):
        self.lbl_status.config(text=f"[{datetime.now().strftime('%H:%M:%S')}] {message}")
        self.root.update_idletasks()

    def load_file(self):
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
        self.log("正在分析文件，请稍候...")
        threading.Thread(target=self.analyze_pdf, daemon=True).start()

    def clean_filename(self, text):
        return re.sub(r'[\\/*?:"<>|]', "", text).strip()

    def is_valid_abc_receipt(self, doc, check_limit=3):
        """
        极速检测是否为农行回单
        :param doc: fitz.Document 对象
        :param check_limit: 最多检查前几页
        :return: (bool, message)
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

    def analyze_pdf(self):
        """核心解析逻辑：高精度定位 + 严格20位编号匹配"""
        # 使用线程安全的方式清空树视图
        self.safe_gui_update(self._clear_tree)

        # --- 新增：指纹校验逻辑 ---
        is_valid, msg = self.is_valid_abc_receipt(self.doc)
        if not is_valid:
            self.safe_gui_update(self._show_analysis_error, msg)
            return
        
        local_company_name = self.entry_local_company.get().strip()
        if local_company_name == self.placeholder_text:
            local_company_name = ""

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
                        """精确提取户名，遇到停止关键词时停止"""
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
                    payer_name = payer_name_text.strip() or "未知付款方"

                    receiver_name_text = extract_name_only(["收款方户名", "收款方", "户名"], search_width=200) or ""
                    receiver_name = receiver_name_text.strip() or "未知收款方"

                    def extract_receipt_no_with_pdfplumber(page_idx, crop_rect):
                        """使用pdfplumber提取回单编号，严格匹配20位数字"""
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
                        """使用PyMuPDF提取回单编号，严格匹配20位数字"""
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
                    
                    r_no = r_no_text if r_no_text else "未知编号"

                    r_amt_text = find_text_from_anchor(["金额（小写）"], search_width=150) or ""
                    r_amt_match = re.search(r'([0-9,]+\.\d{2})', r_amt_text)
                    r_amt = r_amt_match.group(1).replace(",", "") if r_amt_match else "0.00"

                    if r_amt == "0.00":
                        full_text = page.get_text(clip=crop_rect)
                        amt_match = re.search(r'([0-9,]+\.\d{2})', full_text)
                        if amt_match: r_amt = amt_match.group(1).replace(",", "")

                    r_name = payer_name
                    if local_company_name and local_company_name in payer_name:
                        r_name = receiver_name
                    
                    total_receipts += 1
                    item_data = {"page_idx": page_idx, "rect": list(crop_rect), "name": self.clean_filename(r_name), "no": r_no, "amt": r_amt, "seq": total_receipts}
                    
                    status = "正常" if "未知" not in r_name and "未知" not in r_no else "需核对"
                    # 使用线程安全的方式插入数据和更新preview_data
                    self.safe_gui_update(self._insert_tree_item_with_data, item_data, total_receipts, r_no, r_amt, status)

            # 使用线程安全的方式更新状态
            self.safe_gui_update(self._update_analysis_complete, total_receipts)

        except Exception as e:
            error_msg = str(e)
            self.safe_gui_update(self._show_analysis_error, error_msg)

    def _clear_tree(self):
        """清空树视图（在主线程中执行）"""
        self.preview_data = []
        for item in self.tree.get_children():
            self.tree.delete(item)

    def _insert_tree_item_with_data(self, item_data, seq, receipt_no, amount, status):
        """插入树视图项并更新preview_data（在主线程中执行）"""
        # 确保seq一致
        item_data['seq'] = seq
        self.preview_data.append(item_data)
        # 将item_id存储到item_data中，方便后续查找
        item_id = self.tree.insert("", "end", values=(seq, item_data['name'], receipt_no, amount, status))
        item_data['item_id'] = item_id

    def _update_analysis_complete(self, total_receipts):
        """更新分析完成状态（在主线程中执行）"""
        self.log(f"解析完成，共发现 {total_receipts} 条回单。请核对后点击开始拆分。")
        if total_receipts > 0:
            self.btn_process.config(state="normal")

    def _show_analysis_error(self, error_msg):
        """显示分析错误（在主线程中执行）"""
        self.log(f"解析出错: {error_msg}")
        messagebox.showerror("错误", error_msg)

    def start_processing(self):
        output_dir = filedialog.askdirectory(title="选择保存位置")
        if not output_dir: return
        self.btn_process.config(state="disabled")
        self.progress_bar['value'] = 0
        self.progress_bar['maximum'] = len(self.preview_data)
        threading.Thread(target=self.process_and_save, args=(output_dir,), daemon=True).start()

    def process_and_save(self, output_dir):
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
                    filename = f"{item['name']}_{item['no']}_{item['amt']}.pdf"
                    save_path = os.path.join(output_dir, filename)
                    
                    try:
                        counter = 1
                        while os.path.exists(save_path):
                            filename = f"{item['name']}_{item['no']}_{item['amt']}_{counter}.pdf"
                            save_path = os.path.join(output_dir, filename)
                            counter += 1
                        
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
        """更新进度条（在主线程中执行）"""
        self.progress_bar['value'] = current
        self.log(f"正在导出... ({current}/{total})")

    def _show_completion_message(self, success_count, log_filename, output_dir):
        """显示完成消息（在主线程中执行）"""
        self.log(f"处理完成！成功导出 {success_count} 个文件。日志已保存至 {log_filename}")
        messagebox.showinfo("成功", f"已成功拆分并保存 {success_count} 个回单文件！\n日志文件已生成：{log_filename}")
        # 添加异常处理
        try:
            os.startfile(output_dir)
        except Exception as e:
            self.log(f"无法打开文件夹: {str(e)}")

    def _show_export_error(self, error_msg):
        """显示导出错误（在主线程中执行）"""
        self.log(f"导出出错: {error_msg}")
        messagebox.showerror("导出错误", error_msg)

    def _reset_processing_ui(self):
        """重置处理UI（在主线程中执行）"""
        self.btn_process.config(state="normal")
        self.progress_bar['value'] = 0


if __name__ == "__main__":
    root = tk.Tk()
    style = ttk.Style()
    style.theme_use('clam')
    app = ReceiptSplitterApp(root)
    root.mainloop()