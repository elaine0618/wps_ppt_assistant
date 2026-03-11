import win32com.client
import tkinter as tk
from tkinter import ttk
import pythoncom
import threading

class WPSNotesViewer:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("WPSNOTE")
        self.root.geometry("420x650+50+100")
        self.root.attributes('-topmost', True)
        
        # 设置窗口背景
        self.root.configure(bg='#f0f0f0')
        
        # 连接WPS
        self.connect_wps()
        
        self.setup_ui()
        self.refresh_loop()
        self.root.mainloop()
    
    def connect_wps(self):
        """连接WPS演示"""
        try:
            self.wps = win32com.client.GetActiveObject("Kwpp.Application")
            print("已连接到运行中的WPS")
        except:
            try:
                self.wps = win32com.client.Dispatch("Kwpp.Application")
                self.wps.Visible = True
                print("已启动新的WPS实例")
            except Exception as e:
                print(f"连接WPS失败: {e}")
                self.wps = None
    
    def setup_ui(self):
        """设置界面"""
        # 标题栏
        title_frame = tk.Frame(self.root, bg='#2c3e50', height=45)
        title_frame.pack(fill='x')
        title_frame.pack_propagate(False)
        
        tk.Label(
            title_frame, 
            text="📋", 
            bg='#2c3e50', 
            fg='white',
            font=('微软雅黑', 14, 'bold')
        ).pack(side='left', padx=15, pady=10)
        
        # 右侧按钮区域
        button_frame = tk.Frame(title_frame, bg='#2c3e50')
        button_frame.pack(side='right', padx=10)
        
        # 保存按钮
        self.save_btn = tk.Button(
            button_frame,
            text="💾 保存",
            bg='#3498db',
            fg='white',
            font=('微软雅黑', 9),
            relief='flat',
            cursor='hand2',
            command=self.save_edit,
            state='disabled'
        )
        self.save_btn.pack(side='left', padx=2)
        self.save_btn.bind('<Enter>', lambda e: self.save_btn.config(bg='#2980b9') if self.save_btn['state'] == 'normal' else None)
        self.save_btn.bind('<Leave>', lambda e: self.save_btn.config(bg='#3498db') if self.save_btn['state'] == 'normal' else None)
        
        # 取消按钮
        self.cancel_btn = tk.Button(
            button_frame,
            text="✕ 取消",
            bg='#3498db',
            fg='white',
            font=('微软雅黑', 9),
            relief='flat',
            cursor='hand2',
            command=self.cancel_edit,
            state='disabled'
        )
        self.cancel_btn.pack(side='left', padx=2)
        self.cancel_btn.bind('<Enter>', lambda e: self.cancel_btn.config(bg='#2980b9') if self.cancel_btn['state'] == 'normal' else None)
        self.cancel_btn.bind('<Leave>', lambda e: self.cancel_btn.config(bg='#3498db') if self.cancel_btn['state'] == 'normal' else None)
        
        # 生成按钮
        self.generate_btn = tk.Button(
            button_frame,
            text="✨ 生成",
            bg='#3498db',
            fg='white',
            font=('微软雅黑', 9),
            relief='flat',
            cursor='hand2',
            command=self.generate_slides
        )
        self.generate_btn.pack(side='left', padx=2)
        self.generate_btn.bind('<Enter>', lambda e: self.generate_btn.config(bg='#2980b9'))
        self.generate_btn.bind('<Leave>', lambda e: self.generate_btn.config(bg='#3498db'))
        
        # 刷新按钮
        refresh_btn = tk.Button(
            button_frame,
            text="⟳ 刷新",
            bg='#3498db',
            fg='white',
            font=('微软雅黑', 9),
            relief='flat',
            cursor='hand2',
            command=self.refresh_now
        )
        refresh_btn.pack(side='left', padx=2)
        refresh_btn.bind('<Enter>', lambda e: refresh_btn.config(bg='#2980b9'))
        refresh_btn.bind('<Leave>', lambda e: refresh_btn.config(bg='#3498db'))
        
        # ===== 搜索框区域 =====
        search_frame = tk.Frame(self.root, bg='white', height=40)
        search_frame.pack(fill='x', padx=8, pady=(5,0))
        search_frame.pack_propagate(False)
        
        # 搜索图标
        tk.Label(
            search_frame,
            text="🔍",
            bg='white',
            fg='#7f8c8d',
            font=('微软雅黑', 12)
        ).pack(side='left', padx=(5,0))
        
        # 搜索输入框
        self.search_var = tk.StringVar()
        self.search_var.trace('w', lambda *args: self.filter_notes())  # 输入时实时过滤
        
        self.search_entry = tk.Entry(
            search_frame,
            textvariable=self.search_var,
            font=('微软雅黑', 10),
            bg='#f0f0f0',
            fg='#2c3e50',
            relief='flat',
            highlightthickness=1,
            highlightcolor='#bdc3c7',
            highlightbackground='#bdc3c7'
        )
        self.search_entry.pack(side='left', fill='x', expand=True, padx=5, pady=5)
        self.search_entry.bind('<KeyRelease>', lambda e: self.filter_notes())  # 按键释放时过滤
        
        # 清空搜索按钮
        self.clear_btn = tk.Button(
            search_frame,
            text="✕",
            bg='#f0f0f0',
            fg='#7f8c8d',
            font=('微软雅黑', 10),
            relief='flat',
            cursor='hand2',
            command=self.clear_search
        )
        self.clear_btn.pack(side='right', padx=(0,5))
        
        # 搜索提示标签（显示匹配数量）
        self.search_result_label = tk.Label(
            search_frame,
            text="",
            bg='white',
            fg='#7f8c8d',
            font=('微软雅黑', 9)
        )
        self.search_result_label.pack(side='right', padx=10)
        
        # 分隔线
        separator = tk.Frame(self.root, bg='#bdc3c7', height=1)
        separator.pack(fill='x', padx=5, pady=5)
        # ===== 搜索框区域结束 =====
        
        # 主容器
        main_frame = tk.Frame(self.root, bg='white')
        main_frame.pack(fill='both', expand=True, padx=2, pady=2)
        
        # 创建画布和滚动条
        self.canvas = tk.Canvas(main_frame, bg='white', highlightthickness=0)
        scrollbar = tk.Scrollbar(main_frame, orient='vertical', command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg='white')
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw", width=400)
        self.canvas.configure(yscrollcommand=scrollbar.set)
        
        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 绑定鼠标滚轮
        def on_mousewheel(event):
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        
        self.canvas.bind_all("<MouseWheel>", on_mousewheel)
        
        # 状态栏
        self.status = tk.Label(
            self.root, 
            text="就绪", 
            bd=1, 
            relief='sunken', 
            anchor='w',
            bg='#ecf0f1',
            font=('微软雅黑', 9)
        )
        self.status.pack(side='bottom', fill='x')
        
        # 当前编辑中的条目
        self.editing_item = None
        self.is_editing = False  # 编辑状态标志
        
        # 多选相关
        self.selected_items = set()  # 存储选中的幻灯片索引
        self.last_selected = None  # 最后一次选中的索引，用于shift多选
        
        # 绑定Ctrl和Shift键状态
        self.ctrl_pressed = False
        self.shift_pressed = False
        
        def on_key_down(event):
            if event.keysym == 'Control_L' or event.keysym == 'Control_R':
                self.ctrl_pressed = True
            elif event.keysym == 'Shift_L' or event.keysym == 'Shift_R':
                self.shift_pressed = True
        
        def on_key_up(event):
            if event.keysym == 'Control_L' or event.keysym == 'Control_R':
                self.ctrl_pressed = False
            elif event.keysym == 'Shift_L' or event.keysym == 'Shift_R':
                self.shift_pressed = False
        
        self.root.bind('<KeyPress>', on_key_down)
        self.root.bind('<KeyRelease>', on_key_up)
    
    def get_notes(self):
        """获取所有幻灯片的备注（WPS专用方法）"""
        notes = []
        
        if not self.wps:
            return notes
        
        try:
            # 检查是否有打开的演示文稿
            if self.wps.Presentations.Count == 0:
                return notes
            
            pres = self.wps.ActivePresentation
            if not pres:
                return notes
            
            print(f"正在读取 {pres.Slides.Count} 张幻灯片的备注...")
            
            # 遍历所有幻灯片
            for i in range(1, pres.Slides.Count + 1):
                slide = pres.Slides(i)
                
                # ===== WPS备注读取 =====
                notes_text = ""
                
                # 方法1：通过备注页获取
                try:
                    if slide.NotesPage:
                        # 遍历备注页中的所有形状
                        for j in range(1, slide.NotesPage.Shapes.Count + 1):
                            shape = slide.NotesPage.Shapes(j)
                            # 检查是否有文本框
                            if shape.HasTextFrame == -1:  # -1 表示 True
                                if shape.TextFrame.HasText == -1:
                                    text_range = shape.TextFrame.TextRange
                                    if text_range:
                                        notes_text = text_range.Text
                                        if notes_text and notes_text.strip():
                                            print(f"幻灯片 {i} 找到备注: {notes_text[:30]}...")
                                            break
                except Exception as e:
                    print(f"方法1失败: {e}")
                
                # 方法2：直接访问备注占位符
                if not notes_text:
                    try:
                        # WPS中备注通常在第二个占位符
                        if slide.NotesPage.Shapes.Count >= 2:
                            shape = slide.NotesPage.Shapes(2)
                            if shape.HasTextFrame == -1:
                                notes_text = shape.TextFrame.TextRange.Text
                    except:
                        pass
                
                # 方法3：遍历所有形状找文本框
                if not notes_text:
                    try:
                        for j in range(1, slide.NotesPage.Shapes.Count + 1):
                            shape = slide.NotesPage.Shapes(j)
                            if shape.HasTextFrame == -1:
                                try:
                                    text = shape.TextFrame.TextRange.Text
                                    if text and len(text) > 0:
                                        notes_text = text
                                        break
                                except:
                                    pass
                    except:
                        pass
                
                # 如果没有备注，显示提示
                if not notes_text or notes_text.strip() == "":
                    notes_text = "📭 无备注"
                
                notes.append({
                    'index': i,
                    'notes': notes_text
                })
                
        except Exception as e:
            print(f"获取备注时出错: {e}")
            import traceback
            traceback.print_exc()
        
        return notes
    
    def create_note_item(self, note):
        """创建备注项（双击编辑单击选中）"""
        # 背景色（交替颜色）
        bg_color = '#f8f9ff' if note['index'] % 2 == 0 else '#ffffff'
        
        # 主框架
        item_frame = tk.Frame(
            self.scrollable_frame, 
            bg=bg_color,
            relief='solid',
            bd=1,
            highlightbackground='#d0d0d0',
            highlightcolor='#d0d0d0',
            highlightthickness=1
        )
        # 保存备注数据到框架，用于搜索过滤
        item_frame.note_data = note
        item_frame.slide_index = note['index']
        
        item_frame.pack(fill='x', padx=8, pady=4)
        
        # 左侧：序号区域
        left_frame = tk.Frame(item_frame, bg=bg_color)
        left_frame.grid(row=0, column=0, padx=8, pady=8, sticky='n')
        
        # 序号标签
        index_frame = tk.Frame(left_frame, bg='#3498db', width=40, height=40)
        index_frame.pack_propagate(False)
        index_frame.pack()
        
        index_label = tk.Label(
            index_frame,
            text=f"{note['index']}",
            bg='#3498db',
            fg='white',
            font=('微软雅黑', 12, 'bold'),
            width=2,
            height=1
        )
        index_label.pack(expand=True)
        
        # 右侧：备注内容区域
        right_frame = tk.Frame(item_frame, bg=bg_color)
        right_frame.grid(row=0, column=1, padx=(5,10), pady=8, sticky='nsew')
        
        # 备注显示标签
        notes_text = note['notes']
        if notes_text == "📭 无备注":
            note_color = '#95a5a6'
            display_text = notes_text
        else:
            note_color = '#34495e'
            display_text = notes_text
            if len(display_text) > 120:
                display_text = display_text[:117] + "..."
        
        # 创建可点击的标签
        notes_label = tk.Label(
            right_frame,
            text=f"💬 {display_text}",
            bg=bg_color,
            fg=note_color,
            font=('微软雅黑', 11),
            anchor='w',
            justify='left',
            wraplength=250,
            cursor='hand2'
        )
        notes_label.pack(fill='x', expand=True)
        
        # 保存组件引用，方便更新
        item_frame.notes_label = notes_label
        item_frame.full_notes = note['notes']
        item_frame.right_frame = right_frame
        item_frame.bg_color = bg_color
        
        # 绑定点击事件
        def on_click(event, idx=note['index']):
            # 如果正在编辑，不处理点击事件
            if self.is_editing:
                return
            
            # 处理多选
            if self.ctrl_pressed:
                # Ctrl点击：切换选中状态
                if idx in self.selected_items:
                    self.selected_items.remove(idx)
                else:
                    self.selected_items.add(idx)
                self.last_selected = idx
            elif self.shift_pressed and self.last_selected:
                # Shift点击：选择从last_selected到当前的所有项
                start = min(self.last_selected, idx)
                end = max(self.last_selected, idx)
                self.selected_items.clear()
                for i in range(start, end + 1):
                    self.selected_items.add(i)
            else:
                # 普通点击：单选
                self.selected_items.clear()
                self.selected_items.add(idx)
                self.last_selected = idx
            
            # 更新所有条目的背景色
            self.update_selection_display()
            
            # 跳转到选中的幻灯片
            if not self.ctrl_pressed and not self.shift_pressed:
                self.goto_slide(idx)
        
        # 绑定双击事件到标签（编辑）
        notes_label.bind('<Double-Button-1>', lambda e, idx=note['index']: self.start_edit(idx))
        
        # 绑定单击事件到整个条目框架和标签（多选）
        notes_label.bind('<Button-1>', on_click)
        item_frame.bind('<Button-1>', on_click)
        right_frame.bind('<Button-1>', on_click)
        
        # 配置网格列权重
        item_frame.columnconfigure(1, weight=1)
    
    def update_selection_display(self):
        """更新所有条目的选中状态显示"""
        for widget in self.scrollable_frame.winfo_children():
            if hasattr(widget, 'slide_index'):
                idx = widget.slide_index
                if idx in self.selected_items:
                    # 选中状态
                    widget.config(bg='#e6f3ff')
                    widget.notes_label.config(bg='#e6f3ff')
                    widget.right_frame.config(bg='#e6f3ff')
                    widget.left_frame = widget.grid_slaves(row=0, column=0)[0]
                    widget.left_frame.config(bg='#e6f3ff')
                else:
                    # 未选中状态
                    bg_color = '#f8f9ff' if idx % 2 == 0 else '#ffffff'
                    widget.config(bg=bg_color)
                    widget.notes_label.config(bg=bg_color)
                    widget.right_frame.config(bg=bg_color)
                    widget.left_frame = widget.grid_slaves(row=0, column=0)[0]
                    widget.left_frame.config(bg=bg_color)
    
    def generate_slides(self):
        """生成选中的幻灯片到末尾"""
        if not self.selected_items:
            self.status.config(text="请先选择要生成的幻灯片")
            return
        
        try:
            if not self.wps or self.wps.Presentations.Count == 0:
                self.status.config(text="没有打开的PPT")
                return
            
            pres = self.wps.ActivePresentation
            total_slides = pres.Slides.Count
            
            # 获取选中的幻灯片索引并排序
            selected = sorted(list(self.selected_items))
            
            # 复制选中的幻灯片
            for i, idx in enumerate(selected):
                if idx <= total_slides:  # 确保索引有效
                    slide = pres.Slides(idx)
                    # 复制幻灯片到末尾
                    slide.Copy()
                    pres.Slides.Paste()
                    print(f"已复制幻灯片 {idx}")
            
            self.status.config(text=f"已生成 {len(selected)} 张幻灯片")
            
            # 清空选中状态
            self.selected_items.clear()
            self.update_selection_display()
            
            # 刷新显示
            self.refresh_now()
            
        except Exception as e:
            print(f"生成幻灯片时出错: {e}")
            self.status.config(text=f"生成失败: {str(e)[:30]}")
    
    def start_edit(self, slide_index):
        """开始编辑备注"""
        # 如果已经有正在编辑的条目，先保存
        if self.editing_item:
            self.save_edit()
        
        # 找到对应的条目框架
        for widget in self.scrollable_frame.winfo_children():
            if hasattr(widget, 'slide_index') and widget.slide_index == slide_index:
                # 移除原来的标签
                widget.notes_label.pack_forget()
                
                # 创建文本输入框
                text_var = tk.StringVar()
                text_var.set(widget.full_notes if widget.full_notes != "📭 无备注" else "")
                
                entry = tk.Text(
                    widget.right_frame,
                    font=('微软雅黑', 11),
                    height=3,
                    wrap='word',
                    relief='solid',
                    bd=1,
                    padx=5,
                    pady=5,
                    bg='#fff9e6'
                )
                entry.pack(fill='x', expand=True)
                
                # 插入当前文本
                entry.insert('1.0', text_var.get())
                entry.focus_set()
                entry.see('1.0')
                
                # 保存编辑状态
                self.editing_item = {
                    'frame': widget,
                    'entry': entry,
                    'slide_index': slide_index,
                    'original_text': widget.full_notes
                }
                self.is_editing = True  # 设置为编辑状态
                
                # 显示保存和取消按钮
                self.save_btn.config(state='normal')
                self.cancel_btn.config(state='normal')
                
                break
    
    def save_edit(self):
        """保存编辑的内容"""
        if not self.editing_item:
            return
        
        try:
            item = self.editing_item
            slide_index = item['slide_index']
            entry = item['entry']
            frame = item['frame']
            
            # 获取新内容
            new_notes = entry.get('1.0', 'end-1c').strip()
            if not new_notes:
                new_notes = "📭 无备注"
            
            # 更新到WPS
            self.update_notes(slide_index, new_notes)
            
            # 移除输入框
            entry.destroy()
            
            # 重新创建显示标签
            self.update_item_display(frame, new_notes)
            
        except Exception as e:
            print(f"保存编辑时出错: {e}")
        finally:
            self.editing_item = None
            self.is_editing = False  # 退出编辑状态
            # 隐藏保存和取消按钮
            self.save_btn.config(state='disabled')
            self.cancel_btn.config(state='disabled')
    
    def cancel_edit(self):
        """取消编辑"""
        if not self.editing_item:
            return
        
        try:
            item = self.editing_item
            frame = item['frame']
            entry = item['entry']
            original_text = item['original_text']
            
            # 移除输入框
            entry.destroy()
            
            # 重新创建显示标签
            self.update_item_display(frame, original_text)
            
        except Exception as e:
            print(f"取消编辑时出错: {e}")
        finally:
            self.editing_item = None
            self.is_editing = False  # 退出编辑状态
            # 隐藏保存和取消按钮
            self.save_btn.config(state='disabled')
            self.cancel_btn.config(state='disabled')
    
    def update_item_display(self, frame, notes_text):
        """更新条目显示"""
        # 创建新的标签
        if notes_text == "📭 无备注":
            note_color = '#95a5a6'
            display_text = notes_text
        else:
            note_color = '#34495e'
            display_text = notes_text
            if len(display_text) > 120:
                display_text = display_text[:117] + "..."
        
        notes_label = tk.Label(
            frame.right_frame,
            text=f"💬 {display_text}",
            bg=frame.bg_color,
            fg=note_color,
            font=('微软雅黑', 11),
            anchor='w',
            justify='left',
            wraplength=250,
            cursor='hand2'
        )
        notes_label.pack(fill='x', expand=True)
        
        # 绑定双击事件
        notes_label.bind('<Double-Button-1>', lambda e, idx=frame.slide_index: self.start_edit(idx))
        
        # 绑定单击事件（多选）
        def on_click(event, idx=frame.slide_index):
            if self.is_editing:
                return
            
            if self.ctrl_pressed:
                if idx in self.selected_items:
                    self.selected_items.remove(idx)
                else:
                    self.selected_items.add(idx)
                self.last_selected = idx
            elif self.shift_pressed and self.last_selected:
                start = min(self.last_selected, idx)
                end = max(self.last_selected, idx)
                self.selected_items.clear()
                for i in range(start, end + 1):
                    self.selected_items.add(i)
            else:
                self.selected_items.clear()
                self.selected_items.add(idx)
                self.last_selected = idx
            
            self.update_selection_display()
            
            if not self.ctrl_pressed and not self.shift_pressed:
                self.goto_slide(idx)
        
        notes_label.bind('<Button-1>', on_click)
        frame.right_frame.bind('<Button-1>', on_click)
        frame.bind('<Button-1>', on_click)
        
        # 更新框架中的引用
        frame.notes_label = notes_label
        frame.full_notes = notes_text
        if hasattr(frame, 'note_data'):
            frame.note_data['notes'] = notes_text
    
    def goto_slide(self, slide_index):
        """跳转到指定幻灯片"""
        try:
            if self.wps and self.wps.ActiveWindow and self.wps.ActiveWindow.View:
                self.wps.ActiveWindow.View.GotoSlide(slide_index)
        except Exception as e:
            print(f"跳转失败: {e}")
    
    def update_notes(self, slide_index, new_notes):
        """更新WPS中的备注内容"""
        try:
            if not self.wps or self.wps.Presentations.Count == 0:
                return
            
            pres = self.wps.ActivePresentation
            slide = pres.Slides(slide_index)
            
            # 尝试更新备注
            updated = False
            
            # 方法1：通过备注页更新
            try:
                if slide.NotesPage:
                    for j in range(1, slide.NotesPage.Shapes.Count + 1):
                        shape = slide.NotesPage.Shapes(j)
                        if shape.HasTextFrame == -1:
                            if shape.TextFrame.HasText == -1 or new_notes != "📭 无备注":
                                shape.TextFrame.TextRange.Text = new_notes if new_notes != "📭 无备注" else ""
                                updated = True
                                print(f"幻灯片 {slide_index} 备注已更新")
                                break
            except Exception as e:
                print(f"方法1更新失败: {e}")
            
            # 方法2：获取第二个占位符
            if not updated:
                try:
                    if slide.NotesPage.Shapes.Count >= 2:
                        shape = slide.NotesPage.Shapes(2)
                        if shape.HasTextFrame == -1:
                            shape.TextFrame.TextRange.Text = new_notes if new_notes != "📭 无备注" else ""
                            updated = True
                    elif slide.NotesPage.Shapes.Count > 0:
                        shape = slide.NotesPage.Shapes(1)
                        if shape.HasTextFrame == -1:
                            shape.TextFrame.TextRange.Text = new_notes if new_notes != "📭 无备注" else ""
                            updated = True
                except Exception as e:
                    print(f"方法2更新失败: {e}")
            
            if updated:
                # 如果有搜索框内容，重新应用过滤
                if hasattr(self, 'search_var') and self.search_var.get().strip():
                    self.filter_notes()
                
                print(f"幻灯片 {slide_index} 备注更新成功")
            else:
                print(f"幻灯片 {slide_index} 备注更新失败")
                
        except Exception as e:
            print(f"更新备注时出错: {e}")
            import traceback
            traceback.print_exc()
    
    def filter_notes(self):
        """根据搜索关键词过滤备注"""
        keyword = self.search_var.get().strip().lower()
        
        # 遍历所有备注项
        visible_count = 0
        for widget in self.scrollable_frame.winfo_children():
            # 如果是备注项框架（有备注数据）
            if hasattr(widget, 'note_data'):
                note = widget.note_data
                notes_text = note['notes'].lower()
                
                # 如果没有关键词，显示所有；否则检查是否包含关键词
                if keyword == "" or keyword in notes_text:
                    widget.pack(fill='x', padx=8, pady=4)  # 显示
                    visible_count += 1
                else:
                    widget.pack_forget()  # 隐藏
        
        # 更新搜索结果提示
        if keyword:
            self.search_result_label.config(text=f"找到 {visible_count} 项")
        else:
            self.search_result_label.config(text="")
    
    def clear_search(self):
        """清空搜索框"""
        self.search_var.set("")  # 清空搜索词
        self.search_entry.focus_set()  # 焦点回到搜索框
        # filter_notes 会通过 trace 自动调用
    
    def refresh_now(self):
        """立即刷新显示（保留选中状态）"""
        # 如果有正在编辑的，先取消
        if self.editing_item:
            self.cancel_edit()
        
        # 保存当前选中的索引
        selected_before_refresh = self.selected_items.copy() if hasattr(self, 'selected_items') else set()
        last_selected_before_refresh = self.last_selected
        
        # 清空现有内容
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        
        # 获取备注
        notes = self.get_notes()
        
        if not notes:
            # 显示提示信息
            msg_frame = tk.Frame(self.scrollable_frame, bg='white')
            msg_frame.pack(fill='both', expand=True, pady=50)
            
            tk.Label(
                msg_frame,
                text="📂 没有打开的演示文稿",
                fg='#7f8c8d',
                font=('微软雅黑', 12)
            ).pack()
            
            tk.Label(
                msg_frame,
                text="请先在WPS中打开一个PPT文件",
                fg='#95a5a6',
                font=('微软雅黑', 10)
            ).pack(pady=5)
            
            self.status.config(text="未连接演示文稿")
            return
        
        # 显示备注列表
        for note in notes:
            self.create_note_item(note)
        
        # 恢复选中状态（只保留仍然存在的幻灯片索引）
        self.selected_items = set()
        for idx in selected_before_refresh:
            if idx <= len(notes):  # 确保索引仍然有效
                self.selected_items.add(idx)
        
        self.last_selected = last_selected_before_refresh if last_selected_before_refresh and last_selected_before_refresh <= len(notes) else None
        
        # 更新选中状态的显示
        self.update_selection_display()
        
        self.status.config(text=f"共 {len(notes)} 张幻灯片")
        
        # 如果搜索框有内容，重新应用过滤
        if hasattr(self, 'search_var') and self.search_var.get().strip():
            self.filter_notes()
        
    def refresh_loop(self):
        """自动刷新循环（编辑时暂停）"""
        try:
            # 只有在非编辑状态时才自动刷新
            if not self.is_editing:
                self.refresh_now()
        except Exception as e:
            print(f"自动刷新出错: {e}")
        
        # 每2秒刷新一次
        self.root.after(2000, self.refresh_loop)

def main():
    # 初始化COM
    pythoncom.CoInitialize()
    try:
        app = WPSNotesViewer()
    finally:
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    main()