import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
import pyttsx3
import re
from docx import Document

class VocabularyReader:
    def __init__(self, root):
        self.root = root
        self.root.title("Word 单词阅读器与发音工具")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        
        # 初始化文本转语音引擎
        # 主要使用系统发音，pyttsx3作为备用
        self.engine = None
        try:
            # 测试系统发音是否可用
            import subprocess
            result = subprocess.run(['say', '--version'], 
                                  capture_output=True, 
                                  text=True, 
                                  timeout=2)
            if result.returncode == 0:
                print("系统发音功能可用")
            else:
                print("系统发音不可用，将尝试pyttsx3")
        except:
            print("检测系统发音功能时出错，将使用pyttsx3")
        
        # 存储单词和释义
        self.vocabulary = []
        self.all_words = set()  # 存储所有识别到的英文单词
        self.current_file = ""
        self.speech_rate = 150  # 默认语速
        
        # 创建界面组件
        self.create_widgets()
        
    def create_widgets(self):
        # 创建菜单栏
        menubar = tk.Menu(self.root)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="打开 Word 文件", command=self.open_word_file)
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self.root.quit)
        menubar.add_cascade(label="文件", menu=file_menu)
        
        # 添加语音设置菜单
        voice_menu = tk.Menu(menubar, tearoff=0)
        voice_menu.add_command(label="切换语音", command=self.change_voice)
        voice_menu.add_command(label="调整语速", command=self.adjust_rate)
        menubar.add_cascade(label="语音设置", menu=voice_menu)
        
        self.root.config(menu=menubar)
        
        # 创建工具栏
        toolbar = tk.Frame(self.root, bg="#f0f0f0", relief=tk.RAISED, bd=2)
        
        open_btn = tk.Button(toolbar, text="📂 打开 Word 文件", command=self.open_word_file, 
                           font=("SimHei", 10), bg="#e1f5fe", fg="#01579b")
        open_btn.pack(side=tk.LEFT, padx=2, pady=2)
        
        pronounce_btn = tk.Button(toolbar, text="🔊 发音选中单词", command=self.pronounce_selected,
                                font=("SimHei", 10), bg="#e8f5e8", fg="#2e7d32")
        pronounce_btn.pack(side=tk.LEFT, padx=2, pady=2)
        
        # 添加语速控制区域
        speed_frame = tk.Frame(toolbar, bg="#f0f0f0")
        speed_frame.pack(side=tk.LEFT, padx=10, pady=2)
        
        speed_label = tk.Label(speed_frame, text="语速:", font=("SimHei", 9), bg="#f0f0f0")
        speed_label.pack(side=tk.LEFT)
        
        # 语速显示标签
        self.speed_display = tk.Label(speed_frame, text=f"{self.speech_rate}", 
                                    font=("SimHei", 9, "bold"), bg="#fff3e0", 
                                    fg="#e65100", relief=tk.SUNKEN, width=4)
        self.speed_display.pack(side=tk.LEFT, padx=2)
        
        # 语速调整按钮
        speed_down_btn = tk.Button(speed_frame, text="🔽", command=self.decrease_speed,
                                 font=("SimHei", 8), bg="#ffebee", fg="#c62828", width=2)
        speed_down_btn.pack(side=tk.LEFT, padx=1)
        
        speed_up_btn = tk.Button(speed_frame, text="🔼", command=self.increase_speed,
                               font=("SimHei", 8), bg="#e8f5e8", fg="#2e7d32", width=2)
        speed_up_btn.pack(side=tk.LEFT, padx=1)
        
        # 语速重置按钮
        speed_reset_btn = tk.Button(speed_frame, text="↻", command=self.reset_speed,
                                  font=("SimHei", 8), bg="#f3e5f5", fg="#7b1fa2", width=2)
        speed_reset_btn.pack(side=tk.LEFT, padx=1)
        
        # 添加单词统计信息
        self.word_count_label = tk.Label(toolbar, text="单词数: 0", 
                                       font=("SimHei", 10), bg="#f0f0f0")
        self.word_count_label.pack(side=tk.RIGHT, padx=10, pady=2)
        
        toolbar.pack(side=tk.TOP, fill=tk.X)
        
        # 创建主内容区域
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 创建文本显示区域
        self.text_area = scrolledtext.ScrolledText(
            main_frame, wrap=tk.WORD, font=("SimHei", 12),
            bg="white", fg="black", relief=tk.SUNKEN, bd=2
        )
        self.text_area.pack(fill=tk.BOTH, expand=True)
        
        # 绑定鼠标点击事件
        self.text_area.bind("<ButtonRelease-1>", self.on_word_click)
        
        # 状态栏
        self.status_bar = tk.Label(
            self.root, text="就绪", bd=1, relief=tk.SUNKEN, anchor=tk.W
        )
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
    def open_word_file(self):
        """打开 Word 文件并解析单词"""
        file_path = filedialog.askopenfilename(
            defaultextension=".docx",
            filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")]
        )
        
        if file_path:
            try:
                doc = Document(file_path)
                content = []
                for para in doc.paragraphs:
                    content.append(para.text)
                full_content = '\n'.join(content)
                
                self.current_file = file_path
                self.parse_vocabulary(full_content)
                self.display_vocabulary()
                self.status_bar.config(text=f"已打开文件: {file_path.split('/')[-1]}")
            except Exception as e:
                messagebox.showerror("错误", f"打开文件失败: {str(e)}")
    
    def parse_vocabulary(self, content):
        """解析文本内容提取单词和释义"""
        self.vocabulary = []
        self.all_words = set()  # 存储所有识别到的英文单词
        lines = content.split("\n")
        
        # 正则表达式匹配单词格式: 单词 [音标] 词性. 释义
        pattern = r'^(\w+)\s*\[([^\]]+)\]\s*([a-zA-Z.]+)\.\s*(.*)$'
        
        # 英文单词匹配模式 (只匹配纯英文字母，长度2+)
        word_pattern = r'\b[A-Za-z]{2,}\b'
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
                
            match = re.match(pattern, line)
            if match:
                word = match.group(1)
                phonetic = match.group(2)
                pos = match.group(3)
                meaning = match.group(4)
                self.vocabulary.append({
                    "word": word,
                    "phonetic": phonetic,
                    "pos": pos,
                    "meaning": meaning,
                    "full_line": line
                })
                # 添加主单词到单词集合
                self.all_words.add(word.lower())
            else:
                # 如果不符合标准格式，也添加到列表中
                self.vocabulary.append({"full_line": line})
            
            # 从每一行中提取所有英文单词
            words_in_line = re.findall(word_pattern, line)
            for word in words_in_line:
                # 过滤掉一些常见的非单词内容
                if len(word) >= 2 and not word.isdigit():
                    self.all_words.add(word.lower())
    
    def display_vocabulary(self):
        """在文本区域显示单词内容"""
        self.text_area.delete(1.0, tk.END)
        
        # 添加标题
        self.text_area.insert(tk.END, "==== 文档内容 ====\n\n", "title")
        
        for item in self.vocabulary:
            self.text_area.insert(tk.END, item["full_line"] + "\n")
        
        # 添加分隔符和提示
        self.text_area.insert(tk.END, "\n" + "="*50 + "\n")
        self.text_area.insert(tk.END, "💡 提示：蓝色带下划线的单词可以点击发音\n", "tip")
        self.text_area.insert(tk.END, "🔍 鼠标悬停单词会有高亮效果\n", "tip")
        self.text_area.insert(tk.END, "="*50 + "\n\n")
        
        # 更新单词统计
        unique_words = len(self.all_words) if hasattr(self, 'all_words') else 0
        self.word_count_label.config(text=f"识别单词数: {unique_words}")
        
        # 为所有单词添加标签，方便后续识别
        self.tag_words()
    
    def tag_words(self):
        """为文本中的单词添加标签，方便点击识别"""
        # 清除现有标签
        for tag in self.text_area.tag_names():
            if tag.startswith("word_"):
                self.text_area.tag_remove(tag, 1.0, tk.END)
        
        # 配置标题和提示样式
        self.text_area.tag_config("title", font=("SimHei", 14, "bold"), foreground="darkgreen")
        self.text_area.tag_config("tip", font=("SimHei", 10), foreground="gray", justify="center")
        
        # 获取所有文本内容
        content = self.text_area.get("1.0", tk.END)
        
        # 英文单词匹配模式
        word_pattern = r'\b[A-Za-z]{2,}\b'
        
        # 查找所有匹配的单词
        matches = list(re.finditer(word_pattern, content))
        
        print(f"找到 {len(matches)} 个英文单词")
        
        for i, match in enumerate(matches):
            word = match.group()
            start_char = match.start()
            end_char = match.end()
            
            # 精确计算tkinter位置
            start_pos = self._char_to_position(content, start_char)
            end_pos = self._char_to_position(content, end_char)
            
            if start_pos and end_pos:
                # 验证位置是否正确
                actual_text = self.text_area.get(start_pos, end_pos)
                if actual_text == word:
                    # 创建唯一标签
                    tag_name = f"word_{i}_{word}"
                    self.text_area.tag_add(tag_name, start_pos, end_pos)
                    
                    # 设置标签样式
                    self.text_area.tag_config(tag_name, 
                                            foreground="blue", 
                                            underline=1,
                                            font=("SimHei", 12, "bold"))
                    
                    # 绑定点击事件 - 使用函数生成器避免闭包问题
                    self._bind_word_events(tag_name, word)
                    
                    print(f"成功标记单词: '{word}' 位置 {start_pos}-{end_pos}")
                else:
                    print(f"位置错误: 期待'{word}', 实际'{actual_text}'")
    
    def _char_to_position(self, content, char_index):
        """将字符索引转换为tkinter位置"""
        if char_index > len(content):
            return None
        
        # 统计到指定位置之前的行数和列数
        text_before = content[:char_index]
        lines = text_before.split('\n')
        line_num = len(lines)
        col_num = len(lines[-1]) if lines else 0
        
        return f"{line_num}.{col_num}"
    
    def _bind_word_events(self, tag_name, word):
        """绑定单词事件，避免闭包问题"""
        def on_click(event):
            self.on_word_tag_click(event, word)
            return "break"
        
        def on_enter(event):
            self.text_area.tag_config(tag_name, background="lightblue")
        
        def on_leave(event):
            self.text_area.tag_config(tag_name, background="")
        
        self.text_area.tag_bind(tag_name, "<Button-1>", on_click)
        self.text_area.tag_bind(tag_name, "<Enter>", on_enter)
        self.text_area.tag_bind(tag_name, "<Leave>", on_leave)
    
    def on_word_tag_click(self, event, word):
        """处理单词标签点击事件"""
        self.pronounce_word(word)
        self.status_bar.config(text=f"正在发音: {word}")
        # 阻止事件继续传播
        return "break"
    
    def on_word_click(self, event):
        """处理单词点击事件（备用方法）"""
        # 这个方法现在主要作为备用，主要的点击处理由tag绑定完成
        pass
    
    def pronounce_word(self, word):
        """发音指定单词"""
        # 清理单词，只保留字母
        clean_word = re.sub(r'[^A-Za-z]', '', word)
        if not clean_word:
            print(f"无效单词: {word}")
            return
        
        print(f"准备发音: {clean_word}")
        
        # 优先使用系统发音，更稳定
        if self.use_system_voice(clean_word):
            return
        
        # 如果系统发音失败，尝试pyttsx3
        self.use_pyttsx3_voice(clean_word)
    
    def use_system_voice(self, word):
        """使用macOS系统自带的say命令发音"""
        try:
            import subprocess
            print(f"使用系统发音: {word} (语速: {self.speech_rate})")
            
            # 使用系统say命令，设置当前语速
            result = subprocess.run(['say', '-r', str(self.speech_rate), word], 
                                  capture_output=True, 
                                  text=True, 
                                  timeout=10)
            
            if result.returncode == 0:
                print(f"系统发音完成: {word}")
                return True
            else:
                print(f"系统发音失败: {result.stderr}")
                return False
                
        except subprocess.TimeoutExpired:
            print(f"系统发音超时: {word}")
            return False
        except FileNotFoundError:
            print("系统不支持say命令")
            return False
        except Exception as e:
            print(f"系统发音错误: {e}")
            return False
    
    def use_pyttsx3_voice(self, word):
        """使用pyttsx3发音（备用方法）"""
        try:
            print(f"尝试pyttsx3发音: {word} (语速: {self.speech_rate})")
            
            # 每次重新创建引擎实例，避免状态问题
            engine = pyttsx3.init('nsss')
            
            # 设置语音参数
            engine.setProperty('rate', self.speech_rate)
            engine.setProperty('volume', 0.9)
            
            # 尝试设置英语语音
            voices = engine.getProperty('voices')
            for voice in voices:
                if 'en' in voice.id.lower() or 'english' in voice.name.lower():
                    engine.setProperty('voice', voice.id)
                    break
            
            # 发音
            engine.say(word)
            engine.runAndWait()
            
            # 清理引擎
            del engine
            
            print(f"pyttsx3发音完成: {word}")
            
        except Exception as e:
            print(f"pyttsx3发音错误: {e}")
            # 最后的备用方案：播放提示音
            try:
                print(f"\a发音失败，单词: {word}")  # \a 是响铃字符
            except:
                pass
    
    def pronounce_selected(self):
        """发音选中的文本"""
        try:
            selected_text = self.text_area.get(tk.SEL_FIRST, tk.SEL_LAST)
            if selected_text:
                # 使用相同的发音方法
                self.pronounce_word(selected_text)
                self.status_bar.config(text=f"正在发音: {selected_text}")
        except tk.TclError:
            messagebox.showinfo("提示", "请先选中要发音的文本")
    
    def increase_speed(self):
        """增加语速"""
        if self.speech_rate < 300:
            self.speech_rate += 25
            self.speed_display.config(text=f"{self.speech_rate}")
            print(f"语速增加到: {self.speech_rate}")
    
    def decrease_speed(self):
        """减少语速"""
        if self.speech_rate > 50:
            self.speech_rate -= 25
            self.speed_display.config(text=f"{self.speech_rate}")
            print(f"语速减少到: {self.speech_rate}")
    
    def reset_speed(self):
        """重置语速到默认值"""
        self.speech_rate = 150
        self.speed_display.config(text=f"{self.speech_rate}")
        print(f"语速重置到: {self.speech_rate}")
        messagebox.showinfo("语速设置", f"语速已重置为默认值: {self.speech_rate}")
    
    def change_voice(self):
        """切换语音（男声/女声）"""
        # 由于使用系统发音，这个功能暂时不可用
        messagebox.showinfo("语音设置", "当前使用系统默认语音")
    
    def adjust_rate(self):
        """调整语速"""
        # 简化版本，显示当前使用系统发音
        messagebox.showinfo("语速设置", "当前使用系统默认语速")

if __name__ == "__main__":
    root = tk.Tk()
    app = VocabularyReader(root)
    root.mainloop()