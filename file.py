import os
import shutil
import hashlib
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, ttk, messagebox

def get_file_hash(file_path):
    """计算文件的MD5哈希值"""
    hash_md5 = hashlib.md5()
    with open(file_path, "rb") as f:
        for chunk in iter(lambda: f.read(4096), b""):
            hash_md5.update(chunk)
    return hash_md5.hexdigest()

def get_folder_name_from_archive(archive_name):
    """从压缩包文件名获取可能的解压文件夹名"""
    return os.path.splitext(archive_name)[0]

def organize_files(directory, progress_var=None, status_text=None, deleted_files_text=None, root=None):
    # 定义文件类型和对应的文件夹
    file_types = {
        '图片': ['.jpg', '.jpeg', '.png', '.gif', '.bmp'],
        '文档': ['.doc', '.docx', '.pdf', '.txt', '.xlsx', '.ppt', '.pptx'],
        '音频': ['.mp3', '.wav', '.flac', '.m4a'],
        '视频': ['.mp4', '.avi', '.mkv', '.mov'],
        '压缩包': ['.zip', '.rar', '.7z'],
        '文件夹': []  # 用于存放其他文件夹
    }

    # 确保目录存在
    if not os.path.exists(directory):
        messagebox.showerror("错误", f"目录 {directory} 不存在!")
        return

    # 创建分类文件夹
    for folder in file_types.keys():
        folder_path = os.path.join(directory, folder)
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)

    # 用于检测重复文件
    file_hashes = {}
    
    # 存储压缩包文件名，用于后续检查解压文件夹
    archive_names = set()
    
    # 获取文件总数用于进度条
    total_files = len([f for f in os.listdir(directory) if os.path.isfile(os.path.join(directory, f))])
    processed_files = 0
    
    # 记录删除的文件
    deleted_files = []
    
    # 添加更新界面的计数器和阈值
    update_counter = 0
    UPDATE_THRESHOLD = 10  # 每处理10个文件才更新一次界面
    deleted_files_buffer = []  # 用于缓存删除文件的信息
    
    # 第一次遍历收集压缩包信息
    for item in os.listdir(directory):
        item_path = os.path.join(directory, item)
        if os.path.isfile(item_path):
            file_ext = os.path.splitext(item)[1].lower()
            if file_ext in file_types['压缩包']:
                archive_names.add(get_folder_name_from_archive(item))
    
    # 定义更新删除文件显示的函数
    def update_deleted_files():
        if deleted_files_text and deleted_files_buffer:
            if isinstance(deleted_files_text, tk.Text):
                deleted_files_text.delete(1.0, tk.END)
                deleted_files_text.insert(1.0, "\n".join(deleted_files))
    
    # 遍历目录中的所有文件和文件夹
    for item in os.listdir(directory):
        item_path = os.path.join(directory, item)
        
        # 每处理UPDATE_THRESHOLD个文件才更新一次界面
        update_counter += 1
        
        if update_counter % UPDATE_THRESHOLD == 0 and root:
            if status_text:
                status_text.set(f"正在处理: {item}")
            if progress_var:
                progress_var.set(int(processed_files / total_files * 100))
            update_deleted_files()
            root.update()
        
        # 处理文件夹
        if os.path.isdir(item_path):
            if item not in file_types.keys():  # 不移动已创建的分类文件夹
                # 检查是否为解压文件夹
                if item in archive_names:
                    if status_text:
                        status_text.set(f"发现解压文件夹 {item}，正在删除...")
                        root.update()
                    try:
                        shutil.rmtree(item_path)
                        deleted_files.append(f"删除解压文件夹: {item}")
                        deleted_files_buffer.append(f"删除解压文件夹: {item}")
                    except PermissionError:
                        # 遇到拒绝访问时直接跳过当前文件夹的处理
                        if status_text:
                            status_text.set(f"无法访问文件夹 {item}，已跳过")
                            root.update()
                    except Exception as e:
                        messagebox.showwarning("警告", f"删除文件夹 {item} 时出错: {str(e)}")
                    continue  # 移动到这里，只跳过当前的archive处理
                
                destination = os.path.join(directory, '文件夹', item)
                try:
                    shutil.move(item_path, destination)
                    if status_text:
                        status_text.set(f"已移动文件夹 {item} 到文件夹分类")
                        root.update()
                except PermissionError:
                    # 遇到拒绝访问时直接跳过当前文件夹的处理
                    if status_text:
                        status_text.set(f"无法访问文件夹 {item}，已跳过")
                        root.update()
                except Exception as e:
                    messagebox.showwarning("警告", f"移动文件夹 {item} 时出错: {str(e)}")
            continue  # 移动到这里，确保文件夹处理完后继续处理下一个项目

        # 处理文件
        file_ext = os.path.splitext(item)[1].lower()
        
        # 计算文件哈希值
        file_hash = get_file_hash(item_path)
        
        # 检查是否为重复文件
        if file_hash in file_hashes:
            if status_text:
                status_text.set(f"发现重复文件: {item}，正在删除...")
                root.update()
            os.remove(item_path)
            deleted_files.append(f"删除重复文件: {item}")
            deleted_files_buffer.append(f"删除重复文件: {item}")
            processed_files += 1
            continue
        else:
            file_hashes[file_hash] = item_path

        # 根据扩展名分类文件
        moved = False
        for folder, extensions in file_types.items():
            if file_ext in extensions:
                destination = os.path.join(directory, folder, item)
                try:
                    # 检查目标文件夹中是否存在同名文件
                    if os.path.exists(destination):
                        # 如果存在，计算目标文件的哈希值
                        dest_hash = get_file_hash(destination)
                        if dest_hash == file_hash:
                            # 如果是相同文件，删除源文件
                            if status_text:
                                status_text.set(f"目标文件夹已存在相同文件: {item}，删除源文件...")
                                root.update()
                            os.remove(item_path)
                            deleted_files.append(f"删除重复文件: {item}")
                            deleted_files_buffer.append(f"删除重复文件: {item}")
                        else:
                            # 如果是不同文件，重命名后移动
                            base, ext = os.path.splitext(item)
                            new_name = f"{base}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{ext}"
                            destination = os.path.join(directory, folder, new_name)
                            shutil.move(item_path, destination)
                            if status_text:
                                status_text.set(f"已移动并重命名 {item} 到 {folder} 文件夹")
                                root.update()
                    else:
                        # 如果不存在同名文件，直接移动
                        shutil.move(item_path, destination)
                        if status_text:
                            status_text.set(f"已移动 {item} 到 {folder} 文件夹")
                            root.update()
                    moved = True
                except Exception as e:
                    messagebox.showwarning("警告", f"移动 {item} 时出错: {str(e)}")
                break
        
        # 如果文件类型不在预定义列表中，创建其他文件夹
        if not moved:
            other_folder = os.path.join(directory, "其他文件")
            if not os.path.exists(other_folder):
                os.makedirs(other_folder)
            try:
                destination = os.path.join(other_folder, item)
                if os.path.exists(destination):
                    # 检查是否为重复文件
                    dest_hash = get_file_hash(destination)
                    if dest_hash == file_hash:
                        if status_text:
                            status_text.set(f"其他文件夹已存在相同文件: {item}，删除源文件...")
                            root.update()
                        os.remove(item_path)
                        deleted_files.append(f"删除重复文件: {item}")
                        deleted_files_buffer.append(f"删除重复文件: {item}")
                    else:
                        base, ext = os.path.splitext(item)
                        new_name = f"{base}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{ext}"
                        destination = os.path.join(other_folder, new_name)
                        shutil.move(item_path, destination)
                        if status_text:
                            status_text.set(f"已移动并重命名 {item} 到其他文件夹")
                            root.update()
                else:
                    shutil.move(item_path, destination)
                    if status_text:
                        status_text.set(f"已移动 {item} 到其他文件夹")
                        root.update()
            except Exception as e:
                messagebox.showwarning("警告", f"移动 {item} 时出错: {str(e)}")
        
        processed_files += 1
        if progress_var:
            progress_var.set(int(processed_files / total_files * 100))
            root.update()

class FileOrganizerGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("文件整理工具")
        self.root.geometry("900x700")  # 增加窗口大小
        
        # 设置主题样式
        style = ttk.Style()
        style.theme_use('clam')  # 使用 clam 主题
        
        # 自定义样式
        style.configure('TFrame', background='#f0f0f0')
        style.configure('TLabel', background='#f0f0f0', font=('微软雅黑', 10))
        style.configure('TButton', font=('微软雅黑', 10), padding=5)
        style.configure('Header.TLabel', font=('微软雅黑', 14, 'bold'))
        style.configure('Status.TLabel', font=('微软雅黑', 9))
        
        # 创建主框架并添加内边距
        self.main_frame = ttk.Frame(self.root, padding="20", style='TFrame')
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置主窗口的权重，使其可调整大小
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(1, weight=1)
        
        # 创建并配置组件
        self.path_var = tk.StringVar()
        self.status_var = tk.StringVar(value="等待开始...")
        self.progress_var = tk.IntVar()
        
        # 添加标题
        title_label = ttk.Label(self.main_frame, text="文件整理工具", style='Header.TLabel')
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # 路径选择框架
        path_frame = ttk.Frame(self.main_frame, style='TFrame')
        path_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        
        ttk.Label(path_frame, text="目标文件夹:").pack(side=tk.LEFT)
        path_entry = ttk.Entry(path_frame, textvariable=self.path_var, width=60)
        path_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        browse_btn = ttk.Button(path_frame, text="浏览", command=self.browse_directory)
        browse_btn.pack(side=tk.LEFT, padx=(5, 0))
        
        # 状态显示
        status_frame = ttk.Frame(self.main_frame, style='TFrame')
        status_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        ttk.Label(status_frame, text="状态:", style='Status.TLabel').pack(side=tk.LEFT)
        ttk.Label(status_frame, textvariable=self.status_var, style='Status.TLabel').pack(side=tk.LEFT, padx=5)
        
        # 进度条
        self.progress = ttk.Progressbar(self.main_frame, variable=self.progress_var, maximum=100, mode='determinate')
        self.progress.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 20))
        
        # 删除文件记录框架
        deleted_frame = ttk.LabelFrame(self.main_frame, text="删除的文件记录", padding=10)
        deleted_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 20))
        deleted_frame.columnconfigure(0, weight=1)
        deleted_frame.rowconfigure(0, weight=1)
        
        # 文本框和滚动条
        self.deleted_files_text = tk.Text(deleted_frame, height=15, width=70, font=('微软雅黑', 9))
        scrollbar = ttk.Scrollbar(deleted_frame, orient=tk.VERTICAL, command=self.deleted_files_text.yview)
        self.deleted_files_text.configure(yscrollcommand=scrollbar.set)
        
        self.deleted_files_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # 按钮框架
        button_frame = ttk.Frame(self.main_frame, style='TFrame')
        button_frame.grid(row=5, column=0, columnspan=3, pady=(0, 10))
        
        # 开始按钮
        start_btn = ttk.Button(button_frame, text="开始整理", command=self.start_organize, width=15)
        start_btn.pack(side=tk.LEFT, padx=5)
        
        # 取消按钮
        self.cancel_button = ttk.Button(button_frame, text="取消", command=self.cancel_organize, 
                                      state=tk.DISABLED, width=15)
        self.cancel_button.pack(side=tk.LEFT, padx=5)
        
        self.is_running = False
        
    def browse_directory(self):
        directory = filedialog.askdirectory()
        if directory:
            self.path_var.set(directory)
            
    def cancel_organize(self):
        self.is_running = False
        self.status_var.set("操作已取消")
        self.cancel_button.configure(state=tk.DISABLED)
        
    def start_organize(self):
        if not self.path_var.get():
            messagebox.showerror("错误", "请选择要整理的文件夹!")
            return
            
        # 使用线程处理文件操作
        self.is_running = True
        self.cancel_button.configure(state=tk.NORMAL)
        self.progress_var.set(0)
        self.deleted_files_text.delete(1.0, tk.END)
        
        import threading
        thread = threading.Thread(target=self._organize_files_thread)
        thread.daemon = True
        thread.start()
        
    def _organize_files_thread(self):
        try:
            organize_files(
                self.path_var.get(), 
                self.progress_var, 
                self.status_var, 
                self.deleted_files_text,
                self.root
            )
            if self.is_running:  # 只有在未取消的情况下才显示完成消息
                self.root.after(0, lambda: self.status_var.set("文件整理完成!"))
                self.root.after(0, lambda: messagebox.showinfo("完成", "文件整理已完成!"))
        finally:
            self.root.after(0, lambda: self.cancel_button.configure(state=tk.DISABLED))
            self.is_running = False
        
    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = FileOrganizerGUI()
    app.run()
