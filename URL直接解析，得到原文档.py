import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import requests
import time
import os
import threading
from pathlib import Path
import shutil
import tempfile
import uuid
import zipfile
import json


class PDFToRawConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF转原始文件工具 - 直接URL解析")
        self.root.geometry("750x700")
        self.root.minsize(700, 650)

        # MinerU API 信息
        self.token = "eyJ0eXBlIjoiSldUIiwiYWxnIjoiSFM1MTIifQ.eyJqdGkiOiI1MzAwODI3OSIsInJvbCI6IlJPTEVfUkVHSVNURVIiLCJpc3MiOiJPcGVuWExhYiIsImlhdCI6MTc2MzU0NDM5NSwiY2xpZW50SWQiOiJsa3pkeDU3bnZ5MjJqa3BxOXgydyIsInBob25lIjoiMTg0NjAzMDAxOTciLCJvcGVuSWQiOm51bGwsInV1aWQiOiI5NjY3ODRiNC0wMjRjLTQ3NzUtYjE5Ny1kZWY5NTIyZmJjZDciLCJlbWFpbCI6IiIsImV4cCI6MTc2NDc1Mzk5NX0.HPAoPC83v5Xi-ZxjTshshZljtR7zTyTyKAVSt4qSCfCCShaVKWE7_K1bC2lWNrZJWi8r-hpTbv8ym6uRKBCizg"
        self.base_url = "https://mineru.net/api/v4/extract/task"
        self.output_dir = r"D:\Desktop\项目\MinerU输出\原始文件"

        # 初始化
        self.setup_ui()
        self.cleanup_old_temp_files()
        self.is_converting = False

    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)

        title_label = ttk.Label(main_frame, text="PDF转原始文件工具（直接URL解析）", font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, pady=(0, 20))

        # URL输入区域
        url_frame = ttk.LabelFrame(main_frame, text="输入PDF文件URL", padding="15")
        url_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=10)
        url_frame.columnconfigure(0, weight=1)
        
        ttk.Label(url_frame, text="PDF文件URL:").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        self.pdf_url = tk.StringVar()
        url_entry = ttk.Entry(url_frame, textvariable=self.pdf_url, font=("Arial", 10))
        url_entry.grid(row=1, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        
        # URL示例提示
        url_hint = ttk.Label(url_frame, text="示例: https://example.com/document.pdf", foreground="gray", font=("Arial", 9))
        url_hint.grid(row=2, column=0, sticky=tk.W, pady=(5, 0))

        # 转换选项
        options_frame = ttk.LabelFrame(main_frame, text="转换选项", padding="15")
        options_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=10)
        
        # 模型版本选择
        model_frame = ttk.Frame(options_frame)
        model_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        ttk.Label(model_frame, text="模型版本:").grid(row=0, column=0, padx=(0, 10))
        self.model_version = tk.StringVar(value="vlm")
        ttk.Radiobutton(model_frame, text="VLM", variable=self.model_version, value="vlm").grid(row=0, column=1, padx=5)
        ttk.Radiobutton(model_frame, text="Layout", variable=self.model_version, value="layout").grid(row=0, column=2, padx=5)
        
        # 其他选项
        self.enable_ocr = tk.BooleanVar(value=True)
        self.enable_formula = tk.BooleanVar(value=False)
        ttk.Checkbutton(options_frame, text="启用OCR识别", variable=self.enable_ocr).grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Checkbutton(options_frame, text="启用公式识别", variable=self.enable_formula).grid(row=1, column=1, sticky=tk.W, pady=5)

        # 日志区域
        progress_frame = ttk.LabelFrame(main_frame, text="进度与日志", padding="15")
        progress_frame.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        progress_frame.columnconfigure(0, weight=1)
        progress_frame.rowconfigure(1, weight=1)

        self.progress = ttk.Progressbar(progress_frame, mode='indeterminate')
        self.progress.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        self.status_label = ttk.Label(progress_frame, text="就绪", wraplength=650)
        self.status_label.grid(row=1, column=0, sticky=tk.W)
        self.log_text = tk.Text(progress_frame, height=10, width=80, font=("Consolas", 9))
        self.log_text.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar = ttk.Scrollbar(progress_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.grid(row=2, column=1, sticky=(tk.N, tk.S))

        # 按钮
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, pady=20)
        self.convert_btn = ttk.Button(button_frame, text="开始转换", command=self.start_conversion)
        self.convert_btn.grid(row=0, column=0, padx=10, ipadx=20, ipady=5)
        self.cancel_btn = ttk.Button(button_frame, text="取消转换", command=self.cancel_conversion, state='disabled')
        self.cancel_btn.grid(row=0, column=1, padx=10)
        ttk.Button(button_frame, text="清空日志", command=self.clear_log).grid(row=0, column=2, padx=10)
        ttk.Button(button_frame, text="退出", command=self.cleanup_and_quit).grid(row=0, column=3, padx=10)

    # ==============================
    #  转换流程线程
    # ==============================
    def start_conversion(self):
        if not self.pdf_url.get():
            messagebox.showerror("错误", "请输入PDF文件的URL地址")
            return
        
        # 验证URL格式
        url = self.pdf_url.get().strip()
        if not url.startswith(('http://', 'https://')):
            messagebox.showerror("错误", "请输入有效的URL地址（以http://或https://开头）")
            return
        
        self.is_converting = True
        self.convert_btn.config(state='disabled')
        self.cancel_btn.config(state='normal')
        self.progress.start()
        self.status_label.config(text="开始解析PDF文件...")
        threading.Thread(target=self.convert_thread, daemon=True).start()

    def cancel_conversion(self):
        """取消转换"""
        self.is_converting = False
        self.log_message("用户取消转换")
        self.conversion_failed("转换已取消")

    def convert_thread(self):
        try:
            pdf_url = self.pdf_url.get().strip()
            
            # 调用MinerU API
            self.root.after(0, lambda: self.status_label.config(text="提交转换任务..."))
            result = self.call_mineru_api(pdf_url)
            if not result:
                self.root.after(0, lambda: self.conversion_failed("转换失败"))
                return

            # 下载结果
            download_url = result.get("full_zip_url")
            if not download_url:
                self.root.after(0, lambda: self.conversion_failed("未返回下载链接"))
                return

            self.root.after(0, lambda: self.status_label.config(text="下载转换结果..."))
            
            # 生成输出文件名（基于URL）
            file_name = self.generate_filename_from_url(pdf_url)
            success = self.download_and_extract_result(download_url, file_name)
            if success:
                self.root.after(0, self.conversion_success)
            else:
                self.root.after(0, lambda: self.conversion_failed("文件处理失败"))
                
        except Exception as e:
            self.root.after(0, lambda: self.conversion_failed(f"错误: {e}"))

    def generate_filename_from_url(self, url):
        """从URL生成文件名"""
        try:
            # 尝试从URL中提取文件名
            from urllib.parse import urlparse, unquote
            parsed_url = urlparse(url)
            path = unquote(parsed_url.path)
            
            if path and '/' in path:
                filename = path.split('/')[-1]
                if filename and '.' in filename:
                    return filename
                
            # 如果无法从URL提取，使用默认名称
            return f"document_{int(time.time())}"
            
        except:
            return f"document_{int(time.time())}"

    # ==============================
    #  调用 MinerU API
    # ==============================
    def call_mineru_api(self, pdf_url):
        """调用MinerU API并轮询任务状态"""
        headers = {
            "Content-Type": "application/json", 
            "Authorization": f"Bearer {self.token}"
        }
        
        data = {
            "url": pdf_url,
            "model_version": self.model_version.get(),
            "is_ocr": self.enable_ocr.get(),
            "enable_formula": self.enable_formula.get(),
            "output_format": "markdown"  # 固定为markdown格式
        }

        try:
            self.log_message(f"调用 MinerU API (模型版本: {self.model_version.get()})...")
            self.log_message(f"PDF URL: {pdf_url}")
            response = requests.post(self.base_url, headers=headers, json=data, timeout=30)
            
            if response.status_code != 200:
                self.log_message(f"API请求失败，状态码: {response.status_code}")
                self.log_message(f"响应内容: {response.text}")
                return None
                
            result = response.json()
            self.log_message(f"API响应: {result}")
            
            if result.get("code") == 0:
                task_id = result["data"].get("task_id")
                if task_id:
                    self.log_message(f"任务ID: {task_id}")
                    return self.poll_task_status(task_id)
                else:
                    self.log_message("未返回任务ID")
                    return None
            else:
                self.log_message(f"API返回错误: {result.get('message', '未知错误')}")
                return None
                
        except requests.exceptions.Timeout:
            self.log_message("API请求超时")
            return None
        except Exception as e:
            self.log_message(f"API调用出错: {e}")
            return None

    def poll_task_status(self, task_id):
        """轮询任务状态"""
        headers = {"Authorization": f"Bearer {self.token}"}
        status_url = f"https://mineru.net/api/v4/extract/task/{task_id}"
        
        max_attempts = 120
        attempt = 0
        
        while self.is_converting and attempt < max_attempts:
            try:
                attempt += 1
                self.log_message(f"查询任务状态 ({attempt}/{max_attempts})...")
                
                response = requests.get(status_url, headers=headers, timeout=30)
                if response.status_code != 200:
                    self.log_message(f"状态查询失败，状态码: {response.status_code}")
                    time.sleep(5)
                    continue
                    
                result = response.json()
                task_data = result.get("data", {})
                
                state = task_data.get("state")
                self.log_message(f"任务状态: {state}")
                
                if state == "done":
                    download_url = task_data.get("full_zip_url")
                    if download_url:
                        self.log_message("✅ 任务完成！")
                        return task_data
                    else:
                        self.log_message("❌ 任务完成但未返回下载链接")
                        return None
                elif state == "failed":
                    error_msg = task_data.get("err_msg", "未知错误")
                    self.log_message(f"❌ 任务失败: {error_msg}")
                    return None
                elif state == "pending":
                    self.log_message("任务排队中...")
                elif state == "processing":
                    progress = task_data.get("progress", 0)
                    self.log_message(f"处理进度: {progress}%")
                
                time.sleep(5)
                
            except requests.exceptions.Timeout:
                self.log_message("状态查询超时，继续重试...")
                time.sleep(5)
            except Exception as e:
                self.log_message(f"状态查询出错: {e}")
                time.sleep(5)
        
        if attempt >= max_attempts:
            self.log_message("❌ 任务轮询超时")
        return None

    # ==============================
    #  下载和解压原始文件
    # ==============================
    def download_and_extract_result(self, url, original_filename):
        """下载结果并解压到原始文件夹"""
        temp_zip = None
        
        try:
            os.makedirs(self.output_dir, exist_ok=True)
            name = Path(original_filename).stem
            
            # 创建输出文件夹
            output_folder = os.path.join(self.output_dir, f"{name}_原始文件")
            if os.path.exists(output_folder):
                shutil.rmtree(output_folder)
            os.makedirs(output_folder)
            
            # 下载zip文件
            temp_zip = os.path.join(tempfile.gettempdir(), f"mineru_temp_{uuid.uuid4().hex}.zip")
            
            self.log_message("下载转换结果...")
            response = requests.get(url, stream=True, timeout=60)
            response.raise_for_status()
            
            with open(temp_zip, "wb") as file:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        file.write(chunk)
            
            self.log_message("✅ 下载完成，开始解压...")
            
            # 解压到目标文件夹
            with zipfile.ZipFile(temp_zip, 'r') as zip_ref:
                zip_ref.extractall(output_folder)
            
            # 统计解压的文件
            file_count = 0
            for root, dirs, files in os.walk(output_folder):
                file_count += len(files)
            
            self.log_message(f"✅ 解压完成！共 {file_count} 个文件")
            self.log_message(f"📁 原始文件保存在: {output_folder}")
            
            # 列出主要文件
            self.log_message("📄 解压文件列表:")
            for item in os.listdir(output_folder):
                item_path = os.path.join(output_folder, item)
                if os.path.isfile(item_path):
                    size = os.path.getsize(item_path) / 1024  # KB
                    self.log_message(f"   📝 {item} ({size:.1f} KB)")
                else:
                    self.log_message(f"   📁 {item}/")
            
            # 清理临时zip文件
            if temp_zip and os.path.exists(temp_zip):
                os.remove(temp_zip)
                temp_zip = None
            
            return True
                    
        except Exception as e:
            self.log_message(f"❌ 处理失败: {e}")
            # 确保清理临时文件
            try:
                if temp_zip and os.path.exists(temp_zip):
                    os.remove(temp_zip)
            except:
                pass
            return False

    # ==============================
    #  通用工具函数
    # ==============================
    def conversion_success(self):
        """转换成功处理"""
        self.is_converting = False
        self.progress.stop()
        self.convert_btn.config(state='normal')
        self.cancel_btn.config(state='disabled')
        self.status_label.config(text="转换完成！")
        self.log_message("=== 转换完成 ===")
        messagebox.showinfo("成功", f"PDF转换成功！\n原始文件已保存到输出目录")

    def conversion_failed(self, msg):
        """转换失败处理"""
        self.is_converting = False
        self.progress.stop()
        self.convert_btn.config(state='normal')
        self.cancel_btn.config(state='disabled')
        self.status_label.config(text="转换失败")
        self.log_message(f"=== 转换失败: {msg} ===")
        if "取消" not in msg:  # 如果是用户取消，不显示错误对话框
            messagebox.showerror("错误", msg)

    def log_message(self, msg):
        """添加日志消息"""
        timestamp = time.strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {msg}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def clear_log(self):
        """清空日志"""
        self.log_text.delete(1.0, tk.END)

    def cleanup_old_temp_files(self):
        """清理临时文件"""
        temp_dir = tempfile.gettempdir()
        for item in os.listdir(temp_dir):
            if item.startswith("mineru_temp_") or item.startswith("temp_"):
                try:
                    full_path = os.path.join(temp_dir, item)
                    if os.path.isfile(full_path):
                        os.remove(full_path)
                except Exception as e:
                    print(f"清理临时文件失败: {e}")

    def cleanup_and_quit(self):
        """清理资源并退出"""
        self.is_converting = False
        self.cleanup_old_temp_files()
        self.root.quit()
        self.root.destroy()


def main():
    """主函数"""
    try:
        root = tk.Tk()
        app = PDFToRawConverter(root)
        root.protocol("WM_DELETE_WINDOW", app.cleanup_and_quit)
        root.mainloop()
    except Exception as e:
        print(f"程序启动失败: {e}")
        messagebox.showerror("错误", f"程序启动失败: {e}")


if __name__ == "__main__":
    main()

