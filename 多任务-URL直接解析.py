import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
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

class BatchPDFConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("MinerU PDF批量转原始文件工具")
        self.root.geometry("900x750")
        self.root.minsize(800, 700)

        # MinerU API 信息
        self.token = "eyJ0eXBlIjoiSldUIiwiYWxnIjoiSFM1MTIifQ.eyJqdGkiOiI1MzAwODI3OSIsInJvbCI6IlJPTEVfUkVHSVNURVIiLCJpc3MiOiJPcGVuWExhYiIsImlhdCI6MTc2MzU0NDM5NSwiY2xpZW50SWQiOiJsa3pkeDU3bnZ5MjJqa3BxOXgydyIsInBob25lIjoiMTg0NjAzMDAxOTciLCJvcGVuSWQiOm51bGwsInV1aWQiOiI5NjY3ODRiNC0wMjRjLTQ3NzUtYjE5Ny1kZWY5NTIyZmJjZDciLCJlbWFpbCI6IiIsImV4cCI6MTc2NDc1Mzk5NX0.HPAoPC83v5Xi-ZxjTshshZljtR7zTyTyKAVSt4qSCfCCShaVKWE7_K1bC2lWNrZJWi8r-hpTbv8ym6uRKBCizg"
        
        # 批量接口地址
        self.batch_task_url = "https://mineru.net/api/v4/extract/task/batch"
        self.batch_query_url = "https://mineru.net/api/v4/extract-results/batch/{}"
        
        self.output_dir = r"D:\Desktop\项目\MinerU输出\原始文件"

        # 初始化
        self.setup_ui()
        self.cleanup_old_temp_files()
        self.is_converting = False
        self.processed_files = set() # 用于记录批次中已处理完成的 data_id

    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1) # 让URL输入区域可伸缩

        # 标题
        header_frame = ttk.Frame(main_frame)
        header_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        title_label = ttk.Label(header_frame, text="PDF批量转换工具 (URL模式)", font=("微软雅黑", 16, "bold"))
        title_label.pack(side=tk.LEFT)

        # URL输入区域 (改为多行文本框)
        url_frame = ttk.LabelFrame(main_frame, text="输入PDF文件URL列表 (每行一个)", padding="10")
        url_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        url_frame.columnconfigure(0, weight=1)
        url_frame.rowconfigure(0, weight=1)

        self.url_text = scrolledtext.ScrolledText(url_frame, height=8, font=("Consolas", 10))
        self.url_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 示例文本
        example_text = "https://example.com/file1.pdf\nhttps://example.com/file2.pdf"
        self.url_text.insert(tk.END, example_text)
        # 绑定点击清除默认文本事件 (可选，这里简单处理不绑定，让用户自己删)

        # 转换选项
        options_frame = ttk.LabelFrame(main_frame, text="转换配置", padding="10")
        options_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=10)
        
        # 模型版本
        ttk.Label(options_frame, text="模型版本:").pack(side=tk.LEFT, padx=(0, 10))
        self.model_version = tk.StringVar(value="vlm")
        ttk.Radiobutton(options_frame, text="VLM (推荐)", variable=self.model_version, value="vlm").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(options_frame, text="Layout", variable=self.model_version, value="layout").pack(side=tk.LEFT, padx=5)
        
        # 功能开关
        ttk.Separator(options_frame, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=20)
        self.enable_ocr = tk.BooleanVar(value=True)
        self.enable_formula = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="OCR识别", variable=self.enable_ocr).pack(side=tk.LEFT, padx=5)
        ttk.Checkbutton(options_frame, text="公式识别", variable=self.enable_formula).pack(side=tk.LEFT, padx=5)

        # 日志区域
        log_frame = ttk.LabelFrame(main_frame, text="任务日志", padding="10")
        log_frame.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        log_frame.columnconfigure(0, weight=1)
        
        # 进度条
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(log_frame, mode='indeterminate')
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        
        # 状态标签
        self.status_label = ttk.Label(log_frame, text="就绪 - 等待任务提交", foreground="blue")
        self.status_label.grid(row=1, column=0, sticky=tk.W)

        # 日志文本框
        self.log_text = scrolledtext.ScrolledText(log_frame, height=12, state='disabled', font=("Consolas", 9))
        self.log_text.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(5,0))

        # 底部按钮
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=4, column=0, pady=15)
        
        self.convert_btn = ttk.Button(btn_frame, text="开始批量转换", command=self.start_conversion, width=20)
        self.convert_btn.pack(side=tk.LEFT, padx=10)
        
        self.cancel_btn = ttk.Button(btn_frame, text="停止转换", command=self.cancel_conversion, state='disabled', width=15)
        self.cancel_btn.pack(side=tk.LEFT, padx=10)
        
        ttk.Button(btn_frame, text="清空URL", command=lambda: self.url_text.delete(1.0, tk.END)).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="退出", command=self.cleanup_and_quit).pack(side=tk.LEFT, padx=10)

    # ==============================
    #  核心逻辑
    # ==============================

    def start_conversion(self):
        """预处理并启动线程"""
        # 获取并清洗URL
        raw_text = self.url_text.get(1.0, tk.END)
        urls = [line.strip() for line in raw_text.splitlines() if line.strip()]
        
        # 过滤无效URL
        valid_urls = [u for u in urls if u.startswith(('http://', 'https://'))]
        
        if not valid_urls:
            messagebox.showerror("提示", "请至少输入一个有效的URL (以http或https开头)")
            return

        if len(valid_urls) > 200:
            messagebox.showwarning("提示", "单次批量任务不能超过200个URL，将截取前200个。")
            valid_urls = valid_urls[:200]

        self.is_converting = True
        self.processed_files.clear() # 清空已完成记录
        self.toggle_ui_state(processing=True)
        self.progress_bar.start(10)
        
        self.log_message(f"准备提交 {len(valid_urls)} 个文件的转换任务...")
        
        # 启动后台线程
        threading.Thread(target=self.batch_process_thread, args=(valid_urls,), daemon=True).start()

    def cancel_conversion(self):
        self.is_converting = False
        self.log_message("❌ 用户请求停止，正在中断当前操作...")
        self.status_label.config(text="正在停止...")

    def toggle_ui_state(self, processing=True):
        if processing:
            self.convert_btn.config(state='disabled')
            self.cancel_btn.config(state='normal')
            self.url_text.config(state='disabled')
        else:
            self.convert_btn.config(state='normal')
            self.cancel_btn.config(state='disabled')
            self.url_text.config(state='normal')
            self.progress_bar.stop()

    def batch_process_thread(self, urls):
        try:
            # 1. 构造批量请求数据
            files_payload = []
            url_map = {} # data_id -> url (用于日志显示)
            
            for url in urls:
                # 生成唯一的 data_id 用于追踪
                data_id = f"task_{uuid.uuid4().hex[:8]}"
                files_payload.append({
                    "url": url,
                    "data_id": data_id
                })
                url_map[data_id] = url

            # 2. 提交批量任务
            batch_id = self.submit_batch_task(files_payload)
            
            if not batch_id:
                self.root.after(0, lambda: self.finish_conversion("任务提交失败", error=True))
                return

            self.log_message(f"✅ 批量任务提交成功! Batch ID: {batch_id}")
            self.root.after(0, lambda: self.status_label.config(text="任务运行中...正在轮询结果"))

            # 3. 轮询结果
            self.poll_batch_results(batch_id, len(files_payload), url_map)

        except Exception as e:
            self.log_message(f"❌ 发生严重错误: {str(e)}")
            self.root.after(0, lambda: self.finish_conversion("发生异常", error=True))

    def submit_batch_task(self, files_payload):
        """提交批量任务到 API"""
        headers = {
            "Content-Type": "application/json", 
            "Authorization": f"Bearer {self.token}"
        }
        data = {
            "files": files_payload,
            "model_version": self.model_version.get(),
            "enable_ocr": self.enable_ocr.get(),
            "enable_formula": self.enable_formula.get()
        }
        
        try:
            resp = requests.post(self.batch_task_url, headers=headers, json=data, timeout=30)
            result = resp.json()
            
            if resp.status_code == 200 and result.get("code") == 0:
                return result["data"]["batch_id"]
            else:
                self.log_message(f"API提交失败: {result.get('msg', resp.text)}")
                return None
        except Exception as e:
            self.log_message(f"网络请求错误: {e}")
            return None

    def poll_batch_results(self, batch_id, total_count, url_map):
        """轮询批量任务状态"""
        headers = {"Authorization": f"Bearer {self.token}"}
        url = self.batch_query_url.format(batch_id)
        
        while self.is_converting:
            try:
                resp = requests.get(url, headers=headers, timeout=30)
                if resp.status_code != 200:
                    self.log_message(f"轮询请求失败: {resp.status_code}")
                    time.sleep(5)
                    continue

                res_json = resp.json()
                if res_json.get("code") != 0:
                    self.log_message(f"查询出错: {res_json.get('msg')}")
                    time.sleep(5)
                    continue

                # 解析结果列表
                extract_results = res_json["data"].get("extract_result", [])
                
                # 统计状态
                done_count = 0
                failed_count = 0
                running_count = 0
                
                current_round_updates = 0

                for item in extract_results:
                    state = item.get("state")
                    data_id = item.get("data_id")
                    file_name = item.get("file_name", "unknown")
                    
                    # 如果该文件已经处理过（已下载或已报错），跳过
                    if data_id in self.processed_files:
                        if state == "done": done_count += 1
                        elif state == "failed": failed_count += 1
                        continue

                    # 处理新状态
                    if state == "done":
                        # 下载文件
                        dl_url = item.get("full_zip_url")
                        if dl_url:
                            self.log_message(f"📥 文件 [{file_name}] 解析完成，开始下载...")
                            success = self.download_and_extract(dl_url, file_name)
                            if success:
                                self.processed_files.add(data_id)
                                done_count += 1
                                current_round_updates += 1
                        else:
                            self.log_message(f"⚠️ 文件 [{file_name}] 完成但无下载链接")
                            
                    elif state == "failed":
                        err_msg = item.get("err_msg", "未知原因")
                        self.log_message(f"❌ 文件 [{file_name}] 解析失败: {err_msg}")
                        self.processed_files.add(data_id) # 标记为已处理（避免重复报错）
                        failed_count += 1
                        current_round_updates += 1
                        
                    elif state in ["running", "pending", "waiting-file", "converting"]:
                        running_count += 1

                # 更新 UI 状态
                progress_pct = ((done_count + failed_count) / total_count) * 100
                status_msg = f"进度: {done_count + failed_count}/{total_count} (成功: {done_count}, 失败: {failed_count}, 进行中: {running_count})"
                self.root.after(0, lambda: self.status_label.config(text=status_msg))
                
                # 如果所有任务都结束了
                if (done_count + failed_count) >= total_count:
                    self.log_message("✨ 所有任务处理完毕！")
                    self.root.after(0, lambda: self.finish_conversion("所有文件处理完成"))
                    break

                time.sleep(5) # 间隔5秒轮询一次

            except Exception as e:
                self.log_message(f"轮询循环异常: {e}")
                time.sleep(5)

    # ==============================
    #  文件下载与解压 (复用逻辑)
    # ==============================

    def download_and_extract(self, url, filename):
        temp_zip = None
        try:
            # 准备目录
            safe_name = Path(filename).stem
            output_folder = os.path.join(self.output_dir, f"{safe_name}_解析结果")
            os.makedirs(output_folder, exist_ok=True)

            # 下载
            temp_zip = os.path.join(tempfile.gettempdir(), f"mineru_{uuid.uuid4().hex}.zip")
            with requests.get(url, stream=True, timeout=60) as r:
                r.raise_for_status()
                with open(temp_zip, 'wb') as f:
                    for chunk in r.iter_content(chunk_size=8192):
                        f.write(chunk)

            # 解压
            with zipfile.ZipFile(temp_zip, 'r') as z:
                z.extractall(output_folder)
            
            self.log_message(f"✅ 已保存至: {output_folder}")
            return True

        except Exception as e:
            self.log_message(f"❌ 下载解压失败 [{filename}]: {e}")
            return False
        finally:
            if temp_zip and os.path.exists(temp_zip):
                try:
                    os.remove(temp_zip)
                except: pass

    # ==============================
    #  辅助函数
    # ==============================

    def log_message(self, msg):
        timestamp = time.strftime("%H:%M:%S")
        full_msg = f"[{timestamp}] {msg}\n"
        
        def _update():
            self.log_text.config(state='normal')
            self.log_text.insert(tk.END, full_msg)
            self.log_text.see(tk.END)
            self.log_text.config(state='disabled')
        
        self.root.after(0, _update)

    def finish_conversion(self, msg, error=False):
        self.is_converting = False
        self.toggle_ui_state(processing=False)
        self.status_label.config(text=msg)
        
        if error:
            messagebox.showerror("结束", msg)
        else:
            messagebox.showinfo("完成", f"{msg}\n文件已保存到: {self.output_dir}")

    def cleanup_old_temp_files(self):
        temp_dir = tempfile.gettempdir()
        for item in os.listdir(temp_dir):
            if item.startswith("mineru_"):
                try:
                    os.remove(os.path.join(temp_dir, item))
                except: pass

    def cleanup_and_quit(self):
        self.is_converting = False
        self.cleanup_old_temp_files()
        self.root.quit()
        self.root.destroy()

if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = BatchPDFConverter(root)
        root.protocol("WM_DELETE_WINDOW", app.cleanup_and_quit)
        root.mainloop()
    except Exception as e:
        print(f"Error: {e}")
