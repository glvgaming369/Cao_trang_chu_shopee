import os
import json
import threading
import time
import queue
from typing import Optional

import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk

from googleapiclient.discovery import build
import shopee_scraper_core as scraper

# Cấu hình CustomTkinter
ctk.set_appearance_mode("Dark")  # Mặc định là Dark mode để trông premium
ctk.set_default_color_theme("blue")

# Màu sắc chủ đạo Shopee
COLOR_SHOPEE = "#EE4D2D"
COLOR_SHOPEE_HOVER = "#D33A1C"
COLOR_SUCCESS = "#2E7D32"
COLOR_SUCCESS_HOVER = "#1B5E20"

class ShopeeScraperApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Cấu hình cửa sổ chính
        self.title("Shopee Product Scraper & Sheet Sync Tool")
        self.geometry("850x750")
        self.resizable(True, True)
        
        # Biến trạng thái
        self.is_running = False
        self.product_queue = queue.Queue()
        self.config_dir = "config"
        self.client_secret_default = os.path.join(self.config_dir, "client_secret.json")
        self.token_default = os.path.join(self.config_dir, "token.json")
        self.ui_config_path = os.path.join(self.config_dir, "config.json")
        
        # Tạo thư mục config nếu chưa có
        os.makedirs(self.config_dir, exist_ok=True)
        
        # Tạo giao diện
        self.create_widgets()
        
        # Tải cấu hình UI đã lưu
        self.load_ui_config()
        
    def create_widgets(self):
        # Thiết lập grid layout chính (2 cột: Trái là cấu hình, Phải là danh sách link, Dưới là log)
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=0) # Tiêu đề
        self.grid_rowconfigure(1, weight=0) # Panels
        self.grid_rowconfigure(2, weight=1) # Log & Progress
        
        # 1. Tiêu đề ứng dụng
        title_label = ctk.CTkLabel(
            self, 
            text="SHOPEE SCRAPER PRO", 
            font=ctk.CTkFont(size=22, weight="bold"),
            text_color=COLOR_SHOPEE
        )
        title_label.grid(row=0, column=0, columnspan=2, padx=20, pady=15, sticky="ew")
        
        # 2. Khung trái: Cấu hình Google Sheets & Authentication
        left_frame = ctk.CTkFrame(self)
        left_frame.grid(row=1, column=0, padx=(20, 10), pady=10, sticky="nsew")
        left_frame.grid_columnconfigure(0, weight=1)
        
        left_title = ctk.CTkLabel(left_frame, text="⚙️ CẤU HÌNH GOOGLE SHEETS", font=ctk.CTkFont(size=14, weight="bold"))
        left_title.grid(row=0, column=0, columnspan=2, padx=15, pady=(15, 10), sticky="w")
        
        # Client Secret JSON Path
        lbl_secret = ctk.CTkLabel(left_frame, text="Đường dẫn Credentials JSON:", font=ctk.CTkFont(size=11))
        lbl_secret.grid(row=1, column=0, padx=15, pady=(5, 0), sticky="w")
        
        self.entry_secret = ctk.CTkEntry(left_frame, placeholder_text="config/client_secret.json")
        self.entry_secret.grid(row=2, column=0, padx=(15, 5), pady=2, sticky="ew")
        self.entry_secret.insert(0, self.client_secret_default)
        
        btn_browse_secret = ctk.CTkButton(
            left_frame, 
            text="Browse", 
            width=70, 
            command=self.browse_client_secret
        )
        btn_browse_secret.grid(row=2, column=1, padx=(0, 15), pady=2, sticky="ew")
        
        # OAuth Action Buttons
        auth_buttons_frame = ctk.CTkFrame(left_frame, fg_color="transparent")
        auth_buttons_frame.grid(row=3, column=0, columnspan=2, padx=15, pady=8, sticky="ew")
        auth_buttons_frame.grid_columnconfigure(0, weight=1)
        auth_buttons_frame.grid_columnconfigure(1, weight=1)
        
        self.btn_auth = ctk.CTkButton(
            auth_buttons_frame, 
            text="🔑 Xác thực Google", 
            fg_color=COLOR_SUCCESS, 
            hover_color=COLOR_SUCCESS_HOVER,
            command=self.start_authentication
        )
        self.btn_auth.grid(row=0, column=0, padx=(0, 5), pady=0, sticky="ew")
        
        self.btn_clear_token = ctk.CTkButton(
            auth_buttons_frame, 
            text="🗑️ Xoá Token", 
            fg_color="#D32F2F", 
            hover_color="#C62828",
            command=self.clear_token
        )
        self.btn_clear_token.grid(row=0, column=1, padx=(5, 0), pady=0, sticky="ew")
        
        # Spreadsheet ID / URL
        lbl_sheet_url = ctk.CTkLabel(left_frame, text="URL hoặc ID Google Sheet:", font=ctk.CTkFont(size=11))
        lbl_sheet_url.grid(row=4, column=0, columnspan=2, padx=15, pady=(5, 0), sticky="w")
        
        self.entry_sheet_url = ctk.CTkEntry(
            left_frame, 
            placeholder_text="https://docs.google.com/spreadsheets/d/.../edit"
        )
        self.entry_sheet_url.grid(row=5, column=0, columnspan=2, padx=15, pady=2, sticky="ew")
        
        # Sheet Tab Name
        lbl_sheet_tab = ctk.CTkLabel(left_frame, text="Tên Sheet Tab Muốn Ghi:", font=ctk.CTkFont(size=11))
        lbl_sheet_tab.grid(row=6, column=0, columnspan=2, padx=15, pady=(5, 0), sticky="w")
        
        self.entry_sheet_tab = ctk.CTkEntry(left_frame, placeholder_text="Sheet1")
        self.entry_sheet_tab.grid(row=7, column=0, columnspan=2, padx=15, pady=(2, 5), sticky="ew")
        self.entry_sheet_tab.insert(0, "Sheet1")
        
        # Test Connect Button
        self.btn_test_connect = ctk.CTkButton(
            left_frame, 
            text="🧪 Kiểm tra kết nối (Test Connect)", 
            fg_color="#6A1B9A", 
            hover_color="#4A148C",
            command=self.test_connection
        )
        self.btn_test_connect.grid(row=8, column=0, columnspan=2, padx=15, pady=(5, 15), sticky="ew")
        
        # 3. Khung phải: Nhập sản phẩm & Proxy
        right_frame = ctk.CTkFrame(self)
        right_frame.grid(row=1, column=1, padx=(10, 20), pady=10, sticky="nsew")
        right_frame.grid_columnconfigure(0, weight=1)
        
        right_title = ctk.CTkLabel(right_frame, text="📦 THÔNG TIN SẢN PHẨM & PROXY", font=ctk.CTkFont(size=14, weight="bold"))
        right_title.grid(row=0, column=0, columnspan=2, padx=15, pady=(15, 10), sticky="w")
        
        # ID / Link shopee
        lbl_shopee = ctk.CTkLabel(right_frame, text="Danh sách ID hoặc Link sản phẩm Shopee (Mỗi dòng một mục):", font=ctk.CTkFont(size=11))
        lbl_shopee.grid(row=1, column=0, padx=15, pady=(5, 0), sticky="w")
        
        self.txt_shopee_input = ctk.CTkTextbox(right_frame, height=100)
        self.txt_shopee_input.grid(row=2, column=0, padx=(15, 5), pady=2, sticky="ew")
        
        self.btn_import_file = ctk.CTkButton(
            right_frame, 
            text="📂 Import\nFile .txt", 
            width=70, 
            command=self.import_txt_file
        )
        self.btn_import_file.grid(row=2, column=1, padx=(0, 15), pady=2, sticky="ns")
        
        # Proxy (Optional) - NÂNG CẤP THÀNH NHẬP NHIỀU DÒNG
        lbl_proxy = ctk.CTkLabel(right_frame, text="Danh sách Proxy (Mỗi dòng một proxy, optional - ip:port:user:pass hoặc ip:port):", font=ctk.CTkFont(size=11))
        lbl_proxy.grid(row=3, column=0, columnspan=2, padx=15, pady=(5, 0), sticky="w")
        
        self.txt_proxy_input = ctk.CTkTextbox(right_frame, height=60)
        self.txt_proxy_input.grid(row=4, column=0, columnspan=2, padx=15, pady=(2, 10), sticky="ew")
        
        # Filter Frame (Lượt bán tối thiểu & Hoa hồng người bán tối thiểu)
        filter_frame = ctk.CTkFrame(right_frame, fg_color="transparent")
        filter_frame.grid(row=5, column=0, columnspan=2, padx=15, pady=5, sticky="ew")
        filter_frame.grid_columnconfigure(0, weight=1)
        filter_frame.grid_columnconfigure(1, weight=1)
        
        lbl_sold_min = ctk.CTkLabel(filter_frame, text="Lượt bán tối thiểu:", font=ctk.CTkFont(size=11))
        lbl_sold_min.grid(row=0, column=0, padx=(0, 5), pady=0, sticky="w")
        
        self.entry_sold_min = ctk.CTkEntry(filter_frame, placeholder_text="Mặc định: 1")
        self.entry_sold_min.grid(row=1, column=0, padx=(0, 5), pady=2, sticky="ew")
        self.entry_sold_min.insert(0, "1")
        
        lbl_comm_min = ctk.CTkLabel(filter_frame, text="Hoa hồng người bán tối thiểu (đ):", font=ctk.CTkFont(size=11))
        lbl_comm_min.grid(row=0, column=1, padx=(5, 0), pady=0, sticky="w")
        
        self.entry_comm_min = ctk.CTkEntry(filter_frame, placeholder_text="Mặc định: 1")
        self.entry_comm_min.grid(row=1, column=1, padx=(5, 0), pady=2, sticky="ew")
        self.entry_comm_min.insert(0, "1")

        # Controls Frame (Start / Stop)
        controls_frame = ctk.CTkFrame(right_frame, fg_color="transparent")
        controls_frame.grid(row=6, column=0, columnspan=2, padx=15, pady=(5, 15), sticky="ew")
        controls_frame.grid_columnconfigure(0, weight=1)
        controls_frame.grid_columnconfigure(1, weight=1)
        
        self.btn_start = ctk.CTkButton(
            controls_frame, 
            text="▶️ Bắt đầu (Start)", 
            fg_color=COLOR_SHOPEE, 
            hover_color=COLOR_SHOPEE_HOVER,
            font=ctk.CTkFont(size=13, weight="bold"),
            command=self.start_scraping
        )
        self.btn_start.grid(row=0, column=0, padx=(0, 5), pady=0, sticky="ew")
        
        self.btn_stop = ctk.CTkButton(
            controls_frame, 
            text="⏹️ Dừng lại (Stop)", 
            fg_color="#757575", 
            hover_color="#616161",
            font=ctk.CTkFont(size=13, weight="bold"),
            state="disabled",
            command=self.stop_scraping
        )
        self.btn_stop.grid(row=0, column=1, padx=(5, 0), pady=0, sticky="ew")
        
        # 4. Khung dưới: Log và Progress Bar
        bottom_frame = ctk.CTkFrame(self)
        bottom_frame.grid(row=2, column=0, columnspan=2, padx=20, pady=(10, 20), sticky="nsew")
        bottom_frame.grid_columnconfigure(0, weight=1)
        bottom_frame.grid_rowconfigure(2, weight=1)
        
        # Progress Bar & Label
        progress_info_frame = ctk.CTkFrame(bottom_frame, fg_color="transparent")
        progress_info_frame.grid(row=0, column=0, padx=15, pady=(10, 2), sticky="ew")
        progress_info_frame.grid_columnconfigure(0, weight=1)
        
        self.lbl_progress = ctk.CTkLabel(progress_info_frame, text="Tiến trình: Đang chờ...", font=ctk.CTkFont(size=11))
        self.lbl_progress.grid(row=0, column=0, sticky="w")
        
        self.lbl_percentage = ctk.CTkLabel(progress_info_frame, text="0%", font=ctk.CTkFont(size=11, weight="bold"))
        self.lbl_percentage.grid(row=0, column=1, sticky="e")
        
        self.progress_bar = ctk.CTkProgressBar(bottom_frame, progress_color=COLOR_SHOPEE)
        self.progress_bar.grid(row=1, column=0, padx=15, pady=(0, 10), sticky="ew")
        self.progress_bar.set(0)
        
        # Logs View
        lbl_logs = ctk.CTkLabel(bottom_frame, text="Nhật ký Log hệ thống:", font=ctk.CTkFont(size=12, weight="bold"))
        lbl_logs.grid(row=2, column=0, padx=15, pady=(0, 2), sticky="w")
        
        self.txt_logs = ctk.CTkTextbox(bottom_frame, font=ctk.CTkFont(family="Consolas", size=11))
        self.txt_logs.grid(row=3, column=0, padx=15, pady=(0, 15), sticky="nsew")
        self.txt_logs.configure(state="disabled")
        
    # --- Điều hướng File và Config ---
    
    def log(self, message: str, level: str = "INFO"):
        """Ghi log an toàn vào CTkTextbox từ bất kỳ luồng nào."""
        timestamp = time.strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] [{level}] {message}\n"
        
        def write_log():
            self.txt_logs.configure(state="normal")
            self.txt_logs.insert(tk.END, log_entry)
            self.txt_logs.configure(state="disabled")
            self.txt_logs.see(tk.END)
            
        self.after(0, write_log)
        
    def browse_client_secret(self):
        filename = filedialog.askopenfilename(
            initialdir=self.config_dir,
            title="Chọn file credentials JSON",
            filetypes=(("JSON files", "*.json"), ("All files", "*.*"))
        )
        if filename:
            self.entry_secret.delete(0, tk.END)
            self.entry_secret.insert(0, filename)
            self.save_ui_config()
            
    def import_txt_file(self):
        filename = filedialog.askopenfilename(
            title="Chọn file TXT chứa danh sách sản phẩm",
            filetypes=(("Text files", "*.txt"), ("All files", "*.*"))
        )
        if filename:
            try:
                with open(filename, 'r', encoding='utf-8') as f:
                    lines = f.readlines()
                
                # Tách dòng, loại bỏ khoảng trắng và dòng trống
                raw_lines = [line.strip() for line in lines if line.strip()]
                total_raw = len(raw_lines)
                
                # Lọc trùng giữ nguyên thứ tự dòng
                unique_lines = list(dict.fromkeys(raw_lines))
                total_unique = len(unique_lines)
                removed_dup = total_raw - total_unique
                
                if unique_lines:
                    content = "\n".join(unique_lines)
                    self.txt_shopee_input.delete("1.0", tk.END)
                    self.txt_shopee_input.insert("1.0", content)
                    self.log(f"Đã import {total_unique} dòng từ file {os.path.basename(filename)} (Tự động loại bỏ {removed_dup} dòng trùng lặp).")
                else:
                    messagebox.showwarning("Cảnh báo", "File txt này không chứa dữ liệu hợp lệ!")
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể đọc file: {str(e)}")
                
    def load_ui_config(self):
        if os.path.exists(self.ui_config_path):
            try:
                with open(self.ui_config_path, 'r', encoding='utf-8') as f:
                    cfg = json.load(f)
                
                # Điền lại giao diện
                if "client_secret" in cfg:
                    self.entry_secret.delete(0, tk.END)
                    self.entry_secret.insert(0, cfg["client_secret"])
                if "sheet_url" in cfg:
                    self.entry_sheet_url.delete(0, tk.END)
                    self.entry_sheet_url.insert(0, cfg["sheet_url"])
                if "sheet_tab" in cfg:
                    self.entry_sheet_tab.delete(0, tk.END)
                    self.entry_sheet_tab.insert(0, cfg["sheet_tab"])
                if "proxy" in cfg:
                    self.txt_proxy_input.delete("1.0", tk.END)
                    self.txt_proxy_input.insert("1.0", cfg["proxy"])
                if "sold_min" in cfg:
                    self.entry_sold_min.delete(0, tk.END)
                    self.entry_sold_min.insert(0, cfg["sold_min"])
                if "comm_min" in cfg:
                    self.entry_comm_min.delete(0, tk.END)
                    self.entry_comm_min.insert(0, cfg["comm_min"])
                    
                self.log("Đã tải cấu hình UI lưu trước đó.")
            except Exception as e:
                self.log(f"Không thể tải cấu hình UI: {str(e)}", "WARN")
                
    def save_ui_config(self):
        cfg = {
            "client_secret": self.entry_secret.get().strip(),
            "sheet_url": self.entry_sheet_url.get().strip(),
            "sheet_tab": self.entry_sheet_tab.get().strip(),
            "proxy": self.txt_proxy_input.get("1.0", tk.END).strip(),
            "sold_min": self.entry_sold_min.get().strip(),
            "comm_min": self.entry_comm_min.get().strip()
        }
        try:
            with open(self.ui_config_path, 'w', encoding='utf-8') as f:
                json.dump(cfg, f, ensure_ascii=False, indent=4)
        except Exception as e:
            self.log(f"Không thể lưu cấu hình UI: {str(e)}", "WARN")

    def clear_token(self):
        """Xoá file token.json để buộc xác thực lại."""
        if os.path.exists(self.token_default):
            try:
                os.remove(self.token_default)
                self.log("Đã xoá file token.json thành công. Vui lòng Xác thực lại.")
                messagebox.showinfo("Thành công", "Đã xoá Token lưu trữ thành công!")
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể xoá file token: {str(e)}")
        else:
            self.log("Không tìm thấy file token.json trong thư mục config.", "WARN")
            messagebox.showwarning("Cảnh báo", "Chưa lưu token nào trước đó!")

    # --- Triển khai Authentication độc lập ---
    
    def start_authentication(self):
        """Chạy luồng xác thực Google Sheets trên một thread phụ để tránh treo UI."""
        client_secret = self.entry_secret.get().strip()
        if not client_secret or not os.path.exists(client_secret):
            messagebox.showerror("Lỗi", f"Không tìm thấy file JSON credentials tại:\n{client_secret}")
            return
            
        self.btn_auth.configure(state="disabled", text="🔄 Đang xác thực...")
        self.save_ui_config()
        
        def auth_thread():
            try:
                self.log("Bắt đầu khởi chạy luồng xác thực Google OAuth...")
                creds = scraper.authenticate_google(
                    client_secret_path=client_secret,
                    token_path=self.token_default
                )
                if creds:
                    self.log("Xác thực thành công! Token đã được lưu.", "SUCCESS")
                    self.after(0, lambda: messagebox.showinfo("Thành công", "Xác thực Google Sheets thành công và lưu token!"))
            except Exception as e:
                self.log(f"Lỗi xác thực: {str(e)}", "ERROR")
                err_msg = str(e)
                self.after(0, lambda: messagebox.showerror("Lỗi xác thực", err_msg))
            finally:
                self.after(0, lambda: self.btn_auth.configure(state="normal", text="🔑 Xác thực Google"))
                
        threading.Thread(target=auth_thread, daemon=True).start()

    def test_connection(self):
        """Kiểm tra kết nối Google Sheets và quyền ghi."""
        client_secret = self.entry_secret.get().strip()
        sheet_url = self.entry_sheet_url.get().strip()
        sheet_tab = self.entry_sheet_tab.get().strip()
        
        if not client_secret or not os.path.exists(client_secret):
            messagebox.showerror("Lỗi cấu hình", f"Vui lòng cấu hình file credentials JSON hợp lệ tại:\n{client_secret}")
            return
            
        if not sheet_url:
            messagebox.showerror("Lỗi cấu hình", "Vui lòng nhập URL hoặc ID Google Sheet!")
            return
            
        spreadsheet_id = scraper.get_spreadsheet_id(sheet_url)
        if not spreadsheet_id:
            messagebox.showerror("Lỗi cấu hình", "Không trích xuất được ID từ link Google Sheet!")
            return
            
        self.btn_test_connect.configure(state="disabled", text="🔄 Đang kết nối...")
        self.save_ui_config()
        
        def test_thread():
            try:
                self.log("Bắt đầu kiểm tra kết nối Google Sheets...")
                creds = scraper.authenticate_google(client_secret, self.token_default)
                service = build('sheets', 'v4', credentials=creds)
                
                # Gọi API kiểm tra tồn tại và tạo headers
                self.log(f"Đang kiểm tra quyền ghi trên Spreadsheet: {spreadsheet_id}...")
                success = scraper.ensure_sheet_and_headers(service, spreadsheet_id, sheet_tab)
                
                if success:
                    self.log(f"Kết nối thành công! Trang tính '{sheet_tab}' đã sẵn sàng.", "SUCCESS")
                    self.after(0, lambda: messagebox.showinfo("Thành công", f"Kết nối Google Sheets thành công!\nTrang tính '{sheet_tab}' đã sẵn sàng để ghi dữ liệu."))
                else:
                    self.log("Kết nối thất bại. Không thể kiểm tra/tạo tiêu đề trang tính.", "ERROR")
                    self.after(0, lambda: messagebox.showerror("Thất bại", "Không thể ghi tiêu đề vào trang tính. Vui lòng kiểm tra lại quyền ghi (Editor)."))
            except Exception as e:
                self.log(f"Lỗi kết nối: {str(e)}", "ERROR")
                err_msg = str(e)
                self.after(0, lambda: messagebox.showerror("Lỗi kết nối", f"Kết nối thất bại:\n{err_msg}"))
            finally:
                self.after(0, lambda: self.btn_test_connect.configure(state="normal", text="🧪 Kiểm tra kết nối (Test Connect)"))
                
        threading.Thread(target=test_thread, daemon=True).start()

    # --- Triển khai Tiến trình cào và đồng bộ chính với MULTI-WORKER ---
    
    def stop_scraping(self):
        """Kích hoạt cờ dừng và dọn sạch Queue công việc để dừng nhanh nhất."""
        if self.is_running:
            self.is_running = False
            self.log("Đang huỷ hàng đợi và dừng các workers. Vui lòng đợi kết thúc request hiện tại...")
            self.btn_stop.configure(state="disabled", text="🔄 Đang dừng...")
            
            # Xoá sạch các phần tử còn lại trong Queue để giải phóng các thread ngay khi xong tác vụ hiện tại
            while not self.product_queue.empty():
                try:
                    self.product_queue.get_nowait()
                    self.product_queue.task_done()
                except Exception:
                    break

    def start_scraping(self):
        """Bắt đầu tiến trình chính chia việc cho nhiều worker chạy song song."""
        if self.is_running:
            return
            
        # Thu thập và kiểm tra dữ liệu đầu vào
        client_secret = self.entry_secret.get().strip()
        sheet_url = self.entry_sheet_url.get().strip()
        sheet_tab = self.entry_sheet_tab.get().strip()
        input_text = self.txt_shopee_input.get("1.0", tk.END).strip()
        proxy_str = self.txt_proxy_input.get("1.0", tk.END).strip()
        
        # Lấy các điều kiện lọc dữ liệu
        try:
            sold_min = int(self.entry_sold_min.get().strip())
        except ValueError:
            sold_min = 1
            
        try:
            comm_min = int(self.entry_comm_min.get().strip())
        except ValueError:
            comm_min = 1
        
        # 1. Kiểm tra credentials
        if not client_secret or not os.path.exists(client_secret):
            messagebox.showerror("Lỗi cấu hình", f"Vui lòng cấu hình file credentials JSON hợp lệ tại:\n{client_secret}")
            return
            
        # 2. Kiểm tra Google Sheets URL
        if not sheet_url:
            messagebox.showerror("Lỗi cấu hình", "Vui lòng nhập URL hoặc ID Google Sheet!")
            return
            
        spreadsheet_id = scraper.get_spreadsheet_id(sheet_url)
        if not spreadsheet_id:
            messagebox.showerror("Lỗi cấu hình", "Không trích xuất được ID từ link Google Sheet!")
            return
            
        # 3. Kiểm tra danh sách sản phẩm
        if not input_text:
            messagebox.showerror("Lỗi cấu hình", "Vui lòng nhập ID hoặc Link sản phẩm Shopee cần cào!")
            return
            
        # Phân tích danh sách link đầu vào và lọc trùng theo ID sản phẩm
        raw_items = [line.strip() for line in input_text.split("\n") if line.strip()]
        shopee_items = []
        seen_run_ids = set()
        duplicate_count = 0
        
        for ri in raw_items:
            item_id = scraper.extract_item_id(ri)
            if item_id:
                if item_id not in seen_run_ids:
                    seen_run_ids.add(item_id)
                    shopee_items.append((ri, item_id))
                else:
                    duplicate_count += 1
            else:
                self.log(f"Bỏ qua dòng không đúng định dạng sản phẩm Shopee: '{ri[:50]}...'", "WARN")
                
        if duplicate_count > 0:
            self.log(f"Đã tự động loại bỏ {duplicate_count} sản phẩm bị trùng lặp ID trong danh sách đầu vào.", "INFO")
                
        if not shopee_items:
            messagebox.showerror("Lỗi cấu hình", "Không tìm thấy bất kỳ ID hoặc Link sản phẩm Shopee hợp lệ nào!")
            return
            
        # 4. Phân tích danh sách Proxy
        proxies_list = [line.strip() for line in proxy_str.split("\n") if line.strip()]
        if not proxies_list:
            proxies_list = [None]  # Chạy 1 luồng duy nhất không có proxy
            num_workers = 1
        else:
            num_workers = len(proxies_list)
            
        # Lưu lại cấu hình UI
        self.save_ui_config()
        
        # Cập nhật trạng thái UI sang chế độ đang chạy
        self.is_running = True
        self.btn_start.configure(state="disabled", text="🚀 Đang chạy...")
        self.btn_stop.configure(state="normal", text="⏹️ Dừng lại (Stop)", fg_color="#F44336", hover_color="#D32F2F")
        self.progress_bar.set(0)
        self.lbl_progress.configure(text=f"Khởi động {num_workers} workers...")
        self.lbl_percentage.configure(text="0%")
        
        # Dọn dẹp và nạp Queue sản phẩm mới
        while not self.product_queue.empty():
            try:
                self.product_queue.get_nowait()
            except Exception:
                break
                
        for item in shopee_items:
            self.product_queue.put(item)
            
        # Chạy tiến trình cào và đồng bộ đa luồng song song
        def scraper_thread():
            crawled_data = []      # Chứa dữ liệu của lô hiện tại (được xóa đi sau mỗi lần push 100 dòng)
            all_success_data = []  # Chứa toàn bộ dữ liệu cào thành công để dự phòng xuất Excel
            service = None
            existing_ids = set()
            write_lock = threading.Lock()
            sheets_lock = threading.Lock() # Lock riêng để tránh ghi đè đồng thời lên Sheets API
            pushed_items_count = 0
            total_items = len(shopee_items)
            
            try:
                # Bước 1: Kết nối Google Sheets API và chuẩn bị Cache loại trùng
                self.log("Đang khởi tạo kết nối Google Sheets API...")
                self.lbl_progress.configure(text="Đang kết nối Google Sheets...")
                
                creds = scraper.authenticate_google(client_secret, self.token_default)
                service = build('sheets', 'v4', credentials=creds)
                
                self.log("Đang đồng bộ Cache các ID sản phẩm đã có trên Google Sheets...")
                self.lbl_progress.configure(text="Đang tải cache từ Sheets...")
                
                # Check / tạo Tab và lấy danh sách ID đã có trong cột A
                scraper.ensure_sheet_and_headers(service, spreadsheet_id, sheet_tab)
                existing_ids = scraper.get_sheet_existing_ids(service, spreadsheet_id, sheet_tab)
                
                self.log(f"Đã tải thành công cache. Google Sheet đang chứa {len(existing_ids)} sản phẩm.", "SUCCESS")
                self.log(f"Bắt đầu khởi chạy {num_workers} Workers song song...", "INFO")
                
                # Hàm cập nhật tiến trình UI an toàn giữa các luồng
                def update_progress_ui(count: int, total: int):
                    percentage = count / total
                    self.progress_bar.set(percentage)
                    self.lbl_percentage.configure(text=f"{int(percentage * 100)}%")
                    self.lbl_progress.configure(text=f"Đang xử lý {count}/{total} sản phẩm...")
                
                # Hàm thực thi chính của từng Worker
                def worker(worker_id: int, proxy: Optional[str]):
                    nonlocal pushed_items_count
                    self.log(f"Worker #{worker_id} đã kích hoạt với proxy: {proxy if proxy else 'Không dùng proxy'}", "INFO")
                    
                    while self.is_running:
                        try:
                            # Lấy việc từ Queue (chờ tối đa 1s, nếu rỗng thì dừng luồng)
                            item_info = self.product_queue.get(block=True, timeout=1.0)
                        except queue.Empty:
                            break  # Queue đã được xử lý xong
                            
                        raw_str, item_id = item_info
                        
                        # Kiểm tra trùng lặp (Thread-safe)
                        with write_lock:
                            is_duplicate = item_id in existing_ids
                            
                        if is_duplicate:
                            self.log(f"[Worker #{worker_id}] Bỏ qua SP {item_id}: Đã tồn tại trong Google Sheet.", "WARN")
                            with write_lock:
                                pushed_items_count += 1
                                update_progress_ui(pushed_items_count, total_items)
                            self.product_queue.task_done()
                            continue
                            
                        # Gọi API lấy dữ liệu sản phẩm qua Proxy chỉ định cho worker này
                        self.log(f"[Worker #{worker_id}] Đang tải dữ liệu sản phẩm {item_id}...")
                        product = scraper.fetch_product_data(item_id, proxy)
                        
                        # Kiểm tra điều kiện lọc dữ liệu (Lượt bán & Hoa hồng người bán)
                        if product:
                            actual_sales = product.get("Lượt bán", 0)
                            actual_comm = product.get("Hoa hồng người bán (đ)", 0)
                            if actual_sales < sold_min or actual_comm < comm_min:
                                self.log(f"[Worker #{worker_id}] Bỏ qua SP {item_id}: Không đạt bộ lọc (Bán: {actual_sales} < {sold_min} hoặc HH: {actual_comm} < {comm_min})", "WARN")
                                with write_lock:
                                    pushed_items_count += 1
                                    update_progress_ui(pushed_items_count, total_items)
                                self.product_queue.task_done()
                                continue
                        
                        # Ghi nhận kết quả (Thread-safe)
                        with write_lock:
                            pushed_items_count += 1
                            if product:
                                crawled_data.append(product)
                                all_success_data.append(product)
                                existing_ids.add(item_id)
                                self.log(f"[Worker #{worker_id}] Thành công: \"{product['Tên sản phẩm'][:35]}...\"", "SUCCESS")
                                
                                # CƠ CHẾ BATCH PUSH ĐỊNH KỲ:
                                # Tự động đẩy dữ liệu lên Sheets mỗi khi thu thập đủ 100 sản phẩm mới
                                if len(crawled_data) >= 100:
                                    temp_push = list(crawled_data)
                                    crawled_data.clear()
                                    
                                    # Tạo thread phụ để ghi Sheets ngầm, không block tiến trình cào của các worker khác
                                    def push_async(data_to_push):
                                        with sheets_lock:
                                            try:
                                                self.log(f"Đang ghi định kỳ {len(data_to_push)} sản phẩm lên Google Sheets...", "INFO")
                                                scraper.push_to_google_sheet(service, spreadsheet_id, sheet_tab, data_to_push)
                                                self.log(f"Đã ghi định kỳ thành công {len(data_to_push)} sản phẩm.", "SUCCESS")
                                            except Exception as ex:
                                                self.log(f"Lỗi khi ghi định kỳ lên Google Sheets: {str(ex)}", "ERROR")
                                                # Trả lại dữ liệu bị lỗi vào crawled_data để ghi bù vào cuối
                                                with write_lock:
                                                    crawled_data.extend(data_to_push)
                                                    
                                    threading.Thread(target=push_async, args=(temp_push,), daemon=True).start()
                            else:
                                self.log(f"[Worker #{worker_id}] Thất bại khi cào sản phẩm {item_id}", "ERROR")
                                
                            update_progress_ui(pushed_items_count, total_items)
                            
                        self.product_queue.task_done()
                        
                        # Giữ khoảng nghỉ tối thiểu 1.0 giây để tránh Rate Limit trên Proxy này
                        if self.is_running:
                            time.sleep(1.0)
                            
                    self.log(f"Worker #{worker_id} đã kết thúc.", "INFO")
                
                # Khởi tạo và khởi chạy các luồng Worker phụ
                workers = []
                for i, p in enumerate(proxies_list):
                    t = threading.Thread(target=worker, args=(i + 1, p), name=f"Worker-{i+1}", daemon=True)
                    workers.append(t)
                    t.start()
                    
                # Chờ cho đến khi tất cả các luồng Worker hoàn thành nhiệm vụ
                for t in workers:
                    t.join()
                    
                # Bước 3: Đẩy nốt phần dữ liệu còn dư trong crawled_data lên Google Sheets
                with sheets_lock:
                    if crawled_data and self.is_running:
                        self.log(f"Đang ghi {len(crawled_data)} sản phẩm cuối cùng lên Google Sheets...")
                        self.lbl_progress.configure(text="Đang push dữ liệu lên Google Sheets...")
                        
                        pushed_count = scraper.push_to_google_sheet(service, spreadsheet_id, sheet_tab, crawled_data)
                        self.log(f"ĐÃ ĐỒNG BỘ THÀNH CÔNG {pushed_count} sản phẩm cuối cùng lên Google Sheet '{sheet_tab}'!", "SUCCESS")
                        crawled_data.clear()
                    elif not all_success_data:
                        self.log("Không có sản phẩm mới nào cần đồng bộ lên Sheets (Tất cả sản phẩm đã bị bỏ qua hoặc lỗi).", "WARN")
                    else:
                        self.log("Tất cả dữ liệu sản phẩm đã được đồng bộ lên Google Sheets trước đó.", "SUCCESS")
                        
            except Exception as e:
                self.log(f"Đã xảy ra sự cố trong quá trình chạy: {str(e)}", "ERROR")
                
                # Kích hoạt Fallback lưu file Excel dự phòng khi gặp lỗi
                # Sử dụng all_success_data chứa toàn bộ sản phẩm đã cào thành công từ đầu
                if all_success_data:
                    self.log("Kích hoạt chế độ dự phòng (Fallback): Tự động xuất file Excel...", "WARN")
                    self.lbl_progress.configure(text="Đang xuất file Excel fallback...")
                    
                    now = time.strftime("%Y_%m_%d_%Hh%Mm")
                    fallback_file = f"product_fallback_{now}.xlsx"
                    try:
                        scraper.export_to_excel(all_success_data, fallback_file)
                        self.log(f"Đã lưu thành công file Excel dự phòng: {fallback_file}", "SUCCESS")
                        self.after(0, lambda: messagebox.showwarning(
                            "Sự cố Sheets API", 
                            f"Lỗi kết nối Google Sheets! Tool đã tự động tải dữ liệu cào được về file Excel dự phòng:\n{fallback_file}"
                        ))
                    except Exception as ex:
                        self.log(f"Không thể lưu file Excel fallback: {str(ex)}", "ERROR")
                        err_msg = str(ex)
                        self.after(0, lambda: messagebox.showerror("Lỗi nghiêm trọng", f"Không thể lưu cả Excel dự phòng: {err_msg}"))
            finally:
                # Đưa tiến trình về 100% khi kết thúc
                self.progress_bar.set(1.0)
                self.lbl_percentage.configure(text="100%")
                self.lbl_progress.configure(text="Hoàn thành!")
                self.is_running = False
                
                self.after(0, lambda: self.btn_start.configure(state="normal", text="▶️ Bắt đầu (Start)"))
                self.after(0, lambda: self.btn_stop.configure(state="disabled", text="⏹️ Dừng lại (Stop)", fg_color="#757575", hover_color="#616161"))
                self.log("Hệ thống đã kết thúc ca chạy.")
                
        threading.Thread(target=scraper_thread, daemon=True).start()

if __name__ == "__main__":
    app = ShopeeScraperApp()
    app.mainloop()
