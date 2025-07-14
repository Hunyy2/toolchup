import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import multiprocessing
from multiprocessing.pool import Pool
import threading
import time
import re
import os
import base64
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime

# Các thư viện cần thiết cho AI
try:
    from langchain_google_genai import ChatGoogleGenerativeAI
    from langchain_core.messages import HumanMessage

    LANGCHAIN_AVAILABLE = True
except ImportError:
    LANGCHAIN_AVAILABLE = False


# --- CẤU HÌNH CỐ ĐỊNH CHO FORM ---
FORM_FIELD_IDS = {
    "sales_date": "slNgayBanHang",
    "session": "slPhien",
    "full_name": "txtHoTen",
    "day": "txtNgaySinh_Ngay",
    "month": "txtNgaySinh_Thang",
    "year": "txtNgaySinh_Nam",
    "phone_number": "txtSoDienThoai",
    "email": "txtEmail",
    "id_card": "txtCCCD",
    "agree_checkbox": "ckbDongY",
    "captcha_image": "imgCaptcha",  # Thêm ID của ảnh CAPTCHA
    "captcha": "txtCaptcha",
    "submit_button": "btDangKyThamGia",
}

EXCEL_COLUMN_MAPPING = {
    "full_name": "full_name",
    "date_of_birth": "date_of_birth",
    "phone_number": "phone_number",
    "email": "email",
    "id_card": "id_card",
}


# --- CÁC HÀM TIỆN ÍCH ---
def get_chrome_options(headless=True):
    options = webdriver.ChromeOptions()
    if headless:
        options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--log-level=3")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    return options


def format_date_parts(date_input):
    if pd.isna(date_input):
        return {"day": "", "month": "", "year": ""}
    try:
        date_str = str(date_input).split(" ")[0]
        for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
            try:
                dt_obj = datetime.strptime(date_str, fmt)
                return {
                    "day": dt_obj.strftime("%d"),
                    "month": dt_obj.strftime("%m"),
                    "year": dt_obj.strftime("%Y"),
                }
            except ValueError:
                continue
        parts = re.split(r"[-/]", date_str)
        if len(parts) == 3:
            return {"day": parts[0], "month": parts[1], "year": parts[2]}
    except Exception:
        pass
    return {"day": "", "month": "", "year": ""}


def normalize_phone(phone_input):
    if pd.isna(phone_input):
        return ""
    phone = str(phone_input).strip().replace(".0", "")
    if phone.startswith("84"):
        return "0" + phone[2:]
    return phone


def solve_captcha_with_gemini(api_key, image_bytes):
    """Gửi ảnh CAPTCHA đến Gemini và trả về text giải được."""
    if not LANGCHAIN_AVAILABLE:
        return None, "Thư viện 'langchain-google-genai' chưa được cài đặt."
    if not api_key:
        return None, "API Key của Google AI bị thiếu."
    try:
        llm = ChatGoogleGenerativeAI(model="gemini-2.0-flash", google_api_key=api_key)
        message = HumanMessage(
            content=[
                {
                    "type": "text",
                    "text": "Đọc các ký tự trong ảnh này. Chỉ trả về các ký tự dưới dạng một chuỗi duy nhất, không có giải thích hay định dạng gì khác.",
                },
                {
                    "type": "image_url",
                    "image_url": {
                        "url": f"data:image/png;base64,{base64.b64encode(image_bytes).decode('utf-8')}"
                    },
                },
            ]
        )
        response = llm.invoke([message])
        captcha_text = response.content.strip().replace(" ", "")
        # Lọc để chỉ giữ lại chữ và số, phòng trường hợp AI trả về ký tự lạ
        cleaned_text = re.sub(r"[^A-Za-z0-9]", "", captcha_text)
        return cleaned_text, None
    except Exception as e:
        return None, str(e)


# --- TIẾN TRÌNH ĐIỀN FORM (WORKER) ---
def fill_and_submit_process(task_info):
    url = task_info["url"]
    chrome_options = task_info["options"]
    data = task_info["data"]
    process_id = task_info["process_id"]
    is_headless = task_info["is_headless"]
    use_ai_captcha = task_info["use_ai_captcha"]
    api_keys = task_info["api_keys"]

    driver = None
    success = False
    name_for_log = data.get("full_name", "Không tìm thấy tên")

    try:
        service = webdriver.chrome.service.Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        driver.get(url)

        # --- BẮT ĐẦU SỬA LOGIC CHỌN PHIÊN ---
        # 1. Chọn ngày bán hàng
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, FORM_FIELD_IDS["sales_date"]))
        )
        Select(
            driver.find_element(By.ID, FORM_FIELD_IDS["sales_date"])
        ).select_by_visible_text(data["sales_date"])

        # 2. Chờ cho các phiên được tải ra sau khi chọn ngày
        time.sleep(1)  # Chờ một chút để AJAX gọi và tải phiên

        # 3. Chờ cho đến khi dropdown phiên có nhiều hơn 1 lựa chọn (có lựa chọn thật)
        session_dropdown_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, FORM_FIELD_IDS["session"]))
        )
        WebDriverWait(driver, 5).until(
            lambda d: len(Select(session_dropdown_element).options) > 1
        )

        # 4. Chọn phiên dựa trên lựa chọn của người dùng
        session_to_select = data.get("session")
        Select(session_dropdown_element).select_by_visible_text(session_to_select)
        # --- KẾT THÚC SỬA LOGIC CHỌN PHIÊN ---

        fields_to_fill = {
            k: data.get(k, "")
            for k in [
                "full_name",
                "day",
                "month",
                "year",
                "phone_number",
                "email",
                "id_card",
            ]
        }
        for key, value in fields_to_fill.items():
            if value:
                driver.find_element(By.ID, FORM_FIELD_IDS[key]).send_keys(value)

        driver.find_element(By.ID, FORM_FIELD_IDS["agree_checkbox"]).click()

        # Logic giải CAPTCHA (không đổi)
        if use_ai_captcha:
            print(f"[{process_id}] Đang dùng AI giải CAPTCHA cho '{name_for_log}'...")
            try:
                captcha_img_element = WebDriverWait(driver, 5).until(
                    EC.visibility_of_element_located(
                        (By.CSS_SELECTOR, "#dvCaptcha img")
                    )
                )
                image_bytes = captcha_img_element.screenshot_as_png
                captcha_text = None
                last_error = "Không có API key nào được cung cấp."

                for key_index, current_key in enumerate(api_keys):
                    print(
                        f"[{process_id}] Thử giải CAPTCHA với key #{key_index + 1}..."
                    )
                    text, err = solve_captcha_with_gemini(current_key, image_bytes)
                    if text:
                        captcha_text = text
                        last_error = None
                        break
                    else:
                        last_error = err
                        print(f"[{process_id}] Key #{key_index + 1} thất bại: {err}")

                if captcha_text:
                    print(f"[{process_id}] AI giải ra: '{captcha_text}'. Đang điền...")
                    driver.find_element(By.ID, FORM_FIELD_IDS["captcha"]).send_keys(
                        captcha_text
                    )
                else:
                    print(
                        f"[{process_id}] Tất cả API key đều thất bại. Lỗi cuối cùng: {last_error}. Chuyển sang nhập thủ công."
                    )
                    if not is_headless:
                        input(
                            f"\n---> [{process_id}] AI THẤT BẠI. NHẬP CAPTCHA CHO '{name_for_log}' RỒI NHẤN ENTER..."
                        )

            except Exception as e:
                print(
                    f"[{process_id}] Không tìm thấy ảnh CAPTCHA: {e}. Chuyển sang nhập thủ công."
                )
                if not is_headless:
                    input(
                        f"\n---> [{process_id}] NHẬP CAPTCHA CHO '{name_for_log}' RỒI NHẤN ENTER..."
                    )
        else:
            if not is_headless:
                input(
                    f"\n---> [{process_id}] VUI LÒNG NHẬP CAPTCHA CHO '{name_for_log}' TRONG TRÌNH DUYỆT, SAU ĐÓ NHẤN ENTER TẠI ĐÂY..."
                )

        driver.find_element(By.ID, FORM_FIELD_IDS["submit_button"]).click()
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//*[contains(text(), 'ĐĂNG KÝ THÀNH CÔNG')]")
                )
            )
            success = True
        except TimeoutException:
            success = False
        return success, name_for_log
    except Exception as e:
        print(f"[{process_id}] Lỗi trong tiến trình của '{name_for_log}': {e}")
        return False, name_for_log
    finally:
        if driver:
            time.sleep(2)
            driver.quit()


# --- LỚP GIAO DIỆN (GUI) ---
class AutoFillerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Công cụ điền Form tự động (v3.0 - Tích hợp AI)")
        self.geometry("850x700")
        style = ttk.Style(self)
        style.configure(".", font=("Segoe UI", 10))
        self.session_choice_var = tk.StringVar(value="10:00 - 12:00")
        self.create_widgets()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_widgets(self):
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- Frame Cấu hình chung (Không đổi) ---
        config_frame = ttk.LabelFrame(main_frame, text="Cấu hình chung", padding="10")
        config_frame.pack(fill=tk.X, pady=5)
        config_frame.columnconfigure(1, weight=1)
        ttk.Label(config_frame, text="URL Trang Web:").grid(
            row=0, column=0, sticky="w", padx=5, pady=3
        )
        self.url_entry = ttk.Entry(config_frame)
        self.url_entry.grid(row=0, column=1, columnspan=2, sticky="ew", padx=5, pady=3)
        self.url_entry.insert(0, "https://popmartbanahills.xspin.live/popmart")
        ttk.Label(config_frame, text="File Excel:").grid(
            row=1, column=0, sticky="w", padx=5, pady=3
        )
        self.excel_path_entry = ttk.Entry(config_frame, state="readonly")
        self.excel_path_entry.grid(row=1, column=1, sticky="ew", padx=5, pady=3)
        ttk.Button(config_frame, text="Chọn File...", command=self.browse_excel).grid(
            row=1, column=2, padx=5
        )

        # --- BẮT ĐẦU THAY ĐỔI: Thêm Frame chọn Phiên ---
        session_frame = ttk.LabelFrame(
            main_frame, text="Tùy chọn Phiên (Session)", padding="10"
        )
        session_frame.pack(fill=tk.X, pady=5)

        ttk.Radiobutton(
            session_frame,
            text="Phiên 1 (10:00 - 12:00)",
            variable=self.session_choice_var,
            value="10:00 - 12:00",
        ).pack(side=tk.LEFT, padx=10)

        ttk.Radiobutton(
            session_frame,
            text="Phiên 2 (13:30 - 15:30)",
            variable=self.session_choice_var,
            value="13:30 - 15:30",  # Sửa lại giá trị cho đúng với thực tế của web
        ).pack(side=tk.LEFT, padx=10)
        # --- KẾT THÚC THAY ĐỔI ---

        # --- Frame Cấu hình nâng cao (Không đổi) ---
        adv_frame = ttk.LabelFrame(main_frame, text="Cấu hình nâng cao", padding="10")
        adv_frame.pack(fill=tk.X, pady=5)
        adv_frame.columnconfigure(1, weight=1)
        self.headless_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            adv_frame,
            text="Chạy ẩn (Không thể nhập CAPTCHA thủ công)",
            variable=self.headless_var,
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=5, pady=5)
        ttk.Label(adv_frame, text="Số Tab chạy cùng lúc:").grid(
            row=1, column=0, sticky="w", padx=5, pady=3
        )
        self.max_workers_var = tk.StringVar(value=str(min(os.cpu_count() or 1, 8)))
        self.max_workers_spinbox = ttk.Spinbox(
            adv_frame, from_=1, to=50, textvariable=self.max_workers_var, width=10
        )
        self.max_workers_spinbox.grid(row=1, column=1, sticky="w", padx=5, pady=3)

        # --- Frame Cấu hình AI (Sửa label) ---
        ai_frame = ttk.LabelFrame(
            main_frame, text="Cấu hình AI - Giải CAPTCHA", padding="10"
        )
        ai_frame.pack(fill=tk.X, pady=5)
        ai_frame.columnconfigure(1, weight=1)
        self.use_ai_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            ai_frame,
            text="Tự động giải CAPTCHA bằng AI (Gemini)",
            variable=self.use_ai_var,
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=5, pady=5)
        # Sửa label để rõ ràng hơn
        ttk.Label(ai_frame, text="Google AI API Key (cách nhau bằng dấu phẩy):").grid(
            row=1, column=0, sticky="w", padx=5, pady=3
        )
        self.api_key_entry = ttk.Entry(ai_frame, show="*")
        self.api_key_entry.grid(row=1, column=1, sticky="ew", padx=5, pady=3)
        if not LANGCHAIN_AVAILABLE:
            ttk.Label(
                ai_frame,
                text="Lỗi: Cần cài đặt 'langchain-google-genai'",
                foreground="red",
            ).grid(row=2, column=1, sticky="w", padx=5)

        # --- Các nút và log (Không đổi) ---
        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill=tk.X, pady=10, padx=10)
        self.start_button = ttk.Button(
            action_frame, text="Bắt đầu điền Form", command=self.start_automation
        )
        self.start_button.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        log_frame = ttk.LabelFrame(main_frame, text="Nhật ký hoạt động", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        self.log_text = scrolledtext.ScrolledText(
            log_frame, height=10, state="disabled", font=("Courier New", 9)
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def log_message(self, message):
        log_entry = f"[{datetime.now().strftime('%H:%M:%S')}] {message}\n"
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, log_entry)
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")
        self.update_idletasks()

    def browse_excel(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx;*.xls")]
        )
        if file_path:
            self.excel_path_entry.config(state="normal")
            self.excel_path_entry.delete(0, tk.END)
            self.excel_path_entry.insert(0, file_path)
            self.excel_path_entry.config(state="readonly")
            self.log_message(f"Đã chọn file Excel: {os.path.basename(file_path)}")

    def start_automation(self):
        self.start_button.config(state="disabled")
        threading.Thread(target=self.run_automation_logic, daemon=True).start()

    def run_automation_logic(self):
        url = self.url_entry.get()
        excel_file = self.excel_path_entry.get()
        is_headless = self.headless_var.get()
        use_ai = self.use_ai_var.get()
        api_keys_str = self.api_key_entry.get()
        api_keys_list = [key.strip() for key in api_keys_str.split(",") if key.strip()]
        max_workers = int(self.max_workers_var.get())
        # Lấy lựa chọn phiên từ người dùng
        session_choice = self.session_choice_var.get()

        if not url or not excel_file:
            messagebox.showerror("Lỗi", "Vui lòng nhập URL và chọn file Excel.")
            self.start_button.config(state="normal")
            return
        if use_ai and not api_keys_list:
            messagebox.showerror("Lỗi", "Vui lòng nhập API Key để dùng chức năng AI.")
            self.start_button.config(state="normal")
            return
        if use_ai and not LANGCHAIN_AVAILABLE:
            messagebox.showerror("Lỗi", "Cần cài đặt 'langchain-google-genai'.")
            self.start_button.config(state="normal")
            return

        try:
            df = pd.read_excel(excel_file, dtype=str)
            self.log_message(f"Đọc thành công {len(df)} dòng từ file Excel.")
            if "full_name" not in df.columns:
                messagebox.showerror("Lỗi Cột Excel", "Không tìm thấy cột 'full_name'.")
                self.start_button.config(state="normal")
                return

            self.log_message("Đang lấy danh sách ngày bán hàng từ trang web...")
            with webdriver.Chrome(
                service=webdriver.chrome.service.Service(
                    ChromeDriverManager().install()
                ),
                options=get_chrome_options(headless=True),
            ) as temp_driver:
                temp_driver.get(url)
                sales_date_element = WebDriverWait(temp_driver, 10).until(
                    EC.presence_of_element_located(
                        (By.ID, FORM_FIELD_IDS["sales_date"])
                    )
                )
                valid_sales_dates = [
                    opt.text
                    for opt in Select(sales_date_element).options
                    if "--" not in opt.text and opt.get_attribute("value")
                ]

            if not valid_sales_dates:
                self.log_message("Lỗi: Không tìm thấy ngày bán hàng hợp lệ trên web.")
                self.start_button.config(state="normal")
                return
            self.log_message(
                f"Các ngày bán hàng hợp lệ: {', '.join(valid_sales_dates)}"
            )
            self.log_message(f"Đã chọn điền cho phiên: {session_choice}")

            tasks = []
            for _, row in df.iterrows():
                base_data = {}
                for excel_col, form_key in EXCEL_COLUMN_MAPPING.items():
                    if excel_col in row:
                        value = row[excel_col]
                        if form_key == "date_of_birth":
                            base_data.update(format_date_parts(value))
                        elif form_key == "phone_number":
                            base_data[form_key] = normalize_phone(value)
                        else:
                            base_data[form_key] = value

                for s_date in valid_sales_dates:
                    task_data = base_data.copy()
                    task_data["sales_date"] = s_date
                    # Thêm lựa chọn phiên vào task_data
                    task_data["session"] = session_choice
                    tasks.append(
                        {
                            "url": url,
                            "options": get_chrome_options(is_headless),
                            "data": task_data,
                            "process_id": len(tasks) + 1,
                            "is_headless": is_headless,
                            "use_ai_captcha": use_ai,
                            "api_keys": api_keys_list,
                        }
                    )

            num_workers = min(max_workers, len(tasks))
            self.log_message(
                f"Bắt đầu điền {len(tasks)} form với {num_workers} tiến trình..."
            )

            if is_headless and not use_ai:
                self.log_message("Cảnh báo: Chạy ẩn nhưng không bật AI giải CAPTCHA.")

            with Pool(processes=num_workers) as pool:
                async_results = [
                    pool.apply_async(fill_and_submit_process, (task,)) for task in tasks
                ]
                self.process_results(async_results, len(tasks))

        except Exception as e:
            self.log_message(f"Lỗi nghiêm trọng: {e}")
        finally:
            self.start_button.config(state="normal")

    def process_results(self, async_results, total_tasks):
        success_count = 0
        for i, res in enumerate(async_results):
            try:
                success, name = res.get(timeout=180)
                if success:
                    success_count += 1
                self.log_message(
                    f"-> Task [{i+1}/{total_tasks}] - {'Thành công' if success else 'THẤT BẠI'} - {name}"
                )
            except Exception as e:
                self.log_message(f"-> Task [{i+1}/{total_tasks}] - Lỗi: {e}")
        self.log_message("\n----- HOÀN TẤT -----")
        self.log_message(
            f"Tổng kết: {success_count}/{total_tasks} form đã được điền thành công."
        )

    def on_closing(self):
        if messagebox.askokcancel("Thoát", "Bạn có muốn thoát chương trình?"):
            self.destroy()


if __name__ == "__main__":
    multiprocessing.freeze_support()
    app = AutoFillerApp()
    app.mainloop()
