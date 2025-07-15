import json
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
from selenium.common.exceptions import TimeoutException, ElementNotInteractableException
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime

# Các thư viện cần thiết cho AI
try:
    from langchain_google_genai import ChatGoogleGenerativeAI
    from langchain_core.messages import HumanMessage

    LANGCHAIN_AVAILABLE = True
except ImportError:
    LANGCHAIN_AVAILABLE = False


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
                    "text": "Read the characters in this image. Return only the characters as a single string, with no other explanation or formatting.",
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
        cleaned_text = re.sub(r"[^A-Za-z0-9]", "", captcha_text)
        return cleaned_text, None
    except Exception as e:
        return None, str(e)


# --- TIẾN TRÌNH ĐIỀN FORM (WORKER) - PHIÊN BẢN ĐÃ VIẾT LẠI ---
# --- TIẾN TRÌNH ĐIỀN FORM (WORKER) - PHIÊN BẢN ĐÃ VIẾT LẠI ---
def fill_and_submit_process(task_info):
    url = task_info["url"]
    chrome_options = task_info["options"]
    data = task_info["data"]
    process_id = task_info["process_id"]
    is_headless = task_info["is_headless"]
    use_ai_captcha = task_info["use_ai_captcha"]
    api_keys = task_info["api_keys"]
    FORM_FIELD_IDS = task_info["FORM_FIELD_IDS"]
    keep_failed_tab = task_info["keep_failed_tab"]
    driver = None
    success = False
    name_for_log = data.get("full_name", f"Task {process_id}")

    try:
        service = webdriver.chrome.service.Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        driver.get(url)
        wait = WebDriverWait(driver, 5)  # Tăng thời gian chờ lên 5 giây cho ổn định

        # --- VÒNG LẶP ĐIỀN FORM ĐỘNG ---
        field_order = ["sales_date", "session"] + [
            k
            for k in FORM_FIELD_IDS.keys()
            if k not in ["sales_date", "session", "submit_button"]
        ]

        for key in field_order:
            element_id = FORM_FIELD_IDS.get(key)
            if not element_id:
                continue

            try:
                element = wait.until(
                    EC.presence_of_element_located((By.ID, element_id))
                )
                tag = element.tag_name.lower()
                value = data.get(key)

                if tag == "select":
                    if not value:
                        continue

                    select = Select(element)

                    # <<< THAY ĐỔI QUAN TRỌNG BẮT ĐẦU TỪ ĐÂY >>>
                    # Nếu là dropdown 'session', tìm kiếm thông minh hơn
                    if key == "session":
                        print(
                            f"[{process_id}] Finding smart match for '{key}' with value '{value}'..."
                        )
                        found_option = False
                        for option in select.options:
                            # Kiểm tra xem option có chứa chuỗi thời gian mong muốn không
                            if str(value) in option.text:
                                print(
                                    f"[{process_id}] Found match: '{option.text}'. Selecting..."
                                )
                                select.select_by_visible_text(option.text)
                                found_option = True
                                break  # Thoát khỏi vòng lặp khi đã tìm thấy
                        if not found_option:
                            print(
                                f"[{process_id}] WARNING: Could not find any option containing '{value}' for '{key}'."
                            )
                    else:
                        # Giữ nguyên logic cũ cho các dropdown khác như 'sales_date'
                        print(
                            f"[{process_id}] Selecting '{value}' in dropdown '{key}'..."
                        )
                        select.select_by_visible_text(str(value))
                    # <<< KẾT THÚC THAY ĐỔI >>>

                    # Kích hoạt sự kiện 'change' để website nhận diện
                    driver.execute_script(
                        "arguments[0].dispatchEvent(new Event('change'));", element
                    )
                    if key == "sales_date":
                        time.sleep(1)

                elif tag == "input":
                    input_type = element.get_attribute("type").lower()
                    if input_type in ("text", "number", "email", "tel", "password"):
                        if value:
                            print(
                                f"[{process_id}] Filling text field '{key}' with '{value}'..."
                            )
                            element.clear()
                            element.send_keys(str(value))
                    elif input_type in ("checkbox", "radio"):
                        # Chỉ click nếu là checkbox "đồng ý" hoặc có giá trị trong data
                        if key == "agree_checkbox" or value:
                            print(f"[{process_id}] Clicking checkbox/radio '{key}'...")
                            driver.execute_script("arguments[0].click();", element)

                elif tag == "textarea":
                    if value:
                        print(f"[{process_id}] Filling textarea '{key}'...")
                        element.clear()
                        element.send_keys(str(value))

            except Exception as e:
                print(
                    f"[{process_id}] Error processing field '{key}' (ID: {element_id}): {e}"
                )

        # --- XỬ LÝ CAPTCHA ĐỘNG ---
        captcha_input_id = FORM_FIELD_IDS.get("captcha")
        captcha_image_selector = FORM_FIELD_IDS.get("captcha_image_selector")

        if use_ai_captcha and captcha_input_id and captcha_image_selector:
            print(f"[{process_id}] AI solving CAPTCHA for '{name_for_log}'...")
            try:
                captcha_img_element = wait.until(
                    EC.visibility_of_element_located(
                        (By.CSS_SELECTOR, captcha_image_selector)
                    )
                )
                image_bytes = captcha_img_element.screenshot_as_png

                captcha_text, last_error = None, "No API keys provided."
                for i, api_key in enumerate(api_keys):
                    print(f"[{process_id}] Attempting with key #{i+1}...")
                    text, err = solve_captcha_with_gemini(api_key, image_bytes)
                    if text:
                        captcha_text, last_error = text, None
                        break
                    last_error = err
                    print(f"[{process_id}] Key #{i+1} failed: {err}")

                if captcha_text:
                    print(f"[{process_id}] AI result: '{captcha_text}'. Filling...")
                    driver.find_element(By.ID, captcha_input_id).send_keys(captcha_text)
                else:
                    print(
                        f"[{process_id}] All API keys failed. Last error: {last_error}. Manual input required."
                    )
                    if not is_headless:
                        input(
                            f"\n---> [{process_id}] AI FAILED. Enter CAPTCHA for '{name_for_log}' and press ENTER..."
                        )

            except Exception as e:
                print(
                    f"[{process_id}] CAPTCHA image not found: {e}. Manual input required."
                )
                if not is_headless:
                    input(
                        f"\n---> [{process_id}] Enter CAPTCHA for '{name_for_log}' and press ENTER..."
                    )
        else:
            if not is_headless:
                input(
                    f"\n---> [{process_id}] Please enter CAPTCHA for '{name_for_log}' and press ENTER..."
                )

        # --- SUBMIT ĐỘNG ---
        submit_button_id = FORM_FIELD_IDS.get("submit_button")
        if submit_button_id:
            print(f"[{process_id}] Clicking submit button...")
            wait.until(EC.element_to_be_clickable((By.ID, submit_button_id))).click()

            try:
                wait.until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//*[contains(text(), 'ĐĂNG KÝ THÀNH CÔNG')]")
                    )
                )
                success = True
            except TimeoutException:
                success = False
        else:
            print(f"[{process_id}] Submit button ID not found in mapping.")

        return success, name_for_log

    except Exception as e:
        print(f"[{process_id}] CRITICAL ERROR in process for '{name_for_log}': {e}")
        return False, name_for_log

    finally:
        # <<< LOGIC QUYẾT ĐỊNH ĐÓNG TAB >>>
        if driver:
            # Chỉ đóng tab nếu: 1. Thành công, HOẶC 2. Người dùng không muốn giữ lại tab lỗi
            if success or not keep_failed_tab:
                time.sleep(1)  # Chờ 1 giây trước khi đóng
                driver.quit()
            else:
                # Nếu thất bại và người dùng muốn giữ lại tab
                print(
                    f"[{process_id}] THẤT BẠI. Giữ lại tab của '{name_for_log}' để kiểm tra."
                )
                # Để tab lại và không làm gì cả


def analyze_form_with_gemini(api_key, html_content, excel_columns):
    if not LANGCHAIN_AVAILABLE:
        return None, "Thư viện 'langchain-google-genai' chưa được cài đặt."
    if not api_key:
        return None, "API Key của Google AI bị thiếu."

    try:
        llm = ChatGoogleGenerativeAI(model="gemini-2.0-flash", google_api_key=api_key)

        prompt_text = f"""
        Bạn là một công cụ lập trình hỗ trợ điền form web. Nhiệm vụ của bạn là phân tích một đoạn mã HTML của form đăng ký và các tiêu đề cột từ một file Excel, sau đó tạo ra một đối tượng JSON với hai trường: "FORM_FIELD_IDS" và "EXCEL_COLUMN_MAPPING".

        **ĐỊNH DẠNG ĐẦU RA BẮT BUỘC PHẢI LÀ MỘT CHUỖI JSON ĐÚNG CÚ PHÁP, KHÔNG CÓ BẤT KỲ VĂN BẢN NÀO KHÁC TRƯỚC HAY SAU NÓ.**

        **Mô tả các trường:**
        1.  **FORM_FIELD_IDS**: Một đối tượng JSON, ánh xạ tên trường chung sang 'id' HTML tương ứng của nó. Dựa vào nhãn (label) và thuộc tính của các phần tử để xác định ID đúng. Nếu không tìm thấy, hãy gán giá trị rỗng hoặc `null`.
            -   `"full_name"`: Trường Tên đầy đủ.
            -   `"day"`, `"month"`, `"year"`: Các trường cho Ngày, Tháng, Năm sinh.
            -   `"phone_number"`: Số điện thoại.
            -   `"email"`: Địa chỉ Email.
            -   `"id_card"`: Số CCCD/CMND.
            -   `"sales_date"`: Trường dropdown cho Ngày bán hàng.
            -   `"session"`: Trường dropdown cho Phiên.
            -   `"agree_checkbox"`: Checkbox đồng ý điều khoản.
            -   `"captcha"`: Trường nhập mã CAPTCHA.
            -   `"captcha_image_selector"`: **(QUAN TRỌNG)** Một CSS selector để tìm thẻ `<img>` của CAPTCHA. Ví dụ: "#dvCaptcha img" hoặc ".captcha-image-class".
            -   `"submit_button"`: Nút gửi form.

        2.  **EXCEL_COLUMN_MAPPING**: Một đối tượng JSON, ánh xạ tên trường chung sang tên tiêu đề cột Excel tương ứng từ danh sách được cung cấp. Nếu không tìm thấy cột phù hợp, hãy gán giá trị `null`.

        **Đầu vào của bạn:**
        -   **HTML Form:**
        ```html
        {html_content}
        ```

        -   **Excel Columns:**
        ```json
        {excel_columns}
        ```
        """

        message = HumanMessage(content=prompt_text)
        response = llm.invoke([message])

        json_content = (
            response.content.strip().replace("```json", "").replace("```", "")
        )
        mapping_data = json.loads(json_content)

        if "FORM_FIELD_IDS" in mapping_data and "EXCEL_COLUMN_MAPPING" in mapping_data:
            return mapping_data, None
        else:
            return None, "Cấu trúc JSON trả về không đúng."

    except Exception as e:
        return None, str(e)


# --- LỚP GIAO DIỆN (GUI) ---
class AutoFillerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Công cụ điền Form tự động (v4.0 - Hoàn toàn động)")
        self.geometry("850x700")
        style = ttk.Style(self)
        style.configure(".", font=("Segoe UI", 10))
        self.session_choice_var = tk.StringVar(value="10:00 - 12:00")
        self.create_widgets()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_widgets(self):
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- Frame Cấu hình chung ---
        config_frame = ttk.LabelFrame(main_frame, text="Cấu hình chung", padding="10")
        config_frame.pack(fill=tk.X, pady=5)
        config_frame.columnconfigure(1, weight=1)

        ttk.Label(config_frame, text="URL Trang Web:").grid(
            row=0, column=0, sticky="w", padx=5, pady=3
        )
        self.url_entry = ttk.Entry(config_frame)
        self.url_entry.grid(row=0, column=1, columnspan=2, sticky="ew", padx=5, pady=3)
        self.url_entry.insert(0, "http://localhost:3000/registration")

        ttk.Label(config_frame, text="File Excel:").grid(
            row=1, column=0, sticky="w", padx=5, pady=3
        )
        self.excel_path_entry = ttk.Entry(config_frame, state="readonly")
        self.excel_path_entry.grid(row=1, column=1, sticky="ew", padx=5, pady=3)
        ttk.Button(config_frame, text="Chọn File...", command=self.browse_excel).grid(
            row=1, column=2, padx=5
        )

        # --- Frame Nội dung HTML ---
        html_frame = ttk.LabelFrame(
            main_frame, text="Nội dung HTML của Form (Tùy chọn)", padding="10"
        )
        html_frame.pack(fill=tk.X, pady=5)
        html_frame.columnconfigure(0, weight=1)
        ttk.Label(
            html_frame, text="Dán HTML vào đây (để trống nếu muốn tự động lấy từ URL):"
        ).grid(row=0, column=0, sticky="w", padx=5, pady=3)
        self.html_text = scrolledtext.ScrolledText(
            html_frame, height=5, font=("Courier New", 9)
        )
        self.html_text.grid(row=1, column=0, sticky="nsew", padx=5, pady=3)

        # --- Frame chọn Phiên ---
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
            value="13:30 - 15:30",
        ).pack(side=tk.LEFT, padx=10)

        # --- Frame Cấu hình nâng cao ---
        adv_frame = ttk.LabelFrame(main_frame, text="Cấu hình nâng cao", padding="10")
        adv_frame.pack(fill=tk.X, pady=5)
        adv_frame.columnconfigure(1, weight=1)
        self.headless_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            adv_frame,
            text="Chạy ẩn (Không thể nhập CAPTCHA thủ công)",
            variable=self.headless_var,
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=5, pady=5)
        # <<< THÊM ĐOẠN NÀY VÀO >>>
        self.keep_failed_tab_var = tk.BooleanVar(value=True)  # Mặc định bật
        ttk.Checkbutton(
            adv_frame,
            text="Giữ lại tab thất bại để kiểm tra",
            variable=self.keep_failed_tab_var,
        ).grid(row=1, column=0, columnspan=2, sticky="w", padx=5, pady=5)
        # <<< KẾT THÚC ĐOẠN THÊM >>>
        ttk.Label(adv_frame, text="Số Tab chạy cùng lúc:").grid(
            row=2, column=0, sticky="w", padx=5, pady=3
        )
        self.max_workers_var = tk.StringVar(value=str(min(os.cpu_count() or 1, 8)))
        ttk.Spinbox(
            adv_frame, from_=1, to=50, textvariable=self.max_workers_var, width=10
        ).grid(row=2, column=1, sticky="w", padx=5, pady=3)

        # --- Frame Cấu hình AI ---
        ai_frame = ttk.LabelFrame(
            main_frame, text="Cấu hình AI - Phân tích Form & Giải CAPTCHA", padding="10"
        )
        ai_frame.pack(fill=tk.X, pady=5)
        ai_frame.columnconfigure(1, weight=1)
        self.use_ai_var = tk.BooleanVar(value=True)  # Mặc định bật AI
        ttk.Checkbutton(
            ai_frame,
            text="Sử dụng AI để phân tích Form và giải CAPTCHA (Gemini)",
            variable=self.use_ai_var,
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=5, pady=5)
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

        # --- Các nút và log ---
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
        session_choice = self.session_choice_var.get()
        keep_failed_tab = self.keep_failed_tab_var.get()
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
            self.log_message("--- BẮT ĐẦU QUÁ TRÌNH ---")

            # Đọc file Excel
            df = pd.read_excel(excel_file, dtype=str)
            excel_columns = list(df.columns)
            self.log_message(
                f"Đã đọc {len(df)} dòng từ Excel. Các cột: {excel_columns}"
            )

            # Lấy nội dung HTML
            html_content_input = self.html_text.get("1.0", tk.END).strip()
            if not html_content_input:
                self.log_message("Đang tải nội dung HTML từ URL...")
                with webdriver.Chrome(
                    service=webdriver.chrome.service.Service(
                        ChromeDriverManager().install()
                    ),
                    options=get_chrome_options(headless=True),
                ) as temp_driver:
                    temp_driver.get(url)
                    html_content = temp_driver.page_source
            else:
                html_content = html_content_input
                self.log_message("Sử dụng nội dung HTML từ ô nhập liệu.")

            # Phân tích Form bằng AI
            self.log_message("Đang gửi yêu cầu đến Gemini để phân tích form...")
            # mapping_data, error_message = analyze_form_with_gemini(
            #     api_keys_list[0], html_content, excel_columns
            # )
            mapping_data, error_message = None, "No API keys succeeded."
            for i, api_key in enumerate(api_keys_list):
                self.log_message(f"Đang thử phân tích HTML với key #{i+1}...")
                mapping_data, error_message = analyze_form_with_gemini(
                    api_key, html_content, excel_columns
                )
                if mapping_data:
                    self.log_message(f"Phân tích HTML thành công với key #{i+1}.")
                    break
                self.log_message(f"Key #{i+1} thất bại: {error_message}")
            if not mapping_data:
                messagebox.showerror(
                    "Lỗi Gemini", f"Không thể phân tích form: {error_message}"
                )
                self.start_button.config(state="normal")
                return
            if error_message:
                messagebox.showerror(
                    "Lỗi Gemini", f"Không thể phân tích form: {error_message}"
                )
                self.start_button.config(state="normal")
                return

            FORM_FIELD_IDS = mapping_data["FORM_FIELD_IDS"]
            EXCEL_COLUMN_MAPPING = mapping_data["EXCEL_COLUMN_MAPPING"]
            self.log_message("Gemini đã phân tích thành công. Mapping được tạo:")
            self.log_message(
                f"  FORM_FIELD_IDS: {json.dumps(FORM_FIELD_IDS, indent=2)}"
            )
            self.log_message(
                f"  EXCEL_COLUMN_MAPPING: {json.dumps(EXCEL_COLUMN_MAPPING, indent=2)}"
            )

            # Lấy danh sách ngày bán hàng từ web để tạo task
            self.log_message("Đang lấy danh sách ngày bán hàng từ trang web...")
            with webdriver.Chrome(
                service=webdriver.chrome.service.Service(
                    ChromeDriverManager().install()
                ),
                options=get_chrome_options(headless=True),
            ) as temp_driver:
                temp_driver.get(url)
                sales_date_id = FORM_FIELD_IDS.get("sales_date")
                if not sales_date_id:
                    self.log_message(
                        "Lỗi: Không tìm thấy 'sales_date' ID trong mapping của AI."
                    )
                    self.start_button.config(state="normal")
                    return
                sales_date_element = WebDriverWait(temp_driver, 10).until(
                    EC.presence_of_element_located((By.ID, sales_date_id))
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

            # Chuẩn bị tasks
            tasks = []
            for _, row in df.iterrows():
                base_data = {}
                for form_key, excel_col in EXCEL_COLUMN_MAPPING.items():
                    if (
                        excel_col
                        and excel_col in row
                        and form_key not in ["day", "month", "year"]
                    ):
                        if form_key == "phone_number":
                            base_data[form_key] = normalize_phone(row[excel_col])
                        else:
                            base_data[form_key] = str(row[excel_col])

                dob_excel_col = EXCEL_COLUMN_MAPPING.get(
                    "day"
                ) or EXCEL_COLUMN_MAPPING.get("date_of_birth")
                if dob_excel_col and dob_excel_col in row:
                    base_data.update(format_date_parts(row[dob_excel_col]))

                for s_date in valid_sales_dates:
                    task_data = base_data.copy()
                    task_data["sales_date"] = s_date
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
                            "FORM_FIELD_IDS": FORM_FIELD_IDS,
                            "keep_failed_tab": keep_failed_tab,
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
            import traceback

            self.log_message(traceback.format_exc())
        finally:
            self.start_button.config(state="normal")

    def process_results(self, async_results, total_tasks):
        success_count = 0
        for i, res in enumerate(async_results):
            try:
                success, name = res.get(timeout=300)
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
