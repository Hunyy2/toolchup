# Sử dụng image cơ sở
FROM selenium/standalone-chrome:latest

# Thiết lập ngôn ngữ và các biến môi trường
ENV LANG C.UTF-8
ENV LC_ALL C.UTF-8
ENV DEBIAN_FRONTEND noninteractive

# Chuyển sang người dùng root để cài đặt
USER root

# Tạo thư mục cho ứng dụng
WORKDIR /app

# Sao chép các file Python cần thiết vào container
COPY auto_form_filler.py .
COPY requirements.txt .

# Cài đặt các dependencies Python và thư viện GUI
RUN apt-get update && apt-get install -y --no-install-recommends \
    python3 \
    python3-pip \
    python3-tk \
    && rm -rf /var/lib/apt/lists/*

RUN pip3 install --no-cache-dir -r requirements.txt

# Tạo alias cho Python
RUN ln -s /usr/bin/python3 /usr/bin/python

# Trở lại người dùng mặc định của image (seluser) để đảm bảo an toàn
USER seluser

# Lệnh để chạy ứng dụng
CMD ["python3", "auto_form_filler.py"]