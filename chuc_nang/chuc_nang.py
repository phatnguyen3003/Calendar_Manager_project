import os.path
from unittest import expectedFailure
import customtkinter as ctk
import os
import json
import psutil
import time

old_sent, old_recv = 0, 0  # để tính tốc độ mạng

# ==== LOAD / SAVE JSON ====
def check_info():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    data_path = os.path.normpath(os.path.join(base_dir, "..", "data", "dulieu.json"))
    print("Path to JSON:", data_path)

    if os.path.exists(data_path):
        with open(data_path, "r", encoding="utf-8") as f:
            try:
                data = json.load(f)
            except json.JSONDecodeError:
                data = {}
    else:
        data = {}

    return data.get("name", "")


def save_data(value: str, key: str):
    base_dir = os.path.dirname(os.path.abspath(__file__))
    data_path = os.path.normpath(os.path.join(base_dir, "..", "data", "dulieu.json"))
    os.makedirs(os.path.dirname(data_path), exist_ok=True)

    try:
        if os.path.exists(data_path):
            with open(data_path, "r", encoding="utf-8") as f:
                try:
                    data = json.load(f)
                except json.JSONDecodeError:
                    data = {}
        else:
            data = {}

        data[key] = value
        with open(data_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
    except Exception as e:
        print("Lỗi khi lưu dữ liệu:", e)



# ==== HÀM CẬP NHẬT ====
def update_status(frame_parent,label_cpu,label_ram,label_hhd,label_internet_speed):
    global old_sent, old_recv

    # CPU
    hs_cpu = psutil.cpu_percent(interval=0)

    # RAM
    ram = psutil.virtual_memory()
    ram_used = ram.used / (1024 ** 3)
    ram_total = ram.total / (1024 ** 3)
    ram_percent = ram.percent

    # HDD
    disk = psutil.disk_usage("/")
    hhd_used = disk.used / (1024 ** 3)
    hhd_total = disk.total / (1024 ** 3)

    # Internet
    net = psutil.net_io_counters()
    upload_speed = (net.bytes_sent - old_sent) / 1024
    download_speed = (net.bytes_recv - old_recv) / 1024
    old_sent, old_recv = net.bytes_sent, net.bytes_recv

    # Update label
    label_cpu.configure(text=f"CPU: {hs_cpu}%")
    label_ram.configure(text=f"RAM: {ram_used:.2f}/{ram_total:.2f} GB ({ram_percent}%)")
    label_hhd.configure(text=f"HDD: {hhd_used:.2f}/{hhd_total:.2f} GB")
    label_internet_speed.configure(
        text=f"Internet: ↓ {download_speed:.2f} KB/s | ↑ {upload_speed:.2f} KB/s"
    )

    # gọi lại sau 1000ms trên widget cha
    frame_parent.after(1000, lambda: update_status(frame_parent,label_cpu,label_ram,label_hhd,label_internet_speed))


def load_data_json(foldername="Data",filename="chatbot.json",keydata="Data"):
    duong_dan_file=os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)),"..",foldername,filename))
    try:
        with open(duong_dan_file,"r",encoding="utf-8") as f:
            du_lieu=json.load(f)
        if keydata in du_lieu:
            du_lieu_trong_file=du_lieu[keydata]
            return du_lieu_trong_file
        else:
            return None
    except FileNotFoundError:
        print("File Not Found")
    except json.JSONDecodeError:
        print("Json File Is Error Or Empty")

def save_data_json(foldername="Data", filename="chatbot.json", keydata="Data", data_list=None, key_list=None):
    duong_dan_folder = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", foldername))
    duong_dan_file = os.path.join(duong_dan_folder, filename)

    # Tạo folder nếu chưa có
    os.makedirs(duong_dan_folder, exist_ok=True)

    if data_list is None:
        data_list = []
    if key_list is None:
        key_list = []

    # Đọc dữ liệu cũ (nếu có và hợp lệ)
    du_lieu = {}
    if os.path.exists(duong_dan_file):
        try:
            if os.path.getsize(duong_dan_file) > 0:  # file có nội dung
                with open(duong_dan_file, "r", encoding="utf-8") as f:
                    du_lieu = json.load(f)
        except json.JSONDecodeError:
            print("⚠️ JSON file is empty or corrupted, starting fresh...")
            du_lieu = {}

    # Đảm bảo keydata là dict
    if keydata not in du_lieu or not isinstance(du_lieu[keydata], dict):
        du_lieu[keydata] = {}

    # Ghi dữ liệu mới vào
    for k, d in zip(key_list, data_list):
        du_lieu[keydata][k] = d

    # Lưu lại file
    try:
        with open(duong_dan_file, "w", encoding="utf-8") as f:
            json.dump(du_lieu, f, ensure_ascii=False, indent=4)
    except Exception as e:
        print(f"❌ Error saving JSON: {e}")