import customtkinter as ctk
from datetime import datetime
import random
from chuc_nang.chuc_nang import update_status,check_info,save_data,save_data_json,load_data_json
import os
import json
from tkcalendar import DateEntry
import shutil
import win32api, win32con, win32gui
from PIL import Image
from tkinter import filedialog
from PIL import Image, ImageTk
from icoextract import IconExtractor
import pythoncom
import win32com.client
import subprocess
import time
import webbrowser
import pyodbc
#======== GLOBAL ==========

NAME=""
CURRENT_FRAME=""

danh_sach_mau = {
    "Đỏ": "#C0392B",
    "Xanh Lá": "#229954",
    "Xanh Dương": "#2471A3",
    "Vàng": "#F39C12",
    "Cam": "#D35400",
    "Tím": "#6C3483",
    "Hồng": "#E91E63",
    "Xám": "#626567",
    "Xanh Ngọc": "#0E6655",
    "Xanh Biển": "#2874A6",
    "Xanh Lá Mạ": "#1D8348",
    "Xanh Bầu Trời": "#5DADE2",  # Màu mới
    "Hổ Phách": "#F7DC6F",      # Màu mới
    "Xanh Lục": "#239B56",      # Màu mới
    "Tím Oải Hương": "#AF7AC5", # Màu mới
    "Đỏ Tươi": "#D98880",       # Màu mới
    "Xanh Cyan": "#48C9B0",     # Màu mới
    "Đỏ Đô": "#9B59B6",         # Màu mới
    "Vàng Kim": "#E59866",      # Màu mới
    "Xanh Lam": "#3498DB",      # Màu mới
    "Xanh Teal": "#16A085"      # Màu mới
}

#======== DEF ==========

def randomso():
    return random.randint(0,7)

def show_frame(name_frame):
    global CURRENT_FRAME
    for i in frame_list.values():
        i.pack_forget()
    CURRENT_FRAME=name_frame
    frame_list[name_frame].pack(fill="both", expand=True)


def create_frame_1(sub_frame_3):
    global NAME, label_cpu, label_ram, label_hhd, label_internet_speed
    NAME = check_info()

    frame = ctk.CTkFrame(sub_frame_3, width=799, height=499)
    frame.pack_propagate(False)

    # Title
    ctk.CTkLabel(frame, text="Main Screen", font=("Arial", 28)).pack(anchor="n", padx=0)

    # Entry + Save Button
    child_frame = ctk.CTkFrame(frame, fg_color="transparent")
    child_frame.pack(anchor="w", padx=0, pady=5)

    ctk.CTkLabel(child_frame, text="Your name:", font=("Arial", 18)).pack(side="left", padx=0)
    name_entry = ctk.CTkEntry(child_frame, width=200, placeholder_text="Edit your name")
    if NAME:
        name_entry.insert(0, NAME)
    name_entry.pack(side="left", padx=(0, 10))

    hello_label = ctk.CTkLabel(frame, text=f"Hello {NAME}!!!", font=("Arial", 20))

    def on_save():
        nonlocal hello_label, name_entry
        global NAME
        NAME = name_entry.get().strip()
        save_data(NAME, "name")
        hello_label.configure(text=f"Hello {NAME}!!!")

    ctk.CTkButton(child_frame, text="Save Information", command=on_save).pack(side="left", padx=5)

    ctk.CTkLabel(frame, text="*" * 120, font=("Arial", 18)).pack(padx=0, pady=2, fill="x")
    hello_label.pack(pady=5)

    # System information labels
    label_cpu = ctk.CTkLabel(frame, text="CPU: ...", font=("Arial", 22))
    label_cpu.pack(pady=5)

    label_ram = ctk.CTkLabel(frame, text="RAM: ...", font=("Arial", 22))
    label_ram.pack(pady=5)

    label_hhd = ctk.CTkLabel(frame, text="HDD: ...", font=("Arial", 22))
    label_hhd.pack(pady=5)

    label_internet_speed = ctk.CTkLabel(frame, text="Internet Speed: ...", font=("Arial", 22))
    label_internet_speed.pack(pady=5)

    ctk.CTkLabel(frame, text="*" * 120, font=("Arial", 18)).pack(padx=0, pady=2, fill="x")

    # Initial update call
    update_status(sub_frame_3, label_cpu, label_ram, label_hhd, label_internet_speed)

    return frame


def create_frame_2(sub_frame_3):
    global danh_sach_mau, luu_o_tkb

    # --- Child Window (Add/Edit Schedule) ---
    def child_win_calendar():
        DATA_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "tkb.json")

        def add_subject(data=None):
            def doi_mau_khung(ma_mau_chon):
                ma_mau = danh_sach_mau[ma_mau_chon]
                khung_ma_mau.configure(fg_color=ma_mau)

            # Tiếng Anh nhưng vẫn lưu đúng key tiếng Việt
            ngay_trong_tuan = {
                "Thứ Hai": "Monday",
                "Thứ Ba": "Tuesday",
                "Thứ Tư": "Wednesday",
                "Thứ Năm": "Thursday",
                "Thứ Sáu": "Friday",
                "Thứ Bảy": "Saturday",
                "Chủ Nhật": "Sunday"
            }
            tiet_hoc = [str(i) for i in range(1, 13)]

            khung_moi = ctk.CTkFrame(tkb_frame_2, width=1500, height=30)
            khung_moi.pack(padx=0, pady=5)
            khung_moi.pack_propagate(True)

            entry = ctk.CTkEntry(khung_moi, height=30, width=300, placeholder_text="Enter Subject Name")
            entry.grid(row=0, column=0, padx=0)
            
            ctk.CTkLabel(khung_moi, text="classRoom:").grid(row=0, column=1, padx=5)
            room_entry = ctk.CTkEntry(khung_moi, height=30, width=80, placeholder_text="Enter Room")
            room_entry.grid(row=0, column=2, padx=0)

            ctk.CTkLabel(khung_moi, text="Display Color:").grid(row=0, column=3, padx=5)
            khung_ma_mau = ctk.CTkOptionMenu(khung_moi, values=list(danh_sach_mau.keys()), command=doi_mau_khung)
            khung_ma_mau.grid(row=0, column=4, padx=0)

            ctk.CTkLabel(khung_moi, text="Start Date:").grid(row=0, column=5, padx=0)
            startdateentry = DateEntry(khung_moi, width=12, background='darkblue', foreground='white',
                                       borderwidth=2, year=2025, date_pattern='dd/mm/yyyy')
            startdateentry.grid(row=0, column=6, padx=0)

            ctk.CTkLabel(khung_moi, text="End Date:").grid(row=1, column=1, padx=0)
            enddateentry = DateEntry(khung_moi, width=12, background='darkblue', foreground='white',
                                     borderwidth=2, year=2025, date_pattern='dd/mm/yyyy')
            enddateentry.grid(row=1, column=2, padx=0)


            ctk.CTkLabel(khung_moi, text="Day in week:").grid(row=2, column=0, padx=0 , sticky="e")
            ngay_hoc = ctk.CTkOptionMenu(khung_moi, values=list(ngay_trong_tuan.values()))
            ngay_hoc.grid(row=2, column=1, padx=0)

            ctk.CTkLabel(khung_moi, text="Start Period:").grid(row=2, column=2, padx=0)
            tiet_bat_dau = ctk.CTkOptionMenu(khung_moi, values=tiet_hoc)
            tiet_bat_dau.grid(row=2, column=3, padx=0)

            ctk.CTkLabel(khung_moi, text="Number of Periods:").grid(row=2, column=4, padx=0)
            tiet_ket_thuc = ctk.CTkOptionMenu(khung_moi, values=tiet_hoc)
            tiet_ket_thuc.grid(row=2, column=5, padx=0)

            nut_xoa_khung = ctk.CTkButton(khung_moi, text="Delete Row", command=khung_moi.destroy)
            nut_xoa_khung.grid(row=2, column=7, padx=0)

            
            if data:
                entry.insert(0, data.get("mon_hoc", ""))
                room_entry.insert(0, data.get("phong_hoc", ""))
                khung_ma_mau.set(data.get("mau", ""))
                if "ngay_bat_dau" in data:
                    startdateentry.set_date(datetime.strptime(data["ngay_bat_dau"], "%d/%m/%Y"))
                if "ngay_ket_thuc" in data:
                    enddateentry.set_date(datetime.strptime(data["ngay_ket_thuc"], "%d/%m/%Y"))
                # chuyển Thứ Hai -> Monday khi load
                ngay_vi = data.get("ngay_hoc", "")
                ngay_en = ngay_trong_tuan.get(ngay_vi, ngay_vi)
                ngay_hoc.set(ngay_en)
                tiet_bat_dau.set(data.get("tiet_bat_dau", ""))
                tiet_ket_thuc.set(data.get("so_tiet", ""))

            # Khi lưu, ngược lại: English → Vietnamese
            def get_vietnamese_day(en_day):
                for vi, en in ngay_trong_tuan.items():
                    if en == en_day:
                        return vi
                return en_day

            khung_moi.get_data = lambda: {
                "mon_hoc": entry.get(),
                "phong_hoc": room_entry.get(),
                "mau": khung_ma_mau.get(),
                "ngay_bat_dau": startdateentry.get_date().strftime("%d/%m/%Y"),
                "ngay_ket_thuc": enddateentry.get_date().strftime("%d/%m/%Y"),
                "ngay_hoc": get_vietnamese_day(ngay_hoc.get()),
                "tiet_bat_dau": tiet_bat_dau.get(),
                "so_tiet": tiet_ket_thuc.get()
            }

        def luu_tkb(frame_cha):
            all_data = []
            for khung_moi in frame_cha.winfo_children():
                if hasattr(khung_moi, "get_data"):
                    info = khung_moi.get_data()
                    if info:
                        all_data.append(info)

            base_dir = os.path.dirname(os.path.abspath(__file__))
            data_dir = os.path.join(base_dir, "data")
            os.makedirs(data_dir, exist_ok=True)
            file_path = os.path.join(data_dir, "tkb.json")

            with open(file_path, "w", encoding="utf-8") as f:
                json.dump(all_data, f, indent=4, ensure_ascii=False)
            print(f"✅ Saved {len(all_data)} subjects to {file_path}")

        def tai_tkb(frame_cha):
            if not os.path.exists(DATA_PATH):
                return
            with open(DATA_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            for mon in data:
                add_subject(mon)

        child_window_1 = ctk.CTkToplevel(main_window)
        child_window_1.title("Schedule Setup")
        child_window_1.geometry("1200x500")
        child_window_1.resizable(False, False)

        tkb_frame = ctk.CTkFrame(child_window_1)
        tkb_frame.pack(fill="both", expand=True)

        tkb_frame_1 = ctk.CTkFrame(tkb_frame, width=1200, height=50)
        tkb_frame_1.pack(side="top", padx=0, fill="x")
        tkb_frame_2 = ctk.CTkScrollableFrame(tkb_frame, width=1200, height=450, border_width=2)
        tkb_frame_2.pack(side="top", padx=0)

        create_tkb_button = ctk.CTkButton(tkb_frame_1, text="Add Subject", command=add_subject)
        create_tkb_button.pack(side="left", padx=0, pady=5)
        save_tkb_button = ctk.CTkButton(tkb_frame_1, text="Save Schedule", command=lambda: luu_tkb(tkb_frame_2))
        save_tkb_button.pack(side="right", padx=0, pady=5)
        tai_tkb(tkb_frame_2)
    
    # --- Table Display (Main Schedule View) ---
    def load_tkb():
        global luu_o_tkb
        luu_o_tkb = {}
        ngay_trong_tuan = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        tiet_hoc = {
            "1": "7:00-7:50", "2": "7:50-8:40", "3": "9:00-9:50",
            "4": "9:50-10:40", "5": "10:40-11:30", "6": "13:00-13:50",
            "7": "13:50-14:40", "8": "15:00-15:50", "9": "15:50-16:40",
            "10": "16:40-17:30", "11": "17:40-18:30", "12": "18:30-19:20"
        }
        
        for i in range(8):
            if i == 0:
                continue
            frame_moi = ctk.CTkFrame(child_frame_2, width=91, height=20, border_width=1)
            frame_moi.grid(row=0, column=i, padx=2)

            frame_moi.grid_rowconfigure(0, weight=1)
            frame_moi.grid_columnconfigure(0, weight=1)

            ctk.CTkLabel(
                frame_moi,
                text=ngay_trong_tuan[i-1],
                anchor="center"
            ).grid(row=0, column=0, sticky="nsew")

            luu_o_tkb[(0, i)] = frame_moi
            frame_moi.grid_propagate(False)


        for i in range(13):
            if i == 0:
                continue
            frame_moi = ctk.CTkFrame(child_frame_2, width=91, height=80, border_width=1)
            frame_moi.grid(row=i, column=0, padx=2)

            frame_moi.grid_rowconfigure(0, weight=1)
            frame_moi.grid_columnconfigure(0, weight=1)

            ctk.CTkLabel(
                frame_moi,
                text=tiet_hoc.get(str(i)),
                anchor="center"
            ).grid(row=0, column=0, sticky="nsew")

            luu_o_tkb[(i, 0)] = frame_moi
            frame_moi.grid_propagate(False)


        for i in range(1, 13):
            for g in range(1, 8):
                frame_moi = ctk.CTkFrame(child_frame_2, width=91, height=80, border_width=1)
                frame_moi.grid(row=i, column=g, padx=2)
                luu_o_tkb[(i, g)] = frame_moi
                frame_moi.grid_propagate(False)
    
    def lay_mon_theo_ngay(ngay: str):
        vi_ngay_map = {
            "Monday": "Thứ Hai", "Tuesday": "Thứ Ba", "Wednesday": "Thứ Tư",
            "Thursday": "Thứ Năm", "Friday": "Thứ Sáu", "Saturday": "Thứ Bảy", "Sunday": "Chủ Nhật"
        }
        ngay_vi = vi_ngay_map.get(ngay, ngay)
        base_dir = os.path.dirname(os.path.abspath(__file__))
        file_path = os.path.join(base_dir, "data", "tkb.json")
        if not os.path.exists(file_path):
            return []
        
        with open(file_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        
        return [
            (item["mon_hoc"], item.get("phong_hoc", ""), item["tiet_bat_dau"], item["so_tiet"], item["mau"])
            for item in data if item["ngay_hoc"] == ngay_vi
        ]

    def chinh_sua_tkb(thu, tiet_bd, so_tiet, mon_hoc, phong_hoc, mau):
        global danh_sach_mau, luu_o_tkb
        ma_mau = danh_sach_mau.get(mau, "gray")
        for i in range(tiet_bd, tiet_bd + so_tiet):
            frame = luu_o_tkb.get((i, thu))
            if frame:
                frame.configure(fg_color=ma_mau)
                for widget in frame.winfo_children():
                    widget.destroy()

        last_frame = luu_o_tkb.get((tiet_bd + so_tiet - 1, thu))
        if last_frame:
            full_text = f"{mon_hoc}\n({phong_hoc})" if phong_hoc else mon_hoc
            label = ctk.CTkLabel(last_frame, text=full_text, fg_color=ma_mau, wraplength=80, justify="center")
            label.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)

    def cap_nhat_tkb():
        for i in range(1, 13):
            for g in range(1, 8):
                frame = luu_o_tkb.get((i, g))
                if frame:
                    frame.configure(fg_color="transparent")
                    for widget in frame.winfo_children():
                        widget.destroy()

        ngay_trong_tuan = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        for i, ngay in enumerate(ngay_trong_tuan, start=1):
            ds_mon = lay_mon_theo_ngay(ngay)
            for mon in ds_mon:
                mon_hoc, phong_hoc, tiet_bd, so_tiet, mau = mon
                chinh_sua_tkb(i, int(tiet_bd), int(so_tiet), mon_hoc, phong_hoc, mau)
    
    frame = ctk.CTkFrame(sub_frame_3, width=799, height=499)
    ctk.CTkLabel(frame, text="Weekly Schedule", font=("Arial", 28)).pack(side="top", padx=0)
    frame.pack_propagate(False)

    child_frame_1 = ctk.CTkFrame(frame, width=799, height=49)
    child_frame_1.pack(side="top", padx=0)
    child_frame_1.pack_propagate(False)

    thiet_lap_tkb_button = ctk.CTkButton(child_frame_1, text="Open Schedule Setup Window", command=child_win_calendar)
    thiet_lap_tkb_button.pack(side="left", padx=0, pady=5)
    lam_moi_tkb_button = ctk.CTkButton(child_frame_1, text="Refresh Schedule", command=cap_nhat_tkb)
    lam_moi_tkb_button.pack(side="right", padx=0, pady=5)

    child_frame_2 = ctk.CTkScrollableFrame(frame, width=799, height=400)
    child_frame_2.pack(side="top", padx=0)
    
    load_tkb()
    cap_nhat_tkb()

    return frame



def create_frame_3(sub_frame_3):
    def them_duong_dan():
        
        def extract_and_save_icon_icoextract(app_path: str) -> str:
            """
            Trích xuất icon từ file exe/lnk và lưu vào data/icon bằng icoextract.
            Trả về đường dẫn icon .png đã lưu.
            """
            try:
                # Nếu là shortcut .lnk -> lấy target exe thật
                if app_path.lower().endswith(".lnk"):
                    pythoncom.CoInitialize()
                    shell = win32com.client.Dispatch("WScript.Shell")
                    shortcut = shell.CreateShortCut(app_path)
                    app_path = shortcut.Targetpath
            
                    # Kiểm tra xem đường dẫn target có tồn tại không
                    if not os.path.exists(app_path):
                        raise Exception("Đường dẫn file .exe từ shortcut không tồn tại.")

                # Tạo thư mục lưu icon nếu chưa có
                icon_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "icon")
                os.makedirs(icon_dir, exist_ok=True)

                # Trích xuất icon bằng icoextract
                extractor = IconExtractor(app_path)
                icon_data = extractor.get_icon()  # Đã sửa thành get_icon()
                img = Image.open(icon_data)

                # Lưu icon vào file .png
                base_name = name_link_entry.get().strip()
                icon_path = os.path.join(icon_dir, f"{base_name}.png")
                img.save(icon_path, "PNG")
                icon_path=icon_path.replace('\\', '/')
                return icon_path

            except Exception as e:
                print(f"Lỗi khi trích xuất icon: {e}")
                return ""

        def open_choose_link():
            link=filedialog.askopenfilename(title="Chọn File Làm Đường Dẫn",filetypes=[("Tất Cả Các File","*.*"),("File Python","*.py"),("File Văn Bản","*.txt")])
            link_entry.delete(0,ctk.END)
            link_entry.insert(0,link)
            name_link_entry.delete(0,ctk.END)
            name_link_entry.insert(0,os.path.splitext(os.path.basename(link))[0])
        
        def on_save():
            duong_dan = link_entry.get().strip()
            if not duong_dan:
                return

            # ==== đọc dữ liệu cũ ====
            if os.path.exists(json_path):
                with open(json_path, "r", encoding="utf-8") as f:
                    try:
                        data = json.load(f)
                    except:
                        data = {}
            else:
                data = {}

            ten_file = name_link_entry.get().strip()
            icon_path=extract_and_save_icon_icoextract(duong_dan)
            if icon_path is None:
                icon_path = ""

            data[ten_file] = {
                "duong_dan_ung_dung": duong_dan,
                "duong_dan_icon": icon_path
            }

            try:
                with open(json_path, "w", encoding="utf-8") as f:
                    json.dump(data, f, ensure_ascii=False, indent=4)
            except IOError:
                print("Lỗi", "Không thể ghi vào file.")
                return



            cua_so_moi.destroy()

        cua_so_moi = ctk.CTkToplevel()
        cua_so_moi.title("Thêm Đường Dẫn")
        cua_so_moi.geometry("400x150")

        frame_1 = ctk.CTkFrame(cua_so_moi)
        frame_1.pack(fill="x", padx=5, pady=5)

        child_frame_1=ctk.CTkFrame(frame_1)
        child_frame_1.pack(fill="x", padx=5, pady=5)

        link_entry = ctk.CTkEntry(child_frame_1, width=250, placeholder_text="Nhập Đường Dẫn")
        link_entry.pack(side="left", padx=5)

        link_chose_button = ctk.CTkButton(child_frame_1, text="Chọn Đường Dẫn", width=120,command=open_choose_link)
        link_chose_button.pack(side="left", padx=5)


        child_frame_2=ctk.CTkFrame(frame_1)
        child_frame_2.pack(fill="x", padx=5, pady=5)

        name_link_entry=ctk.CTkEntry(child_frame_2, width=250, placeholder_text="Nhập Tên Phần Mềm")
        name_link_entry.pack(side="left",padx=5)

        
        icon_label = ctk.CTkLabel(cua_so_moi, text="")
        icon_label.pack(pady=5)

        frame_2 = ctk.CTkFrame(cua_so_moi)
        frame_2.pack(fill="x", padx=5, pady=5)
        save_link_button = ctk.CTkButton(frame_2, text="Lưu Cài Đặt", command=on_save)
        save_link_button.pack(side="right", padx=5)

        # lưu biến cần thiết để dùng trong chon_duong_dan / on_save
        
        data_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
        json_path = os.path.join(data_dir, "link.json")
        icon_dir = os.path.join(data_dir, "icon")
        os.makedirs(icon_dir, exist_ok=True)

    def mo_ung_dung_bat(link: str):
        try:
            bat_path = os.path.join(os.getcwd(), "chay_ung_dung.bat")
            bat_content = f'@echo off\nstart "" "{link}"\ndel "%~f0"\n'
            bat_name="chay_ung_dung.bat"
            with open(bat_path, "w") as f:
                f.write(bat_content)
            subprocess.Popen([bat_path], shell=True)

        except Exception as e:
            print(f"Không thể mở ứng dụng: {link}")
            print("Lỗi:", e)

    def mo_ung_dung(link):
        try:
            os.startfile(link)
        except Exception as e:
            print(f"Không thể mở ứng dụng: {link}")
            print("Lỗi:", e)

    def lam_moi_giao_dien(father_frame):
        hang=0
        cot=0
        for widget in father_frame.winfo_children():
            widget.destroy()
        data_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
        json_path = os.path.join(data_dir, "link.json")

        try:
            with open(json_path,'r',encoding='utf-8') as f:
                data=json.load(f)
        except FileNotFoundError:
            print(f"Lỗi: Không tìm thấy file {json_path}")
            return
        except json.JSONDecodeError:
            print(f"Lỗi: File {json_path} không hợp lệ.")
            return

        for ten_ung_dung,du_lieu_ung_dung in data.items():
            try:
                duong_dan_ung_dung=du_lieu_ung_dung.get("duong_dan_ung_dung")
                duong_dan_icon=du_lieu_ung_dung.get("duong_dan_icon")
        
                if not os.path.exists(duong_dan_icon):
                    print(f"Lỗi: Không tìm thấy icon cho '{ten_ung_dung}' tại {duong_dan_icon}")
                    continue

                app_icon=ctk.CTkImage(light_image=Image.open(duong_dan_icon),dark_image=Image.open(duong_dan_icon),size=(91,100))

                
                app_frame=ctk.CTkFrame(father_frame, fg_color="transparent",width=91,height=170)
                app_frame.pack(side="left",anchor="n",padx=5,pady=5)
                
                duong_dan_label=ctk.CTkLabel(app_frame,text=duong_dan_ung_dung,font=ctk.CTkFont(size=1),fg_color="transparent",bg_color="transparent")
                duong_dan_label.pack_forget()

                nut_mo_ung_dung=ctk.CTkButton(app_frame,image=app_icon,width=91,height=100,text="",command=lambda lbl=duong_dan_label: mo_ung_dung(lbl.cget("text")))
                nut_mo_ung_dung.pack(padx=0)

                label=ctk.CTkLabel(app_frame,height=30,text=ten_ung_dung,font=("Arial",18))
                label.pack(padx=0,pady=5)

                nut_xoa_kien_ket=ctk.CTkButton(app_frame,width=91,height=30,text="Xóa Liên Kết",command=app_frame.destroy)
                nut_xoa_kien_ket.pack(padx=0,pady=5)

                cot+=1
                if cot>=8:
                    cot=0
                    hang+=1
            except Exception as e:
                print(f"Lỗi khi xử lý ứng dụng '{ten_ung_dung}': {e}")

    frame = ctk.CTkFrame(sub_frame_3, width=799, height=499)
    ctk.CTkLabel(frame, text="Đường Dẫn", font=("Arial", 28)).pack(side="top", padx=0)

    child_frame_1 = ctk.CTkFrame(frame, width=799, height=46, border_width=2)
    child_frame_1.pack(side="top", fill="x", padx=0)
    child_frame_1.pack_propagate(False)
    child_frame_2 = ctk.CTkFrame(frame, width=799, height=400, border_width=1)
    child_frame_2.pack(side="top", fill="both", expand=True, padx=0)
    child_frame_2.pack_propagate(False)



    add_link_button = ctk.CTkButton(child_frame_1, text="Add Link", command=them_duong_dan,state="disabled")
    add_link_button.pack(side="right", padx=3,pady=5)

    ctk.CTkLabel(child_frame_1,text="Under Development. Not Ready",font=("Arial", 16)).pack(side="left",padx=3, pady=5)

    refresh_link_button = ctk.CTkButton(child_frame_1,text="Refresh Interface",command=lambda: lam_moi_giao_dien(child_frame_2),state="disabled")
    refresh_link_button.pack(side="left", padx=3,pady=5)


    return frame

def create_frame_4(sub_frame_3):
    frame=ctk.CTkFrame(sub_frame_3, width=799, height=499)
    frame.pack_propagate(False)

    def mo_cua_so_api():
        new_window=ctk.CTkToplevel()
        new_window.geometry("500x500")
        new_window.title("Thêm API Key")
        frame=ctk.CTkFrame(new_window)
        frame.pack(fill="x",expand="True",padx=0)
        def create_frame(father_frame):
            ai_select=["Chat GPT","Copilot"]
            frame_con=ctk.CTkFrame(father_frame,height=95,width=450)
            frame_con.pack(padx=0,pady=5)
            frame_con.pack_propagate(False)
            frame_con_1=ctk.CTkFrame(frame_con,height=40,width=450)
            frame_con_1.pack(padx=0,pady=5)
            api_entry=ctk.CTkEntry(frame_con_1,height=40,width=300,placeholder_text="Nhập API Key")
            api_entry.grid(row=0,column=0)
            api_dropdown=ctk.CTkOptionMenu(frame_con_1,height=40,width=150,values=ai_select)
            api_dropdown.grid(row=0,column=1)
            check_api_button=ctk.CTkButton(frame_con_1,height=40,width=150,text="Kiểm Tra API Key")
            check_api_button.grid(row=1,column=0,pady=5)
            delete_frame_button=ctk.CTkButton(frame_con_1,height=40,width=150,text="Xóa Tùy Chọn",command=lambda:frame_con.destroy())
            delete_frame_button.grid(row=1,column=1,pady=5)


        frame_1=ctk.CTkFrame(frame,width=500,height=80)
        frame_1.pack(side="top",padx=0)
        frame_2=ctk.CTkScrollableFrame(frame,height=340,width=500)
        frame_2.pack(side="top",padx=0,pady=5)
        frame_3=ctk.CTkFrame(frame,height=80,width=500)
        frame_3.pack(side="top",padx=0,pady=5)

        
        add_frame_button=ctk.CTkButton(frame_1,text="+",font=("arial",22),command=lambda:create_frame(frame_2))
        add_frame_button.pack(side="left",padx=0)

        save_frame_button=ctk.CTkButton(frame_1,text="📝",font=("arial",22))
        save_frame_button.pack(side="left",padx=0)

    def mo_cua_so_data_base():
        connection_data=("DRIVER={ODBC Driver 18 for SQL Server};")
        def open_link(event):
            url = "https://learn.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server?view=sql-server-ver17&utm_source=chatgpt.com" 
            webbrowser.open(url)
        def kiem_tra_driver():
            driver_name="ODBC Driver 18 for SQL Server"
            drivers_installed=pyodbc.drivers()

            if driver_name in drivers_installed:
                return True
            else:
                return False
        def win_authentication():
            if active_flag_var.get():
                entry4.delete(0,"end")
                entry4.configure(state="disabled")
                entry5.delete(0,"end")
                entry5.configure(state="disabled")
            else:
                entry4.configure(state="normal")
                entry5.configure(state="normal")


        def check_condition():
            if entry4.get()=="":
                entry4.configure(border_color="red")
            else:
                entry4.configure(border_color="gray")
            if entry5.get()=="":
                entry5.configure(border_color="red")
            else:
                entry5.configure(border_color="gray")
            if entry5.get()!="" and entry4.get()!="":
                process_save_information()

        def process_save_information():
            nonlocal key_list,data_list
            key_list=["SERVER","DATABASE_NAME"]
            if active_flag_var.get():
                key_list.append("Trusted_Connection")
            else:
                key_list.append("USERID")
                key_list.append("PASSWORD")
            key_list.append("CREATED_DATABASE")
            data_list=[]
            data_list.append(entry2.get())
            data_list.append(entry3.get())
            if active_flag_var.get():
                data_list.append(True)
            else:
                data_list.append(entry4.get())
                data_list.append(entry5.get())
            data_list.append(database_flag_var.get())

            save_data_json("Data","chatbot.json","Data",data_list,key_list)

        
        def process_load_information():
            du_lieu_trong_file_dict=load_data_json("Data","chatbot.json","Data")
            entry2.insert(0, du_lieu_trong_file_dict.get("SERVER", ""))
            entry3.insert(0, du_lieu_trong_file_dict.get("DATABASE_NAME", ""))
            if du_lieu_trong_file_dict.get("Trusted_Connection", True):
                active_flag_var.set(True)
            else:
                entry4.insert(0, du_lieu_trong_file_dict.get("USERNAME", ""))
                entry5.insert(0, du_lieu_trong_file_dict.get("PASSWORD", ""))
            if du_lieu_trong_file_dict.get("CREATED_DATABASE", False):
                database_flag_var.set(True)
            else:
                database_flag_var.set(False)

            



        new_window=ctk.CTkToplevel()
        new_window.geometry("500x500")
        new_window.title("Thêm DataBase")
        frame=ctk.CTkScrollableFrame(new_window,width=500,height=500)
        frame.pack(fill="x",expand="True",padx=0)
        label0=ctk.CTkLabel(frame,text="Open Your Microsoft SQL Server And Get DataBase Information",font=("arial",14)).pack(side="top",padx=5,pady=10)

        frame_1=ctk.CTkFrame(frame,width=500,height=80)
        frame_1.pack(side="top",padx=0)
        frame_1.pack_propagate(False)
        label1=ctk.CTkLabel(frame_1,text="DRIVER:",font=("arial",14)).pack(side="left",padx=5,pady=10)
        entry1=ctk.CTkEntry(frame_1,width=450,height=40)
        entry1.pack(side="left",padx=5,pady=10)

        if kiem_tra_driver():
            entry1.insert(0,"ODBC Driver 18 for SQL Server Found")
        else:
            entry1.insert(0,"Driver Not Found")
        entry1.configure(state="disabled")

        label_link=ctk.CTkLabel(frame,text="If not installed: Click Here")
        label_link.pack(side="top",padx=5,pady=10)
        label_link.bind("<Button-1>", open_link)

        frame_2=ctk.CTkFrame(frame,width=500,height=80)
        frame_2.pack(side="top",padx=0)
        frame_2.pack_propagate(False)
        label2=ctk.CTkLabel(frame_2,text="SERVER:",font=("arial",14)).pack(side="left",padx=5,pady=10)
        entry2=ctk.CTkEntry(frame_2,width=450,height=40)
        entry2.pack(side="left",padx=5,pady=10)

        database_flag_var = ctk.BooleanVar(value=False)
        check_box=ctk.CTkCheckBox(frame,text="Database Created?",variable=database_flag_var)
        check_box.pack(side="top",padx=5,pady=10)

        frame_3=ctk.CTkFrame(frame,width=500,height=80)
        frame_3.pack(side="top",padx=0)
        frame_3.pack_propagate(False)
        label3=ctk.CTkLabel(frame_3,text="DATABASE NAME:",font=("arial",14)).pack(side="left",padx=5,pady=10)
        entry3=ctk.CTkEntry(frame_3,width=450,height=40)
        entry3.pack(side="left",padx=5,pady=10)

        active_flag_var = ctk.BooleanVar(value=False)
        check_box=ctk.CTkCheckBox(frame,text="Windows Authentication",variable=active_flag_var,command=lambda:win_authentication())
        check_box.pack(side="top",padx=5,pady=10)

        frame_4=ctk.CTkFrame(frame,width=500,height=80)
        frame_4.pack(side="top",padx=0)
        frame_4.pack_propagate(False)
        label4=ctk.CTkLabel(frame_4,text="USERID:",font=("arial",14)).pack(side="left",padx=5,pady=10)
        entry4=ctk.CTkEntry(frame_4,width=450,height=40)
        entry4.pack(side="left",padx=5,pady=10)

        frame_5=ctk.CTkFrame(frame,width=500,height=80)
        frame_5.pack(side="top",padx=0)
        frame_5.pack_propagate(False)
        label5=ctk.CTkLabel(frame_5,text="PASSWORD:",font=("arial",14)).pack(side="left",padx=5,pady=10)
        entry5=ctk.CTkEntry(frame_5,width=450,height=40)
        entry5.pack(side="left",padx=5,pady=10)
        process_load_information()


        key_list=None
        data_list=None
        frame_6=ctk.CTkFrame(frame,width=500,height=80)
        frame_6.pack(side="top",padx=0)
        frame_6.pack_propagate(False)
        save_button=ctk.CTkButton(frame_6,text="Save DataBase Information",command=lambda:check_condition())
        save_button.pack(side="top",padx=0)

    hang=0
    cot=1

    def chat(frame_giua):
        nonlocal hang,cot
        frame_user=ctk.CTkFrame(frame_giua,width=395,height=80,border_width=2)
        frame_user.grid(row=hang,column=1,sticky="e",padx=0)
        user_label = ctk.CTkLabel(frame_user, text=ask_entry.get(), anchor="w",wraplength=380)
        frame_user.pack_propagate(False)
        user_label.pack(fill="both", expand=True, padx=5, pady=5)
        ask_entry.delete(0,"end")
        
        hang+=1
        frame_bot=ctk.CTkFrame(frame_giua,width=395,height=80,border_width=2)
        frame_bot.grid(row=hang,column=0,sticky="w",padx=0)
        bot_label = ctk.CTkLabel(frame_bot, text="w", anchor="w",wraplength=380)
        frame_bot.pack_propagate(False)
        bot_label.pack(fill="both", expand=True, padx=5, pady=5)
        hang+=1
        


    frame_tren=ctk.CTkFrame(frame,width=799,height=48,border_width=1)
    frame_tren.pack(side="top",fill="x",padx=0)
    frame_giua=ctk.CTkScrollableFrame(frame,width=799,height=400,border_width=1)
    frame_giua.pack(side="top",fill="x",padx=0)
    frame_duoi=ctk.CTkFrame(frame,width=799,height=48,border_width=1)
    frame_duoi.pack(side="top",fill="x",padx=0)


    nut_mo_thiet_lap_chatbot=ctk.CTkButton(frame_tren,text="Thiết Lập DataBase",command=mo_cua_so_data_base,state="disabled")
    nut_mo_thiet_lap_chatbot.grid(row=0,column=0,padx=0)

    ctk.CTkLabel(frame_tren,text="Under Development. Not Ready",font=("Arial", 16)).grid(row=0,column=1,padx=0)


    nut_mo_thiet_lap_api=ctk.CTkButton(frame_tren,text="Thêm API Key Của ChatBot",command=mo_cua_so_api,state="disabled")
    nut_mo_thiet_lap_api.grid(row=0,column=2,padx=0)

    ask_entry=ctk.CTkEntry(frame_duoi,width=699,height=48,placeholder_text="Nhập Câu Hỏi")
    ask_entry.grid(row=0,column=0)
    send_button=ctk.CTkButton(frame_duoi,width=99,height=48,text="📤",font=("arial",30),command=lambda:chat(frame_giua),state="disabled")
    send_button.grid(row=0,column=1,padx=1)
    return frame

def create_frame_5(sub_frame_3):
    frame=ctk.CTkFrame(sub_frame_3, width=799, height=499)
    label=ctk.CTkLabel(frame,text="test frame 5").pack(side="top",padx=0)
    frame.pack_propagate(False)
    return frame
#======== MAIN MENU ==========

main_window=ctk.CTk()
main_window.geometry("900x700")
main_window.title("Main Menu")
main_window.resizable(False,False)

#======== INSIDE MENU ==========
main_frame=ctk.CTkFrame(main_window, fg_color="gray")
main_frame.pack(fill="both",expand=True)



#======== SUBSTATION FRAME 1 ==========
day_of_week=["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
color_list = ["#E74C3C", "#27AE60", "#2980B9", "#F1C40F", "#E67E22", "#8E44AD", "#EC407A", "#16A085"]

sub_frame_1 = ctk.CTkFrame(main_frame, width=900, height=200,border_width=1,border_color="lightblue")
sub_frame_1.pack(padx=0)

now=datetime.now()
year=now.year
month=now.month
month_name = now.strftime("%B")
day=now.day
thu=day_of_week[now.weekday()]
gio=now.hour
phut=now.minute
giay=now.second

text_ngay = f"{thu} the {day} of {month_name}, {year}"
text_gio=f"{gio} : {phut} : {giay}"

ctk.CTkLabel(sub_frame_1,text="======================================================================================",font=("Arial", 25)).pack(padx=0,pady=10,fill="x")

datenow=ctk.CTkLabel(sub_frame_1,text=text_ngay,font=("Arial", 25))
datenow.pack(padx=0,pady=10,fill="x")

timenow=ctk.CTkLabel(sub_frame_1,text=text_gio,font=("Arial", 25))
timenow.pack(padx=0,pady=10,fill="x")

ctk.CTkLabel(sub_frame_1,text="======================================================================================",font=("Arial", 25)).pack(padx=0,pady=10,fill="x")

def update_tg():
    global color_list
    now=datetime.now()
    year=now.year
    month=now.month
    month_name = now.strftime("%B")
    day=now.day
    thu=day_of_week[now.weekday()]
    gio=now.hour
    phut=now.minute
    giay=now.second

    text_ngay=f"{thu}, the {day} of {month_name}, {year}"
    text_gio=f"{gio:02d} : {phut:02d} : {giay:02d}"

    mau=color_list[randomso()]
    datenow.configure(text=text_ngay,text_color=mau)

    mau=color_list[randomso()]
    timenow.configure(text=text_gio,text_color=mau)

    timenow.after(1000, update_tg)



#======== SUBSTATION FRAME 2 ==========

sub_frame_2 = ctk.CTkFrame(main_frame, width=100, height=500,border_width=1,border_color="lightblue")
sub_frame_2.pack(padx=0,side="left")
sub_frame_2.pack_propagate(False)

menu_button=ctk.CTkButton(sub_frame_2,text="🏠",height=95,font=("Arial", 25),fg_color="#2980B9",hover_color="#2471A3", command=lambda: show_frame("main_frame"))
menu_button.pack(side="top",padx=5,pady=5)

calendar_button=ctk.CTkButton(sub_frame_2,text="📅",height=95,font=("Arial", 25),fg_color="#2980B9",hover_color="#2471A3", command=lambda: show_frame("calendar_frame"))
calendar_button.pack(side="top",padx=5,pady=5)

schedule_button=ctk.CTkButton(sub_frame_2,text="🔗",height=95,font=("Arial", 25),fg_color="#2980B9",hover_color="#2471A3", command=lambda: show_frame("schedule_frame"))
schedule_button.pack(side="top",padx=5,pady=5)

chatbot_button=ctk.CTkButton(sub_frame_2,text="🤖",height=95,font=("Arial", 25),fg_color="#2980B9",hover_color="#2471A3", command=lambda: show_frame("chatbot_frame"))
chatbot_button.pack(side="top",padx=5,pady=5)

programs_button=ctk.CTkButton(sub_frame_2,text="💻",height=95,font=("Arial", 25),fg_color="#2980B9",hover_color="#2471A3", command=lambda: show_frame("programs_frame"))
programs_button.pack(side="top",padx=5,pady=5)
#======== SUBSTATION FRAME 3 ==========

sub_frame_3 = ctk.CTkFrame(main_frame, width=800, height=500,border_width=1,border_color="lightblue")
sub_frame_3.pack(padx=0,side="left")




frame_list={
    "main_frame":create_frame_1(sub_frame_3),
    "calendar_frame":create_frame_2(sub_frame_3),
    "schedule_frame":create_frame_3(sub_frame_3),
    "chatbot_frame":create_frame_4(sub_frame_3),
    "programs_frame":create_frame_5(sub_frame_3)
    }

show_frame("main_frame")



#======== RUN ==========
update_tg()
main_window.mainloop()
