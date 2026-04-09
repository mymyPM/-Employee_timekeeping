import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import numpy as np
import datetime
import threading
import os
import time

class ChamCongProcessor:
    def __init__(self, root):
        self.root = root
        self.root.title("Xử lý Chấm Công - ©️2025 Phan Ngọc My")
        self.root.geometry("700x450")
        self.root.resizable(False, False)
        
        # Biến lưu trữ đường dẫn file
        self.cham_cong_path = tk.StringVar()
        self.ds_ca_path = tk.StringVar()
        self.output_path = tk.StringVar()
        
        self.setup_ui()
    
    def setup_ui(self):
        # Frame chính
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # File chấm công
        ttk.Label(main_frame, text="File Chấm Công:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.cham_cong_path, width=50).grid(row=0, column=1, pady=5, padx=5)
        ttk.Button(main_frame, text="Chọn", command=self.select_cham_cong).grid(row=0, column=2, pady=5)
        
        # File DS Ca
        ttk.Label(main_frame, text="File DS Ca:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.ds_ca_path, width=50).grid(row=1, column=1, pady=5, padx=5)
        ttk.Button(main_frame, text="Chọn", command=self.select_ds_ca).grid(row=1, column=2, pady=5)
        
        # File output
        ttk.Label(main_frame, text="File Output:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.output_path, width=50).grid(row=2, column=1, pady=5, padx=5)
        ttk.Button(main_frame, text="Chọn", command=self.select_output).grid(row=2, column=2, pady=5)
        
        # Thanh tiến trình
        ttk.Label(main_frame, text="Tiến trình:").grid(row=3, column=0, sticky=tk.W, pady=10)
        self.progress = ttk.Progressbar(main_frame, orient=tk.HORIZONTAL, length=400, mode='determinate')
        self.progress.grid(row=3, column=1, pady=10, padx=5)
        self.progress_label = ttk.Label(main_frame, text="0%")
        self.progress_label.grid(row=3, column=2, pady=10)
        
        # Trạng thái
        self.status_label = ttk.Label(main_frame, text="Sẵn sàng")
        self.status_label.grid(row=4, column=0, columnspan=3, sticky=tk.W, pady=5)
        
        # Nút xử lý
        ttk.Button(main_frame, text="Xử lý", command=self.start_processing, width=20).grid(row=5, column=1, pady=20)

        # Hộp văn bản hiển thị log debug
        self.log_text = tk.Text(main_frame, height=10, width=70, wrap=tk.WORD)
        self.log_text.grid(row=6, column=0, columnspan=3, pady=10)
    
    def select_cham_cong(self):
        filename = filedialog.askopenfilename(
            title="Chọn file Chấm Công",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.cham_cong_path.set(filename)
    
    def select_ds_ca(self):
        filename = filedialog.askopenfilename(
            title="Chọn file DS Ca",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.ds_ca_path.set(filename)
    
    def select_output(self):
        filename = filedialog.asksaveasfilename(
            title="Chọn nơi lưu file kết quả",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.output_path.set(filename)
    
    def start_processing(self):
        if not self.cham_cong_path.get() or not self.ds_ca_path.get():
            messagebox.showerror("Lỗi", "Vui lòng chọn đầy đủ các file đầu vào!")
            return
        
        if not self.output_path.get():
            # Tạo tên file output mặc định dựa vào thời gian
            current_time = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            self.output_path.set(f"BaoCaoTongHopChamCong_{current_time}.xlsx")
        
        # Xử lý trong một thread riêng biệt để không đóng băng giao diện
        processing_thread = threading.Thread(target=self.process_data)
        processing_thread.daemon = True
        processing_thread.start()
    
    def update_progress(self, value, message="Đang xử lý..."):
        # Cập nhật thanh tiến trình và nhãn
        self.progress["value"] = value
        self.progress_label["text"] = f"{int(value)}%"
        self.status_label["text"] = message
        self.root.update_idletasks()  # Cập nhật giao diện

    def log(self, message):
        """Ghi log ra hộp văn bản trong giao diện"""
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)  # Tự động cuộn xuống dòng mới nhất
    
    def process_data(self):
        try:
            self.update_progress(5, "Đang đọc dữ liệu...")
            
            # Đọc các file dữ liệu
            df_cham_cong = pd.read_excel(self.cham_cong_path.get(), dtype=str)
            df_ds_ca = pd.read_excel(self.ds_ca_path.get(), dtype=str)
            
            self.update_progress(10, "Đang xử lý dữ liệu chấm công...")
            
            # Xử lý dữ liệu chấm công
            all_staff = []
            
            # Lấy tất cả bác sĩ và KTV từ file chấm công
            staff_columns = [
                ('mabs', 'TenBacSi'), 
                ('mabs_thuoc', 'TenBacSi_Thuoc'), 
                ('mabs_cls', 'TenBacSi_CLS'), 
                ('maktv', 'TenKTV')
            ]
            
            self.update_progress(15, "Đang tổng hợp danh sách nhân viên...")
            
            # Thu thập danh sách nhân viên
            for i, row in df_cham_cong.iterrows():
                for code_col, name_col in staff_columns:
                    if pd.notna(row[code_col]) and row[code_col].strip() != '':
                        all_staff.append({
                            'ma': row[code_col],
                            'ten': row[name_col],
                            'ngay': row['Ngay'],
                            'gio': row['Gio']
                        })
            
            self.update_progress(25, "Đang phân tích thời gian ca làm việc...")
            
            # Xử lý thông tin về ca làm việc
            ca_info = []
            for i, row in df_ds_ca.iterrows():
                ca_name = row['Tên ca']
                time_range = row['Thời gian'].split(' - ')
                hours = row['Số giờ']
                
                start_time = time_range[0].strip()
                end_time = time_range[1].strip()
                
                ca_info.append({
                    'ten_ca': ca_name,
                    'start': start_time,
                    'end': end_time,
                    'hours': hours,
                    'time_range': row['Thời gian']
                })
            
            self.update_progress(45, "Đang xây dựng dữ liệu báo cáo...")
            
            # Tạo báo cáo
            report_data = []
            
            # Định nghĩa các nhóm ca chồng chéo
            overlapping_shift_groups = [
                ['Ca 1', 'Ca 1.1'],   # Nhóm 1: Ca 1 (06:00-07:00) và Ca 1.1 (06:30-07:00)
                ['Ca 2', 'Ca 2.1'],   # Nhóm 2: Ca 2 (07:00-11:30) và Ca 2.1 (08:00-11:30)
                ['Ca 3', 'Ca 4']      # Nhóm 3: Ca 3 (13:30-17:00) và Ca 4 (14:30-17:00)
            ]
            
            # Tạo DataFrame tổng hợp từ danh sách nhân viên
            if all_staff:
                df_staff = pd.DataFrame(all_staff)
                
                # Loại bỏ trùng lặp để có danh sách nhân viên duy nhất
                unique_staff = df_staff[['ma', 'ten']].drop_duplicates().reset_index(drop=True)
                
                # Đối với mỗi nhân viên
                total_staff = len(unique_staff)
                
                for idx, staff in unique_staff.iterrows():
                    self.update_progress(45 + 40 * (idx + 1) / total_staff, 
                                      f"Đang xử lý dữ liệu của {staff['ten']} ({idx+1}/{total_staff})...")
                    
                    # Lọc thông tin chấm công của nhân viên
                    staff_records = df_staff[(df_staff['ma'] == staff['ma']) & (df_staff['ten'] == staff['ten'])]
                    
                    # Chuẩn bị dữ liệu chấm công theo ngày và giờ
                    staff_attendance = {}
                    for _, record in staff_records.iterrows():
                        day = record['ngay']
                        time = record['gio']
                        
                        if pd.notna(day) and pd.notna(time):
                            try:
                                day_num = int(day)
                                if day_num not in staff_attendance:
                                    staff_attendance[day_num] = []
                                staff_attendance[day_num].append(time)
                            except ValueError:
                                continue
                    
                    # Tạo một bản ghi cho mỗi ca cho mỗi nhân viên
                    for ca in ca_info:
                        record = {
                            'Mã': staff['ma'],
                            'Họ và tên': staff['ten'],
                            'Tên ca': ca['ten_ca'],
                            'Thời gian': ca['time_range'],
                            'Số giờ': ca['hours']
                        }
                        
                        # Thêm các cột ngày
                        for day in range(1, 32):
                            record[f'n{day}'] = ''
                        
                        # Đánh dấu chấm công cho từng ngày
                        for day_num, times in staff_attendance.items():
                            record[f'n{day_num}'] = self.determine_shift_mark(day_num, times, ca, ca_info, overlapping_shift_groups)
                        
                        # Thêm bản ghi vào báo cáo
                        report_data.append(record)
            
            self.update_progress(85, "Đang tạo file báo cáo...")
            
            # Tạo DataFrame báo cáo cuối cùng
            if report_data:
                df_report = pd.DataFrame(report_data)
                
                # Sắp xếp theo Mã và Tên ca
                df_report = df_report.sort_values(by=['Mã', 'Tên ca'])
                
                # Lưu file báo cáo
                df_report.to_excel(self.output_path.get(), index=False)
                
                self.update_progress(100, f"Hoàn thành! File được lưu tại: {self.output_path.get()}")
                messagebox.showinfo("Thành công", "Xử lý dữ liệu hoàn tất!")
            else:
                self.update_progress(100, "Không có dữ liệu phù hợp để tạo báo cáo!")
                messagebox.showwarning("Cảnh báo", "Không có dữ liệu phù hợp để tạo báo cáo!")
        
        except Exception as e:
            self.update_progress(0, f"Lỗi: {str(e)}")
            messagebox.showerror("Lỗi", f"Có lỗi xảy ra: {str(e)}")

    def determine_shift_mark(self, day_num, times, ca, ca_info, overlapping_shift_groups):
        """
        Xác định xem có đánh dấu X vào ca hiện tại không.
        - Nếu thuộc nhóm ca chồng chéo → xử lý theo logic cũ.
        - Nếu không thuộc nhóm → đánh dấu X vào ca có thời lượng ngắn nhất mà thời gian chấm công nằm trong đó.
        """
        fmt = '%H:%M'

        # self.log(f"Xử lý ca: {ca['ten_ca']} vào ngày {day_num}")

        # Kiểm tra xem ca hiện tại có thuộc nhóm chồng chéo nào không
        current_group = None
        for group in overlapping_shift_groups:
            if ca['ten_ca'] in group:
                current_group = group
                break

        if current_group:
            # === Xử lý nhóm ca chồng chéo ===
            if current_group in [['Ca 1', 'Ca 1.1'], ['Ca 2', 'Ca 2.1'], ['Ca 3', 'Ca 4']]:
                # self.log("Xử lý nhóm ca chồng chéo và bao trùm thời gian")
                primary_ca = next((c for c in ca_info if c['ten_ca'] == current_group[0]), None)
                secondary_ca = next((c for c in ca_info if c['ten_ca'] == current_group[1]), None)

                primary_only_times = []
                secondary_times = []

                for time in times:
                    time_dt = datetime.datetime.strptime(time, fmt)
                    primary_start_dt = datetime.datetime.strptime(primary_ca['start'], fmt)
                    secondary_start_dt = datetime.datetime.strptime(secondary_ca['start'], fmt)
                    secondary_end_dt = datetime.datetime.strptime(secondary_ca['end'], fmt)

                    if primary_start_dt < time_dt < secondary_start_dt:
                        primary_only_times.append(time)
                    elif secondary_start_dt < time_dt < secondary_end_dt:
                        secondary_times.append(time)

                if ca['ten_ca'] == current_group[0]:
                    if primary_only_times or (primary_only_times and secondary_times):
                        return 'X'
                elif ca['ten_ca'] == current_group[1]:
                    if secondary_times and not primary_only_times:
                        return 'X'
                return ''
            else:
                # self.log("Xử lý nhóm ca chồng chéo khác")
                # Nếu là nhóm chồng chéo khác (nếu có) → xử lý mặc định như sau
                times_in_group = {
                    group_ca: [
                        t for t in times if self.time_in_shift(t, next(c for c in ca_info if c['ten_ca'] == group_ca))
                    ] for group_ca in current_group
                }

                if ca['ten_ca'] == current_group[0]:
                    if times_in_group.get(ca['ten_ca']):
                        return 'X'
                else:
                    if times_in_group.get(ca['ten_ca']) and not times_in_group.get(current_group[0]):
                        return 'X'
                return ''
        
        else:
            # self.log("Xử lý nhóm ca khác")
            # === Xử lý mặc định: chọn ca có thời lượng ngắn nhất khi thời gian thuộc nhiều ca ===
            marked = False

            for time_str in times:
                try:
                    time_dt = datetime.datetime.strptime(time_str, fmt)
                    matched_cas = []

                    for c in ca_info:
                        ca_start = datetime.datetime.strptime(c['start'], fmt)
                        ca_end = datetime.datetime.strptime(c['end'], fmt)

                        if ca_end < ca_start:
                            ca_end += datetime.timedelta(days=1)
                            if time_dt < ca_start:
                                time_dt_adjusted = time_dt + datetime.timedelta(days=1)
                            else:
                                time_dt_adjusted = time_dt
                        else:
                            time_dt_adjusted = time_dt

                        if ca_start < time_dt_adjusted < ca_end:
                            duration = (ca_end - ca_start).total_seconds() / 60
                            matched_cas.append((c, duration))

                    if matched_cas:
                        matched_cas.sort(key=lambda x: x[1])  # ưu tiên ca ngắn nhất
                        best_ca = matched_cas[0][0]
                        if best_ca['ten_ca'] == ca['ten_ca']:
                            marked = True
                except Exception as e:
                    print(f"Lỗi xử lý thời gian {time_str}: {e}")

            return 'X' if marked else ''
    
    def calculate_ca_duration(self, ca):
        """Tính số phút làm việc của ca"""
        fmt = '%H:%M'
        start = datetime.datetime.strptime(ca['start'], fmt)
        end = datetime.datetime.strptime(ca['end'], fmt)
        
        # Xử lý ca qua ngày mới
        if end < start:
            end += datetime.timedelta(days=1)
            
        duration = (end - start).total_seconds() / 60
        return duration
    
    def time_in_shift(self, time_str, ca):
        """Kiểm tra xem thời gian có nằm trong ca không"""
        try:
            fmt = '%H:%M'
            time_dt = datetime.datetime.strptime(time_str, fmt)
            ca_start = datetime.datetime.strptime(ca['start'], fmt)
            ca_end = datetime.datetime.strptime(ca['end'], fmt)
            
            # Xử lý trường hợp qua ngày mới
            if ca_end < ca_start:
                ca_end += datetime.timedelta(days=1)
                if time_dt < ca_start:
                    time_dt += datetime.timedelta(days=1)
            
            # Kiểm tra xem thời gian có nằm trong ca không
            return ca_start < time_dt < ca_end
        except Exception as e:
            print(f"Lỗi khi kiểm tra thời gian trong ca: {e}")
            return False
    
    def find_best_shift_for_time(self, time_str, all_cas):
        """
        Tìm ca phù hợp nhất cho thời gian chấm công
        Chiến lược: Chọn ca ngắn nhất mà chứa thời gian chấm công
        """
        try:
            fmt = '%H:%M'
            time_dt = datetime.datetime.strptime(time_str, fmt)
            
            # Lọc tất cả các ca mà thời gian chấm công nằm trong đó
            valid_cas = []
            
            for ca in all_cas:
                ca_start = datetime.datetime.strptime(ca['start'], fmt)
                ca_end = datetime.datetime.strptime(ca['end'], fmt)
                
                # Xử lý trường hợp qua ngày mới
                if ca_end < ca_start:
                    ca_end += datetime.timedelta(days=1)
                    if time_dt < ca_start:
                        time_dt_adjusted = time_dt + datetime.timedelta(days=1)
                    else:
                        time_dt_adjusted = time_dt
                else:
                    time_dt_adjusted = time_dt

                # Debug
                # self.log(f"[DEBUG] Kiểm tra {time_str} với ca {ca['ten_ca']} ({ca['start']} – {ca['end']})")
                
                # Kiểm tra xem thời gian có nằm trong ca không
                if ca_start < time_dt_adjusted < ca_end:
                    # Debug
                    # self.log(f"[DEBUG] {time_str} nằm TRONG ca {ca['ten_ca']}")
                    # Tính khoảng thời gian của ca
                    ca_duration = (ca_end - ca_start).total_seconds() / 60
                    valid_cas.append((ca, ca_duration))
            
            # Nếu có các ca hợp lệ, chọn ca có thời lượng ngắn nhất
            if valid_cas:
                # Sắp xếp theo thời lượng ca từ ngắn đến dài
                valid_cas.sort(key=lambda x: x[1])
                return valid_cas[0][0]
            
            return None
        except Exception as e:
            print(f"Lỗi khi tìm ca phù hợp nhất: {e}")
            return None
    
    def is_time_inside(self, inner_start, inner_end, outer_start, outer_end):
        """
        Kiểm tra xem khoảng thời gian inner có nằm hoàn toàn trong outer không
        """
        # Chuyển thời gian sang định dạng datetime
        fmt = '%H:%M'
        inner_start_dt = datetime.datetime.strptime(inner_start, fmt)
        inner_end_dt = datetime.datetime.strptime(inner_end, fmt)
        outer_start_dt = datetime.datetime.strptime(outer_start, fmt)
        outer_end_dt = datetime.datetime.strptime(outer_end, fmt)
        
        # Xử lý trường hợp qua ngày mới
        if outer_end_dt < outer_start_dt:
            outer_end_dt += datetime.timedelta(days=1)
        if inner_end_dt < inner_start_dt:
            inner_end_dt += datetime.timedelta(days=1)
        
        # Kiểm tra
        return inner_start_dt >= outer_start_dt and inner_end_dt <= outer_end_dt

def main():
    root = tk.Tk()
    app = ChamCongProcessor(root)
    root.mainloop()

if __name__ == "__main__":
    main()