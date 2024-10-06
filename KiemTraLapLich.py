from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog, QMainWindow, QMessageBox
import pandas as pd
from ui_main import Ui_MainWindow
import concurrent.futures
class MainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super(MainWindow, self).__init__(parent)
        self.setupUi(self)
        self.df_input_tab1 = None
        self.df_CT2 = None
        self.df_cbdl=None
        self.df_alter_subject = None
        self.df_input_tab2=None
        self.df_cbdl_tab2 = None
        self.df_date_tab2 = None
        self.df_object_tab2=None
        self.df_room_tab2=None
        self.MAX_STUDENTS_PER_SHIFT = 1400
        self.THRESHOLD = 50
        self.STUDENTS_PER_ROOM = 40
        self.MAX_SHIFTS_PER_DAY_PER_STUDENTS = 2
        self.LOW_OCCUPANCY_THRESHOLD=30
        self.SHIFT = ['7h', '9h', '13h30', '15h30']
        self.shift_pairs = [('7h', '9h'), ('13h30', '15h30')]
        self.tab1_setup()
        self.tab2_setup()
        # tab1 được hiển thị khi khởi động
        self.tabWidget.setCurrentIndex(0)

    def tab1_setup(self):
        self.fileSVthiHK.clicked.connect(self.load_file_PhanLich)
        self.file2CT.clicked.connect(self.load_file_Student_CT2)
        self.file_alter_subject.clicked.connect(self.load_file_Alter_Subject)
        self.file_cbdl.clicked.connect(self.load_file_ChuanBiDuLieu)
        self.KtrCBDL.clicked.connect(self.Show_KtrDL)
        self.XuatfileCBDL.clicked.connect(self.export_differences)


    def tab2_setup(self):
        self.file_PhanLich.clicked.connect(self.load_file_PhanLich_tab2)
        self.fileCBDL.clicked.connect(self.load_file_CBDL)
        self.file_date.clicked.connect(self.load_file_date)
        self.file_object.clicked.connect(self.load_file_object)
        self.file_room.clicked.connect(self.load_file_room)
        self.KtrPhanLich.clicked.connect(self.Show_KtrPhanLich)
        self.XuatfileKtr.clicked.connect(self.export_file_KtrDl)
        self.XuatfileLichthi.clicked.connect(self.create_summary_excel)


    # Tab 1 methods
    def load_file_PhanLich(self):
        self.load_file_and_update_label('label_SVthiHK', 'Danh sách sinh viên thi học kì')

    def load_file_Student_CT2(self):
        self.load_file_and_update_label('label_2CT', 'Danh sách sinh viên CT2')

    def load_file_Alter_Subject(self):
        self.load_file_and_update_label('label_alter_subject', 'Danh sách học phần thay thế')

    def load_file_ChuanBiDuLieu(self):
        self.load_file_and_update_label('label_cbdl', 'Danh sách tiền xử lý dữ liệu')

    def load_file_and_update_label(self, label_name, dialog_title):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(self, dialog_title, "", "Excel Files (*.xlsx *.xls);;All Files (*)", options=options)
        if file_name:
            label = getattr(self, label_name)
            label.setText(file_name)
    def remove_duplicates(self):
        self.df_input_tab1 = self.df_input_tab1.drop_duplicates()

    def student_CT2(self):
        df_CT2_unique = self.df_CT2.drop_duplicates(subset='MSV_CT1')
        msv_ct1_to_ct2 = df_CT2_unique.set_index('MSV_CT1')['MSV_CT2'].to_dict()

        if 'MSV mở rộng' not in self.df_input_tab1.columns:
            self.df_input_tab1['MSV mở rộng'] = None

        self.df_input_tab1['MSV mở rộng'] = self.df_input_tab1['MSV'].map(msv_ct1_to_ct2).fillna(self.df_input_tab1['MSV'])

    def alter_subject(self):
        subject_mapping = self.df_alter_subject.set_index('Mã học phần')['Mã học phần thay thế'].to_dict()

        if 'Mã học phần mở rộng' not in self.df_input_tab1.columns:
            self.df_input_tab1['Mã học phần mở rộng'] = None

        self.df_input_tab1['Mã học phần mở rộng'] = self.df_input_tab1['Mã học phần'].map(subject_mapping).fillna(self.df_input_tab1['Mã học phần'])

        self.df_input_tab1['Ghi chú'] = self.df_input_tab1.apply(lambda row: f"Mã học phần cũ: {row['Mã học phần']}" if row['Mã học phần'] in subject_mapping else row['Ghi chú'], axis=1)

    def remove_foreign_language_exempted_students(self):
        self.df_input_tab1 = self.df_input_tab1[~self.df_input_tab1['HP miễn ngoại ngữ'].str.contains('Miễn', case=False, na=False)]

    def remove_discontinued_students(self):
        self.df_input_tab1 = self.df_input_tab1[~self.df_input_tab1['Ghi chú'].str.contains('SV tạm ngừng học|SV đã thôi học|SV rút HP', case=False, na=False)]

    def english_exam(self):
        self.df_input_tab1.loc[self.df_input_tab1['Đề thi TA'].str.contains('x', case=False, na=False), 'Mã học phần mở rộng'] += '_Anh'

    # def compare_files(self):
    #     common_columns = list(set(self.df_input_tab1.columns) & set(self.df_cbdl.columns))
    #     differences = pd.concat([self.df_input_tab1[common_columns], self.df_cbdl[common_columns]]).drop_duplicates(keep=False)

    #     if differences.empty:
    #         return "File dữ liệu tiền xử lý đã chính xác.", None
    #     else:
    #         return None, differences
    def compare_data(self, df_input, df_cbdl):
        conditions_ruthp = ['SV đã thôi học', 'SV rút HP', 'SV tạm ngừng học']
        conditions_mienTA = 'Miễn'

        # Tạo tập hợp các tuple (MSV, Mã học phần)
        df_input_ids = set(zip(df_input['MSV'], df_input['Mã học phần']))
        df_cbdl_ids = set(zip(df_cbdl['MSV'], df_cbdl['Mã học phần']))

        missing_in_cbdl = df_input_ids - df_cbdl_ids
        extra_in_cbdl = df_cbdl_ids - df_input_ids

        differences = pd.DataFrame(columns=['Thông điệp', 'MSV', 'Mã học phần', 'Tên học phần', 'Ghi chú'])

        # Xử lý sinh viên thiếu trong file CBdL
        if missing_in_cbdl:
            missing_students_info = df_input[df_input.apply(lambda x: (x['MSV'], x['Mã học phần']) in missing_in_cbdl, axis=1)]
            missing_students_info['Thông điệp'] = 'Sinh viên thiếu trong file chuẩn bị dữ liệu'
            differences = pd.concat([differences, missing_students_info[['Thông điệp', 'MSV', 'Mã học phần', 'Tên học phần', 'Ghi chú']]])

        # Xử lý sinh viên bị trùng lặp trong file CBdL
        if extra_in_cbdl:
            unnecessary_students_info = df_cbdl[df_cbdl.apply(lambda x: (x['MSV'], x['Mã học phần']) in extra_in_cbdl, axis=1)]
            unnecessary_students_info['Thông điệp'] = 'Sinh viên bị trùng lặp trong file chuẩn bị dữ liệu'
            differences = pd.concat([differences, unnecessary_students_info[['Thông điệp', 'MSV', 'Mã học phần', 'Tên học phần', 'Ghi chú']]])

        # Xử lý sinh viên được miễn thi ngoại ngữ
        sv_mienTA = df_input[df_input['HP miễn ngoại ngữ'].str.contains(conditions_mienTA, case=False, na=False)]
        if not sv_mienTA.empty:
            sv_mienTA['Thông điệp'] = 'Sinh viên được miễn ngoại ngữ vẫn còn trong file chuẩn bị dữ liệu'
            differences = pd.concat([differences, sv_mienTA[['Thông điệp', 'MSV', 'Mã học phần', 'Tên học phần', 'Ghi chú']]])

        # Xử lý sinh viên đã thôi học, rút học phần, tạm ngừng học
        sv_ruthp = df_input[df_input['Ghi chú'].str.contains('|'.join(conditions_ruthp), case=False, na=False)]
        if not sv_ruthp.empty:
            sv_ruthp['Thông điệp'] = 'Sinh viên tạm ngừng học | SV đã thôi học | SV rút HP vẫn còn trong file chuẩn bị dữ liệu'
            differences = pd.concat([differences, sv_ruthp[['Thông điệp', 'MSV', 'Mã học phần', 'Tên học phần', 'Ghi chú']]])

        return differences
    def Show_KtrDL(self):
        self.textBrowser.clear()

        # Check if all necessary files are loaded
        if not all([self.label_SVthiHK.text(), self.label_2CT.text(), self.label_alter_subject.text(), self.label_cbdl.text()]):
            QMessageBox.warning(self, "Lỗi", "Vui lòng chọn đầy đủ các file đầu vào.")
            return

        try:
            # Read the input files based on the paths in labels
            self.df_input_tab1 = pd.read_excel(self.label_SVthiHK.text())
            self.df_CT2 = pd.read_excel(self.label_2CT.text())
            self.df_alter_subject = pd.read_excel(self.label_alter_subject.text())
            self.df_cbdl = pd.read_excel(self.label_cbdl.text())

            # Perform data processing steps
            self.remove_duplicates()
            self.student_CT2()
            self.alter_subject()
            self.remove_foreign_language_exempted_students()
            self.remove_discontinued_students()
            self.english_exam()

            # Compare data and retrieve differences
            differences_data = self.compare_data(self.df_input_tab1, self.df_cbdl)

            # Show the results in textBrowser
            if not differences_data.empty:
                self.textBrowser.append("Thông tin chi tiết về sự khác biệt trong dữ liệu:\n")
                output_data = "Thông điệp\t\t\t\tMSV\tMã học phần\n"
                for index, row in differences_data.iterrows():
                    thong_diep = row['Thông điệp']
                    msv = row['MSV']
                    ma_hoc_phan = row['Mã học phần']
                    output_data += f"{thong_diep}\t{msv}\t{ma_hoc_phan}\n"
                self.textBrowser.append(output_data)

                # Store differences data for exporting
                self.differences = differences_data

            else:
                self.textBrowser.append("Không có sự khác biệt nào được tìm thấy.")

            # Display summary information if needed
            if hasattr(self, 'differences'):
                print(self.differences.head())

        except Exception as e:
            QMessageBox.warning(self, "Lỗi", f"Lỗi xử lý dữ liệu: {str(e)}")


    def export_differences(self):
        save_path, _ = QFileDialog.getSaveFileName(None, "Lưu file khác biệt", "", "Excel Files (*.xlsx *.xls)")
        if not hasattr(self, 'differences') or self.differences is None:
            QMessageBox.warning(None, "Cảnh báo", "Không tìm thấy thông tin khác biệt. Vui lòng chạy phương thức 'Kiểm tra file CBDL'  trước khi xuất file.")
            return
        if not save_path:
            QMessageBox.warning(None, "Lỗi", "Chưa chọn nơi lưu file.")
            return
        
        
        try:
            desired_columns = [
                'Thông điệp', 'MSV', 'Mã học phần', 'Tên học phần',
                'Ghi chú'
            ]
            self.differences = self.differences[desired_columns]

            with pd.ExcelWriter(save_path, engine='xlsxwriter') as writer:
                self.differences.to_excel(writer, index=False, sheet_name='Sheet1')

                workbook = writer.book
                worksheet = writer.sheets['Sheet1']

                # Ghi tên cột vào sheet Excel
                for col_num, value in enumerate(self.differences.columns):
                    cell_format = workbook.add_format({'text_wrap': True, 'valign': 'top'})
                    worksheet.write(0, col_num, value, cell_format)
                for i, col in enumerate(desired_columns):
                    max_len = self.differences[col].astype(str).map(len).max()  # Tìm chiều dài lớn nhất của cột
                    max_len = max(max_len, len(col))  # Lấy max với chiều dài tên cột
                    worksheet.set_column(i, i, max_len)  # Đặt chiều rộng tối thiểu cho cột i
            QMessageBox.information(None, "Thông báo", "File đã được xuất thành công.")
        except Exception as e:
            QMessageBox.critical(None, "Lỗi", f"Lỗi khi xuất file: {str(e)}")
    # tab2 methods
    def load_file_PhanLich_tab2(self):
        self.load_file_and_update_label_tab2('label_SVPhanLich', 'Danh sách sinh viên đã phân lịch thi')
    def load_file_CBDL(self):
        self.load_file_and_update_label_tab2('label_cbdl_tab2', 'Danh sách tiền xử lý dữ liệu')
    def load_file_date(self):
        self.load_file_and_update_label_tab2('label_date', 'Danh sách ngày thi')
    def load_file_object(self):
        self.load_file_and_update_label_tab2('label_subject', 'Danh sách học phần')
    def load_file_room(self):
        self.load_file_and_update_label_tab2('label_room', 'Danh sách phòng thi')
    def load_file_and_update_label_tab2(self, label_name, dialog_title):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(self, dialog_title, "", "Excel Files (*.xlsx *.xls);;All Files (*)", options=options)
        if file_name:
            label = getattr(self, label_name)
            label.setText(file_name)
    def add_violation(self, violations, violation_type, result_code, message, subject_id="", subject_name="", student_id="", exam_date="", exam_time="", extra_info=""):
        violation = {
            "Loại kiểm tra": violation_type,
            "Mã kết quả": result_code,
            "Thông điệp": message,
            "Mã học phần mở rộng": subject_id,
            "Tên học phần": subject_name,
            "Mã sinh viên mở rộng": student_id,
            "Ngày thi": exam_date,
            "Giờ thi": exam_time,
            "Ghi chú": extra_info,
        }
        violations.append(violation)

    def check_student_per_shift(self, df):
        count_student = df.groupby(['MSV mở rộng', 'Ngày thi', 'Giờ thi']).size().reset_index(name='count')
        violating_students = count_student[count_student['count'] > 1]
        violations = []

        if violating_students.empty:
            self.add_violation(violations, "check_student_per_shift", 0, "Tất cả sinh viên chỉ thi một môn trong một ca thi")
        else:
            for index, row in violating_students.iterrows():
                # Lấy danh sách các môn học mà sinh viên đã thi trong cùng một ca thi
                subjects_thi = df[(df['MSV mở rộng'] == row['MSV mở rộng']) & 
                                        (df['Ngày thi'] == row['Ngày thi']) & 
                                        (df['Giờ thi'] == row['Giờ thi'])]['Mã học phần'].unique()
                extra_info = f"Mã học phần: {', '.join(subjects_thi)}"
                self.add_violation(violations, "check_student_per_shift", 3, 'Sinh viên thi 2 môn trong 1 ca', 
                                student_id=row['MSV mở rộng'], exam_date=row['Ngày thi'], exam_time=row['Giờ thi'],
                                extra_info=extra_info)

        return pd.DataFrame(violations)

    def check_subject_student_list(self, df, df_cbdl):
        master_subjects = set(df_cbdl['Mã học phần mở rộng'].unique())
        missing_info = []

        for subject in master_subjects:
            current_students = set(df[df['Mã học phần mở rộng'] == subject]['MSV mở rộng'].unique())
            master_students = set(df_cbdl[df_cbdl['Mã học phần mở rộng'] == subject]['MSV mở rộng'].unique())
            missing_students = master_students - current_students
            if missing_students:
                missing_students = [str(student) for student in missing_students]
                for student_id in missing_students:
                    subject_name = df_cbdl[df_cbdl['Mã học phần mở rộng'] == subject]['Tên học phần'].iloc[0]  # Lấy tên học phần từ df_cbdl

                    self.add_violation(missing_info, "check_subject_student_list", 3, "MSV bị thiếu", subject_id=subject, subject_name=subject_name, student_id=student_id)

        missing_subjects = master_subjects - set(df['Mã học phần mở rộng'].unique())
        if missing_subjects:
            for subject_id in missing_subjects:
                subject_name = df_cbdl[df_cbdl['Mã học phần mở rộng'] == subject_id]['Tên học phần'].iloc[0]  # Lấy tên học phần từ df_cbdl

                self.add_violation(missing_info, "check_subject_student_list", 3, "Mã học phần bị thiếu", subject_id=subject_id, subject_name=subject_name)

        if missing_info:
            return pd.DataFrame(missing_info)
        else:
            self.add_violation(missing_info, "check_subject_student_list", 0, "Tất cả các học phần và sinh viên đã đủ")
            return pd.DataFrame(missing_info)

    def check_student_in_room(self, df,df_room_tab2):
        violations = []

        for subject, subject_data in df.groupby('Mã học phần mở rộng'):
            students_of_subject = len(subject_data)
            subject_name = df[df['Mã học phần mở rộng'] == subject]['Tên học phần'].iloc[0]
            rooms_for_students = subject_data.groupby(['Giờ thi', 'Ngày thi', 'Mã phòng']).size().reset_index(name='num_students')

            num_rooms = len(rooms_for_students)
            if num_rooms>1:
                avg_students_per_room = students_of_subject / num_rooms
                if avg_students_per_room < self.LOW_OCCUPANCY_THRESHOLD:
                        room_info_list = []
                        for index, row in rooms_for_students.iterrows():
                            exam_time = row['Giờ thi']
                            exam_date = row['Ngày thi']
                            room = row['Mã phòng']
                            num_students = row['num_students']
                            max_seats = df_room_tab2[df_room_tab2['Mã phòng'] == room]['Chỗ ngồi'].iloc[0]
                            if num_students > max_seats:
                                room_info_list.append(f"Mã phòng {room} có {num_students} sinh viên, vượt quá so với quy mô của phòng {max_seats}")
                            else:
                                room_info_list.append(f"Mã phòng {room} có {num_students} sinh viên") 
                            
                        room_info = ", ".join(room_info_list)
                        self.add_violation(violations,
                                                    "check_student_in_room", 
                                                    1,
                                                    "Số lượng sinh viên của mã học phần ít nhưng được chia nhiều phòng",
                                                    subject_id=subject,
                                                    subject_name=subject_name,
                                                    exam_date=exam_date,
                                                    exam_time=exam_time,
                                                    extra_info=room_info)
                else:
                    for shift, shift_data in subject_data.groupby('Giờ thi'):
                        for room, students_of_room in shift_data.groupby('Mã phòng'):
                            num_students = len(students_of_room)
                            max_seats = df_room_tab2[df_room_tab2['Mã phòng'] == room]['Chỗ ngồi'].iloc[0]
                            if num_students>self.THRESHOLD or num_students > max_seats:
                                room_info = f"Mã phòng {room} có {num_students} sinh viên, vượt quá số chỗ ngồi {max_seats}"
                                self.add_violation(violations,
                                                "check_student_in_room", 
                                                3,
                                                "Phòng thi có số lượng sinh viên vượt quá quy mô của phòng",
                                                subject_id=subject,
                                                exam_date=students_of_room['Ngày thi'].iloc[0],  # Lấy ngày thi từ dòng đầu tiên
                                                exam_time=shift,
                                                extra_info=room_info)

        if violations:
            return pd.DataFrame(violations)
        else:
            self.add_violation(violations, 
                            "check_student_in_room", 
                            0, 
                            "Tất cả các phòng đều không có vượt quá quy mô")
            return pd.DataFrame(violations)

    def check_alter_subjects(self, df):
        alt_subject_dict = df.groupby('Mã học phần mở rộng')['Mã học phần'].apply(list).to_dict()
        notified_subjects = set()
        violations = []
        for main_subject, alt_subjects in alt_subject_dict.items():
        # Lấy dữ liệu các học phần chính và thay thế có cùng Mã học phần mở rộng, Giờ thi, và Ngày thi
            alt_sessions = df[df['Mã học phần mở rộng'].isin([main_subject] + alt_subjects)][['Mã học phần mở rộng', 'Giờ thi', 'Ngày thi']]
            
            # Tìm các học phần có số lượng sinh viên vượt quá MAX_STUDENTS_PER_SHIFT
            large_subjects = alt_sessions['Mã học phần mở rộng'].value_counts()
            large_subjects = large_subjects[large_subjects > self.MAX_STUDENTS_PER_SHIFT].index

            for subject in large_subjects:
                if subject in notified_subjects:
                    continue

                subject_df = alt_sessions[alt_sessions['Mã học phần mở rộng'] == subject]
                subject_name = df[df['Mã học phần mở rộng'] == subject]['Tên học phần'].iloc[0] 
                # Kiểm tra nếu học phần thi cùng ca
                if len(subject_df) <= self.MAX_STUDENTS_PER_SHIFT:
                    if subject_df.groupby(['Giờ thi', 'Ngày thi']).size().gt(1).any():
                        for index, row in subject_df.iterrows():
                            self.add_violation(violations, "check_alter_subjects", 2, "Các học phần thay thế không thi cùng ca", subject_id=subject, subject_name=subject_name, exam_date=row['Ngày thi'], exam_time=row['Giờ thi'])
                        notified_subjects.update(alt_subjects)
                else:
                    for exam_date, daily_sessions in subject_df.groupby('Ngày thi'):
                        sessions = daily_sessions['Giờ thi'].unique()
                        sorted_sessions = sorted(sessions, key=lambda x: self.SHIFT.index(x))
                        for i in range(len(sessions)-1):
                            if not ((sorted_sessions[i] == '7h' and sorted_sessions[i+1] == '9h') or (sorted_sessions[i] == '13h30' and sorted_sessions[i+1] == '15h30')):
                                self.add_violation(violations, "check_alter_subjects", 2, "Số lượng sinh viên vượt quá quy mô của 1 ca và không được chia vào các ca liên tiếp", subject_id=main_subject, subject_name=subject_name, exam_date=exam_date, exam_time=sessions[i])
                    notified_subjects.add(main_subject)

        if violations:
            return pd.DataFrame(violations)
        else:
            self.add_violation(violations, "check_alter_subjects", 0, "Tất cả học phần thay thế nhau thi cùng 1 ca")
            return pd.DataFrame(violations)


    def check_count_room_shift(self, df):
        reused_message_printed = False
        dates = df['Ngày thi'].unique()
        violations = []
        for date in dates:
            daily_sessions = df[df['Ngày thi'] == date]
            shifts = daily_sessions['Giờ thi'].unique()

            for shift_pair in [(self.SHIFT[0], self.SHIFT[1]), (self.SHIFT[2], self.SHIFT[3])]:
                shift1, shift2 = shift_pair
                if shift1 in shifts and shift2 in shifts:
                    shift1_rooms = daily_sessions[daily_sessions['Giờ thi'] == shift1]['Mã phòng'].unique()
                    shift2_rooms = daily_sessions[daily_sessions['Giờ thi'] == shift2]['Mã phòng'].unique()
                    shift1_count=len(shift1_rooms)
                    shift2_count=len(shift2_rooms)
                    shift_rooms_count = abs(len(shift1_rooms) - len(shift2_rooms))
                    if shift_rooms_count > 3:
                        self.add_violation(violations, "check_count_room_shift", 2, "Có sự chênh lệch phòng giữa các ca thi", exam_date=date, extra_info=f"Chênh lệch {shift_rooms_count} phòng giữa ca thi {shift1} và {shift2}. Số phòng ca {shift1}: {shift1_count}, Số phòng ca {shift2}: {shift2_count}.")
                        reused_message_printed = True

        if not violations:
            self.add_violation(violations, "check_count_room_shift", 0, "Số lượng phòng của 2 ca không có sự chênh lệch")

        return pd.DataFrame(violations)

    def check_room_reuse(self, df):
        reused_message_printed = False
        dates = df['Ngày thi'].unique()
        violations = []

        for date in dates:
            daily_sessions = df[df['Ngày thi'] == date]
            for shift_pair in [(self.SHIFT[0], self.SHIFT[1]), (self.SHIFT[2], self.SHIFT[3])]:
                current_shift = daily_sessions[daily_sessions['Giờ thi'] == shift_pair[0]]
                next_shift = daily_sessions[daily_sessions['Giờ thi'] == shift_pair[1]]

                current_shift_rooms = current_shift['Mã phòng'].unique()
                next_shift_rooms = next_shift['Mã phòng'].unique()

                reused_rooms = set(current_shift_rooms).intersection(next_shift_rooms)
                non_reused_rooms = set(next_shift_rooms) - reused_rooms

                if non_reused_rooms:
                    non_reused_rooms_list = ', '.join(map(str, non_reused_rooms))
                    room_count = len(non_reused_rooms)
                    message = f"Ca thi sau không sử dụng lại phòng của ca thi trước trong cùng buổi"
                    extra_info = f"Giờ thi {shift_pair[1]} không sử dụng lại {room_count} phòng từ giờ thi {shift_pair[0]}: {non_reused_rooms_list}"

                    # Thêm vi phạm một lần cho mỗi cặp ngày và ca thi
                    self.add_violation(violations, 
                                    "check_room_reuse", 
                                    2, 
                                    message, 
                                    exam_date=date, 
                                    exam_time=shift_pair[1], 
                                    extra_info=extra_info)
                    reused_message_printed = True

        if not violations:
            self.add_violation(violations, "check_room_reuse", 0, "Tất cả các mã phòng đều được sử dụng lại")

        return pd.DataFrame(violations)

    def check_exam_datetime(self, df, df_date):
        valid_times = df_date['Giờ thi'].tolist()
        valid_dates = pd.to_datetime(df_date['Ngày thi'], dayfirst=True).dt.date.tolist()
        violations = []

        for index, row in df.iterrows():
            exam_date = pd.to_datetime(row['Ngày thi'], dayfirst=True).date()
            exam_time = row['Giờ thi']

            if exam_date not in valid_dates:
                self.add_violation(violations, "check_exam_datetime", 2, "Ngày thi không hợp lệ", subject_id=row['Mã học phần mở rộng'], exam_date=row['Ngày thi'], exam_time=row['Giờ thi'], extra_info=f"Ngày thi {exam_date} không hợp lệ")
            if exam_time not in valid_times:
                self.add_violation(violations, "check_exam_datetime", 2, "Giờ thi không hợp lệ", subject_id=row['Mã học phần mở rộng'], exam_date=row['Ngày thi'], exam_time=row['Giờ thi'], extra_info=f"Ca thi {exam_time} không hợp lệ")

        if not violations:
            self.add_violation(violations, "check_exam_datetime", 0, "Ngày và giờ thi bình thường")

        return pd.DataFrame(violations)


    def calculate_required_rooms(self,df_cbdl):
        required_rooms = {}

        for ma_hoc_phan, group in df_cbdl.groupby('Mã học phần mở rộng'):
            total_students = len(group)

            if total_students <= self.THRESHOLD:
                # required_rooms[ma_hoc_phan] = (1, [total_students])       
                required_rooms[ma_hoc_phan] = 1       
            else:
                rooms_needed = (total_students + self.STUDENTS_PER_ROOM - 1) // self.STUDENTS_PER_ROOM
                if total_students > self.MAX_STUDENTS_PER_SHIFT and rooms_needed % 2 != 0:
                    rooms_needed += 1

                base_students_per_room = total_students // rooms_needed
                students_distribution = [base_students_per_room] * rooms_needed

                for i in range(total_students % rooms_needed):
                    students_distribution[i] += 1

                required_rooms[ma_hoc_phan] = rooms_needed
                # # Initial room calculation
                # rooms_needed = total_students// self.STUDENTS_PER_ROOM
                # if total_students > self.MAX_STUDENTS_PER_SHIFT and rooms_needed % 2 != 0:
                #     rooms_needed += 1
                # # Basic student distribution per room
                # base_students_per_room = total_students // rooms_needed
                # students_distribution = [base_students_per_room] * rooms_needed
                
                # # Calculate remaining students
                # remaining_students = total_students % rooms_needed
                
                # if remaining_students > 0:
                #     if remaining_students <= 20:
                #         for i in range(remaining_students):
                #             students_distribution[i] += 1
                #     else:
                #         rooms_needed += 1
                #         base_students_per_room = total_students // rooms_needed
                #         students_distribution = [base_students_per_room] * rooms_needed
                #         remaining_students = total_students % rooms_needed
                #         for i in range(remaining_students):
                #             students_distribution[i] += 1
                
                # # Ensure the rooms are as evenly distributed as possible
                # if rooms_needed > 1:
                #     while True:
                #         max_students = max(students_distribution)
                #         min_students = min(students_distribution)
                #         if max_students - min_students <= 5:
                #             break
                #         max_index = students_distribution.index(max_students)
                #         min_index = students_distribution.index(min_students)
                #         students_distribution[max_index] -= 1
                #         students_distribution[min_index] += 1
                # required_rooms[ma_hoc_phan] = (rooms_needed, students_distribution)


        return required_rooms


    def check_room_assignment(self, df_cbdl, df):
        required_rooms_count = self.calculate_required_rooms(df_cbdl)
        # for ma_hoc_phan, (num_rooms, students_distribution) in required_rooms_count.items():
        #     print(f"Mã học phần {ma_hoc_phan}: {num_rooms} phòng")
        #     if isinstance(students_distribution, list):
        #         for i, num_students in enumerate(students_distribution):
        #             print(f"  Phòng {i+1}: {num_students} sinh viên")
        #     else:
        #         print(f"  Phòng 1: {students_distribution} sinh viên")
        violations = []

        # Loại bỏ các bản ghi trùng lặp dựa trên 'Ca thi', 'Mã học phần mở rộng' và 'Mã phòng'
        df_unique_rooms = df.drop_duplicates(subset=['Giờ thi', 'Mã học phần mở rộng', 'Mã phòng'])

        # Nhóm các sinh viên theo 'Mã học phần mở rộng' và tính tổng số lượng phòng duy nhất cho mỗi mã học phần mở rộng
        assigned_rooms = df_unique_rooms.groupby(['Mã học phần mở rộng', 'Giờ thi'])['Mã phòng'].nunique().reset_index()
        assigned_rooms_total = assigned_rooms.groupby('Mã học phần mở rộng')['Mã phòng'].sum()

        # for subject, (required_rooms, students_distribution) in required_rooms_count.items():
        for subject, required_rooms in required_rooms_count.items():
            if subject in assigned_rooms_total:
                assigned_rooms_count = assigned_rooms_total[subject]
                subject_rooms = df_unique_rooms[df_unique_rooms['Mã học phần mở rộng'] == subject]['Mã phòng'].unique()
            else:
                assigned_rooms_count = 0
                subject_rooms = []
            subject_name = df_cbdl[df_cbdl['Mã học phần mở rộng'] == subject]['Tên học phần'].iloc[0]
            
            for index, row in assigned_rooms.iterrows():
                if row['Mã học phần mở rộng'] == subject:
                    exam_time = row['Giờ thi']
                    exam_date = df_unique_rooms[(df_unique_rooms['Mã học phần mở rộng'] == subject) & (df_unique_rooms['Giờ thi'] == exam_time)]['Ngày thi'].iloc[0]
                    if assigned_rooms_count < required_rooms:
                        extra_info = f"Có thể cần thêm {abs(required_rooms - assigned_rooms_count)} phòng"
                        self.add_violation(violations, 
                                        "check_room_assignment", 
                                        1, 
                                        "Mã học phần mở rộng cần thêm phòng thi", 
                                        subject_id=subject, 
                                        subject_name=subject_name, 
                                        exam_date=exam_date, 
                                        exam_time=exam_time)
                    elif assigned_rooms_count > required_rooms:
                        excess_rooms = subject_rooms[-(assigned_rooms_count - required_rooms):]
                        remaining_rooms = set(subject_rooms) - set(excess_rooms)
                        extra_info = f"Phòng {' ,'.join(excess_rooms)}, {' ,'.join(remaining_rooms)} có thể gộp phòng"
                        self.add_violation(violations, 
                                        "check_room_assignment", 
                                        1, 
                                        "Các phòng thi của mã học phần mở rộng có thể gộp phòng", 
                                        subject_id=subject, 
                                        subject_name=subject_name, 
                                        exam_date=exam_date, 
                                        exam_time=exam_time, 
                                        extra_info=extra_info)

        if not violations:
            self.add_violation(violations, "check_room_assignment", 0, "Đã gán đủ phòng cho tất cả các mã học phần mở rộng")

        return pd.DataFrame(violations)

    # def check_schedule_per_day(self, df_input):
    #     shift_count = df_input.groupby(['MSV mở rộng', 'Ngày thi','Mã học phần']).size().reset_index(name='số ca thi')
    #     violating_students = shift_count[shift_count['số ca thi'] > self.MAX_SHIFTS_PER_DAY_PER_STUDENTS]
    #     violation_count_per_course = violating_students.groupby(['Mã học phần', 'Ngày thi']).size().reset_index(name='Số lượng vi phạm')

    #     violations = []

    #     for index, row in violation_count_per_course.iterrows():
    #         subject_id = row['Mã học phần']
    #         exam_date = row['Ngày thi']
    #         student_count=row['Số lượng vi phạm']
    #         # Lọc ra tất cả các giờ thi của sinh viên 
    #         violating_times = df_input[(df_input['Mã học phần'] == subject_id) & (df_input['Ngày thi'] == exam_date)]['Giờ thi'].unique()
    #         # violating_times_str = ", ".join(violating_times)
    #         self.add_violation(violations, "check_schedule_per_day", 2, "Sinh viên thi quá 2 ca thi trong 1 ngày ", subject_id=subject_id, exam_date=row['Ngày thi'],exam_time=violating_time,extra_info=f'Có {student_count} sinh viên thi quá 2 ca')

    #     if not violations:
    #         self.add_violation(violations, "check_schedule_per_day", 0, "Không có sinh viên nào thi quá 2 ca trong một ngày")

    #     return pd.DataFrame(violations)

    # def check_schedule_per_day(self, df_input):
    #     # Đếm số ca thi của từng sinh viên trong mỗi ngày
    #     shift_count = df_input.groupby(['MSV mở rộng', 'Ngày thi']).size().reset_index(name='số ca thi')

    #     # Lọc ra các sinh viên thi quá 2 ca trong một ngày
    #     violating_students = shift_count[shift_count['số ca thi'] > 2]

    #     violations = []
    #     notified_violations = set()  

    #     # Nếu có sinh viên vi phạm, thêm mã số sinh viên vào danh sách vi phạm
    #     if not violating_students.empty:
    #         print("Mã số sinh viên thi quá 2 ca trong một ngày:")
    #         for student in violating_students['MSV mở rộng'].unique():
    #             print(student)

    #             # Lấy thông tin các lần thi của sinh viên vi phạm
    #             violating_records = df_input[(df_input['MSV mở rộng'] == student) & (df_input['Ngày thi'].isin(violating_students['Ngày thi']))]
    #             for _, row in violating_records.iterrows():
    #                 subject_id = row['Mã học phần mở rộng']
    #                 subject_name = row['Tên học phần']
    #                 exam_date = row['Ngày thi']
    #                 exam_time = row['Giờ thi']
                    
    #                 violation_key = (student, subject_id, exam_date, exam_time)
    #                 if violation_key not in notified_violations:
    #                     self.add_violation(violations, "check_schedule_per_day", 2, "Sinh viên thi quá 2 ca thi trong 1 ngày", subject_id=subject_id, subject_name=subject_name, student_id=student, exam_date=exam_date, exam_time=exam_time)
    #                     notified_violations.add(violation_key)
    #     else:
    #         self.add_violation(violations, "check_schedule_per_day", 0, "Không có sinh viên nào thi quá 2 ca trong một ngày")

    #     return pd.DataFrame(violations)

    def check_schedule_per_day(self, df_input):
        
        shift_count = df_input.groupby(['MSV mở rộng', 'Ngày thi']).size().reset_index(name='số ca thi')

        # Lọc ra các sinh viên thi quá 2 ca trong một ngày
        violating_students = shift_count[shift_count['số ca thi'] > 2]

        # Đếm số lượng sinh viên vi phạm theo mã học phần mở rộng và ngày thi
        violation_count_per_course = df_input[df_input['MSV mở rộng'].isin(violating_students['MSV mở rộng'])]

        violations = []
        notified_violations = set()  

        # Lấy tất cả các mã học phần mở rộng có sinh viên vi phạm
        violating_courses = violation_count_per_course['Mã học phần mở rộng'].unique()

        for subject_id in violating_courses:
            subject_name = df_input[df_input['Mã học phần mở rộng'] == subject_id]['Tên học phần'].iloc[0]
            # Lọc ra các sinh viên vi phạm trong mã học phần mở rộng đó
            violating_students_in_course = violation_count_per_course[(violation_count_per_course['Mã học phần mở rộng'] == subject_id) & (violation_count_per_course['MSV mở rộng'].isin(violating_students['MSV mở rộng']))]

            # Kiểm tra xem có sinh viên nào thi quá 2 ca trong mã học phần đó không
            if not violating_students_in_course.empty:
                student_count = violating_students_in_course['MSV mở rộng'].nunique()
                violating_students_str = ", ".join(map(str, violating_students_in_course['MSV mở rộng'].unique()))

                # Lấy tất cả các ngày thi và giờ thi của sinh viên vi phạm trong mã học phần đó
                violating_times = violating_students_in_course[['Ngày thi', 'Giờ thi']].drop_duplicates()

                for _, time_row in violating_times.iterrows():
                    exam_date = time_row['Ngày thi']
                    violation_key = (subject_id, subject_name, exam_date, student_count)

                    
                    if violation_key not in notified_violations:
                        extra_info = f'Số lượng sinh viên thi hơn 2 ca thi là: {student_count}'
                        self.add_violation(violations, "check_schedule_per_day", 2, "Sinh viên thi quá 2 ca thi trong 1 ngày", subject_id=subject_id, subject_name=subject_name, exam_date=exam_date, extra_info=extra_info)
                        notified_violations.add(violation_key)

        if not violations:
            self.add_violation(violations, "check_schedule_per_day", 0, "Không có sinh viên nào thi quá 2 ca trong một ngày")

        return pd.DataFrame(violations)

    def read_input_files(self):
        if self.label_SVPhanLich.text():
            self.df_input_tab2 = pd.read_excel(self.label_SVPhanLich.text())
        if self.label_cbdl_tab2.text():
            self.df_cbdl_tab2 = pd.read_excel(self.label_cbdl_tab2.text())
        if self.label_date.text():
            self.df_date_tab2 = pd.read_excel(self.label_date.text())
        if self.label_subject.text():
            self.df_subject_tab2 = pd.read_excel(self.label_subject.text())
        if self.label_room.text():
            self.df_room_tab2=pd.read_excel(self.label_room.text())
        # with concurrent.futures.ThreadPoolExecutor() as executor:
        #         future1 = executor.submit(pd.read_excel, self.label_SVPhanLich.text())
        #         future2 = executor.submit(pd.read_excel, self.label_cbdl_tab2.text())
        #         future3 = executor.submit(pd.read_excel, self.label_date.text())
        #         future4 = executor.submit(pd.read_excel, self.label_subject.text())

        #         self.df_input_tab2 = future1.result()
        #         self.df_cbdl_tab2 = future2.result()
        #         self.df_date_tab2 = future3.result()
        #         self.df_subject_tab2 = future4.result()

    def save_results_to_excel(self, results, output_file, sheet_name='Kết quả'):
        # Tạo DataFrame từ kết quả
        df_results = pd.DataFrame(results)
        columns_to_save = ['Loại kiểm tra', 'Mã kết quả', 'Thông điệp', 'Mã học phần mở rộng',
                        'Tên học phần', 'Mã sinh viên mở rộng', 'Ngày thi', 'Giờ thi', 'Ghi chú']

        # Kiểm tra và điền các cột thiếu vào DataFrame
        for col in columns_to_save:
            if col not in df_results.columns:
                df_results[col] = None 

        df_results.sort_values(by=['Mã kết quả','Loại kiểm tra'], inplace=True)

        try:
            with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                df_results.to_excel(writer, sheet_name=sheet_name, index=False, header=True)
                # Đối tượng workbook và worksheet
                workbook  = writer.book
                worksheet = writer.sheets[sheet_name]

                # Thiết lập tự động điều chỉnh độ rộng của các cột
                for i, col in enumerate(columns_to_save):
                    max_len = df_results[col].astype(str).map(len).max()  # Tìm chiều dài lớn nhất của cột
                    max_len = max(max_len, len(col))  # Lấy max với chiều dài tên cột
                    worksheet.set_column(i, i, max_len)  # Đặt chiều rộng tối thiểu cho cột i
            print(f"Kết quả đã được lưu vào sheet '{sheet_name}' của file Excel: {output_file}")
        except Exception as e:
            print(f"Lỗi khi lưu file Excel: {e}")

    def Show_KtrPhanLich(self):
        self.textBrowser_2.clear()
        self.read_input_files()
        if self.df_input_tab2 is None or self.df_cbdl_tab2 is None or self.df_date_tab2 is None:
            self.textBrowser_2.setText("Chưa có dữ liệu. Vui lòng chọn dữ liệu đầu vào.")
            return
        printed_messages = set()
        results = []
        results.append(self.check_student_per_shift(self.df_input_tab2))
        results.append(self.check_subject_student_list(self.df_input_tab2, self.df_cbdl_tab2))
        results.append(self.check_student_in_room(self.df_input_tab2,self.df_room_tab2))
        results.append(self.check_alter_subjects(self.df_input_tab2))
        results.append(self.check_count_room_shift(self.df_input_tab2))
        results.append(self.check_room_reuse(self.df_input_tab2))
        results.append(self.check_exam_datetime(self.df_input_tab2, self.df_date_tab2))
        results.append(self.check_room_assignment(self.df_cbdl_tab2, self.df_input_tab2))
        results.append(self.check_schedule_per_day(self.df_input_tab2))
        combined_results = pd.concat(results, ignore_index=True)
        # for index, row in combined_results.iterrows():
        #     message = f"Thông điệp: {row['Thông điệp']}, Mã kết quả: {row['Mã kết quả']}"

        #     # Kiểm tra xem thông điệp đã được in ra trước đó chưa
        #     if message not in printed_messages:
        #         self.textBrowser_2.append(message)
        #         printed_messages.append(message)
        # Tạo từ điển ánh xạ từng loại kiểm tra sang mô tả tương ứng
        detail_info_map = {
        "check_student_per_shift": "Kiểm tra sinh viên thi 1 học phần trong 1 ca",
        "check_subject_student_list": "Kiểm tra danh sách sinh viên theo học phần",
        "check_student_in_room": "Kiểm tra sinh viên trong phòng thi",
        "check_alter_subjects": "Kiểm tra các học phần thay thế",
        "check_count_room_shift": "Kiểm tra số lượng phòng giữa các ca thi",
        "check_room_reuse": "Kiểm tra sử dụng lại phòng thi",
        "check_exam_datetime": "Kiểm tra ngày giờ thi hợp lệ",
        "check_room_assignment": "Kiểm tra phân bổ phòng thi",
        "check_schedule_per_day": "Kiểm tra lịch thi của sinh viên"
    }
        count=1

        for index, row in combined_results.iterrows():
            message = f"{row['Loại kiểm tra']}_{row['Mã kết quả']}"
            
            if message not in printed_messages:
                detail_info = f"{count}. {detail_info_map.get(row['Loại kiểm tra'], 'Không rõ')}"
                detail_info += f"\n Thông điệp: {row['Thông điệp']}\n Mã kết quả: {row['Mã kết quả']}"
                self.textBrowser_2.append(detail_info)
                printed_messages.add(message)

                count += 1  

        # Lưu kết quả vào thuộc tính 
        self.combined_results = combined_results
        # self.XuatfileKtr.setEnabled(True)

    def export_file_KtrDl(self):
        if self.combined_results is None:
            QMessageBox.warning(None, "Lỗi", "Chưa có kết quả kiểm tra để xuất.")
            return

        save_path, _ = QFileDialog.getSaveFileName(None, "Lưu file khác biệt", "", "Excel Files (*.xlsx)")

        if not save_path:
            QMessageBox.warning(None, "Lỗi", "Chưa chọn nơi lưu file.")
            return
        # Lưu kết quả vào file Excel
        self.save_results_to_excel(self.combined_results, save_path, sheet_name='Kết quả')

        QMessageBox.information(None, "Thông báo", "File đã được xuất thành công.")

    
    def create_summary_excel(self):
        self.textBrowser_2.clear()
        try:
            self.read_input_files()
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Lỗi khi đọc file đầu vào: {str(e)}")
            return
        
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getSaveFileName(self, "Lưu File lịch thi tổng hợp", "", "Excel Files (*.xlsx *);;All Files (*)", options=options)
        if file_name:
            if self.df_input_tab2 is None or self.df_subject_tab2 is None:
                QMessageBox.warning(self, "Lỗi", "Vui lòng chọn đầy đủ các file đầu vào cần thiết.")
                return
                
            df = self.df_input_tab2.copy()
            df.dropna(subset=['Mã học phần'], inplace=True)

            department_dict = self.df_subject_tab2.set_index('Mã HP')['Đơn vị'].to_dict()

            df['Khoa'] = df['Mã học phần'].map(department_dict)
            df['Mã học phần thay thế'] = df.apply(lambda row: row['Mã học phần mở rộng'] if row['Mã học phần'] != row['Mã học phần mở rộng'] else '', axis=1)
            
            df['Phòng Thi'] = df.groupby(['Mã học phần', 'Ngày thi', 'Giờ thi'])['Mã phòng'].transform(lambda x: ', '.join(x.unique()))

            df['Số lượng sinh viên của học phần'] = df.groupby('Mã học phần')['MSV mở rộng'].transform('nunique')

            df['Số lượng phòng thi của học phần'] = df.groupby(['Mã học phần', 'Ngày thi', 'Giờ thi'])['Mã phòng'].transform('nunique')

            df['Ghi chú'] = df.apply(lambda row: self.generate_notes(row), axis=1)

            summary_df = df.groupby(['Mã học phần', 'Tên học phần', 'Số tín chỉ', 'Ngày thi', 'Giờ thi', 'Khoa', 'Ghi chú', 'Mã học phần thay thế'], as_index=False).agg({
                'Phòng Thi': lambda x: ','.join(sorted(set(x))),
                'Số lượng sinh viên của học phần': 'first',
                'Số lượng phòng thi của học phần': 'first'
            })

            temp_df = df.groupby(['Ngày thi', 'Giờ thi'])['Mã phòng'].nunique().reset_index()
            temp_df.rename(columns={'Mã phòng': 'Tổng số phòng của ca thi'}, inplace=True)

            summary_df = summary_df.merge(temp_df, on=['Ngày thi', 'Giờ thi'], how='left')

            summary_df['Tổng số phòng của ca thi'] = summary_df['Tổng số phòng của ca thi'].fillna('')

            summary_df.sort_values(by=['Ngày thi', 'Giờ thi'], inplace=True)

            summary_df['STT'] = range(1, len(summary_df) + 1)
            summary_df = summary_df[['STT', 'Mã học phần', 'Tên học phần', 'Mã học phần thay thế', 'Số tín chỉ', 'Ngày thi', 'Giờ thi', 'Phòng Thi', 'Số lượng sinh viên của học phần', 'Số lượng phòng thi của học phần', 'Tổng số phòng của ca thi', 'Khoa', 'Ghi chú']]

            with pd.ExcelWriter(file_name, engine='xlsxwriter') as writer:
                summary_df.to_excel(writer, index=False, sheet_name='Tổng hợp')

                workbook = writer.book
                worksheet = writer.sheets['Tổng hợp']
                for idx, col in enumerate(summary_df):
                    max_len = max((summary_df[col].astype(str).map(len).max(), len(str(col)))) + 1
                    worksheet.set_column(idx, idx, max_len)
                for col_num, value in enumerate(summary_df.columns.values):
                    worksheet.write(0, col_num, value)

            summary_info = f"File tổng hợp đã được tạo thành công tại: {file_name}\n"
            summary_info += f"Tổng số học phần: {summary_df['Mã học phần'].nunique()}\n"
            summary_info += f"Tổng số ca thi: {summary_df[['Ngày thi', 'Giờ thi']].drop_duplicates().shape[0]}"
            self.textBrowser_2.append(summary_info)

            QMessageBox.information(self, "Thành công", "File tổng hợp đã được tạo thành công.")

        
            
            # except Exception as e:
            #     QMessageBox.critical(self, "Lỗi", f"Đã xảy ra lỗi khi tạo file tổng hợp: {str(e)}")

    def generate_notes(self, row):
        notes = []
        if pd.notna(row.get('Mã học phần mở rộng', '')):  # Kiểm tra nếu giá trị không phải NaN
            if '_Anh' in str(row['Mã học phần mở rộng']):
                notes.append("Thi đề Tiếng Anh")
        if row['Mã học phần'] in ['HRM2001', 'MIS2002', 'MIS2902', 'MGT1902']:
            notes.append("SV làm bài thi trên Elearning")
        return ', '.join(notes) if notes else ''
if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    mainWindow = MainWindow()
    mainWindow.show()
    sys.exit(app.exec_())


