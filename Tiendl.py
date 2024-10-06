# mainwindow.py
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog, QMainWindow, QMessageBox
import pandas as pd
from ui_main import Ui_MainWindow  # Import lớp được sinh ra từ Qt Designer

class Tab1():
    def __init__(self, parent):
        self.parent = parent

    def load_file_PhanLich(self):
        self.load_file_and_update_label('label_4', 'Danh sách sinh viên thi')

    def load_file_Student_CT2(self):
        self.load_file_and_update_label('label_5', 'Danh sách sinh viên 2 CT')

    def load_file_Alter_Subject(self):
        self.load_file_and_update_label('label_6', 'Danh sách học phần thay thế')

    def load_file_ChuanBiDuLieu(self):
        self.load_file_and_update_label('label_10', 'File Chuẩn bị dữ liệu')

    def load_file_and_update_label(self, label_name, dialog_title):

        file_name, _ = QFileDialog.getOpenFileName(self.parent, dialog_title, "", "Excel Files (*.xlsx *.xls)")
        if file_name:
            self.parent.label_3.setText(file_name)
            self.parent.df_date_tab2 = pd.read_excel(file_name)



    def remove_duplicates(self):
        self.df_input = self.df_input.drop_duplicates()

    def student_CT2(self):
        df_CT2_unique = self.df_CT2.drop_duplicates(subset='MSV_CT1')
        msv_ct1_to_ct2 = df_CT2_unique.set_index('MSV_CT1')['MSV_CT2'].to_dict()

        if 'MSV mở rộng' not in self.df_input.columns:
            self.df_input['MSV mở rộng'] = None

        self.df_input['MSV mở rộng'] = self.df_input['MSV'].map(msv_ct1_to_ct2).fillna(self.df_input['MSV'])

    def alter_subject(self):
        subject_mapping = self.df_alter_subject.set_index('Mã học phần')['Mã học phần thay thế'].to_dict()

        if 'Mã học phần mở rộng' not in self.df_input.columns:
            self.df_input['Mã học phần mở rộng'] = None

        self.df_input['Mã học phần mở rộng'] = self.df_input['Mã học phần'].map(subject_mapping).fillna(self.df_input['Mã học phần'])

        self.df_input['Ghi chú'] = self.df_input.apply(lambda row: f"Mã học phần cũ: {row['Mã học phần']}" if row['Mã học phần'] in subject_mapping else row['Ghi chú'], axis=1)

    def remove_foreign_language_exempted_students(self):
        self.df_input = self.df_input[~self.df_input['HP miễn ngoại ngữ'].str.contains('Miễn', case=False, na=False)]

    def remove_discontinued_students(self):
        self.df_input = self.df_input[~self.df_input['Ghi chú'].str.contains('SV tạm ngừng học|SV đã thôi học|SV rút HP', case=False, na=False)]

    def english_exam(self):
        self.df_input.loc[self.df_input['Đề thi TA'].str.contains('x', case=False, na=False), 'Mã học phần mở rộng'] += '_Anh'

    def compare_files(self):
        common_columns = list(set(self.df_input.columns) & set(self.df_cbdl.columns))
        merged = pd.merge(self.df_input[common_columns], self.df_cbdl[common_columns], on=['MSV', 'Mã học phần'], how='outer', suffixes=('_input', '_cbdl'), indicator=True)

        differences = merged[merged['_merge'] != 'both']
        
        if differences.empty:
            return "File dữ liệu tiền xử lý đã chính xác.", None
        else:
            return None, differences

    def Show_KtrDL(self):
        self.parent.textBrowser.clear()
        if self.parent.df_input_tab1 is None or self.parent.df_CT2 is None or self.parent.df_alter_subject is None:
            self.parent.textBrowser.setText("Chưa có dữ liệu. Vui lòng chọn dữ liệu đầu vào.")
            return

        try:
            # Read the input files based on the paths in labels
            df_input = pd.read_excel(self.parent.label_4.text())
            df_CT2 = pd.read_excel(self.parent.label_5.text())
            df_alter_subject = pd.read_excel(self.parent.label_6.text())
            df_cbdl = pd.read_excel(self.parent.label_10.text())

            # Store the dataframes in the parent object for further use
            self.parent.df_input_tab1 = df_input
            self.parent.df_CT2 = df_CT2
            self.parent.df_alter_subject = df_alter_subject
            self.parent.df_cbdl = df_cbdl

            # Perform data processing steps
            self.remove_duplicates()
            self.student_CT2()
            self.alter_subject()
            self.remove_foreign_language_exempted_students()
            self.remove_discontinued_students()
            self.english_exam()

            # Show the results in textBrowser
            success_message, differences = self.compare_files()

            if success_message:
                self.parent.textBrowser.setText(success_message)
                self.parent.XuatfileKtr.setEnabled(False)
            else:
                source_file = "file đầu vào đã xử lý" if differences.columns.equals(df_input.columns) else "file chuẩn bị dữ liệu"
                self.parent.textBrowser.append(f"Có sự khác biệt về mã sinh viên và mã học phần giữa các file. Thông tin chi tiết được lấy từ {source_file}, được liệt kê dưới đây:\n")

                output = "MSV\t\tMã học phần\t\n"
                for index, row in differences.iterrows():
                    msv = row['MSV']
                    ma_hoc_phan = row['Mã học phần']
                    output += f"{msv}\t\t{ma_hoc_phan}\t\n"

                self.parent.textBrowser.append(output)
                self.parent.XuatfileKtr.setEnabled(True)
                self.parent.differences = differences

        except Exception as e:
            QMessageBox.warning(self.parent, "Lỗi", f"Lỗi xử lý dữ liệu: {str(e)}")

    def export_differences(self):
        save_path, _ = QFileDialog.getSaveFileName(None, "Lưu file khác biệt", "", "Excel Files (*.xlsx)")

        if not save_path:
            QMessageBox.warning(None, "Lỗi", "Chưa chọn nơi lưu file.")
            return

        try:
            desired_order = [
                'Lớp', 'MSV', 'Mã học phần', 'Tên học phần', 'Đề thi TA',
                'Số tín chỉ', 'Lớp học phần', 'Loại đăng ký', 'HP miễn ngoại ngữ',
                'Ghi chú', 'Hình thức thi', 'Mã học phần mở rộng', 'MSV mở rộng'
            ]
            self.differences = self.differences[desired_order]

            with pd.ExcelWriter(save_path, engine='xlsxwriter') as writer:
                self.differences.to_excel(writer, index=False, sheet_name='Sheet1')

                workbook = writer.book
                worksheet = writer.sheets['Sheet1']

                # Ghi tên cột vào sheet Excel
                for col_num, value in enumerate(self.differences.columns):
                    cell_format = workbook.add_format({'text_wrap': True, 'valign': 'top'})
                    worksheet.write(0, col_num, value, cell_format)

            QMessageBox.information(None, "Thông báo", "File đã được xuất thành công.")
        except Exception as e:
            QMessageBox.critical(None, "Lỗi", f"Lỗi khi xuất file: {str(e)}")
