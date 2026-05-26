import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from io import BytesIO
import datetime

# --- 1. HÀM TẠO FILE EXCEL NGẦM ---
def export_phieu_xuat_excel(export_data, selected_date):
    template_path = "phieu_mau.xlsx"
    try:
        wb = openpyxl.load_workbook(template_path)
    except FileNotFoundError:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws['A1'] = "CÔNG TY QUẢN LÝ KHO ĐÔNG THÁP"
        ws.merge_cells('A4:F4')
        ws['A4'] = "PHIẾU XUẤT KHO"
        ws['A4'].font = openpyxl.styles.Font(name="Arial", size=16, bold=True)
        ws['A4'].alignment = openpyxl.styles.Alignment(horizontal="center")
    
    ws = wb.active
    ws.views.sheetView[0].showGridLines = False 
    
    font_regular = Font(name="Arial", size=11)
    font_bold = Font(name="Arial", size=11, bold=True)
    font_italic = Font(name="Arial", size=11, italic=True)
    font_header = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    
    fill_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid") 
    fill_zebra = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")  
    thin_border = Border(
        left=Side(style='thin', color='A0A0A0'),
        right=Side(style='thin', color='A0A0A0'),
        top=Side(style='thin', color='A0A0A0'),
        bottom=Side(style='thin', color='A0A0A0')
    )
    
    date_str = f"Ngày {selected_date.day:02d} tháng {selected_date.month:02d} năm {selected_date.year}"
    ws['C5'] = date_str
    ws['C5'].font = font_italic
    ws['C5'].alignment = Alignment(horizontal="center")
    
    headers = ["STT", "Tên hàng hóa", "Đvt", "Số Lượng", "Diễn Giải", "Ghi Chú"]
    cols = ["A", "B", "C", "D", "E", "F"]
    
    for col_letter, header_text in zip(cols, headers):
        cell = ws[f"{col_letter}8"]
        cell.value = header_text
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
    
    start_row = 9
    for i, item in enumerate(export_data):
        current_row = start_row + i
        ws[f"A{current_row}"] = i + 1
        ws[f"B{current_row}"] = item.get("Tên HH", "")
        ws[f"C{current_row}"] = item.get("Đvt", "")
        ws[f"D{current_row}"] = float(item.get("Số lượng", 0))
        ws[f"E{current_row}"] = item.get("Ghi chú", "")
        ws[f"F{current_row}"] = "" 
        
        ws[f"A{current_row}"].alignment = Alignment(horizontal="center")
        ws[f"B{current_row}"].alignment = Alignment(horizontal="left")
        ws[f"C{current_row}"].alignment = Alignment(horizontal="center")
        ws[f"D{current_row}"].alignment = Alignment(horizontal="right")
        ws[f"D{current_row}"].number_format = '#,##0' 
        ws[f"E{current_row}"].alignment = Alignment(horizontal="left")
        ws[f"F{current_row}"].alignment = Alignment(horizontal="left")
        
        for col_letter in cols:
            cell = ws[f"{col_letter}{current_row}"]
            cell.font = font_regular
            cell.border = thin_border
            if i % 2 == 1:
                cell.fill = fill_zebra
                
    last_data_row = start_row + len(export_data) - 1
    sign_title_row = last_data_row + 2
    sign_name_row = sign_title_row + 4
    
    ws.merge_cells(f'B{sign_title_row}:C{sign_title_row}')
    ws[f"B{sign_title_row}"] = "Người lập phiếu"
    ws.merge_cells(f'E{sign_title_row}:F{sign_title_row}')
    ws[f"E{sign_title_row}"] = "Kế toán / Giám đốc"
    
    for col_letter in ["B", "E"]:
        ws[f"{col_letter}{sign_title_row}"].font = font_bold
        ws[f"{col_letter}{sign_title_row}"].alignment = Alignment(horizontal="center")
        ws[f"{col_letter}{sign_name_row}"] = "(Ký, họ tên)"
        ws[f"{col_letter}{sign_name_row}"].font = font_italic
        ws[f"{col_letter}{sign_name_row}"].alignment = Alignment(horizontal="center")
        
    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 35
    ws.column_dimensions['C'].width = 12
    ws.column_dimensions['D'].width = 15
    ws.column_dimensions['E'].width = 25
    ws.column_dimensions['F'].width = 25
        
    output = BytesIO()
    wb.save(output)
    return output.getvalue()


# --- 2. HÀM HIỂN THỊ GIAO DIỆN (VIEW) ---
def show_print_export_view(service):
    st.subheader("🖨️ In Phiếu Xuất Kho Từ Lịch Sử")
    
    st.markdown("Chọn một ngày để hệ thống tự động gom tất cả các mặt hàng đã **Xuất** trong ngày đó thành một phiếu in.")
    
    # 1. Chọn ngày
    selected_date = st.date_input("Chọn ngày in phiếu:", datetime.date.today())
    
    # 2. Nút Tạo Phiếu
    if st.button("🔄 Tạo Phiếu Xuất", type="primary"):
        # Lấy trực tiếp dữ liệu mới nhất (không dùng cache để đảm bảo cập nhật liên tục)
        st.cache_data.clear() 
        history = service.get_history()
        products = service.get_products()
        
        if not history:
            st.error("Kho dữ liệu trống!")
            return
            
        # Tạo từ điển map Mã hàng -> Đơn vị tính
        dvt_dict = {}
        if products:
            dvt_dict = {str(p[1]): str(p[3]) for p in products} # p[1] là Mã, p[3] là Đvt
            
        # Chuyển lịch sử thành DataFrame để lọc
        if len(history[0]) >= 7:
            df = pd.DataFrame(history, columns=["Ngày", "Mã HH", "Tên hàng hóa", "Loại", "Số Lượng", "Ghi Chú", "Nhân viên"])
        else:
            df = pd.DataFrame(history, columns=["Ngày", "Mã HH", "Loại", "Số Lượng", "Ghi Chú"])
            df["Tên hàng hóa"] = df["Mã HH"] # Tránh lỗi nếu sheet cũ chưa có cột Tên
            
        # ========================================================
        # ĐÂY LÀ PHẦN CODE ĐÃ ĐƯỢC FIX ĐỂ CHỐNG LỖI ĐỊNH DẠNG NGÀY
        # ========================================================
        
        # 1. Chuẩn hóa cột Loại (Đưa về chữ thường "XUẤT" và xóa khoảng trắng thừa)
        df['Loại_chuẩn'] = df['Loại'].astype(str).str.strip().str.upper()
        
        # 2. Xử lý cột Ngày: Pandas tự động đọc mọi định dạng và tách lấy phần Ngày
        df['Ngày_chuẩn'] = pd.to_datetime(df['Ngày'], errors='coerce').dt.date
        
        # 3. Lọc dữ liệu: So sánh khớp Ngày và Loại
        filtered_df = df[(df['Loại_chuẩn'] == 'XUẤT') & (df['Ngày_chuẩn'] == selected_date)]
        
        # ========================================================
        
        if filtered_df.empty:
            st.warning(f"⚠️ Không có giao dịch XUẤT KHO nào được ghi nhận trong ngày {selected_date.strftime('%d/%m/%Y')}.")
        else:
            # Ráp dữ liệu thành List để ném vào hàm Excel
            export_data = []
            for _, row in filtered_df.iterrows():
                ma_hh = str(row.get("Mã HH", ""))
                export_data.append({
                    "Tên HH": row.get("Tên hàng hóa", ma_hh),
                    "Đvt": dvt_dict.get(ma_hh, ""), # Lấy Đvt từ data sản phẩm
                    "Số lượng": row.get("Số Lượng", 0),
                    "Ghi chú": row.get("Ghi Chú", "")
                })
            
            st.success(f"✅ Đã tìm thấy **{len(export_data)}** dòng hàng hóa xuất kho trong ngày này!")
            
            # 4. Tạo file và hiển thị nút tải
            excel_data = export_phieu_xuat_excel(export_data, selected_date)
            
            st.download_button(
                label=f"📥 TẢI PHIẾU XUẤT (Ngày {selected_date.strftime('%d/%m/%Y')})",
                data=excel_data,
                file_name=f"Phieu_Xuat_{selected_date.strftime('%d%m%Y')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
            
            # Hiển thị trước xem đã lấy đúng dữ liệu chưa
            with st.expander("👀 Bấm vào đây để xem trước danh sách hàng hóa sẽ in"):
                st.dataframe(pd.DataFrame(export_data), use_container_width=True, hide_index=True)