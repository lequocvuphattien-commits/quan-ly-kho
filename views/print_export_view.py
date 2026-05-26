import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from io import BytesIO
import datetime

# --- 1. HÀM TẠO FILE EXCEL NGẦM ---
def export_phieu_xuat_excel(export_data, selected_date):
    """
    Hàm xuất dữ liệu giỏ hàng ra file Excel - TỐI ƯU HOÀN HẢO ĐỂ IN KHỔ A4 DỌC
    """
    template_path = "phieu_mau.xlsx"
    try:
        wb = openpyxl.load_workbook(template_path)
    except FileNotFoundError:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws['A1'] = "CÔNG TY TNHH THỦY SẢN PHÁT TIẾN"
        ws.merge_cells('A4:F4')
        ws['A4'] = "PHIẾU XUẤT KHO"
        ws['A4'].font = openpyxl.styles.Font(name="Arial", size=16, bold=True)
        ws['A4'].alignment = openpyxl.styles.Alignment(horizontal="center")
    
    ws = wb.active
    
    # ----------------=======================================----------------
    # CẤU HÌNH TRANG IN PHẢI CÓ ĐỂ IN VỪA KHÍT KHỔ A4
    # ----------------=======================================----------------
    ws.views.sheetView[0].showGridLines = False  # Ẩn đường lưới mờ của Excel để phiếu sạch sẽ
    
    # 1. Khai báo kích thước giấy A4 và hướng dọc (Portrait)
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    
    # 2. Cấu hình tự động ép vừa khít chiều ngang trang giấy A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1   # Ép toàn bộ các cột nằm trọn chiều ngang 1 trang
    ws.page_setup.fitToHeight = 0  # Chiều dọc tự do giãn ra trang sau nếu nhiều hàng
    
    # 3. Cài đặt khoảng cách lề (Margins) đơn vị Inches chuẩn in ấn
    ws.page_margins.left = 0.5
    ws.page_margins.right = 0.5
    ws.page_margins.top = 0.6
    ws.page_margins.bottom = 0.6
    # ----------------=======================================----------------

    # Định nghĩa Font và Phong cách chuyên nghiệp hơn
    font_regular = Font(name="Arial", size=11)
    font_bold = Font(name="Arial", size=11, bold=True)
    font_italic = Font(name="Arial", size=11, italic=True)
    font_header = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    
    fill_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid") # Xanh navy sang trọng
    fill_zebra = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")  # Màu sọc nhẹ tinh tế
    
    thin_border = Border(
        left=Side(style='thin', color='B0B0B0'),
        right=Side(style='thin', color='B0B0B0'),
        top=Side(style='thin', color='B0B0B0'),
        bottom=Side(style='thin', color='B0B0B0')
    )
    
    # --- THÔNG TIN CÔNG TY & TIÊU ĐỀ PHIẾU ---
    ws['A1'] = "CÔNG TY TNHH THỦY SẢN PHÁT TIẾN"
    ws['A1'].font = Font(name="Arial", size=11, bold=True)
    
    ws['A2'] = "Địa chỉ: Lô B3, đường số 2, Cụm CN Mỹ Hiệp, Xã Mỹ Hiệp, Tỉnh Đồng Tháp"
    ws['A2'].font = Font(name="Arial", size=10, italic=True)
    
    ws['A3'] = "Số điện thoại: 02778.553.388 - 02773.918.999"
    ws['A3'].font = Font(name="Arial", size=10, italic=True)

    ws.merge_cells('A4:F4')
    ws['A4'] = "PHIẾU XUẤT KHO"
    ws['A4'].font = Font(name="Arial", size=16, bold=True, color="1F4E78")
    ws['A4'].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[4].height = 30
    
    # --- GHI NGÀY IN PHIẾU (MERGE A5:F5) ---
    date_str = f"Ngày {selected_date.day:02d} tháng {selected_date.month:02d} năm {selected_date.year}"
    ws.merge_cells('A5:F5')
    ws['A5'] = date_str
    ws['A5'].font = font_italic
    ws['A5'].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[5].height = 20
       
    # Thiết lập độ rộng dòng tiêu đề bảng (Dòng số 8) thoáng hơn
    ws.row_dimensions[8].height = 26
    
    headers = ["STT", "Tên hàng hóa", "Đvt", "Số Lượng", "Diễn Giải", "Ghi Chú"]
    cols = ["A", "B", "C", "D", "E", "F"]
    
    for col_letter, header_text in zip(cols, headers):
        cell = ws[f"{col_letter}8"]
        cell.value = header_text
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
    
    # Đổ dữ liệu hàng hóa từ dòng số 9
    start_row = 9
    for i, item in enumerate(export_data):
        current_row = start_row + i
        
        # Thiết lập chiều cao dòng dữ liệu rộng rãi, dễ đọc khi in ra giấy
        ws.row_dimensions[current_row].height = 22
        
        ws[f"A{current_row}"] = i + 1
        ws[f"B{current_row}"] = item.get("Tên HH", "")
        ws[f"C{current_row}"] = item.get("Đvt", "")
        ws[f"D{current_row}"] = float(item.get("Số lượng", 0))
        ws[f"E{current_row}"] = item.get("Ghi chú", "")
        ws[f"F{current_row}"] = "" 
        
        # Cấu hình căn lề kết hợp thuộc tính wrap_text=True (tự động xuống dòng khi chữ quá dài)
        ws[f"A{current_row}"].alignment = Alignment(horizontal="center", vertical="center")
        ws[f"B{current_row}"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws[f"C{current_row}"].alignment = Alignment(horizontal="center", vertical="center")
        ws[f"D{current_row}"].alignment = Alignment(horizontal="right", vertical="center")
        ws[f"D{current_row}"].number_format = '#,##0' 
        ws[f"E{current_row}"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws[f"F{current_row}"].alignment = Alignment(horizontal="left", vertical="center")
        
        for col_letter in cols:
            cell = ws[f"{col_letter}{current_row}"]
            cell.font = font_regular
            cell.border = thin_border
            if i % 2 == 1:
                cell.fill = fill_zebra
                
    # --- PHẦN XỬ LÝ CHỮ KÝ ĐÃ TÙY CHỈNH THEO CỘT VÀ VỊ TRÍ ---
    last_data_row = start_row + len(export_data) - 1
    sign_title_row = last_data_row + 2
    
    # Thiết lập chiều cao cho dòng chức danh ký tên
    ws.row_dimensions[sign_title_row].height = 25
    
    # 1. Kế Toán (Cột B) - Căn trái
    ws[f"B{sign_title_row}"] = "Kế Toán"
    ws[f"B{sign_title_row}"].font = font_bold
    ws[f"B{sign_title_row}"].alignment = Alignment(horizontal="left", vertical="center")
    
    # 2. Thủ Kho (Cột C) - Căn giữa
    ws[f"C{sign_title_row}"] = "Thủ Kho"
    ws[f"C{sign_title_row}"].font = font_bold
    ws[f"C{sign_title_row}"].alignment = Alignment(horizontal="center", vertical="center")
    
    # 3. Người Lập Phiếu (Cột E) - Căn phải
    ws[f"E{sign_title_row}"] = "Người Lập Phiếu"
    ws[f"E{sign_title_row}"].font = font_bold
    ws[f"E{sign_title_row}"].alignment = Alignment(horizontal="right", vertical="center")
    
    # Đã bỏ dòng "(Ký, họ tên)" theo yêu cầu
    # --- KẾT THÚC PHẦN CHỮ KÝ ---
        
    # Định tỷ lệ độ rộng cột tối ưu hoàn hảo cho khổ dọc A4 (Tổng số khoảng ~85-90 là chuẩn đẹp)
    ws.column_dimensions['A'].width = 6   # STT
    ws.column_dimensions['B'].width = 32  # Tên hàng hóa rộng để hạn chế xuống dòng nhiều
    ws.column_dimensions['C'].width = 10  # Đvt
    ws.column_dimensions['D'].width = 14  # Số lượng
    ws.column_dimensions['E'].width = 24  # Diễn giải
    ws.column_dimensions['F'].width = 12  # Ghi chú
        
    output = BytesIO()
    wb.save(output)
    return output.getvalue()


# --- 2. HÀM HIỂN THỊ GIAO DIỆN (VIEW) ---
def show_print_export_view(service):
    st.subheader("🖨️ In Phiếu Xuất Kho")
    
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
        # ĐÂY LÀ PHẦN CODE ĐÃ ĐƯỢC FIX THEO CHỮ "XUẤT" VIẾT HOA VÀ ĐỊNH DẠNG NGÀY
        # ========================================================
        
        # 1. Chuẩn hóa: Ép kiểu chuỗi, xóa khoảng trắng 2 đầu và đưa về chữ IN HOA
        df['Loại_chuẩn'] = df['Loại'].astype(str).str.strip().str.upper()
        
        # 2. Xử lý cột Ngày: Pandas cắt bỏ giờ phút, chỉ lấy ngày gọn gàng
        df['Ngày_chuẩn'] = pd.to_datetime(df['Ngày'], errors='coerce').dt.date
        
        # 3. Lọc dữ liệu: So sánh khớp Ngày và Loại là chữ "XUẤT" in hoa
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