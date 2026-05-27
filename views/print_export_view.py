import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as OpenpyxlImage
import os
from io import BytesIO
import datetime

# --- 1. HÀM TẠO FILE EXCEL NGẦM ---
def export_phieu_xuat_excel(export_data, selected_date):
    """
    Hàm xuất dữ liệu giỏ hàng ra file Excel - CÓ LOGO CÔNG TY & TỐI ƯU IN KHỔ A4 DỌC
    """
    template_path = "phieu_mau.xlsx"
    try:
        wb = openpyxl.load_workbook(template_path)
    except FileNotFoundError:
        wb = openpyxl.Workbook()
        ws = wb.active
    
    # --- CẤU HÌNH TRANG IN (VỪA KHÍT A4 DỌC) ---
    ws.views.sheetView[0].showGridLines = False 
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1 
    ws.page_setup.fitToHeight = 0 
    
    ws.page_margins.left = 0.5
    ws.page_margins.right = 0.5
    ws.page_margins.top = 0.6
    ws.page_margins.bottom = 0.6

    # --- ĐỊNH NGHĨA FONT & STYLE ---
    font_regular = Font(name="Arial", size=11)
    font_bold = Font(name="Arial", size=11, bold=True)
    font_italic = Font(name="Arial", size=11, italic=True)
    font_header = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    
    fill_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_zebra = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid") 
    
    thin_border = Border(
        left=Side(style='thin', color='B0B0B0'),
        right=Side(style='thin', color='B0B0B0'),
        top=Side(style='thin', color='B0B0B0'),
        bottom=Side(style='thin', color='B0B0B0')
    )
    
    # --- XỬ LÝ CHÈN LOGO ---
    logo_path = "logo.png"
    if os.path.exists(logo_path):
        img = OpenpyxlImage(logo_path)
        img.width = 145
        img.height = 85
        ws.add_image(img, 'E1')
        
    ws.row_dimensions[1].height = 18
    ws.row_dimensions[2].height = 18
    ws.row_dimensions[3].height = 18
    
    ws['A1'] = "CÔNG TY TNHH THỦY SẢN PHÁT TIẾN"
    ws['A1'].font = Font(name="Arial", size=11, bold=True)
    ws['A2'] = "Địa chỉ: Lô B3, đường số 2, Cụm CN Mỹ Hiệp, Xã Mỹ Hiệp, Tỉnh Đồng Tháp"
    ws['A2'].font = Font(name="Arial", size=10, italic=True)
    ws['A3'] = "Số điện thoại: 02778.553.388 - 02773.918.999"
    ws['A3'].font = Font(name="Arial", size=10, italic=True)

    # --- TIÊU ĐỀ PHIẾU (DÒNG 5) ---
    ws['C5'] = "PHIẾU XUẤT KHO"
    ws['C5'].font = Font(name="Arial", size=16, bold=True, color="1F4E78")
    ws['C5'].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[5].height = 32
    
    # --- BỘ PHẬN ĐỀ NGHỊ (Yêu cầu 2: Nằm ở Ô A7) ---
    ws['A7'] = "Bộ phận đề nghị: Thành Phẩm 1"
    ws['A7'].font = font_bold
    ws.row_dimensions[7].height = 20
       
    # --- HEADER BẢNG DỮ LIỆU (Yêu cầu 4: 5 Cột) ---
    ws.row_dimensions[9].height = 26
    
    headers = ["STT", "Tên hàng hóa", "Đvt", "Số Lượng", "Ghi Chú"]
    cols = ["A", "B", "C", "D", "E"]
    
    for col_letter, header_text in zip(cols, headers):
        cell = ws[f"{col_letter}9"]
        cell.value = header_text
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
    
    # --- ĐỔ DỮ LIỆU BẢNG ---
    start_row = 10
    current_row = start_row
    
    for i, item in enumerate(export_data):
        current_row = start_row + i
        ws.row_dimensions[current_row].height = 20
        
        ws[f"A{current_row}"] = i + 1
        ws[f"B{current_row}"] = item.get("Tên HH", "")
        ws[f"C{current_row}"] = item.get("Đvt", "")
        ws[f"D{current_row}"] = float(item.get("Số lượng", 0))
        ws[f"E{current_row}"] = item.get("Ghi chú", "")
        
        # Cấu hình căn lề
        ws[f"A{current_row}"].alignment = Alignment(horizontal="center", vertical="center")
        ws[f"B{current_row}"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws[f"C{current_row}"].alignment = Alignment(horizontal="center", vertical="center")
        ws[f"D{current_row}"].alignment = Alignment(horizontal="center", vertical="center")
        ws[f"D{current_row}"].number_format = '#,##0' 
        ws[f"E{current_row}"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        
        for col_letter in cols:
            cell = ws[f"{col_letter}{current_row}"]
            cell.font = font_regular
            cell.border = thin_border
            if i % 2 == 1:
                cell.fill = fill_zebra
                
    # --- CHỮ KÝ & NGÀY THÁNG (Yêu cầu 3: Ngày tháng nằm trên dòng Người lập) ---
    last_data_row = current_row
    date_row = last_data_row + 2
    sign_title_row = date_row + 1
    
    # Ghi ngày tháng (Gộp ô D và E để cân giữa phía trên người lập)
    date_str = f"Ngày {selected_date.day:02d} tháng {selected_date.month:02d} năm {selected_date.year}"
    ws.merge_cells(f'D{date_row}:E{date_row}')
    ws[f'D{date_row}'] = date_str
    ws[f'D{date_row}'].font = font_italic
    ws[f'D{date_row}'].alignment = Alignment(horizontal="center", vertical="center")
    
    ws.row_dimensions[sign_title_row].height = 25
    
    # Chữ ký Kế Toán (Căn trái cột B)
    ws[f"B{sign_title_row}"] = "Kế Toán"
    ws[f"B{sign_title_row}"].font = font_bold
    ws[f"B{sign_title_row}"].alignment = Alignment(horizontal="left", vertical="center")
    
    # Chữ ký Thủ Kho (Căn giữa cột C)
    ws[f"C{sign_title_row}"] = "Thủ Kho"
    ws[f"C{sign_title_row}"].font = font_bold
    ws[f"C{sign_title_row}"].alignment = Alignment(horizontal="center", vertical="center")
    
    # Chữ ký Người lập (Gộp ô D và E, căn giữa)
    ws.merge_cells(f'D{sign_title_row}:E{sign_title_row}')
    ws[f'D{sign_title_row}'] = "Người Lập"
    ws[f'D{sign_title_row}'].font = font_bold
    ws[f'D{sign_title_row}'].alignment = Alignment(horizontal="center", vertical="center")
        
    # --- CHỈNH KÍCH THƯỚC CỘT (Tối ưu cho 5 cột) ---
    ws.column_dimensions['A'].width = 7   # STT
    ws.column_dimensions['B'].width = 38  # Tên hàng hóa
    ws.column_dimensions['C'].width = 10  # Đvt
    ws.column_dimensions['D'].width = 14  # Số lượng
    ws.column_dimensions['E'].width = 22  # Ghi chú
        
    output = BytesIO()
    wb.save(output)
    return output.getvalue()

# --- 2. HÀM HIỂN THỊ GIAO DIỆN (VIEW) ---
def show_print_export_view(service):
    st.subheader("🖨️ In Phiếu Xuất Kho")
    st.markdown("Chọn một ngày để hệ thống tự động gom tất cả các mặt hàng đã **Xuất** trong ngày đó thành một phiếu in.")
    
    # 1. Chọn ngày (Dữ liệu sẽ tự động xử lý ngay khi bạn đổi ngày)
    selected_date = st.date_input("Chọn ngày in phiếu:", datetime.date.today())
    
    # Lấy dữ liệu từ Google Sheets
    history = service.get_history()
    products = service.get_products()
    
    if not history:
        st.info("Kho dữ liệu trống hoặc đang tải...")
        return
        
    dvt_dict = {str(p[1]): str(p[3]) for p in products} if products else {}
        
    # Chuyển đổi dữ liệu
    if len(history[0]) >= 7:
        df = pd.DataFrame(history, columns=["Ngày", "Mã HH", "Tên hàng hóa", "Loại", "Số Lượng", "Ghi Chú", "Nhân viên"])
    else:
        df = pd.DataFrame(history, columns=["Ngày", "Mã HH", "Loại", "Số Lượng", "Ghi Chú"])
        df["Tên hàng hóa"] = df["Mã HH"]
        
    df['Loại_chuẩn'] = df['Loại'].astype(str).str.strip().str.upper()
    df['Ngày_chuẩn'] = pd.to_datetime(df['Ngày'], errors='coerce').dt.date
    
    # Lọc danh sách XUẤT theo NGÀY
    filtered_df = df[(df['Loại_chuẩn'] == 'XUẤT') & (df['Ngày_chuẩn'] == selected_date)]
    
    if filtered_df.empty:
        st.warning(f"⚠️ Không có giao dịch XUẤT KHO nào được ghi nhận trong ngày {selected_date.strftime('%d/%m/%Y')}.")
        return
        
    # --- XỬ LÝ GỘP DÒNG (Yêu cầu 5) ---
    raw_data = []
    for _, row in filtered_df.iterrows():
        ma_hh = str(row.get("Mã HH", ""))
        raw_data.append({
            "Tên HH": row.get("Tên hàng hóa", ma_hh),
            "Đvt": dvt_dict.get(ma_hh, ""),
            "Số lượng": float(row.get("Số Lượng", 0)),
            "Ghi chú": str(row.get("Ghi Chú", ""))
        })
    
    df_export = pd.DataFrame(raw_data)
    
    # Gom nhóm theo Tên hàng, Đvt, Ghi chú và TÍNH TỔNG số lượng
    # dropna=False để giữ nguyên các dòng có Ghi chú là trống
    df_grouped = df_export.groupby(['Tên HH', 'Đvt', 'Ghi chú'], dropna=False, as_index=False)['Số lượng'].sum()
    export_data = df_grouped.to_dict('records')
    
    # Hiển thị LƯỚI DANH SÁCH (Grid) trực quan ra màn hình
    st.success(f"✅ Đã tìm thấy và cộng dồn thành **{len(export_data)}** mặt hàng xuất kho trong ngày.")
    st.dataframe(df_grouped, use_container_width=True, hide_index=True)
    
    # Tạo sẵn File Excel ngầm
    excel_data = export_phieu_xuat_excel(export_data, selected_date)
    
    # Nút bấm Tải xuống hiển thị luôn, không cần nút trung gian "Tạo phiếu"
    st.download_button(
        label=f"📥 TẢI FILE EXCEL PHIẾU XUẤT KHO (Ngày {selected_date.strftime('%d/%m/%Y')})",
        data=excel_data,
        file_name=f"Phieu_Xuat_{selected_date.strftime('%d%m%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True
    )