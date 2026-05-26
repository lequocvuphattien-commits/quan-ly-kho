import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from io import BytesIO
import datetime

# --- 1. HÀM TẠO FILE EXCEL NGẦM ---
def export_phieu_xuat_excel(cart_data, selected_date):
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
    for i, item in enumerate(cart_data):
        current_row = start_row + i
        ws[f"A{current_row}"] = i + 1
        ws[f"B{current_row}"] = item.get("Tên HH", item.get("Mã HH", ""))
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
                
    last_data_row = start_row + len(cart_data) - 1
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
    st.subheader("🖨️ Tạo & In Phiếu Xuất Kho")
    
    # 1. Chọn ngày
    selected_date = st.date_input("Chọn ngày ghi trên phiếu:", datetime.date.today())
    st.divider()

    # 2. Chọn hàng hóa để đưa vào danh sách in
    products = service.get_products()
    if products:
        p_dict = {f"{p[1]} - {p[2]}": {"Mã": p[1], "Tên": p[2], "Đvt": p[3]} for p in products}
        selected = st.selectbox("Chọn hàng hóa:", options=list(p_dict.keys()), index=None)
        
        c1, c2, c3 = st.columns([1, 2, 0.5])
        with c1: 
            qty = st.number_input("Số lượng", min_value=1.0, value=None, step=1.0)
        with c2: 
            note = st.text_input("Diễn giải / Ghi chú")
        with c3:
            st.write("###")
            if st.button("➕ Thêm"):
                if not selected or not qty:
                    st.warning("⚠️ Nhập đủ thông tin!")
                else:
                    # Sử dụng biến 'print_cart' riêng để không trùng lặp với giỏ hàng Nhập/Xuất thật
                    if 'print_cart' not in st.session_state: st.session_state.print_cart = []
                    st.session_state.print_cart.append({
                        "Mã HH": p_dict[selected]["Mã"],
                        "Tên HH": p_dict[selected]["Tên"],
                        "Đvt": p_dict[selected]["Đvt"],
                        "Số lượng": float(qty),
                        "Ghi chú": note
                    })
                    st.rerun()

    # 3. Hiển thị danh sách và Nút in
    if 'print_cart' in st.session_state and st.session_state.print_cart:
        st.markdown("### 📋 Danh sách sẽ in")
        st.data_editor(pd.DataFrame(st.session_state.print_cart), use_container_width=True, hide_index=True)
        
        c_btn1, c_btn2 = st.columns([1, 1])
        with c_btn1:
            excel_data = export_phieu_xuat_excel(st.session_state.print_cart, selected_date)
            st.download_button(
                label="📥 Tải Phiếu Xuất (Excel)",
                data=excel_data,
                file_name=f"Phieu_Xuat_{selected_date.strftime('%d%m%Y')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
        with c_btn2:
            if st.button("🗑️ Xóa danh sách làm lại", use_container_width=True):
                st.session_state.print_cart = []
                st.rerun()