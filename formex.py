import xlsxwriter

# Tạo workbook và worksheet
workbook = xlsxwriter.Workbook('example.xlsx')
wsTotal = workbook.add_worksheet('Total')
wsDetail = workbook.add_worksheet('Details')

# Định dạng cột trang Detail
wsTotal.set_column('A:A', 7)
wsTotal.set_column('B:B', 18)
wsTotal.set_column('C:C', 5)
wsTotal.set_column('D:D', 80)
wsTotal.set_column('E:E', 20)
wsTotal.set_column('F:F', 20)
wsTotal.set_column('G:G', 5)
wsTotal.set_column('H:H', 30)
wsTotal.set_column('I:I', 10)
wsTotal.set_column('J:J', 10)

# Định dạng cột trang Detail
wsDetail.set_column('A:A', 7)
wsDetail.set_column('B:B', 20)
wsDetail.set_column('C:C', 20)
wsDetail.set_column('D:D', 80)
wsDetail.set_column('E:E', 15)
wsDetail.set_column('F:F', 15)
wsDetail.set_column('G:G', 20)
wsDetail.set_column('H:H', 20)
wsDetail.set_column('I:I', 20)
wsDetail.set_column('J:J', 25)

# Định dạng tiêu đề
header1_format = workbook.add_format({
    'bold': True,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'bg_color': '#9FC5E8'  # Màu cho classify
})

header2_format = workbook.add_format({
    'bold': True,
    'bg_color': '#F4CCCC',
    'border': 1,
    'align': 'center',
    'valign': 'vcenter'
})

# Định dạng số
number_format = workbook.add_format({
    'border': 1,
    'num_format': '#,##0.00',
    'bg_color': '#FFFFFF'
    })

# Định dạng sum
sum_format = workbook.add_format({
    'bold': True,
    'bg_color': '#D9EAD3',
    'border': 1,
    'align': 'right',
    'valign': 'vcenter',
    'num_format': '#,##0.00'
})

# Định dạng nhân công
labor_format = workbook.add_format({
    'bold': True,
    'bg_color': '#FFD966',  # Màu cho nhân công và category
    'border': 1,
    'align': 'right',
    'valign': 'vcenter',
    'num_format': '#,##0.00'
})

# Định dạng product
product_format = workbook.add_format({
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'bg_color': '#FFFFFF'
})


# Định dạng với đường viền màu trắng
infoCompany_format = workbook.add_format({
    'align': 'center',
    'valign': 'vcenter',
    'border': 1,
    'border_color': '#FFFFFF'
})

# Định dạng căn giữa và căn lề trái
merge_format_center = workbook.add_format({
    'align': 'center',
    'valign': 'vcenter',
    'border': 1,
    'border_color': '#FFFFFF'
})
merge_format_left = workbook.add_format({
    'align': 'left', 
    'valign': 'vcenter',
    'border': 1,
    'border_color': '#FFFFFF'
})

#-------------------------------------------------------------------------------------------------------
# Dữ liệu infoCompany
infoCompany = [
    'CÔNG TY CỔ PHẦN KIM SƠN TIẾN',
    'Đ/c: Số 16 đường 35, P. An Khánh, TP. Thủ Đức, TP. HCM',
    'Hotline: 0913 699 545',
    'Website: www.kimsontien.com / www.fibarovn.com'
]
form_input = [
    ['Kính gửi quý Khách hàng',':','','Đại diện kinh doanh',':','Hồ Thị Bích Ngà'],
    ['Số điện thoại',':','','Số điện thoại',':','0914 172 812'],
    ['Email',':','','Email',':','bichnga@kimsontien.com'],
    ['Địa chỉ',':','','Số báo giá',':','BGKST11112024/01'],
]

headersD = ['STT', 'Mã Sản Phẩm', 'Hình ảnh', 'Thông tin sản phẩm', 'Thương hiệu', 'DVT', 'Số lượng', 'Đơn giá', 'Thành tiền', 'Ghi chú']
headersT = ['STT', 'DANH MỤC HỆ THỐNG', 'MÔ TẢ HỆ THỐNG']

STT_Total = ['A', 'B', 'C', 'D', 'E','F', 'G', 'H']
STT_Total1 = ['1', '2', '3', '4', '5','6','7','8','9','10','11','12','13']

# Data mới
data = [
    {'classify': 'Hạng mục thông minh fibaro', 'categories': [{'category': 'Bộ trung tâm', 'products': [{'productId': '9', 'productName': 'HC3', 'productDescription': 'DSFSDFG3', 'productPrice': 12313, 'productQuantity': 1, 'productTotal': 12313}]}, {'category': 'Điều khiển chiếu sáng', 'products': [{'productId': '8', 'productName': 'sw', 'productDescription': '24ddd', 'productPrice': 234, 'productQuantity': 1, 'productTotal': 234}]}]},
    {'classify': 'Hạng mục camera an ninh và thiết bị mạng internet', 'categories': [{'category': 'Camera An ninh', 'products': [{'productId': '11', 'productName': 'Camera nice Indoor', 'productDescription': '345sd', 'productPrice': 3453, 'productQuantity': 1, 'productTotal': 3453}]}]},
    {'classify': 'Hạng mục âm thanh đa vùng', 'categories': [{'category': 'Loa', 'products': [{'productId': '12', 'productName': 'Lo SONOS', 'productDescription': '232', 'productPrice': 252, 'productQuantity': 1, 'productTotal': 252}]}]}
]



value_sum_cellsT = []
value_allsum_cellsT = []
value_nhancong_cellsT = []


#--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
rowid = 0
colid = 0

# Sử dụng vòng lặp for để in ra thông tin
for classify in data:
    
    wsDetail.set_row(rowid, 40)  # Đặt chiều cao dòng cho classify
    wsDetail.merge_range(f'A{rowid + 1}:J{rowid + 1}', classify['classify'], header1_format)
    rowid += 1
    
    # Write the headers with the defined format
    for col_num, header in enumerate(headersD):
        wsDetail.write(rowid, col_num, header, header2_format)
    
    rowid += 1  # Di chuyển xuống hàng kế tiếp sau khi viết header

    sum_cells = []  # Danh sách lưu trữ địa chỉ ô của các ô tổng tiền
    letter = 65  # Khởi đầu với ký tự 'A' (ASCII 65)

    for category_data in classify['categories']:
        wsDetail.set_row(rowid, 30)  # Đặt chiều cao dòng cho category
        
        # Đánh dấu dòng bắt đầu của các sản phẩm trong category này
        start_row = rowid + 1  
        
        # Dòng chứa tên category
        wsDetail.write(rowid, colid, chr(letter), labor_format)  # Điền chữ cái A, B, C...
        wsDetail.merge_range(f'B{rowid + 1}:H{rowid + 1}', category_data['category'], labor_format)
        wsDetail.write(rowid, colid + 9, '', labor_format)  # Ghi chú (chưa có dữ liệu)

        # Dòng chứa công thức SUM
        end_row = start_row + len(category_data['products']) - 1  # Dòng cuối cùng của các sản phẩm
        sum_cell = f'I{rowid + 1}'
        wsDetail.write_formula(rowid, colid + 8, f'SUM(I{start_row+1}:I{end_row + 1})', labor_format)
        
        sum_cells.append(sum_cell)  # Lưu lại địa chỉ của ô tổng tiền
        value_sum_cellsT.append(sum_cell) 
        rowid += 1  # Chuyển đến hàng tiếp theo cho sản phẩm




        
        for index, product in enumerate(category_data['products']):
            wsDetail.set_row(rowid, 60)  # Đặt chiều cao dòng cho product
            wsDetail.write(rowid, colid, index + 1, product_format)  # STT
            wsDetail.write(rowid, colid + 1, product['productName'], product_format)  # Mã Sản Phẩm
            wsDetail.write(rowid, colid + 2, '', product_format)  # Hình ảnh (chưa có dữ liệu)
            wsDetail.write(rowid, colid + 3, product['productDescription'], product_format)  # Thông tin sản phẩm
            wsDetail.write(rowid, colid + 4, '', product_format)  # Thương hiệu (chưa có dữ liệu)
            wsDetail.write(rowid, colid + 5, '', product_format)  # DVT (chưa có dữ liệu)
            wsDetail.write(rowid, colid + 6, product['productQuantity'], product_format)  # Số lượng
            wsDetail.write(rowid, colid + 7, product['productPrice'], number_format)  # Đơn giá
            wsDetail.write(rowid, colid + 8, product['productTotal'], number_format)  # Thành tiền
            wsDetail.write(rowid, colid + 9, '', product_format)  # Ghi chú (chưa có dữ liệu)
            
            rowid += 1  # Chuyển đến hàng tiếp theo cho sản phẩm
        
        # rowid += 1  # Thêm khoảng trống giữa các danh mục   -------- CHƯA CẦN THIẾT
        letter += 1  # Tăng ký tự cho category tiếp theo
    
    wsDetail.set_row(rowid, 30)  # Đặt chiều cao dòng cho chi phí nhân công
    wsDetail.write(rowid, colid, chr(letter), labor_format)  # Điền chữ cái cho chi phí nhân công
    wsDetail.merge_range(f'B{rowid + 1}:H{rowid + 1}', "NHÂN CÔNG LẮP ĐẶT, PHỤ KIỆN VÀ CẤU HÌNH HỆ THỐNG", labor_format)
    wsDetail.write(rowid, colid + 9, '', labor_format)  # Ghi chú (chưa có dữ liệu)
    # Thêm dòng hiển thị nhân công cấu hình dưới cùng của classify
    if sum_cells:
        sum_formula = ','.join(sum_cells)
        labor_row = rowid  # Lưu hàng này để sử dụng trong SUM tổng
        labor_formula = f"=ROUND(SUM({sum_formula})*13%,-3)"
        wsDetail.write_formula(rowid, colid + 8, labor_formula, labor_format)
        nhancong_cell = f'I{rowid+1}'
        value_sum_cellsT.append(nhancong_cell)
        value_nhancong_cellsT.append(nhancong_cell)
        rowid += 1  # Tăng rowid để chuẩn bị cho classify tiếp theo
    
    # Thêm dòng SUM dưới dòng chi phí nhân công
    if sum_cells:
        wsDetail.set_row(rowid, 30)  # Đặt chiều cao dòng cho tổng chi phí
        final_sum_formula = f"=SUM(I{labor_row+1},{','.join(sum_cells)})"
        wsDetail.merge_range(f'A{rowid + 1}:H{rowid + 1}', "TỔNG CỘNG", sum_format)
        wsDetail.write_formula(rowid, colid + 8, final_sum_formula, sum_format)
        wsDetail.write(rowid, colid + 9, '', sum_format)  # Ghi chú (chưa có dữ liệu)
        sumall_cell = f'I{rowid+1}'
        value_allsum_cellsT.append(sumall_cell)
        rowid += 1  # Tăng rowid để chuẩn bị cho classify tiếp theo
    

     




    rowid += 1  # Thêm khoảng trống giữa các classify
    

#--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
print(value_allsum_cellsT)
print(value_sum_cellsT)
print(value_nhancong_cellsT)
# Áp dụng định dạng cho tất cả các ô từ A1 đến J8
for row in range(8):  # Dòng từ 1 đến 8 (index 0 đến 7)
    for col in range(10):  # Cột từ A đến J (index 0 đến 9)
        wsTotal.write(row, col, '', infoCompany_format)


# Merge các ô và điền dữ liệu
for i in range(4):
    # Hợp nhất các ô E1:J1, E2:J2, E3:J3, E4:J4
    wsTotal.merge_range(i, 4, i, 9, infoCompany[i], infoCompany_format)


for i, row_data in enumerate(form_input):
    row = 4 + i  # Dòng bắt đầu từ dòng 5 (index 4 trong xlsxwriter)

    # Hợp nhất các ô và điền dữ liệu cho từng cột
    wsTotal.merge_range(row, 0, row, 1, row_data[0], merge_format_left)  # A5:B5, A6:B6, ...
    wsTotal.write(row, 2, row_data[1], merge_format_center)               # C5, C6, ...
    wsTotal.merge_range(row, 3, row, 4, row_data[2], merge_format_center)  # D5:E5, D6:E6, ...
    wsTotal.write(row, 5, row_data[3], merge_format_left)
    wsTotal.write(row, 6, row_data[4], merge_format_center)                   # F5, F6, ...
    wsTotal.merge_range(row, 7, row, 9, row_data[5], merge_format_center)  # G5:J5, G6:J6, ...
wsTotal.merge_range(8, 0, 8, 9, 'Công ty Cổ Phần Kim Sơn Tiến xin trân gửi đến anh Bảng Báo gía hệ thống Smarthome, Camera an ninh, thiết bị mạng và hệ thống âm thanh cho công trình, cụ thể như sau:', merge_format_left)
wsTotal.merge_range(9, 0, 9, 9, 'BẢNG TỔNG HỢP GIÁ TRỊ HỆ THỐNG NHÀ THÔNG MINH FIBARO - KIM SƠN TIẾN', header1_format)




# Thêm các tiêu đề vào các ô tương ứng
wsTotal.write(10,0, headersT[0], header1_format)   # STT vào A11
wsTotal.merge_range(10, 1, 10, 4, headersT[1], header1_format)  # DANH MỤC HỆ THỐNG vào B11:E11
wsTotal.merge_range(10, 5, 10, 9, headersT[2], header1_format)  # MÔ TẢ HỆ THỐNG vào F11:J11

rowTotal = 11
indx = 0
idx2 =0
for classify in data:
    wsTotal.write(rowTotal,0, STT_Total[indx], header1_format)   # STT vào A11
    wsTotal.merge_range(rowTotal, 1, rowTotal, 3, classify['classify'], header1_format)
    wsTotal.write(rowTotal,4, f'=Details!{value_allsum_cellsT[indx]}', header1_format)  
    wsTotal.merge_range(rowTotal, 5, rowTotal, 9, '', header1_format) 
    indx +=1
    rowTotal+=1
    
    for category_data in classify['categories']:
        wsTotal.write(rowTotal,0, STT_Total1[idx2], header1_format)   # STT vào A11
        wsTotal.merge_range(rowTotal, 1, rowTotal, 3, category_data['category'], header1_format)
        wsTotal.write(rowTotal,4, f'=Details!{value_sum_cellsT[idx2]}', header1_format)  
        wsTotal.merge_range(rowTotal, 5, rowTotal, 9, '', header1_format) 
        idx2 +=1
        rowTotal+=1
    wsTotal.write(rowTotal,0, STT_Total1[idx2], header1_format)   # STT vào A11
    wsTotal.merge_range(rowTotal, 1, rowTotal, 3, "Nhân công lắp đặt", header1_format)
    wsTotal.write(rowTotal,4, f'=Details!{value_sum_cellsT[idx2]}', header1_format)  
    wsTotal.merge_range(rowTotal, 5, rowTotal, 9, '', header1_format) 
    rowTotal += 1  # Chuyển đến dòng tiếp theo
    idx2 +=1



final_sum1 = f"=SUM({','.join([f'Details!{cell}' for cell in value_allsum_cellsT])})"

wsTotal.merge_range(rowTotal, 0, rowTotal, 3, "TỔNG TRƯỚC THUẾ", header1_format)
wsTotal.write(rowTotal,4, final_sum1, header1_format)
wsTotal.merge_range(rowTotal, 5, rowTotal, 9, "", header1_format)

arr_device = list(set(value_sum_cellsT) - set(value_nhancong_cellsT))
numper_d = 10
final_vat_device = f"=SUM({','.join([f'Details!{cell}' for cell in arr_device])})*{numper_d}%"
vat_device_with_plus = final_vat_device.replace('=', '+', 1)
wsTotal.merge_range(rowTotal+1, 0, rowTotal+1, 3, "VAT THIẾT BỊ 10%", header1_format)
wsTotal.write(rowTotal+1,4, final_vat_device, header1_format)
wsTotal.merge_range(rowTotal+1, 5, rowTotal+1, 9, "", header1_format)


numper_p = 8
final_vat_person = f"=SUM({','.join([f'Details!{cell}' for cell in arr_device])})*{numper_p}%"
vat_person_with_plus = final_vat_person.replace('=', '+', 1)
wsTotal.merge_range(rowTotal+2, 0, rowTotal+2, 3, "VAT NHÂN CÔNG 8%", header1_format)
wsTotal.write(rowTotal+2,4, final_vat_person, header1_format)
wsTotal.merge_range(rowTotal+2, 5, rowTotal+2, 9, "", header1_format)

all_sum = final_sum1 + vat_device_with_plus + vat_person_with_plus
wsTotal.merge_range(rowTotal+3, 0, rowTotal+3, 3, "TỔNG THANH TOÁN", header1_format)
wsTotal.write(rowTotal+3,4, all_sum, header1_format)
wsTotal.merge_range(rowTotal+3, 5, rowTotal+3, 9, "", header1_format)
# Đóng workbook
workbook.close()
