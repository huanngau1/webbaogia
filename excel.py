import xlsxwriter

# Tạo workbook và worksheet
workbook = xlsxwriter.Workbook('example.xlsx')
wsTotal = workbook.add_worksheet('Total')
wsDetail = workbook.add_worksheet('Detail')

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

headers = ['STT', 'Mã Sản Phẩm', 'Hình ảnh', 'Thông tin sản phẩm', 'Thương hiệu', 'DVT', 'Số lượng', 'Đơn giá', 'Thành tiền', 'Ghi chú']

# Data mới
data = [
    {'classify': 'Hạng mục thông minh fibaro', 'categories': [{'category': 'Bộ trung tâm', 'products': [{'productId': '9', 'productName': 'HC3', 'productDescription': 'DSFSDFG3', 'productPrice': 12313, 'productQuantity': 1, 'productTotal': 12313}]}, {'category': 'Điều khiển chiếu sáng', 'products': [{'productId': '8', 'productName': 'sw', 'productDescription': '24ddd', 'productPrice': 234, 'productQuantity': 1, 'productTotal': 234}]}]},
    {'classify': 'Hạng mục camera an ninh và thiết bị mạng internet', 'categories': [{'category': 'Camera An ninh', 'products': [{'productId': '11', 'productName': 'Camera nice Indoor', 'productDescription': '345sd', 'productPrice': 3453, 'productQuantity': 1, 'productTotal': 3453}]}]},
    {'classify': 'Hạng mục âm thanh đa vùng', 'categories': [{'category': 'Loa', 'products': [{'productId': '12', 'productName': 'Lo SONOS', 'productDescription': '232', 'productPrice': 252, 'productQuantity': 1, 'productTotal': 252}]}]}
]

rowid = 0
colid = 0

# Sử dụng vòng lặp for để in ra thông tin
for classify in data:
    wsDetail.set_row(rowid, 40)  # Đặt chiều cao dòng cho classify
    wsDetail.merge_range(f'A{rowid + 1}:J{rowid + 1}', classify['classify'], header1_format)
    rowid += 1
    
    # Write the headers with the defined format
    for col_num, header in enumerate(headers):
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
        rowid += 1  # Tăng rowid để chuẩn bị cho classify tiếp theo
    
    # Thêm dòng SUM dưới dòng chi phí nhân công
    if sum_cells:
        wsDetail.set_row(rowid, 30)  # Đặt chiều cao dòng cho tổng chi phí
        final_sum_formula = f"=SUM(I{labor_row+1},{','.join(sum_cells)})"
        wsDetail.merge_range(f'A{rowid + 1}:H{rowid + 1}', "TỔNG CỘNG", sum_format)
        wsDetail.write_formula(rowid, colid + 8, final_sum_formula, sum_format)
        wsDetail.write(rowid, colid + 9, '', sum_format)  # Ghi chú (chưa có dữ liệu)
        rowid += 1  # Tăng rowid để chuẩn bị cho classify tiếp theo
    
    rowid += 1  # Thêm khoảng trống giữa các classify

# Đóng workbook
workbook.close()
