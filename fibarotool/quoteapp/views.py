from django.shortcuts import render
from .models import Classify, Brand
from io import BytesIO
import pandas as pd
from django.http import HttpResponse
from django.views.decorators.csrf import csrf_exempt
import json
from collections import defaultdict
import xlsxwriter
import requests  # Thêm thư viện này để tải ảnh từ URL
from django.templatetags.static import static
def home(request):
    classifys = Classify.objects.all()
    brands = Brand.objects.all()
    context = {
        'classifys': classifys,
        'brands': brands
    }
    return render(request, 'quoteapp/home.html', context)

@csrf_exempt  # Chỉ dùng khi bạn đã hiểu rõ về bảo mật
def download_excel(request):
    if request.method == 'POST':
        data = json.loads(request.body)
        classify_map = defaultdict(list)
        classifys = data.get('classifys', [])
        FormInfo = data.get('customerInfo', [])

        print(FormInfo)
        for entry in classifys:
            classify = entry['classify']
            classify_map[classify].extend(entry['categories'])

        result = [{'classify': classify, 'categories': categories} for classify, categories in classify_map.items()]
        print(result)
        buffer = BytesIO()
        workbook = xlsxwriter.Workbook(buffer, {'in_memory': True})
        wsTotal = workbook.add_worksheet('Total')
        wsDetail = workbook.add_worksheet('Details')

        # Định dạng cột trang Total
        wsTotal.set_column('A:A', 7)
        wsTotal.set_column('B:B', 25)
        wsTotal.set_column('C:C', 5)
        wsTotal.set_column('D:D', 80)
        wsTotal.set_column('E:E', 20)
        wsTotal.set_column('F:F', 25)
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
#-----------------------------           BEGIN ĐỊNH DẠNG                -----------------------------#
        #ĐỊNH DẠNG SHEET TOTAL
        def create_format(workbook, bold=False, align='left', valign='top', font_color='black', font_size=11, border=0, border_color='#000000', bg_color='#FFFFFF', font_name='Times New Roman',num_format=None):
            format_options = {
                'bold': bold,
                'align': align,
                'valign': valign,
                'font_color': font_color,
                'font_size': font_size,
                'border': border,
                'border_color': border_color,
                'bg_color': bg_color,
                'font_name': font_name
            }

            if num_format:
                format_options['num_format'] = num_format
            
            return workbook.add_format(format_options)

        fontRed_18_center_bold = create_format(workbook, bold=True, align='center', valign='vcenter', font_color='red', font_size=18, border=1, border_color='#FFFFFF')
        fontBlack_14_center_bold = create_format(workbook, bold=True, align='center', valign='vcenter', font_color='black', font_size=14, border=1, border_color='#FFFFFF')
        fontRed_14_center_bold = create_format(workbook, bold=True, align='center', valign='vcenter', font_color='red', font_size=14, border=1, border_color='#FFFFFF')
        format_info = [fontRed_18_center_bold, fontBlack_14_center_bold, fontRed_14_center_bold, fontBlack_14_center_bold]  


        fontBlack_13_left_bold = create_format(workbook, bold=True, align='left', valign='vcenter', font_color='black', font_size=13, border=1, border_color='#FFFFFF')
        fontBlack_13_left = create_format(workbook, bold=False, align='left', valign='vcenter', font_color='black', font_size=13, border=1, border_color='#FFFFFF')

        

        fontWhite_13_center_bold = create_format(workbook, bold=True, align='center', valign='vcenter', font_color='white', font_size=13, border=1, bg_color='#3f3f3f')

        fontBlack_13_center_bold_bg = create_format(workbook, bold=True, align='center', valign='vcenter', font_color='black', font_size=13, border=1, bg_color='#d6dce4')
        fontBlack_13_left_bold_bg = create_format(workbook, bold=True, align='left', valign='vcenter', font_color='black', font_size=13, border=1, bg_color='#d6dce4')
        fontBlack_13_right_bold_bg = create_format(workbook, bold=True, align='right', valign='vcenter', font_color='black', font_size=13, border=1, bg_color='#d6dce4')
        fontBlack_13_right_bold_bg_num = create_format(workbook, bold=True, align='right', valign='vcenter', font_color='black', font_size=13, border=1, bg_color='#d6dce4', num_format='#,##0')

        fontBlack_13_center_nonbg = create_format(workbook, bold=False, align='center', valign='vcenter', font_color='black', font_size=13, border=1)
        fontBlack_13_left_nonbg = create_format(workbook, bold=False, align='left', valign='vcenter', font_color='black', font_size=13, border=1)
        fontBlack_13_right_nonbg = create_format(workbook, bold=False, align='right', valign='vcenter', font_color='black', font_size=13, border=1)
        fontBlack_13_right_nonbg_num = create_format(workbook, bold=False, align='right', valign='vcenter', font_color='black', font_size=13, border=1, num_format='#,##0')

        fontRed_13_center_bold_nonbg = create_format(workbook, bold=True, align='center', valign='vcenter', font_color='red', font_size=13, border=1)
        fontRed_13_left_bold_nonbg = create_format(workbook, bold=True, align='left', valign='vcenter', font_color='red', font_size=13, border=1)
        fontRed_13_right_bold_nonbg = create_format(workbook, bold=True, align='right', valign='vcenter', font_color='red', font_size=13, border=1)
        fontRed_13_right_bold_nonbg_num = create_format(workbook, bold=True, align='right', valign='vcenter', font_color='red', font_size=13, border=1, num_format='#,##0')
           
              
        


        # Định dạng với đường viền màu trắng
        infoCompany_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'border_color': '#FFFFFF'
        })
        
#-----------------------------           END ĐỊNH DẠNG                -----------------------------#


#-----------------------------           BEGIN CÁC BIẾN                 -----------------------------#
        # BIẾN SHEET TOTAL
        # Dữ liệu infoCompany
        infoCompany = [
            'CÔNG TY CỔ PHẦN KIM SƠN TIẾN',
            'Đ/c: Số 16 đường 35, P. An Khánh, TP. Thủ Đức, TP. HCM',
            'Hotline: 0913 699 545',
            'Website: www.kimsontien.com / www.fibarovn.com'
        ]


        data_pairs = [
            ('Kính gửi quý Khách hàng', FormInfo['customerName'], 'Đại diện kinh doanh', FormInfo['representativeName']),
            ('Số điện thoại', FormInfo['customerPhone'], 'Số điện thoại', FormInfo['representativePhone']),
            ('Email', FormInfo['customerEmail'], 'Email', FormInfo['representativeEmail']),
            ('Địa chỉ', FormInfo['customerAddress'], 'Số báo giá', FormInfo['quoteCode']),
        ]

        headersT = ['STT', 'DANH MỤC HỆ THỐNG', 'MÔ TẢ HỆ THỐNG']

        STT_Total = ['I', 'II', 'III', 'IV', 'V','VI', 'VII', 'VIII']
        STT_Total1 = ['A', 'B', 'C', 'D', 'E','F','G','H','J','K','L']

        value_sum_cellsT = []
        value_allsum_cellsT = []
        value_nhancong_cellsT = []

        # BIẾN SHEET DETAILS
        headers = ['STT', 'Mã Sản Phẩm', 'Hình ảnh', 'Thông tin sản phẩm', 'Thương hiệu', 'DVT', 'Số lượng', 'Đơn giá', 'Thành tiền', 'Ghi chú']

        rowid = 0
        colid = 0
#-----------------------------           END CÁC BIẾN                 -----------------------------#
        for classify in result:
            wsDetail.set_row(rowid, 40)
            wsDetail.merge_range(f'A{rowid + 1}:J{rowid + 1}', classify['classify'], fontRed_18_center_bold)
            rowid += 1

            for col_num, header in enumerate(headers):
                wsDetail.write(rowid, col_num, header, fontWhite_13_center_bold)
            rowid += 1

            sum_cells = []
            letter = 65  # ASCII 65 là 'A'

            for category_data in classify['categories']:
                wsDetail.set_row(rowid, 30)
                start_row = rowid + 1
                wsDetail.write(rowid, colid, chr(letter), fontBlack_13_center_bold_bg)
                wsDetail.merge_range(f'B{rowid + 1}:H{rowid + 1}', category_data['category'], fontBlack_13_left_bold_bg)
                wsDetail.write(rowid, colid + 9, '', fontBlack_13_right_bold_bg_num)

                end_row = start_row + len(category_data['products']) - 1
                sum_cell = f'I{rowid + 1}'
                wsDetail.write_formula(rowid, colid + 8, f'SUM(I{start_row+1}:I{end_row + 1})', fontBlack_13_right_bold_bg_num)

                sum_cells.append(sum_cell)

                #----------------------------   TOTAL    ---------------------------------------
                value_sum_cellsT.append(sum_cell)  # Add vào arr value_sum_cellsT cho việc hiển thị sheet TOTAL
                #----------------------------  END TOTAL    ---------------------------------------

                rowid += 1

                for index, product in enumerate(category_data['products']):
                    wsDetail.set_row(rowid, 120)
                    wsDetail.write(rowid, colid, index + 1, fontBlack_13_center_nonbg)
                    wsDetail.write(rowid, colid + 1, product['productName'], fontBlack_13_left_nonbg)

                    # Tải và chèn ảnh vào file Excel
                    image_url = product['productImageUrl']
                    image_data = requests.get(image_url).content
                    image_stream = BytesIO(image_data)
                    wsDetail.insert_image(rowid, colid + 2, image_url, {'image_data': image_stream, 'x_scale': 0.4, 'y_scale': 0.4, 'x_offset': 20, 'y_offset': 35})

                    wsDetail.write(rowid, colid + 3, product['productDescription'], fontBlack_13_left_nonbg)
                    wsDetail.write(rowid, colid + 4, product['productBrand'], fontBlack_13_center_nonbg)
                    wsDetail.write(rowid, colid + 5, product['productUnit'], fontBlack_13_center_nonbg)
                    wsDetail.write(rowid, colid + 6, product['productQuantity'], fontBlack_13_center_nonbg)
                    wsDetail.write(rowid, colid + 7, product['productPrice'], fontBlack_13_right_nonbg_num)
                    wsDetail.write(rowid, colid + 8, product['productTotal'], fontBlack_13_right_nonbg_num)
                    wsDetail.write(rowid, colid + 9, '', fontBlack_13_left_nonbg)

                    rowid += 1

                letter += 1

            wsDetail.set_row(rowid, 30)
            wsDetail.write(rowid, colid, chr(letter), fontBlack_13_center_bold_bg)
            wsDetail.merge_range(f'B{rowid + 1}:H{rowid + 1}', "NHÂN CÔNG LẮP ĐẶT PHẦN CỨNG, PHỤ KIỆN VÀ CẤU HÌNH HỆ THỐNG", fontBlack_13_left_bold_bg)
            wsDetail.write(rowid, colid + 9, '', fontBlack_13_right_bold_bg_num)

            if sum_cells:
                sum_formula = ','.join(sum_cells)
                labor_row = rowid
                if FormInfo['laborFee'] == '':
                    labor_formula = f"=ROUND(SUM({sum_formula})*13%,-3)"
                labor_formula = f"=ROUND(SUM({sum_formula})*{FormInfo['laborFee']}%,-3)"
                wsDetail.write_formula(rowid, colid + 8, labor_formula, fontBlack_13_right_bold_bg_num)
                #----------------------------   TOTAL    ---------------------------------------
                nhancong_cell = f'I{rowid+1}'
                value_sum_cellsT.append(nhancong_cell)
                value_nhancong_cellsT.append(nhancong_cell)
                #----------------------------  END TOTAL    ---------------------------------------
                rowid += 1

            if sum_cells:
                wsDetail.set_row(rowid, 30)
                final_sum_formula = f"=SUM(I{labor_row+1},{','.join(sum_cells)})"
                wsDetail.merge_range(f'A{rowid + 1}:H{rowid + 1}', "TỔNG TRƯỚC THUẾ", fontRed_13_center_bold_nonbg)
                wsDetail.write_formula(rowid, colid + 8, final_sum_formula, fontRed_13_right_bold_nonbg_num)
                wsDetail.write(rowid, colid + 9, '', fontRed_13_right_bold_nonbg_num)

                #----------------------------   TOTAL    ---------------------------------------
                sumall_cell = f'I{rowid+1}'
                value_allsum_cellsT.append(sumall_cell)
                #----------------------------  END TOTAL    ---------------------------------------
                rowid += 1

            rowid += 1

        

        # Áp dụng định dạng cho tất cả các ô từ A1 đến J8
        for row in range(8):  # Dòng từ 1 đến 8 (index 0 đến 7)
            for col in range(10):  # Cột từ A đến J (index 0 đến 9)
                wsTotal.write(row, col, '', infoCompany_format)

        # Tải và chèn ảnh vào file Excel
        imgTotal_url = static("/quoteapp/images/kstbaogia.png")
        full_imgTotal_url = request.build_absolute_uri(imgTotal_url)
        imgTotal_data = requests.get(full_imgTotal_url).content
        imgTotal_stream = BytesIO(imgTotal_data)
        wsTotal.insert_image(0, 1, image_url, {'image_data': imgTotal_stream, 'x_scale': 0.8, 'y_scale': 0.5, 'x_offset': 1, 'y_offset': 2})
        # Merge các ô và điền dữ liệu
        for i in range(4):
            wsDetail.set_row(i, 30)
            # Hợp nhất các ô E1:J1, E2:J2, E3:J3, E4:J4
            wsTotal.merge_range(i, 4, i, 9, infoCompany[i], format_info[i])
        
        row = 4
        for pair in data_pairs:
            
            wsTotal.merge_range(row, 0, row, 1, pair[0], fontBlack_13_left_bold)
            wsTotal.write(row, 2, ':', fontBlack_13_left_bold)
            wsTotal.merge_range(row, 3, row, 4, pair[1], fontBlack_13_left)
            wsTotal.write(row, 5, pair[2], fontBlack_13_left_bold)
            wsTotal.write(row, 6, ':', fontBlack_13_left_bold)
            wsTotal.merge_range(row, 7, row, 9, pair[3], fontBlack_13_left)
            
            # Chuyển sang dòng tiếp theo
            row += 1


        wsTotal.merge_range(8, 0, 8, 9, 'Công ty Cổ Phần Kim Sơn Tiến xin trân gửi đến anh Bảng Báo gía hệ thống Smarthome, Camera an ninh, thiết bị mạng và hệ thống âm thanh cho công trình, cụ thể như sau:', fontBlack_13_left_bold)
        wsTotal.merge_range(9, 0, 9, 9, 'BẢNG TỔNG HỢP GIÁ TRỊ HỆ THỐNG NHÀ THÔNG MINH FIBARO - KIM SƠN TIẾN', fontRed_18_center_bold)


        #----------------------------------------   TOTAL    ---------------------------------------

        # Thêm các tiêu đề vào các ô tương ứng
        wsTotal.write(10,0, headersT[0], fontWhite_13_center_bold)   # STT vào A11
        wsTotal.merge_range(10, 1, 10, 4, headersT[1], fontWhite_13_center_bold)  # DANH MỤC HỆ THỐNG vào B11:E11
        wsTotal.merge_range(10, 5, 10, 9, headersT[2], fontWhite_13_center_bold)  # MÔ TẢ HỆ THỐNG vào F11:J11

        rowTotal = 11
        indx = 0
        
        for classify in result:
            wsTotal.write(rowTotal,0, STT_Total[indx], fontBlack_13_center_bold_bg)   # STT vào A11
            wsTotal.merge_range(rowTotal, 1, rowTotal, 3, classify['classify'], fontBlack_13_left_bold_bg)
            wsTotal.write(rowTotal,4, f'=Details!{value_allsum_cellsT[indx]}', fontBlack_13_right_bold_bg_num)  
            wsTotal.merge_range(rowTotal, 5, rowTotal, 9, '', fontBlack_13_right_bold_bg_num) 
            indx +=1
            rowTotal+=1
            idx2 =0
            for category_data in classify['categories']:
                wsTotal.write(rowTotal,0, STT_Total1[idx2], fontBlack_13_center_nonbg)   # STT vào A11
                wsTotal.merge_range(rowTotal, 1, rowTotal, 3, category_data['category'], fontBlack_13_left_nonbg)
                wsTotal.write(rowTotal,4, f'=Details!{value_sum_cellsT[idx2]}', fontBlack_13_right_nonbg_num)  
                wsTotal.merge_range(rowTotal, 5, rowTotal, 9, '', fontBlack_13_left_nonbg) 
                idx2 +=1
                rowTotal+=1
            wsTotal.write(rowTotal,0, STT_Total1[idx2], fontBlack_13_center_nonbg)   # STT vào A11
            wsTotal.merge_range(rowTotal, 1, rowTotal, 3, "NHÂN CÔNG LẮP ĐẶT PHẦN CỨNG, PHỤ KIỆN VÀ CẤU HÌNH HỆ THỐNG", fontBlack_13_left_nonbg)
            wsTotal.write(rowTotal,4, f'=Details!{value_sum_cellsT[idx2]}', fontBlack_13_right_nonbg_num)  
            wsTotal.merge_range(rowTotal, 5, rowTotal, 9, '', fontBlack_13_right_nonbg_num) 
            rowTotal += 1  # Chuyển đến dòng tiếp theo
            idx2 +=1



        final_sum1 = f"=SUM({','.join([f'Details!{cell}' for cell in value_allsum_cellsT])})"

        wsTotal.merge_range(rowTotal, 0, rowTotal, 3, "TỔNG TRƯỚC THUẾ", fontRed_13_center_bold_nonbg)
        wsTotal.write(rowTotal,4, final_sum1, fontRed_13_right_bold_nonbg_num)
        wsTotal.merge_range(rowTotal, 5, rowTotal, 9, "", fontRed_13_right_bold_nonbg_num)

        arr_device = list(set(value_sum_cellsT) - set(value_nhancong_cellsT))
        numper_d = 10
        final_vat_device = f"=SUM({','.join([f'Details!{cell}' for cell in arr_device])})*{numper_d}%"
        vat_device_with_plus = final_vat_device.replace('=', '+', 1)
        wsTotal.merge_range(rowTotal+1, 0, rowTotal+1, 3, "VAT THIẾT BỊ 10%", fontRed_13_center_bold_nonbg)
        wsTotal.write(rowTotal+1,4, final_vat_device, fontRed_13_right_bold_nonbg_num)
        wsTotal.merge_range(rowTotal+1, 5, rowTotal+1, 9, "", fontRed_13_right_bold_nonbg_num)


        numper_p = 8
        final_vat_person = f"=SUM({','.join([f'Details!{cell}' for cell in arr_device])})*{numper_p}%"
        vat_person_with_plus = final_vat_person.replace('=', '+', 1)
        wsTotal.merge_range(rowTotal+2, 0, rowTotal+2, 3, "VAT NHÂN CÔNG 8%", fontRed_13_center_bold_nonbg)
        wsTotal.write(rowTotal+2,4, final_vat_person, fontRed_13_right_bold_nonbg_num)
        wsTotal.merge_range(rowTotal+2, 5, rowTotal+2, 9, "", fontRed_13_right_bold_nonbg_num)

        all_sum = final_sum1 + vat_device_with_plus + vat_person_with_plus
        wsTotal.merge_range(rowTotal+3, 0, rowTotal+3, 3, "TỔNG THANH TOÁN", fontRed_13_center_bold_nonbg)
        wsTotal.write(rowTotal+3,4, all_sum, fontRed_13_right_bold_nonbg_num)
        wsTotal.merge_range(rowTotal+3, 5, rowTotal+3, 9, "", fontRed_13_right_bold_nonbg_num)

        #----------------------------  END TOTAL    ---------------------------------------

        workbook.close()

        buffer.seek(0)
        response = HttpResponse(buffer, content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        response['Content-Disposition'] = 'attachment; filename=bao_gia.xlsx'
        return response

    return HttpResponse(status=400)
