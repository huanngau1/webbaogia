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

        for entry in classifys:
            classify = entry['classify']
            classify_map[classify].extend(entry['categories'])

        result = [{'classify': classify, 'categories': categories} for classify, categories in classify_map.items()]
        print(result)
        buffer = BytesIO()
        workbook = xlsxwriter.Workbook(buffer, {'in_memory': True})
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
            'bg_color': '#9FC5E8'
        })

        header2_format = workbook.add_format({
            'bold': True,
            'bg_color': '#F4CCCC',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })

        number_format = workbook.add_format({
            'border': 1,
            'num_format': '#,##0.00',
            'bg_color': '#FFFFFF'
        })

        sum_format = workbook.add_format({
            'bold': True,
            'bg_color': '#D9EAD3',
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'num_format': '#,##0.00'
        })

        labor_format = workbook.add_format({
            'bold': True,
            'bg_color': '#FFD966',
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'num_format': '#,##0.00'
        })

        product_format = workbook.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#FFFFFF',
        })

        productInfo_format = workbook.add_format({
            'border': 1,
            'bg_color': '#FFFFFF',
            'text_wrap': True
        })

        headers = ['STT', 'Mã Sản Phẩm', 'Hình ảnh', 'Thông tin sản phẩm', 'Thương hiệu', 'DVT', 'Số lượng', 'Đơn giá', 'Thành tiền', 'Ghi chú']

        rowid = 0
        colid = 0

        for classify in result:
            wsDetail.set_row(rowid, 40)
            wsDetail.merge_range(f'A{rowid + 1}:J{rowid + 1}', classify['classify'], header1_format)
            rowid += 1

            for col_num, header in enumerate(headers):
                wsDetail.write(rowid, col_num, header, header2_format)
            rowid += 1

            sum_cells = []
            letter = 65  # ASCII 65 là 'A'

            for category_data in classify['categories']:
                wsDetail.set_row(rowid, 30)
                start_row = rowid + 1
                wsDetail.write(rowid, colid, chr(letter), labor_format)
                wsDetail.merge_range(f'B{rowid + 1}:H{rowid + 1}', category_data['category'], labor_format)
                wsDetail.write(rowid, colid + 9, '', labor_format)

                end_row = start_row + len(category_data['products']) - 1
                sum_cell = f'I{rowid + 1}'
                wsDetail.write_formula(rowid, colid + 8, f'SUM(I{start_row+1}:I{end_row + 1})', labor_format)

                sum_cells.append(sum_cell)
                rowid += 1

                for index, product in enumerate(category_data['products']):
                    wsDetail.set_row(rowid, 120)
                    wsDetail.write(rowid, colid, index + 1, product_format)
                    wsDetail.write(rowid, colid + 1, product['productName'], product_format)

                    # Tải và chèn ảnh vào file Excel
                    image_url = product['productImageUrl']
                    image_data = requests.get(image_url).content
                    image_stream = BytesIO(image_data)
                    wsDetail.insert_image(rowid, colid + 2, image_url, {'image_data': image_stream, 'x_scale': 0.2, 'y_scale': 0.2, 'x_offset': 15, 'y_offset': 25})

                    wsDetail.write(rowid, colid + 3, product['productDescription'], productInfo_format)
                    wsDetail.write(rowid, colid + 4, product['productBrand'], product_format)
                    wsDetail.write(rowid, colid + 5, product['productUnit'], product_format)
                    wsDetail.write(rowid, colid + 6, product['productQuantity'], product_format)
                    wsDetail.write(rowid, colid + 7, product['productPrice'], number_format)
                    wsDetail.write(rowid, colid + 8, product['productTotal'], number_format)
                    wsDetail.write(rowid, colid + 9, '', product_format)

                    rowid += 1

                letter += 1

            wsDetail.set_row(rowid, 30)
            wsDetail.write(rowid, colid, chr(letter), labor_format)
            wsDetail.merge_range(f'B{rowid + 1}:H{rowid + 1}', "NHÂN CÔNG LẮP ĐẶT, PHỤ KIỆN VÀ CẤU HÌNH HỆ THỐNG", labor_format)
            wsDetail.write(rowid, colid + 9, '', labor_format)

            if sum_cells:
                sum_formula = ','.join(sum_cells)
                labor_row = rowid
                labor_formula = f"=ROUND(SUM({sum_formula})*13%,-3)"
                wsDetail.write_formula(rowid, colid + 8, labor_formula, labor_format)
                rowid += 1

            if sum_cells:
                wsDetail.set_row(rowid, 30)
                final_sum_formula = f"=SUM(I{labor_row+1},{','.join(sum_cells)})"
                wsDetail.merge_range(f'A{rowid + 1}:H{rowid + 1}', "TỔNG CỘNG", sum_format)
                wsDetail.write_formula(rowid, colid + 8, final_sum_formula, sum_format)
                wsDetail.write(rowid, colid + 9, '', sum_format)
                rowid += 1

            rowid += 1

        workbook.close()

        buffer.seek(0)
        response = HttpResponse(buffer, content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        response['Content-Disposition'] = 'attachment; filename=bao_gia.xlsx'
        return response

    return HttpResponse(status=400)
