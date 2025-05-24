import requests
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import urllib3
from gspread_formatting import *
import json # Cần import json cho hàm mới

# Tắt cảnh báo InsecureRequestWarning
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# --- Khôi phục hàm get_all_products gốc ---
def get_all_products(site_url, consumer_key, consumer_secret):
    page = 1
    all_products = []
    
    while True:
        products_url = f"{site_url}/wp-json/wc/v3/products?per_page=100&page={page}"
        
        try:
            response = requests.get(
                products_url,
                auth=(consumer_key, consumer_secret),
                verify=False
            )
            
            if response.status_code == 200:
                products = response.json()
                if not products:
                    break
                    
                for product in products:
                    date_modified = product.get('date_modified', '')
                    if date_modified:
                        from datetime import datetime
                        date_obj = datetime.strptime(date_modified, "%Y-%m-%dT%H:%M:%S")
                        date_modified = date_obj.strftime("%d/%m/%Y %H:%M")
                    
                    modified_by = product.get('modified_by', '')
                    
                    product_info = {
                        'id': product.get('id', ''),
                        'name': product.get('name', ''),
                        'status': product.get('status', ''),
                        'stock_status': 'Còn hàng' if product.get('stock_status') == 'instock' else 'Hết hàng',
                        'date_modified': date_modified,
                        'modified_by': modified_by,
                        'permalink': product.get('permalink', ''),
                        'description': product.get('description', '')
                    }
                    all_products.append(product_info)
                
                print(f"Đã lấy {len(all_products)} sản phẩm...")
                page += 1
            else:
                print(f"Lỗi khi lấy danh sách sản phẩm: {response.text}")
                break
                
        except requests.exceptions.RequestException as e:
            print(f"Lỗi kết nối khi lấy danh sách sản phẩm: {e}")
            break
    
    return all_products

# --- Hàm mới để kiểm tra API sản phẩm đơn lẻ ---
def check_single_product_api(site_url, consumer_key, consumer_secret, product_id):
    product_url = f"{site_url}/wp-json/wc/v3/products/{product_id}"
    print(f"\n--- Đang kiểm tra API cho sản phẩm ID: {product_id} ---")
    try:
        response = requests.get(
            product_url,
            auth=(consumer_key, consumer_secret),
            verify=False
        )

        if response.status_code == 200:
            product_data = response.json()
            print("Yêu cầu thành công. Dữ liệu nhận được:")
            # In toàn bộ dữ liệu sản phẩm với định dạng dễ đọc
            print(json.dumps(product_data, indent=4))
            # Kiểm tra trực tiếp trường description trong dữ liệu này
            description = product_data.get('description', '')
            print(f"\nTrường 'description' có tồn tại: {'description' in product_data}")
            # Sửa lỗi cú pháp f-string ở dòng này
            print(f'100 ký tự đầu của description: {description[:100]}...') # Sử dụng dấu nháy đơn cho f-string
            # Sửa lỗi cú pháp f-string ở dòng này
            print(f'Description có chứa \'<table class="cauhinh"\': {\'<table class="cauhinh"\' in description}') # Sử dụng dấu nháy đơn cho f-string và thoát dấu nháy đơn trong chuỗi cố định

        else:
            print(f"Yêu cầu thất bại. Status code: {response.status_code}")
            print(f"Response text: {response.text}")

    except requests.exceptions.RequestException as e:
        print(f"Lỗi kết nối khi kiểm tra sản phẩm ID {product_id}: {e}")

    print("-----------------------------------------------\n")

def export_to_sheets(products, spreadsheet_url):
    # Thiết lập quyền truy cập Google Sheets
    scope = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']
    
    credentials = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
    client = gspread.authorize(credentials)
    
    # Mở spreadsheet
    spreadsheet = client.open_by_url(spreadsheet_url)
    
    # Tạo worksheet mới tên "MKCOM" hoặc sử dụng nếu đã tồn tại
    try:
        worksheet = spreadsheet.worksheet("MKCOM")
    except:
        worksheet = spreadsheet.add_worksheet(title="MKCOM", rows="300", cols="8") # Vẫn tạo 8 cột
    
    # Chuẩn bị dữ liệu với 8 cột
    data = [["Bài Post", "Trạng thái", "Hiển thị", "Số lượng", "Cập nhật lần cuối", "Tài khoản", "Link sửa", "Edit"]] # Header 8 cột

    # Thêm thông tin sản phẩm vào dữ liệu
    for product in products:
        # Tạo link sửa từ ID sản phẩm
        edit_link = f"https://minhkhoicomputer.com/wp-admin/post.php?post={product.get('id', '')}&action=edit" # Sử dụng .get() phòng trường hợp thiếu ID
        # Thay đổi hiển thị số lượng
        stock_status = "-" if product.get('stock_status') == 'Hết hàng' else product.get('stock_status', '') # Sử dụng .get()

        # Lấy trạng thái gốc từ API
        original_status = product.get('status', '')
        
        # Xử lý cột Trạng thái với icon
        status_icon = "-"
        if original_status == 'publish':
            status_icon = '👀'  # Icon con mắt cho trạng thái hiển thị
        elif original_status == 'pending':
            status_icon = '⏰'  # Icon đồng hồ cho trạng thái chờ duyệt
        elif original_status == 'private':
            status_icon = '🔒'  # Icon khóa cho trạng thái riêng tư
        elif original_status == 'draft':
            status_icon = '📝'  # Icon nháp cho trạng thái nháp
        
        # Xử lý cột Hiển thị (vẫn dùng văn bản)
        display_status = "-"
        if original_status == 'publish' or original_status == 'pending':
            display_status = 'Công khai'
        elif original_status == 'private':
            display_status = 'Riêng tư'
        # Thông tin 'Được bảo vệ bằng mật khẩu' không có trong trường 'status' của API WooCommerce.
        
        # Kết hợp tên sản phẩm và permalink vào một ô
        combined_post_info = f"{product.get('name', '')}\n{product.get('permalink', '')}"

        # Lấy description
        description = product.get('description', '')

        # Kiểm tra sự hiện diện của cả hai chuỗi
        has_cauhinh_table = description and '<table class="cauhinh"' in description
        has_notcauhinh_table = description and '<table class="notcauhinh"' in description

        # --- Logic mới cho cột Edit ---
        edited_status = '-' # Mặc định là '-'
        if has_cauhinh_table:
            edited_status = '✏️' # Nếu có cauhinh, dùng icon bút
        elif has_notcauhinh_table:
            edited_status = '✏️' # Nếu không có cauhinh nhưng có notcauhinh, dùng icon thông tin (bạn có thể đổi icon này)
        # --- Kết thúc logic mới ---

        data.append([
            combined_post_info, # Sử dụng thông tin kết hợp
            status_icon,  # Sử dụng icon cho cột Trạng thái
            display_status, # Sử dụng văn bản cho cột Hiển thị
            stock_status,
            product.get('date_modified', ''), # Sử dụng .get()
            product.get('modified_by', ''), # Sử dụng .get()
            edit_link,
            edited_status # Sử dụng trạng thái Edit với logic mới
        ])
    
    # Xóa dữ liệu cũ và cập nhật dữ liệu mới
    worksheet.clear()
    
    # Cập nhật dữ liệu theo từng batch (từ A đến H)
    batch_size = 100
    for i in range(0, len(data), batch_size):
        batch = data[i:i + batch_size]
        range_name = f'A{i+1}:H{i+len(batch)}' # Cập nhật phạm vi batch
        worksheet.update(range_name, batch)
        print(f"Đã cập nhật {i+1} đến {i+len(batch)} dòng")
    
    # Định dạng độ rộng cột
    header_format = CellFormat(
        backgroundColor=Color(0.2, 0.6, 0.86),
        textFormat=TextFormat(bold=True, fontFamily='Arial', fontSize=11, foregroundColor=Color(1,1,1)),
        padding=Padding(left=5, top=5, right=5, bottom=5)
    )
    format_cell_range(worksheet, 'A1:H1', header_format) # Áp dụng cho A1:H1

    wrap_format = CellFormat(
        wrapStrategy='WRAP',
        padding=Padding(left=5, top=5, right=5, bottom=5)
    )
    format_cell_range(worksheet, 'A2:A300', wrap_format) # Chỉ áp dụng cho cột A

    padding_format = CellFormat(
        padding=Padding(left=5, top=5, right=5, bottom=5)
    )
    format_cell_range(worksheet, 'A2:H300', padding_format) # Áp dụng cho A2:H300

    center_format = CellFormat(
        horizontalAlignment='CENTER',
    )
    format_cell_range(worksheet, 'B1:F300', center_format) # Căn giữa B đến F
    format_cell_range(worksheet, 'H1:H300', center_format) # Căn giữa H (Edit)

    left_format = CellFormat(
        horizontalAlignment='LEFT',
    )
    format_cell_range(worksheet, 'A1:A300', left_format) # Căn trái A
    format_cell_range(worksheet, 'G1:G300', left_format) # Căn trái G

    middle_format = CellFormat(
        verticalAlignment='MIDDLE'
    )
    format_cell_range(worksheet, 'A1:H300', middle_format) # Căn giữa dọc A đến H

    worksheet.freeze(rows=1)
    
    # Điều chỉnh độ rộng cột (Thêm cột H)
    set_column_width(worksheet, 'A', 800) # Bài Post (tên + link)
    set_column_width(worksheet, 'B', 90)  # Trạng thái
    set_column_width(worksheet, 'C', 90)  # Hiển thị
    set_column_width(worksheet, 'D', 90)  # Số lượng
    set_column_width(worksheet, 'E', 150) # Cập nhật lần cuối
    set_column_width(worksheet, 'F', 100) # Tài khoản
    set_column_width(worksheet, 'G', 200) # Link sửa
    set_column_width(worksheet, 'H', 50)  # Edit

    print(f"Đã xuất {len(products)} sản phẩm lên Google Sheet!")


def main():
    # Thông tin WordPress
    site_url = "https://minhkhoicomputer.com"
    consumer_key = "ck_7ab1e2ba831d6a6f35bd4d66efdb431aa38ad067"
    consumer_secret = "cs_220abb9c658e837cd10c9d4eb268dbbdda52909f"
    
    # URL của Google Sheet
    spreadsheet_url = "https://docs.google.com/spreadsheets/d/11rK_Z1g4q8E0monnd-AAtBizT5lcbhD8ggMX8GiLgx4/edit?gid=0"
    
    # --- Comment hoặc xóa dòng gọi check_single_product_api sau khi xác nhận đã fix ---
    # test_product_id = 2087
    # check_single_product_api(site_url, consumer_key, consumer_secret, test_product_id)

    # Lấy danh sách sản phẩm (sử dụng hàm Edit)
    print("Đang lấy danh sách sản phẩm...")
    products = get_all_products(site_url, consumer_key, consumer_secret)
    
    # Xuất lên Google Sheet
    print("Đang xuất dữ liệu lên Google Sheet...")
    export_to_sheets(products, spreadsheet_url)

if __name__ == "__main__":
    main()
