from PIL import Image
import io
from xlsxwriter import Workbook

#input: file and crop coordinates | output: series, serial_num and img_snippet
def process_image(reader, file, left, top, right, bottom):

    crop_coords = (left, top, right, bottom)

    img = Image.open(file)
    
    cropped_img = img.crop(crop_coords)

    img_byte_arr = io.BytesIO()
    cropped_img.save(img_byte_arr, format='JPEG')
    cropped_img_bytes = img_byte_arr.getvalue()
    result = reader.readtext(cropped_img_bytes, allowlist='ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789')[0]
    processed_text = process_text(result[1])

    return {"image_file": file.filename, "series": processed_text.get('series'), "serial": processed_text.get('serial'), "confidence_level": result[2], "image": cropped_img}

def process_text(input_text):
    series = input_text[:-6]
    series.replace('4', 'A')
    series.replace('6', 'G')
    series.replace('0','O')
    series.replace('2','Z')

    serial = input_text[-6:]
    serial.replace('O','0')
    serial.replace('l','1')
    serial.replace('i','1')
    serial.replace('I','1')

    return {"series":series, "serial": serial}

def create_excel_file(results):
    excel_bytes_array = io.BytesIO()
    wb = Workbook(excel_bytes_array, {'in_memory': True})
    
    ws = wb.add_worksheet()

    # Add headers to the Excel file
    ws.write(0,0,'Image File Name')
    ws.write(0,1,'Confidence Level')
    ws.write(0,2,'Series')
    ws.write(0,3,'Serial')
    ws.write(0,4,'PB35_Listing')
    ws.write(0,5,'Image')

    for row, result in enumerate(results, start=1):
        ws.write(row, 0, result.get('image_file'))
        ws.write(row, 1, result.get('confidence_level'))
        ws.write(row, 2, result.get('series'))
        ws.write(row, 3, result.get('serial'))
        ws.write(row, 4, result.get('series') + result.get('serial'))

        # Add the image to the Excel file
        img_byte_arr = io.BytesIO()
        image = result.get('image')
        image.save(img_byte_arr, format='JPEG')
        img_width, img_height = image.size

        img_byte_arr.seek(0)

        ws.insert_image(f'F{row+1}', 'image.jpg', {
            'image_data': img_byte_arr,
            'x_scale': 64/ img_width,
            'y_scale': 20/ img_height
            })
    wb.close()
    excel_bytes_array.seek(0)
    return excel_bytes_array