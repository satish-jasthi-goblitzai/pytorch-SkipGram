import os
os.environ["OMP_THREAD_LIMIT"] = '1'

import fitz
from pytesseract import*
import concurrent.futures
import platform
import glob


MAX_WORKERS = 1 if os.cpu_count() == 1 or os.cpu_count() == 2 else ((os.cpu_count() / 2) + 1)

if platform.system() == 'Windows':
    import win32process

    procAff, sysAff = win32process.GetProcessAffinityMask(win32process.GetCurrentProcess())
    mask_set = [ones for ones in bin(procAff)[2:] if ones == '1'] 
    MAX_WORKERS = len(mask_set)
elif platform.system() == 'Linux':
    MAX_WORKERS = len(os.sched_getaffinity(0))
else:
    print('Unknown OS')


input_foler=r'D:\Work\R&D\CTFiles'
img_folder=r'D:\Work\R&D\CTFiles\img'
output_folder=r'D:\Work\R&D\CTFiles\txt'

job_files = glob.glob(os.path.join(input_foler, '*.pdf'))

if platform.system() == 'Windows':
    pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'


def process_and_extract(img_info):
    text = ''

    try:
        img=img_info[0]
        if img is None:
            return text
        print(f'OCR Start : page_num {img_info[1]}', flush=True)
        text = pytesseract.image_to_string(img)
        print(f'OCR End : page_num {img_info[1]}', flush=True)
    except Exception as e:
        print(e) 
        
    return text


dpi = 200
dpi_matrix = fitz.Matrix(dpi / 72, dpi / 72)


for job_file in job_files:
    head, file_name = os.path.split(job_file)
    img_file_name = file_name.split('.')[0]

    proc_data_list=[]

    with fitz.open(job_file) as pdf_file:
        print(f'Total pages: {len(pdf_file)}')
        for page in pdf_file:
            # Get all pixel of page
            page_pixel = page.get_pixmap(matrix = dpi_matrix)
            page_pixel.set_dpi(dpi, dpi)
            #images.append(page_pixel)
            img_file = f'{img_file_name}_{page.number+1}.png'
            img_path = os.path.join(img_folder, img_file)
            page_pixel.save(img_path)
            proc_data = [img_path, page.number + 1, file_name]
            proc_data_list.append(proc_data)

            del page_pixel
            del page
            del proc_data

    if 'pdf_file' in locals() or 'pdf_file' in globals():
        del pdf_file


    fitz.TOOLS.store_shrink(100)

    print('Convert to Images completed')

    ocr_output=[]

    try:     
        with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
            for process_data, output_data in zip(proc_data_list, executor.map(process_and_extract, proc_data_list)):
                ocr_output.append([process_data[1],output_data])
                del process_data

    except Exception as ex:
        print('error', ex)

    print(f'OCR completed for file {file_name}')

    ocr_output.sort(key=lambda x: x[0])

    with open(os.path.join(output_folder,img_file_name+'.txt'), "w") as text_file:
        for page in ocr_output:
            text_file.write(page[1])
            text_file.write('\n')

del dpi_matrix

print('Completed')


