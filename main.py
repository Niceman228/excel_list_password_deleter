import os
import zipfile
import re
import shutil

flag_of_file = False
file_name = ''

def zip_folder(folder_path, output_path):
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_STORED) as zipf:
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, folder_path)
                zipf.write(file_path, arcname)

current_directory = os.path.dirname(os.path.abspath(__file__))
file_list = os.listdir(current_directory)

for i in file_list:
    if str(i).count('xlsx') != 0:
        file_name = str(i)
        print('Файл: ' + file_name + ' обнаружен')
        flag_of_file = True


if flag_of_file == True:
    old_file_name = file_name
    file_name = file_name.replace('xlsx','zip')
    os.rename(current_directory+'/'+old_file_name, current_directory+'/'+file_name)

    os.mkdir(current_directory+'/'+file_name[:-4])

    zip_file = zipfile.ZipFile(current_directory+'/'+file_name, 'r')
    zip_file.extractall(current_directory+'/'+file_name[:-4])

    os.remove(current_directory+'/'+file_name)

    path_to_files = str(current_directory+'/'+file_name[:-4]+'/xl/worksheets')

    list_of_sheets = os.listdir(path_to_files)

    list_of_shit = [str(i) for i in list_of_sheets if str(i).count('sheet') != 0]

    for i in list_of_shit:
        with open(path_to_files + "/" + i, 'r') as file:
            xml_string = file.read()

            updated_xml_string = re.sub(r'<sheetProtection password="DEAA" sheet="1" objects="1" scenarios="1" />', '', xml_string)
        with open(path_to_files + '/' + i, "w", encoding="utf-8") as file:
            file.write(updated_xml_string)

    zip_folder(current_directory+'/'+file_name[:-4], old_file_name)

    shutil.rmtree(current_directory+'/'+file_name[:-4])
    print('Пароли с листов сняты')
else:
    print('Вы не добавили файл с расширением .xlsx')

