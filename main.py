# The program reads certain data from pdf files. Also pre-unpacks from zip archives

# pdf files. The results are recorded in an Excel file.
# Программа производит чтение определенных данных из pdf файлов. Также предварительно распаковывает из zip архивов

# pdf файлы. Результаты записывает в Excel Файл.

import os

import PyPDF2

import zipfile

import openpyxl

import time

 

start = time.time()

 

 

 

# ------------Вытаскивает весь текст из PDF файла--------------

def find_text_in_pdf(pdf_file_path):

    with open(pdf_file_path, 'rb') as file:

        reader = PyPDF2.PdfReader(file)

        text = ""

        for page_num in range(len(reader.pages)):

            page = reader.pages[page_num]

            text += page.extract_text()

        return text

 

 

# --------Читает слова или слово после указанного ключевого слова

def extract_value_after_keyword(text, keyword):

    #print(keyword)

 

    for i, word in enumerate(text):

        len_word = len(word.split())

        #if word.find('Площадь:')>=0 or word.find('Площадь,')>=0:

        #    print(len_word, word)

 

        # --------------------------

 

        if word == keyword:

            if len_word == 3 and (word.find('основание государственной регистрации:') >= 0):

                k = i

                foot = ''

                while text[k + 2].split()[1] not in ('Договоры', 'Заявленные', 'Сведения', 'сведения', 'полное'):

                    foot = foot + ' ' + text[k + 2].split()[1]

                    k += 1

                return foot

            # --определяем вид обременения

            if len_word == 1 and (word.find("вид:") >= 0):

                #print(word)

                k = i

                view = ''

                while text[k + 1] != 'дата':

                    view = view + ' ' + text[k + 1]

                    k += 1

                return view

            # --Определяем кадастровую стоимость

            if len_word == 2 and (keyword == 'Кадастровая стоимость,' or keyword == 'Кадастровая стоимость:'):

                return text[i+2].split()[1]

            # -- Определяем местоположение или адрес

            if len_word == 1 and (keyword == 'Местоположение:' or keyword == 'Адрес:'):

                k = i

                address = ''

                while (text[k+1] != 'Площадь:') and (text[k+1] != 'Площадь,'):

                    address = address + ' ' + text[k+1]

                    #print (text[k+1], address)

                    k += 1

                return address

            # ---------------

            if len_word == 1:

                #print(word, text[i+1], text[i+2], text[i+3])

                return text[i+1]

            # -----------------

            if len_word == 2:

                return text[i+1].split()[1]

            # ---------------------

            if len_word == 3 and (keyword == 'дата государственной регистрации:'

                                  or keyword == 'номер государственной регистрации:'):

                return text[i+2].split()[1]

 

            # ---------------

            # -определяем основание ограничения

            if len_word == 7:

                k = i

                face = ''

                while (text[k + 10].split()[0] != 'основание') and (text[k + 10].split()[0] != 'сведения') and \

                        (text[k + 10].split()[0] != 'полное'):

                    face = face + ' ' + text[k + 10].split()[0]

                    #print(text[k + 1].split()[0])

                    k += 1

                return face

            # -------------------------

            if len_word == 4:

                return text[i+1].split()[3]

            # ---Определяем есть ограничение прав и обременение

            if len_word == 6 and (keyword == '3.Ограничение прав и обременение объекта недвижимости:'

                or keyword == 'Ограничение прав и обременение объекта недвижимости:'

                                  or keyword == 'Ограничение права и обременение объекта недвижимости:'):

                sp = text[i+6].split()

                if sp[0] == 'не':

                    return sp[0] + ' ' + sp[1]

                if sp[1] == 'вид:':

                    return sp[2] + ' ' + sp[3] + ' ' + sp[4] + ' ' + sp[5]

                return sp[1] + ' ' + sp[2] + ' ' + sp[3] + ' ' + sp[4]

            # ----определяем срок на которое установлено ограничение прав

            if len_word == 6 and (keyword == 'срок, на который установлено ограничение прав' or

            keyword == 'срок, на который установлены ограничение прав' or

            keyword == 'срок, на который установлено ограничение права'):

                #print(word, keyword)

                k = i

                term = ''

                while text[k + 6].split()[3] != 'лицо,' and ((k+6)<len(text)-1):

                    #print(text[k+6].split()[3])

                    term = term + ' ' + text[k + 6].split()[3]

                    k += 1

                return term

            # ------------------------------------

            if len_word == 14 and (keyword == 'Сведения о том, что \

                земельный участок полностью расположен в границах зоны с особыми условиями'):

                k = i

                zone = 'участок в зоне'

                #while text[k + 14].split()[3] != 'лицо,':

                #    zone = zone + ' ' + text[k + 6].split()[3]

                #    k += 1

                return zone

            # ------------------------------------

 

    return None

 

 

# ------Для анализа делит текст на группы по 1, 2, 3, chunk_size слова для сравнения с ключевым выражением -----

def generate_all_word_combinations(text, chunk_size):

    words = text.split()

    num_words = len(words)

    # Если количество слов в тексте меньше чем chunk_size, нет возможности создать комбинации

    if num_words < chunk_size:

        return []

    result = []

    for i in range(num_words - chunk_size + 1):

        chunk = ' '.join(words[i:i + chunk_size])

        result.append(chunk)

    return result

 

 

# -----Распаковывает  zip файлы--------

def extract_pdfs_from_zip(zip_file_path, output_directory):

    with zipfile.ZipFile(zip_file_path, 'r') as zip_file:

        count =0

        for file_info in zip_file.infolist():

            if file_info.filename.lower().endswith('.pdf'):

                count += 1

                zip_file.extract(file_info, output_directory)

                parent_zip = os.path.basename(zip_file_path)[0:-4]

                new_name_path = os.path.dirname(output_directory) + '/files/' + parent_zip + str(count) + '.pdf'

                old_name = output_directory + '/' +file_info.filename

                if not os.path.isfile(new_name_path):

                    os.rename(old_name, new_name_path)

 

# ----- Записывает результаты в EXCEL Файл

def write_to_excel(extracted_data, output_file):

    wb = openpyxl.Workbook()

    ws = wb.active

    headers = ["Название файла", "Кадастровый номер", "Дата присвоения кадастрового номера", "Местоположение",

               "Площадь", "Кадастровая стоимость", "Ограничение прав и обременение объектов недвижимости",

               "Вид", "Дата государственной регистрации",

               "Номер государственной регистрации", "Срок, на который установлено ограничение прав и обременение объекта недвижимости",

               "Лицо, в пользу которого установлено ограничение прав и обременение объекта недвижимости",

               "Основание государственной регистрации",

               "Сведения о том, что земельный участок полностью расположен в границах зоны с особыми условиями \

                использования территории, территории объекта культурного наследия, публичного сервитута"]

    ws.append(headers)

    for data in extracted_data:

        # Applying remove_non_numeric_suffix to specific columns

        data[1] = remove_non_numeric_suffix(data[1])  # Кадастровый номер

        data[2] = remove_non_numeric_suffix(data[2])  # Дата присвоения кадастрового номера

        data[4] = remove_non_numeric_suffix(data[4])  # Площадь

        data[5] = remove_non_numeric_suffix(data[5])  # Кадастровая стоимость

        #print(data[10])

        data[10] = remove_word_suffix(data[10])  #

        data[11] = remove_word_suffix(data[11])  #

        #print(data[10])

 

        ws.append(data)

    wb.save(output_file)

 

 

# ---Удаляет лишние буквы в конце если есть----

def remove_non_numeric_suffix(input_string):

    if input_string:

        while not input_string[-1].isdigit() and len(input_string)>2:

            input_string = input_string[:-1]

 

#        print(input_string[1:13])

        #if input_string.find('недвижимости')>=0:

            #rint(input_string, len(input_string))

        #    input_string = input_string[(input_string.find('недвижимости')+13):]

            #print(input_string)

        #for i, word in enumerate(input_string):

        #    print(i, word)

    return input_string

 

# ---Удаляет слово недвижимости если есть----

def remove_word_suffix(input_string):

    if input_string:

        if input_string.find('недвижимости')>=0:

            #print(input_string, len(input_string))

            input_string = input_string[(input_string.find('недвижимости')+13):]

            #print(input_string)

        #for i, word in enumerate(input_string):

        #    print(i, word)

    return input_string

 

if __name__ == "__main__":

    extracted_data = []

    directory_path = "U:/Documents/PycharmProjects/ParsingPDF/files"  # Укажите путь к каталогу с PDF-файлами

    zip_files = [file for file in os.listdir(directory_path) if file.lower().endswith('.zip')]

 

    if zip_files:

        for zip_file in zip_files:

            zip_file_path = os.path.join(directory_path, zip_file)

            extract_pdfs_from_zip(zip_file_path, directory_path)

            os.remove(zip_file_path)

 

    pdf_files = [file for file in os.listdir(directory_path) if file.lower().endswith('.pdf')]

 

    if not pdf_files:

        print("В указанном каталоге нет PDF-файлов.")

    else:

        for pdf_file in pdf_files:

            pdf_file_path = os.path.join(directory_path, pdf_file)

            text = find_text_in_pdf(pdf_file_path)

            cadastral_number = extract_value_after_keyword(generate_all_word_combinations(text, 2), "Кадастровый номер:")

            date_of_assignment = extract_value_after_keyword(generate_all_word_combinations(text, 4), "Дата присвоения кадастрового номера:")

            location = extract_value_after_keyword(generate_all_word_combinations(text, 1), "Местоположение:")

            location2 = extract_value_after_keyword(generate_all_word_combinations(text, 1), "Адрес:")

            area = extract_value_after_keyword(generate_all_word_combinations(text, 1), "Площадь:")

            area2 = extract_value_after_keyword(generate_all_word_combinations(text, 1), "Площадь,")

            cadastral_value = extract_value_after_keyword(generate_all_word_combinations(text, 2), "Кадастровая стоимость,")

            cadastral_value2 = extract_value_after_keyword(generate_all_word_combinations(text, 2), "Кадастровая стоимость:")

            limitation = extract_value_after_keyword(generate_all_word_combinations(text, 6), "3.Ограничение прав и обременение объекта недвижимости:")

            limitation2 = extract_value_after_keyword(generate_all_word_combinations(text, 6), "Ограничение прав и обременение объекта недвижимости:")

            limitation3 = extract_value_after_keyword(generate_all_word_combinations(text, 6), "Ограничение права и обременение объекта недвижимости:")

            if limitation != 'не зарегистрировано' or limitation2 != 'не зарегистрировано':

                view_limit = extract_value_after_keyword(generate_all_word_combinations(text, 1), "вид:")

                date_gos_reg = extract_value_after_keyword(generate_all_word_combinations(text, 3), "дата государственной регистрации:")

                num_gos_reg = extract_value_after_keyword(generate_all_word_combinations(text, 3), "номер государственной регистрации:")

                term_gos_reg = extract_value_after_keyword(generate_all_word_combinations(text, 6), "срок, на который установлено ограничение прав")

                term_gos_reg2 = extract_value_after_keyword(generate_all_word_combinations(text, 6), "срок, на который установлены ограничение прав")

                term_gos_reg3 = extract_value_after_keyword(generate_all_word_combinations(text, 6), "срок, на который установлено ограничение права")

                face_gos_reg = extract_value_after_keyword(generate_all_word_combinations(text, 7), "лицо, в пользу которого установлено ограничение прав")

                face_gos_reg2 = extract_value_after_keyword(generate_all_word_combinations(text, 7), "лицо, в пользу которого установлены ограничение прав")

                foot_gos_reg = extract_value_after_keyword(generate_all_word_combinations(text, 3), "основание государственной регистрации:")

                zone = extract_value_after_keyword(generate_all_word_combinations(text, 14),

                "Сведения о том, что земельный участок полностью расположен в границах зоны с особыми условиями")

 

                extracted_data.append([pdf_file, cadastral_number, date_of_assignment, location or location2, area or area2,

                                   cadastral_value or cadastral_value2, limitation or limitation2 or limitation3, view_limit, date_gos_reg, num_gos_reg,

                                       term_gos_reg or term_gos_reg2 or term_gos_reg3, face_gos_reg or face_gos_reg2, foot_gos_reg, zone])

            else:

                extracted_data.append(

                    [pdf_file, cadastral_number, date_of_assignment, location or location2, area or area2,

                     cadastral_value or cadastral_value2, limitation or limitation2 or limitation3])

            output_excel_file = "output.xlsx"

            write_to_excel(extracted_data, output_excel_file)

            print(f"Data has been written to '{output_excel_file}'.")

            print(f"Из файла '{pdf_file}':")

            print(f"Кадастровый номер: {cadastral_number}")

            print(f"Дата присвоения кадастрового номера: {date_of_assignment}")

            print(f"Местоположение: {location}")

            print(f"Адрес: {location2}")

            print(f"Площадь: {area}")

            print(f"Площадь, {area2}")

            print(f"Кадастровая стоимость: {cadastral_value}")

            print(f"Кадастровая стоимость: {cadastral_value2}")

            print(f"Ограничение прав и обременение объекта недвижимости: {limitation}")

            print(f"Ограничение прав и обременение объекта недвижимости: {limitation2}")

            print(f"Ограничение прав и обременение объекта недвижимости: {limitation3}")

            print(f"Вид: {view_limit}")

            print(f"Дата государственной регистрации: {date_gos_reg}")

            print(f"Номер государственной регистрации: {num_gos_reg}")

            print(f"Срок, на который установлено ограничение прав и обременение объекта недвижимости: {term_gos_reg}")

            print(f"Срок, на который установлены ограничение прав и обременение объекта недвижимости: {term_gos_reg2}")

            print(f"Срок, на который установлены ограничение прав и обременение объекта недвижимости: {term_gos_reg3}")

            print(f"Лицо, в пользу которого установлено ограничение прав и обременение объекта недвижимости: {face_gos_reg}")

            print(f"Лицо, в пользу которого установлены ограничение прав и обременение объекта недвижимости: {face_gos_reg2}")

            print(f"Основание государственной регистрации: {foot_gos_reg}")

            print(f"Сведения о том, что земельный участок полностью расположен в границах зоны с особыми условиями: {zone}")

            print()

 

            end = time.time()

            print(end-start)
