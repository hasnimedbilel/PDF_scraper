import pandas as pd
import numpy as np
import tika
tika.initVM()
from tika import parser
from pdf2docx import parse
import io, sys, os
import re
from bs4 import BeautifulSoup


def convert_FirstPagePdf_to_docx(pdf_input_path, docx_output_path):
    # This method converts the first page of the input PDF into Docx
    # returns a response Flag : True --> Successful conversion
    parse(pdf_input_path, docx_output_path, start=0, end=1)    
    return True

def convert_firstPageDocx_to_txt(docx_input_path, txt_output_path):
    tika_parsed_content = parser.from_file(docx_input_path, xmlContent=True)
    body = tika_parsed_content['content'].split('<body>')[1].split('</body>')[0]

    with io.open(txt_output_path, 'w', encoding='utf8') as txt_file:
        txt_file.write(body.strip())

def remove_tags_from_string(raw_html):
    cleanr = re.compile('<.*?>')
    cleantext = re.sub(cleanr, '', raw_html)
    return cleantext.strip()

def get_paragraph_posology(paragraph):
    # This function accepts a pragraph and returns two lists : 
    # Underlined Posology & Non underlined Posology
    
    # The input should be a paragraph : example --> list_tags[0]
    list_posology_up = []
    list_posology_down = []
    
    one_posology = {}
    
    one_paragraph_items = list(paragraph.find_all())
    i = 0
    while i<len(one_paragraph_items):
        temp_index_posology_up = 0
        temp_txt = str(one_paragraph_items[i])
        if temp_txt.startswith("<b><u>"):
            temp_posology = str(one_paragraph_items[i])
            temp_list_posology_down = []
            list_posology_up.append(remove_tags_from_string(temp_posology))
            temp_index_posology_up = i+1
            temp_posology_down = str(one_paragraph_items[temp_index_posology_up])
            while (("<b><u>" not in temp_posology_down) and 
                   (temp_index_posology_up < len(one_paragraph_items))):
                if temp_posology_down.startswith("<b>"):
                    list_posology_down.append(remove_tags_from_string(temp_posology_down))
                    temp_list_posology_down.append(remove_tags_from_string(temp_posology_down))
                temp_index_posology_up += 1
                if temp_index_posology_up < len(one_paragraph_items):
                    temp_posology_down = str(one_paragraph_items[temp_index_posology_up])
                
            one_posology.update({remove_tags_from_string(temp_posology) : temp_list_posology_down})
            
            i = temp_index_posology_up
        else:
            i += 1

    return list_posology_up, list_posology_down, one_posology

def get_page_posology(one_page_list_tags):
    all_posology_dict = {}
    for i in one_page_list_tags:
        _, __, temp_posology_dict = get_paragraph_posology(i)
        for pos in temp_posology_dict.items():
            all_posology_dict.update({pos[0]:pos[1]})
            
    return all_posology_dict

def get_applicant(one_page_list_tags):
    for i in one_page_list_tags:
        if i.getText().strip().lower().startswith("laboratoire"):
            temp_applicant = i.getText().strip()
            temp_applicant = temp_applicant.replace("Laboratoire", "").strip()
            temp_applicant = temp_applicant.replace("laboratoire", "").strip()
            return temp_applicant

def get_french_dates(text):
    # this method accepts a string
    # returns a string : date --> if it matches :
    # two numbers + character name of month + 4 numbers.
    # if not match : returns an mepty string.
    
    date_pattern = r'\d{1,2}? (?i)(Jan(?:vier)?|Fév(?:rier)?|Mar(?:s)?|Avr(?:il)?|Mai|Jui(?:n)?|Juil(?:let)?|Aoû(?:t)?|Sep(?:tembre)?|Oct(?:obre)?|Nov(?:embre)?|Déc(?:embre)?)? \d{4}' 
    match = re.search(date_pattern, text)
    if match is not None:
        return match.group(0)
    else:
        return False

def get_date(one_page_list_tags):
    # this method returns the first occurence of the date
    for i in one_page_list_tags:
        if get_french_dates(i.getText()):
            return get_french_dates(i.getText())

def get_one_doc_csv(date_string, posology_dict, applicant_string):
    posology_strings = []
    for i in posology_dict.items():
        one_posology_string = str(i[0])
        for j in i[1]:
            one_posology_string = one_posology_string + "\\n" + str(j)
        posology_strings.append(one_posology_string)

    out_dataframe = pd.DataFrame({
        'Date' : pd.Series(date_string),
        'Posologies' : pd.Series(posology_strings),
        'Applicant' : pd.Series(applicant_string)
    })
    
    return out_dataframe




if __name__ == "__main__":
    
    pdf_dir = sys.argv[1]
    output_dir = sys.argv[2]
    for root, dirs, files in os.walk(pdf_dir):
        print("root : {}".format(root))
        print("dirs : {}".format(dirs))
        print("paths : {}".format(files))
        for file in files:
            path_to_pdf = os.path.join(root, file)
            [stem, ext] = os.path.splitext(path_to_pdf)
            filename_without_ext = stem.split("\\")[-1]
            if ext == '.pdf':
                print("Processing " + path_to_pdf + "...")

                docx_path = stem + ".docx"
                txt_path = stem + ".txt"
                conversion_to_docx_response = convert_FirstPagePdf_to_docx(path_to_pdf, docx_path)
                convert_firstPageDocx_to_txt(docx_path, txt_path)

                with open(txt_path, "r", encoding="utf8") as file:
                    tika_html_file = file.read()
                bs_file = BeautifulSoup(tika_html_file, 'html.parser')
                list_tags = bs_file.find_all("p")

                posology_dict = get_page_posology(list_tags)
                applicant = get_applicant(list_tags)
                date = get_date(list_tags)
                df = get_one_doc_csv(date_string=date, posology_dict=posology_dict, applicant_string=applicant)

                path_to_excel = os.path.join(os.getcwd(), output_dir)
                # print("Path to EXCEL : {}".format(path_to_excel))
                path_to_excel = os.path.join(path_to_excel, filename_without_ext+".xlsx")
                # print("Path to EXCEL : {}".format(path_to_excel))
                print("Document finished Processing Successfuly !")

                df.to_excel(path_to_excel, index=False)

                # Remove temporary inetrmidiaire Files : 
                os.remove(docx_path)
                os.remove(txt_path)
                print("Document Saved Successfuly !")

