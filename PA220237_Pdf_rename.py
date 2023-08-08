__author__ = 'Shankar TN'

# -*- coding: utf-8 -*-
# -*- encoding: utf-8 -*-
#PA210067

inputpath=r'\\srtif007.in623.corpintra.net\WebFS_Data_Maintainance\WebfsDownload\PDF_Rename_content\Copy of PDF_data_preparationTemplate.xlsx'
#outputpath=r'C:\Users\Kalespo\Downloads\PDF_segregation\th_TH'
output_segregated=r'C:\Users\KIKUMA\Desktop\PDF Rename Final\Output'
file_extension=r'*.pdf'
panumber='PA220237'
#in_excel=r'C:\2022\Data\MBIO_Scriprts\Final_PDF_changes\Input\modify_result_nl.xlsx'
#in_excel=r'C:\2022\Data\MBIO_Scriprts\Final_PDF_changes\Input\modify_result_es_ES_connect_mbux.xlsx'
#in_excel=r'C:\2022\Data\MBIO_Scriprts\Final_PDF_changes\Input\sv_yellow_data.xlsx'
#in_excel=r'Y:\WebfsDownload\MBIO\UnchangedData\zip_contents\tr_TR\Excel_ip\tr_TR_Data.xlsx'
in_excel=r'\\srtif007.in623.corpintra.net\WebFS_Data_Maintainance\WebfsDownload\PDF_Rename_content\Copy of PDF_data_preparationTemplate.xlsx'


import pikepdf
import os
from warnings import warn
import pandas as pd
import time


class get_pdf_data():
    def get_pdf_filelist(self,in_path):
        cur_path=os.getcwd()
        files=os.listdir(in_path)
        path_list=[]
        for file in files:
            temp_path=os.path.join(in_path,file)
            path_list.append(temp_path)
        return path_list

    def warn_or_raise(msg, e=None):
        warn(msg)

    def updatemeta(self, pdf_name,pdftitle,pdfsubject,pdf_file,lang_code):
        no_extracting = pikepdf.Permissions(extract=False,modify_assembly=False,modify_other=False)
        # print(pdf_file)

        pdf = pikepdf.open(pdf_file,allow_overwriting_input=True)

        try:
            with pdf.open_metadata(set_pikepdf_as_editor=True,update_docinfo=True) as meta_obj:
                if pdf.is_encrypted:
                    #print(meta_obj['dc:title'][0])
                    # old_title=meta_obj['dc:title']
                    # print('title before change:{}'.format(old_title))

                    meta_obj['dc:title']=pdftitle
                    meta_obj['dc:description']=pdfsubject
                else:
                    print('pdf not encrypted:{}'.format(pdf_file))

            # pdf.save(os.path.join(outputpath, pdf_name),
            #                  encryption=pikepdf.Encryption(owner='Star2020' ,allow=no_extracting))
            output_pa = ''
            output_move=os.path.join(output_segregated, lang_code)
            print(output_move)
            if not os.path.exists(output_move):
                os.mkdir(output_move)

            else:
                print("Lang floder already exists")
            output_pa = os.path.join(output_move, panumber)
            print(output_pa)
            os.mkdir(output_pa)
            pdf.save(os.path.join(output_pa, pdf_name),
                     encryption=pikepdf.Encryption(owner='Star2020', allow=no_extracting))

        except (ValueError, AttributeError, NotImplementedError) as e:
            print(self.warn_or_raise)
        pdf.close()









"""#main Logoc start from here"""
variablepath = os.path.dirname(__file__)
os.chdir(variablepath)
main_path = os.getcwd()

df=pd.read_excel((in_excel),sheet_name='PA220237')

get_obj=get_pdf_data()
print(df.columns)
print(df.shape[0])

#to get the current script path
cur_path=os.getcwd()

for i,row in df.iterrows():
    pdf_title=row['PDF-Dokument_Titel']
    pdf_name=row['PDF-Dateiname_m._Zähler']
    pdf_subject=row['PDF-Dokument_Thema']
    file_path=row['source_path']
    lang_code=row['Langauge Code']
   # print('PA-Nr:{}\nPDFTitle:{}\nPDF_Name:{}'.format(row['Abgleich_PA-Nr.'],pdf_title,pdf_name))
    print('PDFTitle:{}\nPDF_Name:{}'.format(pdf_title,pdf_name))

    #Go to input pdf file path
    # os.chdir(inputpath)
    if file_path != 'empty':
        # print(file_path)
        try:
            get_obj.updatemeta(pdf_name, pdf_title, pdf_subject, file_path,lang_code)
        except FileNotFoundError as e:
            print(f"FileNotFoundError successfully handled\n"
                  f"{e}")
    else:
        print('record missing from the excel:{}'.format(pdf_title))

#back to script path
os.chdir(cur_path)