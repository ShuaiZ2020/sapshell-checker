from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from os import path
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
import checker 
import pandas as pd
from tqdm import tqdm

def main():
    project_path = checker.request_project_path()
    docx_path = checker.get_docx_path(project_path)
    document = Document(docx_path)
    df_table = pd.DataFrame({'text':[],
            'style_name':[],
            'good_font_name':[],
            'good_font_size':[],
            'is_cn_en':False,
            'content_type':[],
            'column_id':[],
            'row_id': []
            })
    for i in tqdm(range(len(document.element.body))):
        element = document.element.body[i]
        if isinstance(element, CT_P):
            para = Paragraph(element, document)
            df1 = checker.get_rows_from_para(para)

        elif isinstance(element, CT_Tbl):
            table = Table(element, document)
            df1 = checker.get_df_from_table(table)
        else:
            df1 = checker.get_rows_from_unknowobj(element) 
            
        df_table = pd.concat([df_table, df1], ignore_index=True)

    data_path = checker.save_df_to_datadir_excel(project_path, df_table)
    print(f"Data saved to {data_path}")
    

if __name__ == "__main__":
    main()

