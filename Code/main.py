from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from os import path
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
import checker 
import polars as pl
from tqdm import tqdm
from sys import argv

def main():
    project_path = argv[1]
    print(f'Project path: {project_path}')

    docx_path = checker.get_docx_path(project_path)
    if not docx_path:
        print("No docx file found. Exiting.")
        return

    document = Document(docx_path)

    # Preallocate empty Polars DataFrame
    df_table = pl.DataFrame(schema={
        "text": pl.Utf8,
        "style_name": pl.Utf8,
        "good_font_name": pl.Boolean,
        "good_font_size": pl.Boolean,
        "is_cn_en": pl.Boolean,
        "content_type": pl.Utf8,
        "column_id": pl.Int64,
        "row_id": pl.Int64,
        "para_id": pl.Int32,
    })  

    for i, element in enumerate(tqdm(document.element.body)):
        if isinstance(element, CT_P):
            para = Paragraph(element, document)
            df1 = checker.get_rows_from_para(para)
        elif isinstance(element, CT_Tbl):
            table = Table(element, document)
            df1 = checker.get_df_from_table(table)
        else:
            df1 = checker.get_rows_from_unknowobj(element)

        # Add para_id column in Polars
        df1 = df1.with_columns(pl.lit(i).alias("para_id"))

        # Concatenate vertically
        df_table = pl.concat([df_table, df1], how="vertical")

    # Save DataFrame
    data_path = checker.save_df_to_datadir_excel(project_path, df_table)
    print(f"Data saved to {data_path}")


if __name__ == "__main__":
    main()
