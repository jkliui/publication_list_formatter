import pandas as pd
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

# target author
target_author = "Xxx, X"

def format_authors(authors_str):
    authors = authors_str.split(';')
    formatted = []
    for author in authors:
        author = author.strip()
        if ',' in author:
            last, first = author.split(',', 1)
            first_initial = first.strip()[0] if first.strip() else ''
            formatted.append(f"{last.strip()}, {first_initial}.")
        else:
            formatted.append(author.strip())  # fallback
    return '; '.join(formatted)

# load csv file
df = pd.read_csv("publication_list.csv")

# replace NaN to '' and change types
df.fillna('', inplace=True)
df['Volume'] = pd.to_numeric(df['Volume'], errors='coerce').fillna(0).astype(int)
df['Year'] = pd.to_numeric(df['Year'], errors='coerce').fillna(0).astype(int)
df = df.astype({col: str for col in df.columns if col not in ['Volume', 'Year']})

# format authors
df['Formatted Authors'] = df['Authors'].apply(format_authors)

# extract information and sort with year in descending order
df_selected = df[['Formatted Authors', 'Title', 'Publication', 'Year', 'Volume', 'Pages']]
df_selected = df_selected.sort_values(by='Year', ascending=False).reset_index(drop=True)

# creat word
doc = Document()

# font style
style = doc.styles['Normal']
font = style.font
font.name = 'Times New Roman'
font.size = Pt(12)

# total publication number
total = len(df_selected)

# iterate each publication
for _, row in df_selected.iterrows():
    author = row['Formatted Authors']
    title = row['Title']
    journal = row['Publication']
    year = row['Year']
    volume = row['Volume']
    pages = row['Pages']

    # add paragraph and format
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    paragraph.paragraph_format.line_spacing = 1.0
    paragraph.paragraph_format.space_after = Pt(6)
    paragraph.paragraph_format.left_indent = Inches(0)
    paragraph.paragraph_format.first_line_indent = Inches(-0.3)

    # write order number
    paragraph.add_run(f"{total})  ")

    # write in authors
    if target_author in author:
        parts = author.split(target_author)
        for i, part in enumerate(parts):
            if part:
                paragraph.add_run(part)
            if i < len(parts) - 1:
                run = paragraph.add_run(target_author)
                run.bold = True
                run.underline = True
    else:
        paragraph.add_run(author)

    paragraph.add_run(" ")  # add space

    # Title
    paragraph.add_run(f"{title}. ")

    # Journal name (italic)
    paragraph.add_run(journal).italic = True
    paragraph.add_run(". ")

    # Year (bold)
    year_run = paragraph.add_run(str(year))
    year_run.bold = True
    paragraph.add_run(", ")

    # Volume (italic)
    volume_run = paragraph.add_run(str(volume))
    volume_run.italic = True
    paragraph.add_run(f", {pages}.")

    total -= 1

# save word
doc.save("highlighted_publications.docx")
print("Task completed：highlighted_publications.docx")
