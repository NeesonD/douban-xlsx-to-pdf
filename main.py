import pandas as pd
from bs4 import BeautifulSoup
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.platypus.tableofcontents import TableOfContents
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from datetime import datetime
import os

def create_pdf(content, output_path):
    doc = SimpleDocTemplate(output_path, pagesize=A4)
    styles = getSampleStyleSheet()
    
    # 注册中文字体
    pdfmetrics.registerFont(TTFont('SimSun', 'SimSun.ttf'))
    styles.add(ParagraphStyle(name='ChineseTitle', fontName='SimSun', fontSize=16, leading=20, alignment=1))  # 标题居中，字体稍小
    styles.add(ParagraphStyle(name='ChineseLink', fontName='SimSun', fontSize=10, leading=12, alignment=1))  # 链接字体更小
    styles.add(ParagraphStyle(name='ChineseBody', fontName='SimSun', fontSize=12, leading=20))
    
    story = []

    # 添加目录
    toc = TableOfContents()
    toc.levelStyles = [
        ParagraphStyle(fontName='SimSun', fontSize=14, name='TOCHeading1', leftIndent=20, firstLineIndent=-20, spaceBefore=5),
    ]
    story.append(Paragraph("目录", styles['Title']))
    story.append(toc)
    story.append(PageBreak())

    for i, chapter in enumerate(content):
        story.append(PageBreak())  # 新的标题起新的一页
        title = Paragraph(chapter['title'], styles['ChineseTitle'])
        link = Paragraph(chapter['link'], styles['ChineseLink'])
        body_paragraphs = [Paragraph(p, styles['ChineseBody']) for p in chapter['body'].split('\n') if p.strip()]
        story.append(Paragraph(f'<a name="chapter{i}"/>', styles['ChineseTitle']))  # 添加锚点
        story.append(title)
        story.append(link)
        story.append(Spacer(1, 12))
        for paragraph in body_paragraphs:
            story.append(paragraph)
            story.append(Spacer(1, 12))
        toc.addEntry(0, chapter['title'], i + 1)  # 添加到目录

    doc.build(story)

def read_xlsx(file_path):
    df = pd.read_excel(file_path)
    print(df.columns)  # 打印列名
    content = []
    for index, row in df.iterrows():
        # 去除HTML标记
        soup = BeautifulSoup(row['内容'], 'html.parser')
        text_content = soup.get_text()
        chapter = {
            'title': f"{row['创建时间']} - {row['标题']}",
            'link': row['链接'],
            'body': text_content
        }
        content.append(chapter)
    return content

if __name__ == "__main__":
    # 获取当前目录下面的xlsx文件
    current_dir = os.path.dirname(os.path.abspath(__file__))
    xlsx_files = [f for f in os.listdir(current_dir) if f.endswith('.xlsx')]
    
    if not xlsx_files:
        raise FileNotFoundError("当前目录下没有找到xlsx文件")
    
    xlsx_path = xlsx_files[0]  # 假设只有一个xlsx文件
    # 名字加个时间
    current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_filename = os.path.splitext(xlsx_path)[0]
    output_path = f"{output_filename}_{current_time}.pdf"
    content = read_xlsx(xlsx_path)
    create_pdf(content, output_path)