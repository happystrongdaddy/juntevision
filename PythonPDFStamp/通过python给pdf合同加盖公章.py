from PyPDF2 import PdfFileWriter, PdfFileReader


def extract_information(pdf_path):
    with open(pdf_path, 'rb') as f:
        pdf = PdfFileReader(f)
        information = pdf.getDocumentInfo()
        number_of_pages = pdf.getNumPages()

    txt = f"""
    Information about {pdf_path}: 

    Author: {information.author}
    Creator: {information.creator}
    Producer: {information.producer}
    Subject: {information.subject}
    Title: {information.title}
    Number of pages: {number_of_pages}
    """

    print(txt)
    return information


def create_watermark(input_pdf, output, watermark):
    pdf_reader = PdfFileReader(input_pdf)
    pdf_writer = PdfFileWriter()
    watermark_obj = PdfFileReader(watermark)
    input_pdf_pages = pdf_reader.getNumPages()
    for i in range(input_pdf_pages):
        page = pdf_reader.getPage(i)
        watermark_page = watermark_obj.getPage(i)
        page.mergePage(watermark_page)
        pdf_writer.addPage(page)
    with open(output, 'wb') as out:
        pdf_writer.write(out)


def get_order_pages(input_pdf_path):
    pdf_reader = PdfFileReader(input_pdf_path)
    return pdf_reader.getNumPages()


def get_watermark_file(order_pages):
    if order_pages == 2:
        watermark_path = r"K:\GithubCode\juntevision\PythonPDF盖章\盖海康的合同2页版本水印.pdf"
    elif order_pages == 3:
        watermark_path = r"K:\GithubCode\juntevision\PythonPDF盖章\盖海康的合同3页版本水印.pdf"
    return watermark_path


if __name__ == '__main__':
    #本例提供的是给2页或者3页的海康的合同盖章，
    #input_pdf_path是2页的海康合同
    input_pdf_path = r'K:\GithubCode\juntevision\PythonPDF盖章\海康合同测试3页-副本.pdf'
    #output_pdf_path是输出合同的路径
    output_pdf_path = r'K:\GithubCode\juntevision\PythonPDF盖章\盖章之后合同.pdf'
    # 获取海康进货合同页数
    input_pdf_pages = get_order_pages(input_pdf_path)
    #watermark是水印文件的路径
    watermark_path = get_watermark_file(input_pdf_pages)
    create_watermark(input_pdf=input_pdf_path,
                     output=output_pdf_path,
                     watermark=watermark_path)
