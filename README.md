# PDF-Tool
Compress and convert PDF files to various other file formats using Aspose.PDF library.

Note: you have to purchase a valid licence from aspose.com!

## Examlpe
```python
from pdf_tool import compress_pdf, html_to_pdf, pdf_to_svg, pdf_to_documents, pdf_to_images

# Usages
compress_pdf(input_path="input.pdf", output_path="D:\\Dev\\Python\\ConverterToolkit\\compressed.pdf", quality=100) # Compress pdf file by quality 0-100

html_to_pdf(input_path="input.html", output_path="D:\\Dev\\Python\\ConverterToolkit\\html_to_pdf.pdf") # Convert html content to pdf file

pdf_to_svg(input_path="input.pdf", output_path="D:\\Dev\\Python\\ConverterToolkit\\pdf_to_svg.svg") # Convert pdf file to svg

pdf_to_documents(input_path="input.pdf", output_path="D:\\Dev\\Python\\ConverterToolkit", doc_type="docx") # Convert pdf file to MS documents: ["doc", "docx", "xls", "xlsx", "pptx"]

pdf_to_images(input_path="input.pdf", output_path="D:\\Dev\\Python\\ConverterToolkit", img_type="jpg") # Convert pdf file to images: ["bmp", "emf", "jpg", "png", "gif"]

# All function will return True if convert/compress successfull otherwise False.
```

Make sure to install aspose-pdf version: 24.6.0 
```
pip install aspose-pdf==24.6.0
```

###

<h2 align="left">Support</h2>

###

<p align="left">If you'd like to support my ongoing efforts in sharing fantastic open-source projects, you can contribute by making a donation via PayPal.</p>

###

<div align="center">
  <a href="https://www.paypal.com/paypalme/iamironman0" target="_blank">
    <img src="https://img.shields.io/static/v1?message=PayPal&logo=paypal&label=&color=00457C&logoColor=white&labelColor=&style=flat" height="40" alt="paypal logo"  />
  </a>
</div>

###
