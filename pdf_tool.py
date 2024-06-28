import io
import os
from pathlib import Path

import aspose.pdf as ap


class InvalidPercentageError(Exception):
    def __init__(self, percentage: int):
        super().__init__(f"Invalid compression percentage: {percentage}. Please enter a value between 0 and 100.")


class InvalidDocTypeError(Exception):
    def __init__(self, doc_type: str):
        super().__init__(f'Invalid document type: {doc_type}. Valid documents: ["doc", "docx", "xls", "xlsx", "pptx"].')


class InvalidImgTypeError(Exception):
    def __init__(self, img_type: str):
        super().__init__(f'Invalid image type: {img_type}. Valid images: ["bmp", "emf", "jpg", "png", "gif"].')


def path_exists(path: str) -> bool:
    return os.path.exists(path)


def is_valid_compression_percentage(compression_percentage: int) -> bool:
    return 0 <= compression_percentage <= 100


def is_file(path: str) -> bool:
    return os.path.isfile(path)


def compress_pdf(input_path: str, output_path: str, quality: int) -> bool:
    if not is_file(input_path):
        raise FileExistsError(f"File {input_path} does not exists.")

    if not is_valid_compression_percentage(quality):
        raise InvalidPercentageError(quality)

    try:
        compressPdfDocument = ap.Document(input_path)

        pdfoptimizeOptions = ap.optimization.OptimizationOptions()
        pdfoptimizeOptions.image_compression_options.compress_images = True
        pdfoptimizeOptions.image_compression_options.image_quality = quality
        compressPdfDocument.optimize_resources(pdfoptimizeOptions)

        compressPdfDocument.save(output_path)

        return True
    except Exception as e:
        print(f"Error compress {input_path}: {e}")
        return False


def html_to_pdf(input_path: str, output_path: str) -> bool:
    if not is_file(input_path):
        raise FileExistsError(f"File {input_path} does not exists.")

    try:
        options = ap.HtmlLoadOptions()
        document = ap.Document(input_path, options)

        document.save(output_path)
        return True
    except Exception as e:
        print(f"Error convert {input_path}: {e}")
        return False


def pdf_to_svg(input_path: str, output_path: str) -> bool:
    if not is_file(input_path):
        raise FileExistsError(f"File {input_path} does not exists.")

    try:
        document = ap.Document(input_path)
        saveOptions = ap.SvgSaveOptions()

        saveOptions.compress_output_to_zip_archive = False
        saveOptions.treat_target_file_name_as_directory = True

        document.save(output_path, saveOptions)
        return True
    except Exception as e:
        print(f"Error convert {input_path}: {e}")
        return False


# Convert to MS documents
def pdf_to_doc(input_path: str, output_path: str) -> None:
    document = ap.Document(input_path)
    document.save(output_path, ap.SaveFormat.DOC)


def pdf_to_docx(input_path: str, output_path: str) -> None:
    document = ap.Document(input_path)

    save_options = ap.DocSaveOptions()
    save_options.format = ap.DocSaveOptions.DocFormat.DOC_X
    save_options.mode = ap.DocSaveOptions.RecognitionMode.FLOW
    save_options.relative_horizontal_proximity = 2.5
    save_options.recognize_bullets = True

    document.save(output_path, save_options)


def pdf_to_xls(input_path: str, output_path: str) -> None:
    document = ap.Document(input_path)

    save_option = ap.ExcelSaveOptions()
    save_option.format = ap.ExcelSaveOptions.ExcelFormat.XML_SPREAD_SHEET2003

    document.save(output_path, save_option)


def pdf_to_xlsx(input_path: str, output_path: str) -> None:
    document = ap.Document(input_path)
    save_option = ap.ExcelSaveOptions()

    document.save(output_path, save_option)


def pdf_to_pptx(input_path: str, output_path: str) -> None:
    document = ap.Document(input_path)
    save_option = ap.PptxSaveOptions()
    save_option.slides_as_images = True
    document.save(output_path, save_option)


def pdf_to_documents(input_path: str, output_path: str, doc_type: str) -> bool:
    if not is_file(input_path):
        raise FileExistsError(f"File {input_path} does not exists.")

    if not path_exists(output_path):
        raise NotADirectoryError(f"Output folder {output_path} does not exists.")

    doc_type = doc_type.lower().strip()
    valid_doc_types = ["doc", "docx", "xls", "xlsx", "pptx"]
    if doc_type not in valid_doc_types:
        raise InvalidDocTypeError(doc_type)

    filename, extension = os.path.splitext(os.path.basename(input_path))
    output_name = f"{filename}.{doc_type}"
    output_path = str(Path(output_path) / output_name)

    try:
        if doc_type == "doc":
            pdf_to_doc(input_path, output_path)
        elif doc_type == "docx":
            pdf_to_docx(input_path, output_path)
        elif doc_type == "xls":
            pdf_to_xls(input_path, output_path)
        elif doc_type == "xlsx":
            pdf_to_xlsx(input_path, output_path)
        elif doc_type == "pptx":
            pdf_to_pptx(input_path, output_path)
        return True
    except Exception as e:
        print(f"Error convert {input_path}: {e}")
        return False


# Convert to images
def convert_pdf_bmp(input_path: str, output_path: str) -> None:
    document = ap.Document(input_path)
    resolution = ap.devices.Resolution(300)
    device = ap.devices.BmpDevice(resolution)

    for i in range(0, len(document.pages)):
        imageStream = io.FileIO(
            output_path + "_page_" + str(i + 1) + "_out.bmp", 'x'
        )
        device.process(document.pages[i + 1], imageStream)
        imageStream.close()


def convert_pdf_emf(input_path: str, output_path: str) -> None:
    document = ap.Document(input_path)
    resolution = ap.devices.Resolution(300)
    device = ap.devices.EmfDevice(resolution)

    for i in range(0, len(document.pages)):
        imageStream = io.FileIO(
            output_path + "_page_" + str(i + 1) + "_out.emf", 'x'
        )
        device.process(document.pages[i + 1], imageStream)
        imageStream.close()


def convert_pdf_jpg(input_path: str, output_path: str) -> None:
    document = ap.Document(input_path)
    resolution = ap.devices.Resolution(300)
    device = ap.devices.JpegDevice(resolution)

    for i in range(0, len(document.pages)):
        imageStream = io.FileIO(
            output_path + "_page_" + str(i + 1) + "_out.jpeg", "x"
        )
        device.process(document.pages[i + 1], imageStream)
        imageStream.close()


def convert_pdf_png(input_path: str, output_path: str) -> None:
    document = ap.Document(input_path)
    resolution = ap.devices.Resolution(300)
    device = ap.devices.PngDevice(resolution)

    for i in range(0, len(document.pages)):
        imageStream = io.FileIO(
            output_path + "_page_" + str(i + 1) + "_out.png", 'x'
        )
        device.process(document.pages[i + 1], imageStream)
        imageStream.close()


def convert_pdf_gif(input_path: str, output_path: str) -> None:
    document = ap.Document(input_path)
    resolution = ap.devices.Resolution(300)

    device = ap.devices.GifDevice(resolution)

    for i in range(0, len(document.pages)):
        imageStream = io.FileIO(
            output_path + "_page_" + str(i + 1) + "_out.gif", 'x'
        )
        device.process(document.pages[i + 1], imageStream)
        imageStream.close()


def pdf_to_images(input_path: str, output_path: str, img_type: str) -> bool:
    if not is_file(input_path):
        raise FileExistsError(f"File {input_path} does not exists.")

    if not path_exists(output_path):
        raise NotADirectoryError(f"Output folder {output_path} does not exists.")

    img_type = img_type.lower().strip()
    valid_img_types = ["bmp", "emf", "jpg", "png", "gif"]
    if img_type not in valid_img_types:
        raise InvalidImgTypeError(img_type)

    filename, extension = os.path.splitext(os.path.basename(input_path))
    output_path = str(Path(output_path) / filename)

    try:
        if img_type == "bmp":
            convert_pdf_bmp(input_path, output_path)
        elif img_type == "emf":
            convert_pdf_emf(input_path, output_path)
        elif img_type == "jpg":
            convert_pdf_jpg(input_path, output_path)
        elif img_type == "png":
            convert_pdf_png(input_path, output_path)
        elif img_type == "gif":
            convert_pdf_gif(input_path, output_path)
        return True
    except Exception as e:
        print(f"Error convert {input_path}: {e}")
        return False
