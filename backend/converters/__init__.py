from .md_to_office import MdToOfficeConverter
from .office_to_md import OfficeToMdConverter
from .diagram_to_png import DiagramToPngConverter
from .excel_to_code import ExcelToCodeConverter

CONVERTER_REGISTRY = {
    'md-to-docx': MdToOfficeConverter,
    'md-to-pdf': MdToOfficeConverter,
    'md-to-html': MdToOfficeConverter,
    'md-to-pptx': MdToOfficeConverter,
    'office-to-md': OfficeToMdConverter,
    'diagram-to-png': DiagramToPngConverter,
    'excel-to-code': ExcelToCodeConverter,
}