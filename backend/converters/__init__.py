from .md_to_office import MdToOfficeConverter
from .office_to_md import OfficeToMdConverter
from .diagram_to_png import DiagramToPngConverter

try:
    from .batch_processor import (
        BatchProcessor,
        ProgressCallback,
        ProgressInfo,
        BatchResult,
        FileTask,
        MemoryMonitor,
        PagedFileProcessor,
        StreamingFileHandler,
        ProcessingStatus
    )
    BATCH_PROCESSOR_AVAILABLE = True
except ImportError:
    BATCH_PROCESSOR_AVAILABLE = False
    BatchProcessor = None
    ProgressCallback = None
    ProgressInfo = None
    BatchResult = None
    FileTask = None
    MemoryMonitor = None
    PagedFileProcessor = None
    StreamingFileHandler = None
    ProcessingStatus = None

CONVERTER_REGISTRY = {
    'md-to-docx': MdToOfficeConverter,
    'md-to-pdf': MdToOfficeConverter,
    'md-to-html': MdToOfficeConverter,
    'md-to-pptx': MdToOfficeConverter,
    'office-to-md': OfficeToMdConverter,
    'diagram-to-png': DiagramToPngConverter,
}