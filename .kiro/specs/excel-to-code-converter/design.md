# Design Document

## Overview

The Excel to Code converter will transform the existing autocoder.py script into a CLI-based converter that integrates with the Markdown Hub extension's converter system. The design maintains all existing functionality while adapting the configuration mechanism from JSON file to CLI parameters.

## Architecture

### Component Structure
```
backend/
├── converters/
│   ├── base_converter.py (existing)
│   ├── excel_to_code.py (new - main converter)
│   └── utils/
│       ├── sysinfo_extractor.py (new - JSON generator)
│       └── autocoder_core.py (refactored from original)
├── utils/ (existing)
└── tempscript/ (existing - will be deprecated)
```

### Integration Points
- **VS Code Extension**: Right-click menu integration
- **Base Converter**: Inherits from BaseConverter class
- **Python Service**: Called via existing pythonService.ts
- **Command Handler**: New command registration in commandHandler.ts

## Components and Interfaces

### ExcelToCodeConverter Class
```python
class ExcelToCodeConverter(BaseConverter):
    def __init__(self, output_dir: str, **kwargs):
        # CLI parameter mapping
        self.debug_level = kwargs.get('debug_level', 'info')
        self.language = kwargs.get('language', 'english')
        self.mask_style = kwargs.get('mask_style', 'nxp5777m')
        self.reg_short_description = kwargs.get('reg_short_description', True)
        
    def convert(self, input_path: str) -> List[str]:
        # Main conversion logic
        pass
```

### CLI Parameter Mapping
Original config.json → CLI parameters:
- `debug_level` → `--debug-level`
- `language` → `--language`
- `reg_short_description` → `--reg-short-description`
- `regdef_filename` → input file path (positional)
- `mask_style` → `--mask-style`

### SysInfo Extractor
```python
class SysInfoExtractor:
    def extract_to_json(self, xlsx_path: str, output_path: str) -> str:
        # Extract Ark_sysinfo.xlsx to JSON format
        pass
```

## Data Models

### Configuration Model
```python
@dataclass
class AutocoderConfig:
    debug_level: str = 'info'
    language: str = 'english'
    reg_short_description: bool = True
    mask_style: str = 'nxp5777m'
    output_dir: str = './converted_markdown_files'
```

### SysInfo JSON Structure
```json
{
    "baseinfo": {
        "module_name": "string",
        "product_prefix": "string",
        "base_address": "string"
    },
    "reglist": [...],
    "functions": [...],
    "structures": [...]
}
```

## Error Handling

### CLI Parameter Validation
- Validate file paths exist
- Validate enum values for language, debug_level
- Provide helpful error messages for invalid combinations

### Excel File Processing
- Check file format and structure
- Handle missing sheets gracefully
- Provide detailed error messages for malformed data

### Output Generation
- Ensure output directory exists
- Handle file write permissions
- Clean up partial files on failure

## Testing Strategy

### Unit Tests
- CLI parameter parsing
- Configuration validation
- SysInfo extraction
- Core autocoder functionality

### Integration Tests
- End-to-end conversion workflow
- VS Code extension integration
- File system operations

### Regression Tests
- Compare output with original autocoder.py
- Test all supported Excel formats
- Verify all language/style combinations

## Migration Strategy

### Phase 1: Core Refactoring
1. Extract core logic from autocoder.py
2. Create CLI parameter interface
3. Implement BaseConverter integration

### Phase 2: SysInfo Integration
1. Create SysInfo extractor
2. Generate JSON from Ark_sysinfo.xlsx
3. Update autocoder to use JSON data

### Phase 3: VS Code Integration
1. Add converter to registry
2. Implement right-click menu
3. Update command handlers

### Phase 4: Testing & Cleanup
1. Comprehensive testing
2. Documentation updates
3. Remove deprecated tempscript files