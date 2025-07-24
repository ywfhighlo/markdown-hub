# Requirements Document

## Introduction

This feature transforms the existing autocoder.py script from a config.json-based tool into a CLI-based converter that integrates with the Markdown Hub extension. The converter will transform Excel files into code files while maintaining all existing functionality.

## Requirements

### Requirement 1

**User Story:** As a developer, I want to convert Excel files to code using a right-click menu option, so that I can easily generate code from Excel specifications.

#### Acceptance Criteria

1. WHEN user right-clicks on an Excel file THEN system SHALL display "Excel to Code" menu option
2. WHEN user selects "Excel to Code" option THEN system SHALL execute the autocoder converter
3. WHEN conversion is complete THEN system SHALL generate code files in the output directory

### Requirement 2

**User Story:** As a developer, I want the autocoder script to accept CLI parameters instead of config.json, so that it can be integrated with the converter system.

#### Acceptance Criteria

1. WHEN autocoder is executed THEN system SHALL accept CLI parameters for all configuration options
2. WHEN no CLI parameters are provided THEN system SHALL use sensible default values
3. WHEN invalid parameters are provided THEN system SHALL display helpful error messages
4. WHEN help is requested THEN system SHALL display all available CLI options

### Requirement 3

**User Story:** As a developer, I want system information from Ark_sysinfo.xlsx to be available as JSON, so that the autocoder can access it programmatically.

#### Acceptance Criteria

1. WHEN system starts THEN system SHALL extract data from Ark_sysinfo.xlsx
2. WHEN data is extracted THEN system SHALL convert it to JSON format
3. WHEN JSON is created THEN system SHALL make it available to autocoder script
4. WHEN Ark_sysinfo.xlsx is updated THEN system SHALL regenerate JSON automatically

### Requirement 4

**User Story:** As a developer, I want the autocoder to be integrated into the converters directory structure, so that it follows the same patterns as other converters.

#### Acceptance Criteria

1. WHEN autocoder is moved THEN system SHALL place it in backend/converters directory
2. WHEN autocoder is integrated THEN system SHALL follow BaseConverter interface pattern
3. WHEN autocoder runs THEN system SHALL use the same logging and error handling as other converters
4. WHEN conversion completes THEN system SHALL return output file paths like other converters

### Requirement 5

**User Story:** As a developer, I want all existing autocoder functionality preserved, so that existing workflows continue to work.

#### Acceptance Criteria

1. WHEN autocoder is transformed THEN system SHALL maintain all existing code generation features
2. WHEN Excel files are processed THEN system SHALL produce identical output to original script
3. WHEN different register styles are used THEN system SHALL handle them correctly
4. WHEN language options are specified THEN system SHALL generate appropriate localized output