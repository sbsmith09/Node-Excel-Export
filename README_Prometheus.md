# Node-Excel-Export: Lightweight XLSX Generation for Node.js

## Project Overview

Node-Excel-Export is a lightweight Node.js library designed to simplify the process of programmatically generating Excel (XLSX) files with robust data export capabilities.

### Key Features
- Simple and straightforward Excel file generation
- Support for multiple sheet creation
- Flexible data formatting options
- Handles various data types (strings, numbers, dates, booleans)
- Custom column styling capabilities
- Memory-efficient shared string processing

### Purpose
The library addresses the common challenge of generating structured Excel spreadsheets directly from Node.js applications. It provides developers with an intuitive interface to export data sets into professional, well-formatted Excel files without complex dependencies or external services.

### Use Cases
- Data reporting and analytics
- Generating financial or statistical spreadsheets
- Bulk data exports from databases or APIs
- Creating templated Excel documents with dynamic content

### Core Benefits
- Lightweight implementation with minimal external dependencies
- No need for Microsoft Excel or other external software
- Cross-platform compatibility
- Supports custom cell styling and formatting
- Efficient memory management through shared string optimization

## Getting Started, Installation, and Setup

### Prerequisites

- Node.js (version 10.x or higher recommended)
- npm (Node Package Manager)

### Installation

Install the package using npm:

```bash
npm install excel-export
```

### Quick Start

Here's a basic example of how to use the excel-export package:

```javascript
const ExcelExport = require('excel-export');

// Prepare your data
const data = [
  ['Header1', 'Header2'],
  ['Value1', 'Value2'],
  ['Value3', 'Value4']
];

// Export to Excel
const result = ExcelExport.convert(data);
```

### Usage

#### Basic Export

```javascript
const ExcelExport = require('excel-export');

// Create a dataset
const conf = {};
conf.cols = [
  {caption:'String', type:'string'},
  {caption:'Number', type:'number'}
];
conf.rows = [
  ['foo', 1],
  ['bar', 2]
];

// Generate Excel file
const result = ExcelExport.execute(conf);
```

### Development

#### Running Tests

To run the project tests:

```bash
npm test
```

### Dependencies

The project relies on the following key dependencies:
- `collections`: ^3.0.0
- `node-zip`: 1.x

### Platform Compatibility

This package is compatible with:
- Node.js (10.x and above)
- Windows, macOS, and Linux platforms

### Troubleshooting

- Ensure you have the latest version of Node.js installed
- Check that all dependencies are correctly installed
- Verify your data format matches the expected input structure

## API Reference

### Public Functions and Methods

#### `executeAsync(config, callBack)`
Asynchronously generates an Excel (.xlsx) file.

- **Parameters**:
  - `config`: Configuration object or array of configuration objects for creating Excel sheets
  - `callBack`: Function to be called with the generated Excel file result

- **Example**:
```javascript
exports.executeAsync({
  name: 'My Sheet', 
  cols: [...], 
  rows: [...]
}, function(result) {
  // Handle generated Excel file
});
```

#### `execute(config)`
Synchronously generates an Excel (.xlsx) file.

- **Parameters**:
  - `config`: Configuration object or array of configuration objects for creating Excel sheets

- **Returns**: Generated Excel file as a buffer

- **Example**:
```javascript
let excelFile = exports.execute({
  name: 'My Sheet', 
  cols: [...], 
  rows: [...]
});
```

### Sheet Configuration Object

The configuration object for creating Excel sheets includes the following properties:

- `name` (optional): Name of the sheet (defaults to 'sheet1', 'sheet2', etc.)
- `cols`: Array defining column configurations
  - `caption`: Column header text
  - `type`: Cell data type ('string', 'number', 'date', 'bool')
  - `width` (optional): Column width
  - `captionStyleIndex` (optional): Style index for column headers
  - `beforeCellWrite` (optional): Function to transform cell data before writing
- `rows`: 2D array containing sheet data
- `stylesXmlFile` (optional): Path to a custom XML styles file

### Utility Extensions

#### `Date.prototype.getJulian()`
Calculates the Julian date for the current date.

- **Returns**: Julian date as a number

#### `Date.prototype.oaDate()`
Converts the date to an Office Automation (OA) date.

- **Returns**: OA date as a number representing days since December 30, 1899

### Internal Methods (Not Typically Used Directly)

- `generateMultiSheets(configs, xlsx)`: Generate multiple sheets in an Excel file
- `generateContentType(configs, xlsx)`: Generate content type XML
- `generateRel(configs, xlsx)`: Generate relationships XML
- `generateWorkbook(configs, xlsx)`: Generate workbook XML
- `generateSharedStringsFile(xlsx)`: Generate shared strings XML

### Notes
- The library supports generating Excel files with multiple sheets
- Supports various cell types: string, number, date, boolean
- Allows custom styling and cell transformations
- Uses shared strings for optimized file size
- Works synchronously and asynchronously

## Project Structure

The project follows a standard Node.js module structure with the following key directories and files:

```
.
├── index.js           # Main entry point for the Excel export module
├── sheet.js           # Core implementation for generating Excel worksheet
├── package.json       # Project metadata and dependency configuration
│
└── example/           # Example usage directory
    ├── app.js         # Sample application demonstrating module usage
    ├── package.json   # Example project dependencies
    └── styles.xml     # Optional custom Excel styles configuration
│
└── test/              # Unit testing directory
    └── main.js        # Test suite for the module
```

#### Key Files

- `index.js`: Primary module file containing the core export functionality
  - Manages Excel file generation
  - Handles multiple sheet creation
  - Provides async and synchronous export methods

- `sheet.js`: Worksheet generation logic
  - Manages individual Excel sheet creation
  - Handles cell type conversion
  - Supports custom styling and data transformations

- `package.json`: Defines project metadata
  - Lists dependencies like `collections` and `node-zip`
  - Specifies version and repository information

#### Directories

- `example/`: Demonstrates practical usage of the module
- `test/`: Contains unit tests to validate module functionality

## Technologies Used

### Core Technologies
- **Node.js**: Primary runtime environment for the project

### Libraries and Dependencies
- **collections**: Data structure library
- **node-zip**: Zip file creation and manipulation library

### Development and Testing
- **Mocha**: Testing framework
- **should**: Assertion library for testing

### File Formats
- **XLSX**: Excel file export functionality
- **XML**: Configuration and styling support (evidenced by styles.xml)

### Programming Languages
- **JavaScript**: Primary programming language

## Additional Notes

### Performance and Memory Considerations

This library is designed for server-side Excel export operations and has specific performance characteristics:

- Uses in-memory processing for Excel file generation
- Supports generating single and multiple sheet exports
- Efficient string handling with shared string optimization to reduce file size

### Compatibility and Limitations

- Works with Node.js environments
- Generates `.xlsx` files compatible with modern spreadsheet software
- Does not support reading or editing existing Excel files
- Limited styling options compared to full-featured Excel libraries

### Security Notes

- Always sanitize and validate input data before export
- Escapes special XML characters to prevent potential XML injection
- No external file dependencies beyond npm packages

### Error Handling

- Silent failure for null or undefined cell values
- No built-in error validation for data types
- Recommends client-side data validation before export

### Versioning and Maintenance

- Maintained as an open-source project
- Minimal dependencies (`collections` and `node-zip`)
- Version 0.5.1 indicates an early-stage library with potential for future enhancements

### Usage Recommendations

- Best suited for simple, programmatically generated Excel exports
- Ideal for scenarios requiring server-side spreadsheet generation
- Consider alternative libraries for complex Excel manipulation

## Contributing

We welcome contributions to the Node Excel Export project! To ensure a smooth and collaborative process, please follow these guidelines:

### Contribution Process

1. Fork the repository and create your branch from `main`.
2. Ensure any new code is well-documented and follows the existing code style.
3. Write or update tests to cover any changes you make.

### Development Setup

- The project uses Mocha for testing
- Run tests using the command: `npm test`

### Contribution Guidelines

#### Code Style
- Follow the existing code formatting in the project
- Use clear, descriptive variable and function names
- Add comments to explain complex logic

#### Testing
- All contributions must include corresponding test cases
- Existing tests in `test/main.js` provide a reference for test structure
- Ensure all tests pass before submitting a pull request

#### Pull Request Process
- Provide a clear description of your changes
- Link any related issues
- Ensure all CI checks pass
- Your code will be reviewed by the maintainers

### Reporting Issues
- Use the GitHub Issues section to report bugs or suggest improvements
- Include detailed information about the issue, including:
  - Steps to reproduce
  - Expected behavior
  - Actual behavior
  - Your environment details

### Questions?
If you have any questions about contributing, please open an issue for discussion.

## License

The project is licensed under the BSD License. 

For the full license text, please refer to the standard BSD License terms. The BSD License is a permissive free software license that allows for reuse within both free and proprietary software.

### Key Permissions
- Commercial use is permitted
- Modification and distribution are allowed
- An attribution to the original author is required
- Comes with no warranty

The complete license details can be found in the standard BSD License documentation.