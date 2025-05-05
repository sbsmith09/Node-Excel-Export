# Node-Excel-Export: Effortless Excel File Generation in Node.js

## Project Overview

Node-Excel-Export is a lightweight and efficient Node.js library designed to simplify the programmatic generation of Excel (XLSX) files with robust data export capabilities. The library addresses the common challenge of creating structured spreadsheets directly within Node.js applications, providing developers with an intuitive and straightforward solution for Excel file generation.

### Core Purpose
The primary objective of Node-Excel-Export is to enable seamless, server-side Excel spreadsheet creation with minimal complexity. It allows developers to transform data sets into professional, well-formatted Excel files without requiring external dependencies or complex integrations.

### Key Features
- Simple and intuitive Excel file generation
- Support for creating multiple sheets in a single workbook
- Flexible data formatting for various data types (strings, numbers, dates, booleans)
- Custom column styling and formatting capabilities
- Memory-efficient shared string processing
- Cross-platform compatibility

### Benefits
- Lightweight implementation with minimal external dependencies
- No requirement for Microsoft Excel or additional external software
- Supports both synchronous and asynchronous file generation
- Efficient memory management through optimized processing
- Easy integration into Node.js applications
- Suitable for a wide range of use cases including data reporting, analytics, and bulk data exports

### Ideal Use Cases
- Generating financial or statistical spreadsheets
- Creating dynamic reports from databases or APIs
- Programmatically generating templated Excel documents
- Bulk data exports with custom formatting
- Server-side spreadsheet generation in Node.js applications

## Getting Started, Installation, and Setup

## Quick Start

### Installation

Install the package using npm:

```bash
npm install excel-export
```

### Basic Usage

#### Generating a Simple Excel File

```javascript
const nodeExcel = require('excel-export');

// Define column configuration
const conf = {
  cols: [
    { 
      caption: 'Column 1', 
      type: 'string' 
    },
    { 
      caption: 'Column 2', 
      type: 'date' 
    }
  ],
  rows: [
    ['Value 1', new Date()],
    ['Value 2', new Date()]
  ]
};

// Generate Excel file
const result = nodeExcel.execute(conf);
```

### Configuration Options

#### Columns

Each column can have the following properties:
- `caption`: Column header text
- `type`: Data type ('string', 'date', 'bool', 'number')
- `width`: Column width
- `beforeCellWrite`: Optional function to transform cell data
- `captionStyleIndex`: Optional styling index for column header

#### Supported Data Types
- `string`: Text values
- `date`: Date objects
- `bool`: Boolean values
- `number`: Numeric values

### Supported Node.js Environments

- Compatible with Node.js versions 8.x and above
- Tested on Windows, macOS, and Linux platforms

### Dependencies

The package requires the following dependencies:
- `collections`: Data structure management
- `node-zip`: ZIP file generation

### Performance Considerations

- Suitable for generating Excel files with moderate data sizes
- Memory usage increases with larger datasets
- For very large datasets, consider streaming or batch processing

### Error Handling

Ensure proper error handling when generating Excel files:

```javascript
try {
  const result = nodeExcel.execute(conf);
} catch (error) {
  console.error('Excel generation failed:', error);
}
```

## Features / Capabilities

### Core Features

#### Excel File Generation
- Generate XLSX files dynamically from data sets
- Support for multiple sheets in a single workbook
- Flexible column configuration and styling

#### Data Type Support
- Export data across multiple cell types:
  - Strings
  - Numbers
  - Dates
  - Boolean values

#### Customization Options
- Configurable column widths
- Custom cell styling
- Pre-processing of cell data through callback functions
- Dynamic column captions

#### Performance and Compatibility
- Built for Node.js environments
- Lightweight and dependency-minimal
- Uses base64-encoded template for quick file generation

### Usage Examples

#### Basic Export
```javascript
var nodeExcel = require('excel-export');
var conf = {
    cols: [
        { caption: 'Name', type: 'string' },
        { caption: 'Age', type: 'number' }
    ],
    rows: [
        ['John Doe', 28],
        ['Jane Smith', 32]
    ]
};
var result = nodeExcel.execute(conf);
```

#### Multiple Sheets
```javascript
var configs = [
    {
        name: 'Employees',
        cols: [ ... ],
        rows: [ ... ]
    },
    {
        name: 'Departments',
        cols: [ ... ],
        rows: [ ... ]
    }
];
var result = nodeExcel.execute(configs);
```

#### Advanced Cell Configuration
```javascript
{
    cols: [{
        caption: 'Salary',
        type: 'number',
        beforeCellWrite: function(row, cellData, eObj) {
            // Custom formatting or transformation
            return cellData * 1.1; // Example: 10% increase
        }
    }]
}
```

## Usage Examples

### Basic Excel Export

```javascript
const nodeExcel = require('excel-export');

// Define column configuration
const conf = {
  cols: [{
    caption: 'String Column',
    type: 'string',
    width: 15
  }, {
    caption: 'Date Column',
    type: 'date',
    width: 20
  }, {
    caption: 'Boolean Column',
    type: 'bool'
  }, {
    caption: 'Number Column',
    type: 'number',
    width: 30
  }],
  
  // Define data rows
  rows: [
    ['Hello', new Date(2023, 4, 1), true, 3.14159],
    ['World', new Date(2023, 5, 15), false, 2.7182]
  ]
};

// Generate Excel file
const excelBuffer = nodeExcel.execute(conf);
```

### Generating Large Excel Files

```javascript
const nodeExcel = require('excel-export');

// Create configuration for a large spreadsheet
const conf = {
  cols: [],
  rows: []
};

// Add 100 columns
for (let i = 0; i < 100; i++) {
  conf.cols.push({
    caption: `Column ${i}`,
    type: 'string'
  });
}

// Add 1000 rows of data
for (let j = 0; j < 1000; j++) {
  const row = [];
  for (let k = 0; k < 100; k++) {
    row.push(`Data ${j}-${k}`);
  }
  conf.rows.push(row);
}

// Generate large Excel file
const largeExcelBuffer = nodeExcel.execute(conf);
```

### Advanced Cell Formatting

```javascript
const nodeExcel = require('excel-export');

const conf = {
  cols: [{
    caption: 'Custom String',
    type: 'string',
    beforeCellWrite: function(row, cellData) {
      // Convert cell data to uppercase
      return cellData.toUpperCase();
    },
    width: 15
  }, {
    caption: 'Formatted Date',
    type: 'date',
    beforeCellWrite: function() {
      const originDate = new Date(Date.UTC(1899, 11, 30));
      return function(row, cellData, options) {
        // Optional styling based on row number
        if (options.rowNum % 2) {
          options.styleIndex = 1;
        }
        
        if (cellData === null) {
          options.cellType = 'string';
          return 'N/A';
        }
        
        return (cellData - originDate) / (24 * 60 * 60 * 1000);
      };
    }()
  }],
  
  rows: [
    ['custom text', new Date(2023, 4, 1)],
    ['another text', new Date(2023, 5, 15)]
  ]
};

const excelBuffer = nodeExcel.execute(conf);
```

### Using Custom Styles

```javascript
const nodeExcel = require('excel-export');

const conf = {
  // Optional: Path to a custom styles.xml file
  stylesXmlFile: 'path/to/custom/styles.xml',
  
  cols: [{
    caption: 'Styled Column',
    type: 'string',
    captionStyleIndex: 1,
    width: 20
  }],
  
  rows: [
    ['Styled Content']
  ]
};

const excelBuffer = nodeExcel.execute(conf);
```

### Notes
- The `execute()` method returns a buffer that can be saved to a file or sent as a response
- Column types include: `string`, `number`, `date`, `bool`
- Use `beforeCellWrite` for custom cell transformations
- Optional `width` and `captionStyleIndex` for column customization

## Project Structure

The project is organized with the following key directories and files:

```
.
├── example/              # Example usage and demonstration files
│   ├── app.js            # Sample application showing library usage
│   ├── package.json      # Project dependencies for example
│   └── styles.xml        # XML styling configuration
├── test/                 # Unit testing directory
│   └── main.js           # Test suite for the library
├── index.js              # Main entry point of the library
├── sheet.js              # Core functionality for sheet manipulation
├── package.json          # Project configuration and dependencies
└── .gitignore            # Git ignore file for excluding unnecessary files
```

### Key Components
- `index.js`: The primary entry point that exports the main library functionality
- `sheet.js`: Contains core logic for Excel sheet manipulation
- `example/`: Provides sample code demonstrating library usage
- `test/`: Contains unit tests to ensure library reliability

### Configuration
- `package.json` defines project metadata, dependencies, and scripts
  - Main entry point is set to `index.js`
  - Test script uses Mocha for running tests in `test/main.js`

## Technologies Used

### Languages
- JavaScript (Node.js)

### Core Dependencies
- [collections](https://www.npmjs.com/package/collections): Data structure library for advanced collection handling
- [node-zip](https://www.npmjs.com/package/node-zip): Utility for creating ZIP archives

### Development and Testing
- [Mocha](https://mochajs.org/): Testing framework for JavaScript
- [should.js](https://github.com/shouldjs/should.js): Assertion library for Node.js testing

### Target Platform
- Node.js

### File Formats
- Excel XLSX

## Additional Notes

### Library Maturity and Scope
This is an early-stage library (version 0.5.1) focused on providing a lightweight solution for Excel file generation in Node.js environments. While functional, it is best suited for straightforward export scenarios.

### Performance Characteristics
- In-memory Excel file generation
- Efficient processing for single and multiple sheet exports
- Shared string optimization to minimize file size
- Minimal overhead due to lightweight implementation

### Potential Limitations
- Supports programmatic Excel file creation only
- Cannot read or modify existing Excel files
- Limited advanced styling capabilities
- Relies on specific data structure for export

### Data Handling
- Supports basic data types: strings, numbers, dates, and booleans
- Requires client-side data validation and sanitization
- Gracefully handles null or undefined cell values

### Environment Compatibility
- Requires Node.js 10.x or higher
- Cross-platform support (Windows, macOS, Linux)
- No external software dependencies beyond npm packages

### Best Use Cases
- Generating simple, programmatically created spreadsheets
- Data reporting and analytics exports
- Bulk data transformation to Excel format
- Server-side spreadsheet generation with minimal complexity

### Extensibility
While the library provides core functionality, developers can extend its capabilities through:
- Custom column configurations
- Transformation functions for cell data
- Optional custom XML style definitions

### Community and Support
- Open-source project maintained on GitHub
- Community-driven development
- Minimal dependencies ensures long-term stability

## Contributing

We welcome contributions to the Node Excel Export project! To ensure a smooth contribution process, please follow these guidelines:

### Contribution Process

1. Fork the repository and create your branch from `main`.
2. If you've added code that should be tested, add tests to the `test/` directory.
3. Ensure the test suite passes by running `npm test`.

### Development Setup

- The project uses Mocha for testing
- Dependencies can be installed via `npm install`
- Use the existing test structure in `test/main.js` as a reference for new tests

### Code Style and Standards

- Follow the existing code structure and formatting
- Write clear, concise comments
- Ensure new code is well-documented

### Reporting Issues

- Use GitHub Issues to report bugs or suggest features
- Provide detailed information, including:
  - Expected behavior
  - Actual behavior
  - Steps to reproduce
  - Node.js and package version

### Pull Request Guidelines

- Provide a clear description of your changes
- Link any related issues
- Ensure all tests pass before submitting
- Include appropriate test coverage for new features

### Testing

- Run existing tests with `npm test`
- Add new tests for any added functionality
- Ensure 100% test coverage for new code paths

### Code of Conduct

- Be respectful and considerate of other contributors
- Collaborate openly and constructively

### Licensing

By contributing, you agree that your contributions will be licensed under the project's BSD license.

## License

This project is licensed under the BSD License. 

#### License Details
The BSD License is a permissive free software license that allows for extensive use, modification, and redistribution with minimal restrictions. Key provisions include:

- Commercial and non-commercial use is permitted
- Modification and distribution are allowed
- Attribution to the original author is required
- No warranty is provided

### Disclaimer
This software is provided "as is" without warranties of any kind. The copyright holder shall not be liable for any damages arising from its use.

For the full license text, please refer to the standard BSD License terms available at: [BSD License](https://opensource.org/licenses/BSD-3-Clause)