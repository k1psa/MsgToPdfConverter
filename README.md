# MsgToPdfConverter

## Overview
MsgToPdfConverter is a WPF application that allows users to select multiple Outlook .msg files and convert them into PDF format. This application utilizes the MsgReader library for parsing .msg files and a suitable PDF generation library for exporting the content to PDF.

## Features
- Select multiple .msg files using a user-friendly file dialog.
- Convert selected .msg files to PDF format.
- Batch processing of multiple files for efficient conversion.

## Getting Started

### Prerequisites
- .NET Framework (version compatible with WPF applications)
- Visual Studio or any compatible IDE for building and running the application

### Installation
1. Clone the repository or download the source code.
2. Open the solution file `MsgToPdfConverter.sln` in your IDE.
3. Restore the NuGet packages by running the following command in the Package Manager Console:
   ```
   Update-Package
   ```
4. Ensure that the required libraries for MsgReader and PDF generation are listed in `packages.config`.

### Usage
1. Run the application.
2. Click on the "Select Files" button to open the file dialog.
3. Choose the .msg files you wish to convert.
4. Click on the "Convert" button to start the conversion process.
5. The converted PDF files will be saved in the specified output directory.

## Contributing
Contributions are welcome! Please feel free to submit a pull request or open an issue for any enhancements or bug fixes.

## License
This project is licensed under the MIT License. See the LICENSE file for more details.