# Copilot Instructions for MsgToPdfConverter

## Project Overview
- **MsgToPdfConverter** is a WPF/.NET Framework 4.8+ application for converting Outlook .msg files (and attachments) to PDF, supporting drag-and-drop from Outlook, folders, and files.
- The app handles complex attachment types: nested .msg, Office docs, images, ZIP/7z archives, and more, with recursive extraction and conversion.
- Main UI logic is in `MainWindow.xaml`/`MainWindowViewModel.cs`. Core processing is in `AttachmentService.cs` and `EmailConverterService.cs`.

## Architecture & Patterns
- **MVVM pattern**: ViewModels (e.g., `MainWindowViewModel.cs`) handle UI logic, commands, and data binding.
- **AttachmentService**: Central for recursive attachment extraction, file type handling, and PDF conversion. Handles hierarchy-aware headers and progress reporting.
- **EmailConverterService**: Builds HTML from .msg content, manages inline images, and supports conversion to PDF.
- **PDF generation**: Uses iText and PdfSharp-GDI (see NuGet dependencies). Office conversion uses external tools or libraries (see `_tryConvertOfficeToPdf`).
- **Drag-and-drop**: Special handling for Outlook data formats in `MainWindowViewModel.cs`.
- **Temp files**: All intermediate files are written to `%TEMP%\MsgToPdfConverter` and cleaned up after processing.

## Developer Workflows
- **Build**: Requires .NET Framework 4.8+. Use `dotnet restore` then `dotnet build -c Release` or build via Visual Studio.
- **Run**: Start `MsgToPdfConverter.exe` (GUI) or use CLI arguments for batch/automated conversion.
- **NuGet**: Restore packages before build. PdfSharp-GDI is required (`dotnet add package PdfSharp-GDI --version 1.50.5147`).
- **Native dependency**: `libwkhtmltox.dll` must be present in the output directory for HTML-to-PDF conversion.
- **Debug logging**: Use `#if DEBUG` blocks and `DebugLogger` for diagnostics. Logs are written to `SingleInstanceManager.log`.

## Project-Specific Conventions
- **Attachment deduplication**: Only true duplicates within the same attachment group are skipped; identical names in different containers are processed independently.
- **Hierarchy headers**: Each PDF page for an attachment includes a header showing the parent chain (see `CreateHierarchyHeaderText`).
- **Progress reporting**: Use `progressTick` and `progressTotalCallback` for UI updates during batch/recursive processing.
- **Signature image detection**: Small images are heuristically skipped as likely signatures (see `IsLikelySignatureImage`).
- **Error handling**: On conversion failure, a placeholder PDF with an error message is generated and included in the output.

## Key Files & Directories
- `MainWindow.xaml`, `MainWindowViewModel.cs`: UI and user interaction logic
- `AttachmentService.cs`: Core recursive extraction and conversion logic
- `EmailConverterService.cs`: Email body and inline image handling
- `PdfService.cs`, `HierarchyImageService.cs`: PDF and hierarchy diagram generation
- `fonts/`: Custom fonts for PDF output
- `bin/`, `obj/`: Build outputs (ignore in source control)

## External Integrations
- **iText** and **PdfSharp-GDI** for PDF creation/merging
- **wkhtmltopdf** (via `libwkhtmltox.dll`) for HTML-to-PDF
- **SevenZipSharp** or similar for 7z archive extraction

## Example: Adding a New Attachment Type
1. Update `AttachmentService.ProcessSingleAttachmentWithHierarchy` to handle the new file extension.
2. Add conversion logic and error handling as per existing patterns.
3. Update progress and header generation as needed.

---

For more, see `README.md` and comments in `AttachmentService.cs`.
