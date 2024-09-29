# Bulk Document Generator

A simple tool to generate bulk Documents from a template

## installation

```bash
npm install
```

## Usage

1. Open the application.
2. Select the template file (.docx).

   Note: The template file should contain placeholders for the recipients' names in the following format: {name}

3. Enter the names of the recipients in a text box, one per line.
4. Click the "Generate Document" button.
5. The application will generate the documents in the "pdf" and "docx" folders in the same directory as the template file.

## Acknowledgments

- The [docxtemplater](https://github.com/dolanmiu/docxtemplater) library for generating the documents.
- The [LibreOffice](https://www.libreoffice.org/) application for converting the document to PDF.
