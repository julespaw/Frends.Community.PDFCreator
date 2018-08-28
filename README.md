# Frends.Community.PDFWriter
Frends task for creating PDF documents

- [Installing](#installing)
- [Tasks](#tasks)
  - [Create Pdf](#createpdf)
- [License](#license)
- [Building](#building)
- [Contributing](#contributing)
- [Change Log](#change-log)

# Installing
You can install the task via FRENDS UI Task View or you can find the nuget package from the following nuget feed
'Nuget feed coming at later date'

Tasks
=====

## CreatePdf

### Task Properties

### Output File

Settings for writing the PDF file.

| Property             | Type                 | Description                          | Example |
| ---------------------| ---------------------| ------------------------------------ | ----- |
| Directory | string | Destination folder for the PDF file created. | C:\Pdf_Output |
| File name | string | File name for the PDF file created. | my_pdf_file.pdf |
| File Exists Action | enum {Error, Overwrite, Rename} | What to do if destination file already exists. Error: throw exception. Overwrite: Replaces existing file. Rename: Renames file by adding '\_(1)' to the end. (pdf_file.pdf --> pdf_file\_(1).pdf) |
| Unicode | bool | If true, Unicode support and font embedding are enabled. otherwise ANSI characters are supported. Further [documentation](http://www.pdfsharp.net/wiki/Unicode-sample.ashx). | false |

### Document Settings

Settings for the pdf Document.

| Property             | Type                 | Description                          | Example |
| ---------------------| ---------------------| ------------------------------------ | ----- |
| Title | string | Optional title for the PDF document | 'My title' |
| Author | string |  Optional author of the PDF document | 'John Doe' |
| Size | enum { A0, A1, A2, A3, A4 ...} | PDF document page size | A4 |
| Orientation | enum { Portrait, Landscape } | PDF document page orientation | Portrait |
| Margin left in cm | double | Set left margin in centimeters for all pages. | 2.5 |
| Margin top in cm | double | Set top margin in centimeters for all pages. | 2 |
| Margin right in cm | double | Set right margin in centimeters for all pages. | 2.5 |
| Margin bottom in cm | double | Set bottom margin in centimeters for all pages. | 2 |

### Content

Actual PDF document content

| Property             | Type                 | Description                          | Example |
| ---------------------| ---------------------| ------------------------------------ | ----- |
| Content type | enum { Paragraph, Image, PageBreak } | The type on content added to PDF document | Paragraph |

#### Paragraph

| Property             | Type                 | Description                          | Example |
| ---------------------| ---------------------| ------------------------------------ | ----- |
| Text | string | Text content to write in paragraph | 'This is example content for paragraph.' |
| Font Family | string | Font family to be used in this paragraph | Times New Roman |
| Font Size | int | Paragraph font size (pt) | 11 |
| Font Style | enum { Regular, Bold, Italic, BoldItalic, Underline } | Text formatting style | Regular |
| Line Spacing | int | Spacing between lines. | 14 |
| Paragraph Alignment | enum {Left, Center, Justify, Right } | Alignment of the paragraph. | Left |
| Spacing Before In Pt | int | Space added before paragraph. | 8 |
| Spacing After In Pt | int | Space added after paragraph. | 0 |

#### Image

| Property             | Type                 | Description                          | Example |
| ---------------------| ---------------------| ------------------------------------ | ----- |
| Image Path | string | Full path to image file. Only support PNG. | c:\my_images\example_image.png |
| Image Alignment | enum { Left, Center, Right } | Image alignment in page. | Center |
| Spacing Before In Pt | int | Space added before image. | 8 |
| Spacing After In Pt | int | Space added after image. | 0 |

#### PageBreak

Adds page break to PDF document.

### Options

| Property             | Type                 | Description                          | Example |
| ---------------------| ---------------------| ------------------------------------ | ----- |
| Use Given Credentials | bool | If set, allows you to give the user credentials to use to write the PDF file on remote hosts. | true|
| User Name | string | Domain and username. | 'mydomain\username' |
| Password | string | Password of user. | |
| Throw Error on failure | bool | True: Throws error if PDF writer Task fails. False: Returns object { Success = false } if Task fails. | true |


### Result
| Property             | Type                 | Description                          | Example |
| ---------------------| ---------------------| ------------------------------------ | ----- |
| Success | bool | Task execution result status. | true |
| Message | string | Error message (if Options.ThrowErrorOnFailure is false) | |
| FileName | string | Full path to PDF document created. | c:\output\example_pdf.pdf |


# License

This project is licensed under the MIT License - see the LICENSE file for details

# Building

Clone a copy of the repo

`git clone https://github.com/CommunityHiQ/Frends.Community.PDFWriter.git`

Restore dependencies

`nuget restore frends.community.pdfwriter`

Rebuild the project

Run Tests with nunit3. Tests can be found under

`Frends.Community.PDFWriter.Tests\bin\Release\Frends.Community.PDFWriter.Tests.dll`

Create a nuget package

`nuget pack nuspec/Frends.Community.PDFWriter.nuspec`

# Contributing
When contributing to this repository, please first discuss the change you wish to make via issue, email, or any other method with the owners of this repository before making a change.

1. Fork the repo on GitHub
2. Clone the project to your own machine
3. Commit changes to your own branch
4. Push your work back up to your fork
5. Submit a Pull request so that we can review your changes

NOTE: Be sure to merge the latest from "upstream" before making a pull request!

# Change Log

| Version             | Changes                 |
| ---------------------| ---------------------|
| 1.0.0 | Initial version of PDFWriter |
| 1.1.0 | Fixed bug where multiple whitespaces were trimmed. |
| 1.2.0 | Renamed task class to PDFWriterTask, added ChangeLog to README |
| 1.3.0 | Changed Frends.Task.Attributes from 1.2.1 to 1.2.0 because of a bug in the newer version |
| 1.4.0 | Changed target .net framework to 4.5.2 |
| 1.5.0 | Changed Unicode default value to true |
| 1.6.0 | System.ComponentModel is now used instead of Frends.Tasks.Attributes |
