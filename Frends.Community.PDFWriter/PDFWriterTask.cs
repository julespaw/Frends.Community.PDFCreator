using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Shapes;
using MigraDoc.Rendering;
using PdfSharp.Pdf;
using SimpleImpersonation;
using System;
using System.IO;
using System.ComponentModel;
using MigraDoc.DocumentObjectModel.Tables;
using System.Collections.Generic;
using Newtonsoft.Json;
using System.Linq;

#pragma warning disable 1591


namespace Frends.Community.PDFWriter
{
    public class PDFWriterTask
    {
        /// <summary>
        /// Created PDF document from given content. See https://github.com/CommunityHiQ/Frends.Community.PDFWriter
        /// </summary>
        /// <param name="outputFile"></param>
        /// <param name="documentSettings"></param>
        /// <param name="content"></param>
        /// <param name="options"></param>
        /// <returns>Object { bool Success, string FileName }</returns>
        public static Output CreatePdf([PropertyTab]FileProperties outputFile,
            [PropertyTab]DocumentSettings documentSettings,
            [PropertyTab]DocumentContent content,
            [PropertyTab]Options options)
        {
            try
            {
                var document = new Document();
                if (!string.IsNullOrWhiteSpace(documentSettings.Title))
                {
                    document.Info.Title = documentSettings.Title;
                }
                if (!string.IsNullOrWhiteSpace(documentSettings.Author))
                {
                    document.Info.Author = documentSettings.Author;
                }

                // Get the selected page size
                Unit width, height;
                PageSetup.GetPageSize(documentSettings.Size.ConvertEnum<PageFormat>(), out width, out height);

                var section = document.AddSection();
                SetupPage(section.PageSetup, width, height, documentSettings);

                // index for stylename
                var elementNumber = 0;
                // add page elements
                foreach (var pageElement in content.Contents)
                {
                    var styleName = $"style_{elementNumber}";
                    var style = document.Styles.AddStyle(styleName, "Normal");
                    switch (pageElement.ContentType)
                    {
                        case ElementType.Image:
                            AddImage(section, pageElement, width, style);
                            break;
                        case ElementType.PageBreak:
                            section = document.AddSection();
                            SetupPage(section.PageSetup, width, height, documentSettings);
                            break;
                        case ElementType.Header:
                            SetFont(style, pageElement);
                            SetParagraphStyle(style, pageElement);
                            AddHeaderFooterContent(section, pageElement, style, true);
                            break;
                        case ElementType.Footer:
                            SetFont(style, pageElement);
                            SetParagraphStyle(style, pageElement);
                            AddHeaderFooterContent(section, pageElement, style, false);
                            break;
                        case ElementType.Table:
                            AddTable(section, pageElement);
                            break;
                        default:
                            SetFont(style, pageElement);
                            SetParagraphStyle(style, pageElement);
                            AddTextContent(section, pageElement, style);
                            break;
                    }

                    elementNumber++;
                }

                string fileName = Path.Combine(outputFile.Directory, outputFile.FileName);
                int fileNameIndex = 1;
                while (File.Exists(fileName) && outputFile.FileExistsAction != FileExistsActionEnum.Overwrite)
                {
                    switch (outputFile.FileExistsAction)
                    {
                        case FileExistsActionEnum.Error:
                            throw new Exception($"File {fileName} already exists.");
                        case FileExistsActionEnum.Rename:
                            fileName = Path.Combine(outputFile.Directory, $"{Path.GetFileNameWithoutExtension(outputFile.FileName)}_({fileNameIndex}){Path.GetExtension(outputFile.FileName)}");
                            break;
                    }
                    fileNameIndex++;
                }
                // save document

                var pdfRenderer = new PdfDocumentRenderer(outputFile.Unicode)
                {
                    Document = document
                };
                
                pdfRenderer.RenderDocument();

                if(!options.UseGivenCredentials)
                    pdfRenderer.PdfDocument.Save(fileName);
                else
                {
                    var domainAndUserName = GetDomainAndUserName(options.UserName);
                    using(Impersonation.LogonUser(domainAndUserName[0], domainAndUserName[1], options.Password, LogonType.NewCredentials))
                    {
                        pdfRenderer.PdfDocument.Save(fileName);
                    }
                }

                return new Output { Success = true, FileName = fileName };

            }
            catch (Exception ex)
            {
                if (options.ThrowErrorOnFailure)
                    throw ex;

                return new Output { Success = false };
            }
        }

        private static string[] GetDomainAndUserName(string username)
        {
            var domainAndUserName = username.Split('\\');
            if (domainAndUserName.Length != 2)
            {
                throw new ArgumentException($@"UserName field must be of format domain\username was: {username}");
            }
            return domainAndUserName;
        }

        /// <summary>
        /// Set page settings
        /// </summary>
        /// <param name="setup"></param>
        /// <param name="pageWidth"></param>
        /// <param name="pageHeight"></param>
        /// <param name="documentSettings"></param>
        private static void SetupPage(PageSetup setup, Unit pageWidth, Unit pageHeight, DocumentSettings documentSettings)
        {
            setup.Orientation = documentSettings.Orientation.ConvertEnum<Orientation>();
            setup.PageHeight = pageHeight;
            setup.PageWidth = pageWidth;
            setup.LeftMargin = new Unit(documentSettings.MarginLeftInCm, UnitType.Centimeter);
            setup.TopMargin = new Unit(documentSettings.MarginTopInCm, UnitType.Centimeter);
            setup.RightMargin = new Unit(documentSettings.MarginRightInCm, UnitType.Centimeter);
            setup.BottomMargin = new Unit(documentSettings.MarginBottomInCm, UnitType.Centimeter);
        }

        /// <summary>
        /// Set paragraph style
        /// </summary>
        /// <param name="style"></param>
        /// <param name="pageContent"></param>
        private static void SetParagraphStyle(Style style, PageContentElement pageContent)
        {
            style.ParagraphFormat.LineSpacing = new Unit(pageContent.LineSpacingInPt, UnitType.Point);
            style.ParagraphFormat.LineSpacingRule = LineSpacingRule.Exactly;
            style.ParagraphFormat.Alignment = pageContent.ParagraphAlignment.ConvertEnum<MigraDoc.DocumentObjectModel.ParagraphAlignment>();
            style.ParagraphFormat.SpaceBefore = new Unit(pageContent.SpacingBeforeInPt, UnitType.Point);
            style.ParagraphFormat.SpaceAfter = new Unit(pageContent.SpacingAfterInPt, UnitType.Point);
        }
        
        /// <summary>
        /// Adds an image to the page. Only PNG images are supported out of the box.
        /// </summary>
        /// <param name="section"></param>
        /// <param name="pageContent"></param>
        /// <param name="pageWidth"></param>
        /// <param name="style"></param>
        private static void AddImage(Section section, PageContentElement pageContent, Unit pageWidth, Style style)
        {
            Unit originalImageWidthInches;
            // work around to get image dimensions
            using (System.Drawing.Image userImage = System.Drawing.Image.FromFile(pageContent.ImagePath))
            {
                // get image width in inches
                var imageInches = userImage.Width / userImage.VerticalResolution;
                originalImageWidthInches = new Unit(imageInches, UnitType.Inch);
            }

            // add image
            Image image = section.AddImage(pageContent.ImagePath);

            // Calculate Image size: 
            // if actual image size is larger than PageWidth - margins, set image width as page width - margins
            Unit actualPageContentWidth = new Unit((pageWidth.Inch - section.PageSetup.LeftMargin.Inch - section.PageSetup.RightMargin.Inch), UnitType.Inch);
            if (originalImageWidthInches > actualPageContentWidth)
            {
                image.Width = actualPageContentWidth; 
            }
            image.LockAspectRatio = true;
            image.Left = pageContent.ImageAlignment.ConvertEnum<ShapePosition>();
        }

        private static void AddTextContent(Section section, PageContentElement pageContent, Style style)
        {
            // skip if text content if empty
            if (string.IsNullOrWhiteSpace(pageContent.Text))
            {
                return;
            }

            var paragraph = section.AddParagraph();
            paragraph.Style = style.Name;
            paragraph.Format.Font.Color = Colors.Black;

            //read text line by line
            string line;
            using(var reader = new StringReader(pageContent.Text))
            {
                while((line = reader.ReadLine()) != null)
                {
                    // read text one char at a time, so that multiple whitespaces are added correctly
                    foreach (var c in line.ToCharArray())
                    {
                        if (Char.IsWhiteSpace(c))
                            paragraph.AddSpace(1);
                        else
                            paragraph.AddChar(c, 1);
                    }
                    // add newline
                    paragraph.AddLineBreak();
                }
            }
        }

        /// <summary>
        /// Adds a header or a footer to the document. Header/footer size is optimized for A4 size, portrait oriented page.
        /// Header/footer can have a logo field, text field, and page numbers.
        /// </summary>
        /// <param name="section"></param>
        /// <param name="pageContent"></param>
        /// <param name="style"></param>
        /// <param name="isHeader"></param>
        private static void AddHeaderFooterContent(Section section, PageContentElement pageContent, Style style, bool isHeader)
        {
            // skip if text content if empty
            if (string.IsNullOrWhiteSpace(pageContent.Text))
            {
                return;
            }

            Table table;

            if (isHeader)
            {
                table = section.Headers.Primary.AddTable();
            }
            else
            {
                table = section.Footers.Primary.AddTable();
            }

            Row row;

            Paragraph textField;
            Paragraph pagenumField;

            switch (pageContent.HeaderFooterStyle)
            {
                case HeaderFooterStyleEnum.Text:
                    table.AddColumn("16cm");

                    row = table.AddRow();
                    row.VerticalAlignment = VerticalAlignment.Center;

                    textField = row.Cells[0].AddParagraph();
                    break;
                case HeaderFooterStyleEnum.TextPagenum:
                    table.AddColumn("12cm");
                    table.AddColumn("4cm");

                    row = table.AddRow();
                    row.VerticalAlignment = VerticalAlignment.Center;

                    textField = row.Cells[0].AddParagraph();

                    pagenumField = row.Cells[1].AddParagraph();
                    FormatPagenumField(style, pagenumField);
                    break;
                case HeaderFooterStyleEnum.LogoText:
                    table.AddColumn("5cm");
                    table.AddColumn("11cm");

                    row = table.AddRow();
                    row.VerticalAlignment = VerticalAlignment.Center;

                    FormatHeaderFooterLogo(pageContent, row);

                    textField = row.Cells[1].AddParagraph();
                    break;
                case HeaderFooterStyleEnum.LogoTextPagenum:
                    table.AddColumn("5cm");
                    table.AddColumn("7cm");
                    table.AddColumn("4cm");

                    row = table.AddRow();
                    row.VerticalAlignment = VerticalAlignment.Center;

                    FormatHeaderFooterLogo(pageContent, row);

                    textField = row.Cells[1].AddParagraph();

                    pagenumField = row.Cells[2].AddParagraph();
                    FormatPagenumField(style, pagenumField);
                    break;
                default:
                    throw new Exception($"Cannot insert header without proper style choice.");
            }

            textField.Style = style.Name;
            textField.Format.Font.Color = Colors.Black;
            textField.AddText(pageContent.Text);

            if (pageContent.BorderWidthInPt > 0 && isHeader)
            {
                table.Borders.Bottom.Width = new Unit(pageContent.BorderWidthInPt, UnitType.Point);
            }
            else if (pageContent.BorderWidthInPt > 0 && !isHeader)
            {
                table.Borders.Top.Width = new Unit(pageContent.BorderWidthInPt, UnitType.Point);
            }
        }


        private static void AddTable(Section section, PageContentElement pageContent)
        {
            TableDefinition tableData = JsonConvert.DeserializeObject<TableDefinition>(pageContent.Table);

            Table table;

            switch (tableData.TableType)
            {
                case TableTypeEnum.Header:
                    table = section.Headers.Primary.AddTable();
                    break;
                case TableTypeEnum.Footer:
                    table = section.Footers.Primary.AddTable();
                    break;
                default:
                    table = section.AddTable();
                    break;
            }

            foreach (var column in tableData.Columns)
            {
                table.AddColumn(column.WidthInCm.ToString() + "cm");
            }
            

            if (tableData.HasHeaderRow)
            {
                var columnHeaders = tableData.Columns.Select(column => column.Name).ToList();
                var headerColumnDefinitions = new List<TableColumnDefinition>();
                for (int i = 0; i < columnHeaders.Count; i++)
                {
                    headerColumnDefinitions.Add(new TableColumnDefinition { Type = TableColumnType.Text });
                }
                ProcessRow(table, headerColumnDefinitions, columnHeaders, tableData.StyleSettings);
            }

            foreach (var dataRow in tableData.RowData)
            {
                var data = dataRow.Select(row => row.Value).ToList();
                ProcessRow(table, tableData.Columns, data, tableData.StyleSettings);
            }

            if (tableData.StyleSettings.BorderWidthInPt > 0)
            {
                switch (tableData.StyleSettings.BorderStyle)
                {
                    case TableBorderStyle.Top:
                        table.Borders.Top.Width = new Unit(tableData.StyleSettings.BorderWidthInPt, UnitType.Point);
                        break;
                    case TableBorderStyle.Bottom:
                        table.Borders.Bottom.Width = new Unit(tableData.StyleSettings.BorderWidthInPt, UnitType.Point);
                        break;
                    case TableBorderStyle.All:
                        table.Borders.Width = new Unit(tableData.StyleSettings.BorderWidthInPt, UnitType.Point);
                        break;
                    case TableBorderStyle.None:
                        break;
                    default:
                        break;
                }
            }
        }

        private static void ProcessRow(Table table, List<TableColumnDefinition> columns, List<string> data, TableStyle style)
        {
            var row = table.AddRow();
            row.VerticalAlignment = VerticalAlignment.Center;

            for (int i = 0; i < data.Count; i++)
            {
                switch (columns[i].Type)
                {
                    case TableColumnType.Text:
                        var textField = row.Cells[i].AddParagraph();
                        SetParagraphStyle(textField, style);
                        textField.AddText(data[i]);
                        break;
                    case TableColumnType.Image:
                        if (string.IsNullOrWhiteSpace(data[i]) || !File.Exists(data[i]))
                        {
                            throw new FileNotFoundException($"Path to header graphics was empty or the file does not exist.");
                        }

                        var logo = row.Cells[i].AddImage(data[i]);
                        logo.Height = new Unit(columns[i].HeightInCm, UnitType.Centimeter);
                        logo.LockAspectRatio = true;
                        logo.Top = ShapePosition.Top;
                        logo.Left = ShapePosition.Left;
                        break;
                    case TableColumnType.PageNum:
                        var pagenumField = row.Cells[i].AddParagraph();
                        SetParagraphStyle(pagenumField, style);
                        pagenumField.AddPageField();
                        pagenumField.AddText(" (");
                        pagenumField.AddNumPagesField();
                        pagenumField.AddText(")");
                        break;
                }
            }
        }

        private static void SetParagraphStyle(Paragraph pg, TableStyle style)
        {
            pg.Format.Font.Color = Colors.Black;
            pg.Format.Font.Name = style.FontFamily;
            pg.Format.Font.Size = new Unit(style.FontSizeInPt, UnitType.Point);
            
            switch (style.FontStyle)
            {
                case FontStyleEnum.Bold:
                    pg.Format.Font.Bold = true;
                    break;
                case FontStyleEnum.Italic:
                    pg.Format.Font.Italic = true;
                    break;
                case FontStyleEnum.BoldItalic:
                    pg.Format.Font.Bold = true;
                    pg.Format.Font.Italic = true;
                    break;
                case FontStyleEnum.Underline:
                    pg.Format.Font.Underline = Underline.Single;
                    break;
            }

            pg.Format.LineSpacing = new Unit(style.LineSpacingInPt, UnitType.Point);
            pg.Format.SpaceBefore = new Unit(style.SpacingBeforeInPt, UnitType.Point);
            pg.Format.SpaceAfter = new Unit(style.SpacingAfterInPt, UnitType.Point);
        }

        /// <summary>
        /// Helper method to set styles for header/footer graphics.
        /// </summary>
        /// <param name="pageContent"></param>
        /// <param name="row"></param>
        private static void FormatHeaderFooterLogo(PageContentElement pageContent, Row row)
        {
            if (string.IsNullOrWhiteSpace(pageContent.ImagePath) || !File.Exists(pageContent.ImagePath))
            {
                throw new FileNotFoundException($"Path to header graphics was empty or the file does not exist.");
            }

            var logo = row.Cells[0].AddImage(pageContent.ImagePath);
            logo.Height = new Unit(pageContent.ImageHeightInCm, UnitType.Centimeter);
            logo.LockAspectRatio = true;
            logo.Top = ShapePosition.Top;
            logo.Left = ShapePosition.Left;
        }

        /// <summary>
        /// Helper method to set styles for page numbers in header/footer.
        /// Page numbers are always aligned to the right.
        /// </summary>
        /// <param name="style"></param>
        /// <param name="pagenumField"></param>
        private static void FormatPagenumField(Style style, Paragraph pagenumField)
        {
            pagenumField.Style = style.Name;
            pagenumField.Format.Alignment = ParagraphAlignment.Right;
            pagenumField.Format.Font.Color = Colors.Black;
            pagenumField.AddPageField();
            pagenumField.AddText(" (");
            pagenumField.AddNumPagesField();
            pagenumField.AddText(")");
        }

        private static void SetFont(Style style, PageContentElement textElement)
        {
            style.Font.Name = textElement.FontFamily;
            style.Font.Size = new Unit(textElement.FontSize, UnitType.Point);

            switch (textElement.FontStyle)
            {
                case FontStyleEnum.Bold:
                    style.Font.Bold = true;
                    break;
                case FontStyleEnum.BoldItalic:
                    style.Font.Bold = true;
                    style.Font.Italic = true;
                    break;
                case FontStyleEnum.Italic:
                    style.Font.Italic = true;
                    break;
                case FontStyleEnum.Underline:
                    style.Font.Underline = Underline.Single;
                    break;
            }
        }
    }


}
