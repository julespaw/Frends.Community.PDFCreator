using Frends.Tasks.Attributes;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Shapes;
using MigraDoc.Rendering;
using PdfSharp.Pdf;
using SimpleImpersonation;
using System;
using System.IO;

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
        public static Output CreatePdf([CustomDisplay(DisplayOption.Tab)]FileProperties outputFile,
            [CustomDisplay(DisplayOption.Tab)]DocumentSettings documentSettings,
            [CustomDisplay(DisplayOption.Tab)]DocumentContent content,
            [CustomDisplay(DisplayOption.Tab)]Options options)
        {
            try
            {
                var document = new Document();
                if (!string.IsNullOrWhiteSpace(documentSettings.Title))
                    document.Info.Title = documentSettings.Title;
                if (!string.IsNullOrWhiteSpace(documentSettings.Author))
                    document.Info.Author = documentSettings.Author;

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
                var pdfRenderer = new PdfDocumentRenderer(outputFile.Unicode, PdfFontEmbedding.Always)
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
                image.Width = actualPageContentWidth;
            image.LockAspectRatio = true;
            image.Left = pageContent.ImageAlignment.ConvertEnum<ShapePosition>();
        }

        private static void AddTextContent(Section section, PageContentElement pageContent, Style style)
        {
            // skip if text content if empty
            if (string.IsNullOrWhiteSpace(pageContent.Text))
                return;

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
