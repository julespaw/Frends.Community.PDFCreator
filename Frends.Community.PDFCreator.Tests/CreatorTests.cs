using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;

namespace Frends.Community.PDFCreator.Tests
{
    [TestFixture]
    public class CreatorTests
    {
        private FileProperties _fileProperties;
        private DocumentSettings _docSettings;

        private PageContentElement _header;
        private PageContentElement _footer;
        private PageContentElement _title;
        private PageContentElement _paragraphContent;
        private PageContentElement _tableContent;

        private Options _options;

        private string _destinationFullPath;
        private readonly string _fileName = "test_output.pdf";
        private string _folder;

        [SetUp]
        public void TestSetup()
        {
            _folder = Path.Combine(Path.GetTempPath(), "pdfcreator_tests");
            _destinationFullPath = Path.Combine(_folder, _fileName);

            if (!Directory.Exists(_folder))
            {
                Directory.CreateDirectory(_folder);
            }

            _fileProperties = new FileProperties { Directory = _folder, FileName = _fileName, FileExistsAction = FileExistsActionEnum.Error, Unicode = true };
            _docSettings = new DocumentSettings { MarginBottomInCm = 2, MarginLeftInCm = 2.5, MarginRightInCm = 2.5, MarginTopInCm = 5, Orientation = PageOrientationEnum.Portrait, Size = PageSizeEnum.A4 };

            var logoPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Files\logo.png");
            _header = new PageContentElement { ContentType = ElementType.Header, FontFamily = "Times New Roman", FontSize = 8, FontStyle = FontStyleEnum.Regular, LineSpacingInPt = 11, ParagraphAlignment = ParagraphAlignmentEnum.Right, SpacingAfterInPt = 0, SpacingBeforeInPt = 8, ImagePath = logoPath, HeaderFooterStyle = HeaderFooterStyleEnum.LogoText, BorderWidthInPt = 0.5, ImageHeightInCm = 0.5 };
            _footer = new PageContentElement { ContentType = ElementType.Footer, FontFamily = "Times New Roman", FontSize = 8, FontStyle = FontStyleEnum.Regular, LineSpacingInPt = 11, ParagraphAlignment = ParagraphAlignmentEnum.Center, SpacingAfterInPt = 0, SpacingBeforeInPt = 8, HeaderFooterStyle = HeaderFooterStyleEnum.TextPagenum, BorderWidthInPt = 0.0 };
            _title = new PageContentElement { ContentType = ElementType.Paragraph, FontFamily = "Times New Roman", FontSize = 16, FontStyle = FontStyleEnum.Bold, LineSpacingInPt = 11, ParagraphAlignment = ParagraphAlignmentEnum.Left, SpacingAfterInPt = 0, SpacingBeforeInPt = 8 };
            _paragraphContent = new PageContentElement { ContentType = ElementType.Paragraph, FontFamily = "Times New Roman", FontSize = 11, FontStyle = FontStyleEnum.Regular, LineSpacingInPt = 11, ParagraphAlignment = ParagraphAlignmentEnum.Left, SpacingAfterInPt = 0, SpacingBeforeInPt = 8 };

            var tablePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Files\ContentDefinition.json");
            var tableDefinition = File.ReadAllText(tablePath);
            _tableContent = new PageContentElement { ContentType = ElementType.Table, Table = tableDefinition };

            _options = new Options { UseGivenCredentials = false, ThrowErrorOnFailure = true };
        }

        [TearDown]
        public void TestTearDown()
        {
            Directory.Delete(_folder, true);
        }

        private Output CallCreatePdf(PageContentElement[] contents)
        {
            return PDFCreatorTask.CreatePdf(_fileProperties, _docSettings, new DocumentContent { Contents = contents }, _options);
        }

        [Test]
        public void WritePDF()
        {
            _fileProperties.FileExistsAction = FileExistsActionEnum.Overwrite;

            _header.Text = @"This is a header";
            _footer.Text = @"This is a footer";
            _title.Text = @"The Document Title";
            _paragraphContent.Text = @"Some text           for testing
with some tab
    one
        two
            three. Then end with scandic letter ö and russian word код.";

            var result = CallCreatePdf(new PageContentElement[] { _header, _footer, _title, _paragraphContent, new PageContentElement { ContentType = ElementType.PageBreak }, _tableContent });

            Assert.IsTrue(File.Exists(_destinationFullPath));
            Assert.IsTrue(result.Success);
        }

        [Test]
        public void WritePdf_DoesNotFailIfContentIsEmpty()
        {
            _paragraphContent.Text = null;
            var result = CallCreatePdf(new PageContentElement[] { _paragraphContent });

            Assert.IsTrue(File.Exists(_destinationFullPath));
            Assert.IsTrue(result.Success);
        }

        [Test]
        public void WritePdf_ThrowsExceptionIfFileExists()
        {
            _options.ThrowErrorOnFailure = true;
            _fileProperties.FileExistsAction = FileExistsActionEnum.Error;

            //run once so file exists
            CallCreatePdf(new PageContentElement[] { _paragraphContent });

            var result = Assert.Throws<Exception>(() => CallCreatePdf(new PageContentElement[] { _paragraphContent }));
        }

        [Test]
        public void WritePdf_RenamesFilesIfAlreadyExists()
        {
            _fileProperties.FileExistsAction = FileExistsActionEnum.Rename;

            // Create 3 pdf's
            var result1 = CallCreatePdf(new PageContentElement[] { _paragraphContent });
            var result2 = CallCreatePdf(new PageContentElement[] { _paragraphContent });
            var result3 = CallCreatePdf(new PageContentElement[] { _paragraphContent });

            Assert.IsTrue(File.Exists(result1.FileName));
            Assert.AreEqual(_destinationFullPath, result1.FileName);
            Assert.IsTrue(File.Exists(result2.FileName));
            Assert.IsTrue(result2.FileName.Contains("_(1)"));
            Assert.IsTrue(File.Exists(result3.FileName));
            Assert.IsTrue(result3.FileName.Contains("_(2)"));
        }

        [Test]
        public void WritePdf_LogoNotFound()
        {
            _fileProperties.FileExistsAction = FileExistsActionEnum.Overwrite;

            var logoPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Files\no_such_logo.png");
            var header = new PageContentElement { ContentType = ElementType.Header, FontFamily = "Times New Roman", FontSize = 8, FontStyle = FontStyleEnum.Regular, LineSpacingInPt = 11, ParagraphAlignment = ParagraphAlignmentEnum.Right, SpacingAfterInPt = 0, SpacingBeforeInPt = 8, ImagePath = logoPath, HeaderFooterStyle = HeaderFooterStyleEnum.LogoText, BorderWidthInPt = 0.5, ImageHeightInCm = 0.5 };
            header.Text = @"This is a header";

            var result = Assert.Throws<FileNotFoundException>(() => CallCreatePdf(new PageContentElement[] { header }));
            Assert.IsFalse(File.Exists(_destinationFullPath));
        }

        [Test]
        public void WritePdf_TableWidthTooLarge()
        {
            var tooWideTable = @"{ ""HasHeaderRow"": true, ""TableType"": ""Table"", ""Columns"": [ { ""Name"": ""Sarake 1"", ""WidthInCm"": 21, ""HeightInCm"": 0, ""Type"": ""Text"" } ], ""RowData"": [] }";
            var table1 = new PageContentElement { ContentType = ElementType.Table, Table = tooWideTable };
            var result1 = Assert.Throws<Exception>(() => CallCreatePdf(new PageContentElement[] { table1 }));
            Assert.IsFalse(File.Exists(_destinationFullPath));
        }
    }
}
