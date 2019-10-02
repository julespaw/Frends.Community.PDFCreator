using NUnit.Framework;
using System;
using System.IO;

namespace Frends.Community.PDFWriter.Tests
{
    [TestFixture]
    public class WriterTests
    {
        private FileProperties _fileProperties;
        private DocumentSettings _docSettings;
        private PageContentElement _pageContent;
        private Options _options;

        private string _destinationFullPath;
        private readonly string _fileName = "test_output.pdf";
        private string _folder;

        [SetUp]
        public void TestSetup()
        {
            _folder = Path.Combine(Path.GetTempPath(), "pdfwriter_tests");
            _destinationFullPath = Path.Combine(_folder, _fileName);

            if (!Directory.Exists(_folder))
                Directory.CreateDirectory(_folder);

            _fileProperties = new FileProperties { Directory = _folder, FileName = _fileName, FileExistsAction = FileExistsActionEnum.Error, Unicode = true };
            _docSettings = new DocumentSettings { MarginBottomInCm = 2, MarginLeftInCm = 2.5, MarginRightInCm = 2.5, MarginTopInCm = 5, Orientation = PageOrientationEnum.Portrait, Size = PageSizeEnum.A4 };
            _pageContent = new PageContentElement { ContentType = ElementType.Paragraph, FontFamily = "Times New Roman", FontSize = 11, FontStyle = FontStyleEnum.Regular, LineSpacingInPt = 11, ParagraphAlignment = ParagraphAlignmentEnum.Left, SpacingAfterInPt = 0, SpacingBeforeInPt = 8 };
            _options = new Options { UseGivenCredentials = false, ThrowErrorOnFailure = true };
        }

        [TearDown]
        public void TestTearDown()
        {
            Directory.Delete(_folder, true);
        }

        private Output CallCreatePdf()
        {
            return PDFWriterTask.CreatePdf(_fileProperties, _docSettings, new DocumentContent { Contents = new[] { _pageContent } }, _options);
        }

        [Test]
        public void WritePDF()
        {
            _fileProperties.FileExistsAction = FileExistsActionEnum.Overwrite;
            _pageContent.Text = @"Some text           for testing
with some tab
    one
        two
            three. Then end with scandic letter ö and russian word код.";

            var result = CallCreatePdf();

            Assert.IsTrue(File.Exists(_destinationFullPath));
            Assert.IsTrue(result.Success);
        }

        [Test]
        public void WritePdf_DoesNotFailIfContentIsEmpty()
        {
            _pageContent.Text = null;
            var result = CallCreatePdf();

            Assert.IsTrue(File.Exists(_destinationFullPath));
            Assert.IsTrue(result.Success);
        }

        [Test]
        public void WritePdf_ThrowsExceptionIfFileExists()
        {
            _options.ThrowErrorOnFailure = true;
            _fileProperties.FileExistsAction = FileExistsActionEnum.Error;

            //run once so file exists
            CallCreatePdf();

            var result = Assert.Throws<Exception>(() => CallCreatePdf());
        }

        [Test]
        public void WritePdf_RenamesFilesIfAlreadyExists()
        {
            _fileProperties.FileExistsAction = FileExistsActionEnum.Rename;

            // Create 3 pdf's
            var result1 = CallCreatePdf();
            var result2 = CallCreatePdf();
            var result3 = CallCreatePdf();

            Assert.IsTrue(File.Exists(result1.FileName));
            Assert.AreEqual(_destinationFullPath, result1.FileName);
            Assert.IsTrue(File.Exists(result2.FileName));
            Assert.IsTrue(result2.FileName.Contains("_(1)"));
            Assert.IsTrue(File.Exists(result3.FileName));
            Assert.IsTrue(result3.FileName.Contains("_(2)"));
        }

        [Test]
        public void WritePdf_CreateHeaderAndFooter()
        {
            _fileProperties.FileExistsAction = FileExistsActionEnum.Overwrite;

            var logoPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Files\logo.png");
            var header = new PageContentElement { ContentType = ElementType.Header, FontFamily = "Times New Roman", FontSize = 8, FontStyle = FontStyleEnum.Regular, LineSpacingInPt = 11, ParagraphAlignment = ParagraphAlignmentEnum.Right, SpacingAfterInPt = 0, SpacingBeforeInPt = 8, ImagePath = logoPath, HeaderFooterStyle = HeaderFooterStyleEnum.LogoText, BorderWidthInPt = 0.5, ImageHeightInCm = 0.5 };
            header.Text = @"Company, Inc. Official Document";

            var footer = new PageContentElement { ContentType = ElementType.Footer, FontFamily = "Times New Roman", FontSize = 8, FontStyle = FontStyleEnum.Regular, LineSpacingInPt = 11, ParagraphAlignment = ParagraphAlignmentEnum.Center, SpacingAfterInPt = 0, SpacingBeforeInPt = 8, HeaderFooterStyle = HeaderFooterStyleEnum.TextPagenum, BorderWidthInPt = 0.0};
            footer.Text = @"Company, Inc. 2019";
            
            var title = new PageContentElement { ContentType = ElementType.Paragraph, FontFamily = "Times New Roman", FontSize = 16, FontStyle = FontStyleEnum.Bold, LineSpacingInPt = 11, ParagraphAlignment = ParagraphAlignmentEnum.Left, SpacingAfterInPt = 0, SpacingBeforeInPt = 8 };
            title.Text = "An Official Document";

            var paragraphContent = new PageContentElement { ContentType = ElementType.Paragraph, FontFamily = "Times New Roman", FontSize = 11, FontStyle = FontStyleEnum.Regular, LineSpacingInPt = 11, ParagraphAlignment = ParagraphAlignmentEnum.Left, SpacingAfterInPt = 0, SpacingBeforeInPt = 8 };
            
            paragraphContent.Text = @"Insert official text here.

Line breaks create new text blocks.

Espoo, 1.10.2019";

            var pageBreak = new PageContentElement { ContentType = ElementType.PageBreak };

            var result = PDFWriterTask.CreatePdf(_fileProperties, _docSettings, new DocumentContent { Contents = new[] { header, footer, title, paragraphContent, pageBreak} }, _options);

            Assert.IsTrue(File.Exists(_destinationFullPath));
            Assert.IsTrue(result.Success);
        }
    }
}
