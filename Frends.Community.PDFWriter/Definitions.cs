using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
#pragma warning disable 1591

namespace Frends.Community.PDFWriter
{
    public enum FileExistsActionEnum { Error, Overwrite, Rename };
    public enum PageSizeEnum { A0, A1, A2, A3, A4, A5, A6, B5, Ledger, Legal, Letter };
    public enum PageOrientationEnum { Portrait, Landscape };
    public enum ElementType { Paragraph, Image, PageBreak, Header, Footer, Table };
    public enum ParagraphAlignmentEnum { Left, Center, Justify, Right };
    public enum ImageAlignmentEnum { Left, Center, Right };
    public enum FontStyleEnum { Regular, Bold, Italic, BoldItalic, Underline };
    public enum HeaderFooterStyleEnum { Text, TextPagenum, LogoText, LogoTextPagenum };

    public class FileProperties
    {
        /// <summary>
        /// PDF document destination Directory
        /// </summary>
        [DisplayFormat(DataFormatString = "Text")]
        [DefaultValue(@"C:\Output")]
        public string Directory { get; set; }

        /// <summary>
        /// Filename for created PDF file
        /// </summary>
        [DisplayFormat(DataFormatString = "Text")]
        [DefaultValue("example_file.pdf")]
        public string FileName { get; set; }

        /// <summary>
        /// What to do if destination file already exists
        /// </summary>
        [DefaultValue(FileExistsActionEnum.Error)]
        public FileExistsActionEnum FileExistsAction { get; set; }

        /// <summary>
        /// Use Unicode text (true) or ANSI (false).
        /// </summary>
        [DefaultValue(true)]
        public bool Unicode { get; set; }
    }

    public class DocumentSettings
    {
        /// <summary>
        /// Optional PDF document title
        /// </summary>
        [DisplayFormat(DataFormatString = "Text")]
        public string Title { get; set; }

        /// <summary>
        /// Optional PDF document Author
        /// </summary>
        [DisplayFormat(DataFormatString = "Text")]
        public string Author { get; set; }

        /// <summary>
        /// Document page size
        /// </summary>
        [DefaultValue(PageSizeEnum.A4)]
        public PageSizeEnum Size { get; set; }

        [DefaultValue(PageOrientationEnum.Portrait)]
        public PageOrientationEnum Orientation { get; set; }

        /// <summary>
        /// Page margin left in cm
        /// </summary>
        [DefaultValue(2.5)]
        public double MarginLeftInCm { get; set; }

        /// <summary>
        /// Page margin top in cm
        /// </summary>
        [DefaultValue(2)]
        public double MarginTopInCm { get; set; }

        /// <summary>
        /// Page margin right in cm
        /// </summary>
        [DefaultValue(2.5)]
        public double MarginRightInCm { get; set; }

        /// <summary>
        /// Page margin bottom in cm
        /// </summary>
        [DefaultValue(2)]
        public double MarginBottomInCm { get; set; }

    }

    public class DocumentContent
    {
        /// <summary>
        /// Document content
        /// </summary>
        [DisplayName("Document content")]
        public PageContentElement[] Contents { get; set; }
    }

    public class PageContentElement
    {
        [DefaultValue(ElementType.Paragraph)]
        public ElementType ContentType { get; set; }

        /// <summary>
        /// Full path to image
        /// </summary>
        [UIHint(nameof(ContentType), "", ElementType.Image, ElementType.Header, ElementType.Footer)]
        public string ImagePath { get; set; }

        /// <summary>
        /// Text written to document
        /// </summary>
        [UIHint(nameof(ContentType), "", ElementType.Paragraph, ElementType.Header, ElementType.Footer)]
        [DisplayFormat(DataFormatString = "Text")]
        public string Text { get; set; }

        /// <summary>
        /// Font family name
        /// </summary>
        [UIHint(nameof(ContentType), "", ElementType.Paragraph, ElementType.Header, ElementType.Footer)]
        [DisplayFormat(DataFormatString = "Text")]
        [DefaultValue("Times New Roman")]
        public string FontFamily { get; set; }

        /// <summary>
        /// Font size in pt
        /// </summary>
        [UIHint(nameof(ContentType), "", ElementType.Paragraph, ElementType.Header, ElementType.Footer)]
        [DefaultValue(11)]
        public int FontSize { get; set; }

        /// <summary>
        /// Font style
        /// </summary>
        [UIHint(nameof(ContentType), "", ElementType.Paragraph, ElementType.Header, ElementType.Footer)]
        [DefaultValue(FontStyleEnum.Regular)]
        public FontStyleEnum FontStyle { get; set; }

        /// <summary>
        /// Space between lines
        /// </summary>
        [UIHint(nameof(ContentType), "", ElementType.Paragraph, ElementType.Header, ElementType.Footer)]
        [DefaultValue(14)]
        public int LineSpacingInPt { get; set; }


        [UIHint(nameof(ContentType), "", ElementType.Paragraph, ElementType.Header, ElementType.Footer)]
        [DefaultValue(ParagraphAlignmentEnum.Left)]
        [DisplayName("Alignment")]
        public ParagraphAlignmentEnum ParagraphAlignment { get; set; }

        [UIHint(nameof(ContentType), "", ElementType.Image)]
        [DefaultValue(ImageAlignmentEnum.Left)]
        [DisplayName("Alignment")]
        public ImageAlignmentEnum ImageAlignment { get; set; }
        
        /// <summary>
        /// Amount of space added above this element in pt
        /// </summary>
        [UIHint(nameof(ContentType), "", ElementType.Image, ElementType.Paragraph, ElementType.Header, ElementType.Footer)]
        [DefaultValue(8)]
        public int SpacingBeforeInPt { get; set; }

        /// <summary>
        /// Amount of space added after this element in pt
        /// </summary>
        [UIHint(nameof(ContentType), "", ElementType.Image, ElementType.Paragraph, ElementType.Header, ElementType.Footer)]
        [DefaultValue(0)]
        public int SpacingAfterInPt { get; set; }

        /// <summary>
        /// Header or footer type: only text, or additional graphics and/or pagenumbers
        /// </summary>
        [UIHint(nameof(ContentType), "", ElementType.Header, ElementType.Footer)]
        [DefaultValue(HeaderFooterStyleEnum.Text)]
        public HeaderFooterStyleEnum HeaderFooterStyle { get; set; }

        /// <summary>
        /// Width of header's lower (or footer's upper) border line in pt
        /// </summary>
        [UIHint(nameof(ContentType), "", ElementType.Header, ElementType.Footer)]
        [DefaultValue(0.0)]
        public double BorderWidthInPt { get; set; }

        /// <summary>
        /// Height of the header/footer graphics in cm. Image's aspect ratio is preserved when scaling
        /// </summary>
        [UIHint(nameof(ContentType), "", ElementType.Header, ElementType.Footer)]
        [DefaultValue(2.5)]
        public double ImageHeightInCm { get; set; }


        [UIHint(nameof(ContentType), "", ElementType.Table)]
        [DisplayFormat(DataFormatString = "Json")]
        [DefaultValue("{}")]
        public string Table { get; set; }
    }

    public class Options
    {
        /// <summary>
        /// If set, allows you to give the user credentials to use to write the PDF file on remote hosts.
        /// If not set, the agent service user credentials will be used.
        /// </summary>
        [DefaultValue(false)]
        public bool UseGivenCredentials { get; set; }

        /// <summary>
        /// This needs to be of format domain\username
        /// </summary>
        [UIHint(nameof(UseGivenCredentials), "", true)]
        [DisplayFormat(DataFormatString = "Text")]
        [DefaultValue(@"domain\username")]
        public string UserName { get; set; }

        [PasswordPropertyText]
        [UIHint(nameof(UseGivenCredentials), "", true)]
        public string Password { get; set; }

        /// <summary>
        /// True: Throws error on failure
        /// False: Returns object{ Success = false }
        /// </summary>
        [DefaultValue(true)]
        public bool ThrowErrorOnFailure { get; set; }
    }

    public class Output
    {
        public bool Success { get; set; }

        public string FileName { get; set; }
    }

    
    public enum TableTypeEnum { Table, Header, Footer };
    public enum TableColumnType { Text, Image, PageNum };
    public enum TableBorderStyle { None, Top, Bottom, All };

    public class TableColumnDefinition
    {
        public string Name { get; set; }
        public double WidthInCm { get; set; }
        public double HeightInCm { get; set; }
        public TableColumnType Type { get; set; }
    }

    public class TableStyle
    {
        public string FontFamily { get; set; }
        public double FontSizeInPt { get; set; }
        public FontStyleEnum FontStyle { get; set; }
        public double LineSpacingInPt { get; set; }
        public double SpacingBeforeInPt { get; set; }
        public double SpacingAfterInPt { get; set; }
        public double BorderWidthInPt { get; set; }
        public TableBorderStyle BorderStyle { get; set; }
    }
    public class TableDefinition
    {
        public bool HasHeaderRow { get; set; }
        public TableTypeEnum TableType { get; set; }
        public TableStyle StyleSettings { get; set; }
        public List<TableColumnDefinition> Columns { get; set; }
        public List<Dictionary<string, string>> RowData { get; set; }
    }
}
