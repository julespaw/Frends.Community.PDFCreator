using Frends.Tasks.Attributes;
using System.ComponentModel;

#pragma warning disable 1591

namespace Frends.Community.PDFWriter
{
    public enum FileExistsActionEnum { Error, Overwrite, Rename };
    public enum PageSizeEnum { A0, A1, A2, A3, A4, A5, A6, B5, Ledger, Legal, Letter };
    public enum PageOrientationEnum { Portrait, Landscape };
    public enum ElementType { Paragraph, Image, PageBreak };
    public enum ParagraphAlignmentEnum { Left, Center, Justify, Right };
    public enum ImageAlignmentEnum { Left, Center, Right };

    public enum FontStyleEnum { Regular, Bold, Italic, BoldItalic, Underline };

    public class FileProperties
    {
        /// <summary>
        /// PDF document destination Directory
        /// </summary>
        [DefaultDisplayType(DisplayType.Text)]
        [DefaultValue(@"C:\Output")]
        public string Directory { get; set; }

        /// <summary>
        /// Filename for created PDF file
        /// </summary>
        [DefaultDisplayType(DisplayType.Text)]
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
        [DefaultDisplayType(DisplayType.Text)]
        public string Title { get; set; }

        /// <summary>
        /// Optional PDF document Author
        /// </summary>
        [DefaultDisplayType(DisplayType.Text)]
        public string Author { get; set; }


        /// <summary>
        /// PDF file content
        /// </summary>
        //public PageProperties[] Pages { get; set; }

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
        [ConditionalDisplay(nameof(ContentType), ElementType.Image)]
        public string ImagePath { get; set; }

        /// <summary>
        /// Text written to document
        /// </summary>
        [ConditionalDisplay(nameof(ContentType), ElementType.Paragraph)]
        [DefaultDisplayType(DisplayType.Expression)]
        public string Text { get; set; }

        /// <summary>
        /// Font family name
        /// </summary>
        [ConditionalDisplay(nameof(ContentType), ElementType.Paragraph)]
        [DefaultDisplayType(DisplayType.Text)]
        [DefaultValue("Times New Roman")]
        public string FontFamily { get; set; }

        /// <summary>
        /// Font size in pt
        /// </summary>
        [ConditionalDisplay(nameof(ContentType), ElementType.Paragraph)]
        [DefaultValue(11)]
        public int FontSize { get; set; }

        /// <summary>
        /// Font style
        /// </summary>
        [ConditionalDisplay(nameof(ContentType), ElementType.Paragraph)]
        [DefaultValue(FontStyleEnum.Regular)]
        public FontStyleEnum FontStyle { get; set; }

        /// <summary>
        /// Space between lines
        /// </summary>
        [ConditionalDisplay(nameof(ContentType), ElementType.Paragraph)]
        [DefaultValue(14)]
        public int LineSpacingInPt { get; set; }


        [ConditionalDisplay(nameof(ContentType), ElementType.Paragraph)]
        [DefaultValue(ParagraphAlignmentEnum.Left)]
        [DisplayName("Alignment")]
        public ParagraphAlignmentEnum ParagraphAlignment { get; set; }

        [ConditionalDisplay(nameof(ContentType), ElementType.Image)]
        [DefaultValue(ImageAlignmentEnum.Left)]
        [DisplayName("Alignment")]
        public ImageAlignmentEnum ImageAlignment { get; set; }
        
        /// <summary>
        /// Amount of space added above this element in pt
        /// </summary>
        [ConditionalDisplay(nameof(ContentType), ElementType.Image, ElementType.Paragraph)]
        [DefaultValue(8)]
        public int SpacingBeforeInPt { get; set; }

        /// <summary>
        /// Amount of space added after this element in pt
        /// </summary>
        [ConditionalDisplay(nameof(ContentType), ElementType.Image, ElementType.Paragraph)]
        [DefaultValue(0)]
        public int SpacingAfterInPt { get; set; }

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
        [ConditionalDisplay(nameof(UseGivenCredentials), true)]
        [DefaultDisplayType(DisplayType.Text)]
        [DefaultValue(@"domain\username")]
        public string UserName { get; set; }

        [PasswordPropertyText]
        [ConditionalDisplay(nameof(UseGivenCredentials), true)]
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
}
